"""
Phoneline+ Bulk Customer Processor
Reads customer data from Excel input sheet and creates customers via API
Supports both Production and UAT environments
Supports creating multiple users per customer
"""

import pandas as pd
import json
import shutil
import time
import os
import hashlib
import random
from typing import Dict, List, Tuple, Optional, Set
from datetime import datetime
from pathlib import Path
from phoneline_plus_jwt_auth import PhonelinePlusAuth
from phoneline_plus_create_customer import PhonelinePlusCustomer
from phoneline_plus_create_user import PhonelinePlusUser
from graph_mailbox_check import send_order_report_email


def format_phone_number(value):
    """
    Format phone number, preserving leading zeros that Excel may have stripped
    
    Args:
        value: Phone number from Excel (may be numeric or string)
        
    Returns:
        Properly formatted phone number string with leading zeros
    """
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and remove any decimals Excel may have added
    phone_str = str(value).replace('.0', '').strip()
    
    # If it's a UK mobile number without leading 0, add it
    if phone_str.isdigit() and len(phone_str) == 10 and phone_str.startswith('7'):
        return '0' + phone_str
    
    # If it's a UK landline without leading 0, add it
    if phone_str.isdigit() and len(phone_str) == 10 and phone_str[0] in '1238':
        return '0' + phone_str
    
    return phone_str


class PhonelinePlusBulkProcessor:
    """Handles bulk Hardware Only customer creation from Excel input sheet"""
    
    # Delay between customer processing cycles (seconds)
    # Allows API time to prepare next available number
    CYCLE_DELAY = 3.5
    USER_NUMBER_TYPE = "standard_geographic"
    USER_AREA_CODE = "44161"
    USER_NUMBER_MAX_RETRIES = 8
    USER_NUMBER_RETRY_DELAY_SECONDS = 2.0
    USER_INFLIGHT_RETRY_ATTEMPTS = 10  # Try up to 10 numbers before giving up
    CUSTOMER_INFLIGHT_RETRY_ATTEMPTS = 3
    NUMBER_POOL_SIZE = 20  # Pre-fetch 20 numbers for user creation

    @staticmethod
    def _is_inflight_inventory_error(error_text: Optional[str]) -> bool:
        """
        Check whether API error indicates selected number is in-flight and unavailable.
        Detects both HTTP 400 and HTTP 409 in-flight errors.

        Args:
            error_text: Error message returned by API

        Returns:
            True when the error matches in-flight inventory conflict
        """
        if not error_text:
            return False

        normalized_error = str(error_text).lower()
        
        # Check for in-flight order pattern (HTTP 400 or 409)
        is_inflight = (
            "existing in-flight order" in normalized_error
            or "in-flight order" in normalized_error
        ) and (
            "cannot currently have a change or terminate order" in normalized_error
            or "cannot currently" in normalized_error
        )
        
        return is_inflight

    @staticmethod
    def _normalize_number_for_excel_text(value) -> str:
        """
        Normalize number-like values for Excel text cells.

        Args:
            value: Raw value from report row

        Returns:
            String value suitable for text-formatted Excel cells
        """
        if value is None or pd.isna(value):
            return ""

        if isinstance(value, int):
            return str(value)

        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return format(value, 'f').rstrip('0').rstrip('.')

        value_str = str(value).strip()
        if value_str.lower() == 'nan':
            return ""

        if value_str.endswith('.0') and value_str[:-2].isdigit():
            return value_str[:-2]

        return value_str

    def _request_distinct_available_number(
        self,
        user_manager: PhonelinePlusUser,
        excluded_numbers: Set[str]
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Request an available number that is not in excluded_numbers.

        Args:
            user_manager: User API manager with auth token
            excluded_numbers: Numbers already known to fail/used in this retry flow

        Returns:
            Tuple of (success, number, error_message)
        """
        last_error = None

        for fetch_attempt in range(1, self.USER_NUMBER_MAX_RETRIES + 1):
            while self._available_number_pool:
                candidate_number = self._available_number_pool.pop(0)
                if candidate_number not in excluded_numbers:
                    return True, candidate_number, None

            numbers_success, available_numbers, number_error = user_manager.get_available_numbers(
                number_type=self.USER_NUMBER_TYPE,
                area_code=self.USER_AREA_CODE,
                max_retries=self.USER_NUMBER_MAX_RETRIES,
                retry_delay_seconds=self.USER_NUMBER_RETRY_DELAY_SECONDS
            )

            if not numbers_success:
                last_error = number_error
                continue

            if not available_numbers:
                last_error = "No available numbers returned from API"
                continue

            self._available_number_pool.extend(available_numbers)

        return False, None, last_error or "Unable to fetch a distinct available number"
    
    def __init__(self, input_file: str, environment: str = "uat"):
        """
        Initialize the bulk processor
        
        Args:
            input_file: Path to Excel input file
            environment: Either 'production' or 'uat' (default: 'uat')
        """
        self.input_file = input_file
        self.environment = environment.lower()
        self.df = None
        self.results = []
        self.auth_token = None
        self._available_number_pool = []
    
    
    def load_input_file(self) -> Tuple[bool, Optional[str]]:
        """
        Load and validate the input Excel file
        
        Returns:
            Tuple of (success, error_message)
        """
        try:
            print(f"Loading input file: {self.input_file}")
            
            # Read Excel file
            self.df = pd.read_excel(self.input_file)
            
            # Required columns for Hardware Only accounts
            required_columns = [
                'keyID', 'secret', 'companyName', 'fullName', 'number',
                'contactNumber', 'plan', 'premises', 'street', 'town', 'county', 'postcode'
            ]
            
            # Check for missing columns
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                error_msg = f"Missing required columns: {', '.join(missing_columns)}"
                print(f"✗ {error_msg}")
                return False, error_msg
            
            print(f"✓ File loaded successfully: {len(self.df)} row(s) found")
            print(f"  Columns: {', '.join(self.df.columns.tolist())}")
            
            # DEBUG: Show first row of data (excluding credentials)
            if len(self.df) > 0:
                print(f"\n  DEBUG: Sample data from first row:")
                first_row = self.df.iloc[0]
                for col in ['companyName', 'fullName', 'email', 'contactNumber', 'number', 'plan']:
                    if col in first_row.index:
                        print(f"    {col}: {repr(first_row[col])}")
            
            return True, None
            
        except FileNotFoundError:
            error_msg = f"File not found: {self.input_file}"
            print(f"✗ {error_msg}")
            return False, error_msg
        except Exception as e:
            error_msg = f"Error loading file: {str(e)}"
            print(f"✗ {error_msg}")
            return False, error_msg
    
    def authenticate(self, key_id: str, secret: str) -> Tuple[bool, Optional[str]]:
        """
        Authenticate and get JWT token
        
        Args:
            key_id: Partner API Key ID
            secret: Partner API Secret
            
        Returns:
            Tuple of (success, error_message)
        """
        auth = PhonelinePlusAuth(environment=self.environment)
        success, token, error = auth.generate_token(key_id, secret)
        
        if success:
            self.auth_token = token
            return True, None
        else:
            return False, error
    
    def _parse_users_from_row(self, row: pd.Series) -> List[Dict]:
        """
        Parse user information from a row
        Dynamically detects all user columns (user1_, user2_, user3_, ..., userN_)
        
        Args:
            row: Pandas Series representing a row from the input sheet
            
        Returns:
            List of user dictionaries
        """
        users = []
        
        # Dynamically find all user indices by looking for user*_fullName columns
        user_indices = set()
        for col in row.index:
            if col.startswith('user') and '_fullName' in col:
                # Extract the number from "user123_fullName"
                try:
                    user_num = int(col.split('_')[0].replace('user', ''))
                    user_indices.add(user_num)
                except (ValueError, IndexError):
                    continue
        
        print(f"  DEBUG: Found user columns for user numbers: {sorted(user_indices)}")
        
        # Process each user found
        for i in sorted(user_indices):
            prefix = f"user{i}_"
            
            # Check if this user has required fields
            if pd.notna(row.get(f"{prefix}fullName")) and pd.notna(row.get(f"{prefix}email")):
                print(f"  DEBUG: Processing user{i}: {row.get(f'{prefix}fullName')}")
                user = {
                    "fullName": str(row[f"{prefix}fullName"]).strip(),
                    "email": str(row[f"{prefix}email"]).strip(),
                    "type": str(row.get(f"{prefix}type", "standard")).strip().lower(),
                    "address": {
                        "premises": str(row.get(f"{prefix}premises", "")).strip(),
                        "street": str(row.get(f"{prefix}street", "")).strip(),
                        "town": str(row.get(f"{prefix}town", "")).strip(),
                        "county": str(row.get(f"{prefix}county", "")).strip(),
                        "postcode": str(row.get(f"{prefix}postcode", "")).strip().replace(" ", "")
                    }
                }
                
                # Add phone number if provided
                if pd.notna(row.get(f"{prefix}phoneNumber")) and str(row[f"{prefix}phoneNumber"]).strip() != '':
                    user["phoneNumber"] = format_phone_number(row[f"{prefix}phoneNumber"])
                
                users.append(user)
        
        return users

    def _create_additional_users(self, customer_id: str, users_to_create: List[Dict], result: Dict) -> None:
        """
        Create additional users for a customer.
        If a user has no phone number in input, fetch an available number and include it in payload.

        Args:
            customer_id: Customer ID to attach users to
            users_to_create: Parsed user payloads from input row
            result: Result dictionary to update with user creation outcomes
        """
        result['users_created'] = 0
        result['users_failed'] = 0
        result['user_results'] = []

        print(f"  Creating {len(users_to_create)} user(s) for this customer...")
        user_manager = PhonelinePlusUser(environment=self.environment, auth_token=self.auth_token)

        for user_idx, user in enumerate(users_to_create, 1):
            assigned_number = user.get('phoneNumber')

            if not assigned_number:
                user_result = {
                    'fullName': user['fullName'],
                    'email': user['email'],
                    'success': False,
                    'user_id': None,
                    'assigned_number': '',
                    'error': "No phone number available in enriched payload"
                }
                result['user_results'].append(user_result)
                result['users_failed'] += 1
                print(f"    User {user_idx}: {user['fullName']} - Missing phone number in enriched payload")
                continue

            user_success = False
            user_data = None
            user_error = None
            attempted_numbers = {assigned_number}

            for attempt in range(1, self.USER_INFLIGHT_RETRY_ATTEMPTS + 1):
                print(f"    User {user_idx}: Attempt {attempt}/{self.USER_INFLIGHT_RETRY_ATTEMPTS} - Creating user with number {assigned_number}")
                
                user_success, user_data, user_error = user_manager.create_user(
                    customer_id=customer_id,
                    full_name=user['fullName'],
                    email=user['email'],
                    address=user['address'],
                    user_type=user.get('type', 'standard'),
                    phone_number=assigned_number
                )

                if user_success:
                    print(f"    User {user_idx}: ✓ Successfully created on attempt {attempt}")
                    break

                # Check if this is an in-flight inventory error
                is_inflight_error = self._is_inflight_inventory_error(user_error)
                print(f"    User {user_idx}: ✗ Failed on attempt {attempt}. In-flight error: {is_inflight_error}")
                
                if not is_inflight_error:
                    print(f"    User {user_idx}: Not an in-flight error, stopping retries")
                    break

                if attempt >= self.USER_INFLIGHT_RETRY_ATTEMPTS:
                    print(f"    User {user_idx}: Max retry attempts reached")
                    break

                # Try to get a different number from the pool
                print(
                    f"    User {user_idx}: ⚠ Number {assigned_number} blocked by in-flight order. "
                    f"Fetching replacement number (attempt {attempt}/{self.USER_INFLIGHT_RETRY_ATTEMPTS})..."
                )
                
                number_success, replacement_number, number_error = self._request_distinct_available_number(
                    user_manager=user_manager,
                    excluded_numbers=attempted_numbers
                )

                if number_success and replacement_number:
                    print(f"    User {user_idx}: Got replacement number {replacement_number} (Pool: {len(self._available_number_pool)} remaining)")
                    assigned_number = replacement_number
                    user['phoneNumber'] = replacement_number
                    attempted_numbers.add(replacement_number)
                    continue
                else:
                    print(f"    User {user_idx}: ✗ Failed to get replacement number: {number_error}")
                    user_error = (
                        f"{user_error} | Failed to get replacement number after in-flight conflict: {number_error}"
                    )
                    break

            # Debug: Print user_data to see what's returned
            if user_data:
                print(f"    DEBUG: User data received: {user_data}")

            # Extract user ID - try multiple possible field names
            extracted_user_id = None
            if user_data:
                extracted_user_id = user_data.get('ID') or user_data.get('id') or user_data.get('userId') or user_data.get('userID') or user_data.get('user_id') or user_data.get('customerId')

            # CRITICAL VALIDATION: API said success but no user ID returned
            # This indicates the user may not have been actually created despite success response
            if user_success and not extracted_user_id and not (user_data and user_data.get('alreadyExists')):
                print(f"    ⚠ WARNING: API returned success but no user ID found in response!")
                print(f"    ⚠ This suggests the user may not have been created. Please verify in portal.")
                user_success = False
                if not user_error:
                    user_error = "API returned success (200/201) but no user ID in response - user creation may have failed silently"
                else:
                    user_error = f"{user_error} | No user ID returned despite success"

            user_result = {
                'fullName': user['fullName'],
                'email': user['email'],
                'success': user_success,
                'user_id': extracted_user_id if extracted_user_id else '',
                'assigned_number': assigned_number,
                'error': user_error,
                'needs_verification': user_success and not extracted_user_id  # Flag for manual verification
            }

            print(f"    User {user_idx}: {user['fullName']} - Success: {user_success}, ID: {extracted_user_id or 'MISSING'}")
            if user_result.get('needs_verification'):
                print(f"    ⚠ NEEDS VERIFICATION - check portal manually")
            
            result['user_results'].append(user_result)

            if user_success:
                result['users_created'] += 1
            else:
                result['users_failed'] += 1
            
            # Small delay between user creations to avoid API rate limiting
            if user_idx < len(users_to_create):
                time.sleep(0.5)

        print(f"  Users: {result['users_created']} created, {result['users_failed']} failed")

    def _enrich_users_with_available_numbers(self, users_to_create: List[Dict], result: Dict) -> bool:
        """
        Enrich users with phone numbers before customer/user creation.
        Pre-fetches a batch of available numbers to handle in-flight retries.

        Args:
            users_to_create: Parsed users from input sheet
            result: Result dictionary to capture enrichment failures

        Returns:
            True when all users have phone numbers, otherwise False
        """
        if not users_to_create:
            return True

        user_manager = PhonelinePlusUser(environment=self.environment, auth_token=self.auth_token)
        enrichment_failed = False

        # Pre-fetch a batch of available numbers for user assignment and retries
        print(f"  Pre-fetching {self.NUMBER_POOL_SIZE} available numbers for user creation...")
        numbers_success, available_numbers, number_error = user_manager.get_available_numbers(
            number_type=self.USER_NUMBER_TYPE,
            area_code=self.USER_AREA_CODE,
            max_retries=self.USER_NUMBER_MAX_RETRIES,
            retry_delay_seconds=self.USER_NUMBER_RETRY_DELAY_SECONDS
        )

        if not numbers_success or not available_numbers:
            print(f"  ✗ Failed to fetch available numbers: {number_error}")
            for user_idx, user in enumerate(users_to_create, 1):
                if user.get('phoneNumber'):
                    continue
                user_result = {
                    'fullName': user['fullName'],
                    'email': user['email'],
                    'success': False,
                    'user_id': None,
                    'assigned_number': '',
                    'error': f"Unable to fetch available numbers: {number_error}"
                }
                if 'user_results' not in result:
                    result['user_results'] = []
                result['user_results'].append(user_result)
            return False

        # Add fetched numbers to pool
        self._available_number_pool.extend(available_numbers)
        print(f"  ✓ Fetched {len(available_numbers)} available numbers. Pool size: {len(self._available_number_pool)}")

        # Enrich users that don't have phone numbers
        for user_idx, user in enumerate(users_to_create, 1):
            if user.get('phoneNumber'):
                print(f"  User {user_idx}: Using provided number {user['phoneNumber']}")
                continue

            if not self._available_number_pool:
                enrichment_failed = True
                user_result = {
                    'fullName': user['fullName'],
                    'email': user['email'],
                    'success': False,
                    'user_id': None,
                    'assigned_number': '',
                    'error': "Number pool exhausted during enrichment"
                }
                if 'user_results' not in result:
                    result['user_results'] = []
                result['user_results'].append(user_result)
                print(f"  User {user_idx}: {user['fullName']} - Failed enrichment (pool exhausted)")
                continue

            available_number = self._available_number_pool.pop(0)
            user['phoneNumber'] = available_number
            print(f"  User {user_idx}: Enriched with available number {available_number} (Pool: {len(self._available_number_pool)} remaining)")

        return not enrichment_failed
    
    def process_customers(self) -> List[Dict]:
        """
        Process all customers from the input sheet
        
        Returns:
            List of result dictionaries
        """
        if self.df is None:
            print("✗ No data loaded. Call load_input_file() first.")
            return []
        
        print("\n" + "=" * 60)
        print(f"Processing {len(self.df)} customer(s)")
        print("=" * 60)
        
        for index, row in self.df.iterrows():
            row_num = index + 2  # +2 because Excel is 1-indexed and has header row
            
            print(f"\n{'='*60}")
            print(f"[Row {row_num}] Processing: {row['companyName']}")
            
            # Reset number pool for each customer to ensure fresh numbers
            self._available_number_pool = []
            
            users_to_create = []
            
            result = {
                'row': row_num,
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                # Input data
                'companyName': row['companyName'],
                'fullName': row.get('fullName', ''),
                'email': row.get('email', ''),
                'contactNumber': row.get('contactNumber', ''),
                'number': row.get('number', ''),
                'plan': row.get('plan', ''),
                'premises': row.get('premises', ''),
                'street': row.get('street', ''),
                'town': row.get('town', ''),
                'county': row.get('county', ''),
                'postcode': row.get('postcode', ''),
                # Results
                'success': False,
                'customer_id': None,
                'error': None,
                'needs_verification': False,
                'response_data': {}
            }
            
            try:
                # Get credentials from this row (allows different credentials per row if needed)
                key_id = str(row['keyID']).strip()
                secret = str(row['secret']).strip()
                
                # Authenticate if token not already available
                if not self.auth_token:
                    print(f"  Authenticating with provided credentials...")
                    success, error = self.authenticate(key_id, secret)
                    if not success:
                        result['error'] = f"Authentication failed: {error}"
                        print(f"  ✗ {result['error']}")
                        self.results.append(result)
                        continue
                
                # Create customer manager
                customer_manager = PhonelinePlusCustomer(
                    environment=self.environment,
                    auth_token=self.auth_token
                )
                
                # Build address dictionary
                address = {
                    "premises": str(row['premises']).strip(),
                    "street": str(row['street']).strip(),
                    "town": str(row['town']).strip(),
                    "county": str(row['county']).strip(),
                    "postcode": str(row['postcode']).strip().replace(" ", "")
                }
                
                # Get plan type (hardware_only or hardware_only_cp)
                plan = str(row['plan']).strip().lower()
                if plan not in ['hardware_only', 'hardware_only_cp']:
                    result['error'] = f"Invalid plan type '{plan}'. Must be 'hardware_only' or 'hardware_only_cp'"
                    print(f"  ✗ {result['error']}")
                    self.results.append(result)
                    continue
                
                # Get number field (required for hardware_only, optional for hardware_only_cp)
                number = None
                if pd.notna(row.get('number')) and str(row['number']).strip() != '':
                    number = str(row['number']).strip()
                    # Remove .0 if Excel added it
                    if number.endswith('.0'):
                        number = number[:-2]
                elif plan == 'hardware_only':
                    # Number is required for hardware_only plan
                    result['error'] = f"Number is required for 'hardware_only' plan (found: '{row.get('number')}')"
                    print(f"  ✗ {result['error']}")
                    self.results.append(result)
                    continue
                # For hardware_only_cp, number is optional (customer provides later)
                
                # Get optional email field
                email = str(row['email']).strip() if pd.notna(row.get('email')) and str(row['email']).strip() != '' else None
                
                # DEBUG: Log extracted email value
                print(f"  DEBUG: Email field from Excel:")
                print(f"    Raw value: {repr(row.get('email'))}")
                print(f"    Processed value: {repr(email)}")

                # Parse and enrich additional users before creating customer
                users_to_create = self._parse_users_from_row(row)
                if users_to_create:
                    print(f"  Enriching {len(users_to_create)} user(s) with required phone numbers...")
                    enrichment_success = self._enrich_users_with_available_numbers(users_to_create, result)
                    if not enrichment_success:
                        result['error'] = "Unable to allocate available numbers for one or more users. Customer creation skipped."
                        print(f"  ✗ {result['error']}")
                        self.results.append(result)
                        continue
                
                # Create Hardware Only customer
                number_for_customer = number
                attempted_customer_numbers = {number_for_customer} if number_for_customer else set()
                user_number_manager = PhonelinePlusUser(environment=self.environment, auth_token=self.auth_token)

                print(f"\n  === CUSTOMER CREATION RETRY LOOP START ===")
                print(f"  Max attempts: {self.CUSTOMER_INFLIGHT_RETRY_ATTEMPTS}")
                print(f"  Row index: {index}, Row number: {row_num}")
                
                for customer_attempt in range(1, self.CUSTOMER_INFLIGHT_RETRY_ATTEMPTS + 1):
                    print(f"\n  [ATTEMPT {customer_attempt}/{self.CUSTOMER_INFLIGHT_RETRY_ATTEMPTS}] Creating customer...")
                    # DEBUG: Log parameters being sent to create_customer
                    print(f"  DEBUG: Creating customer with parameters:")
                    print(f"    company_name: {repr(str(row['companyName']).strip())}")
                    print(f"    full_name: {repr(str(row['fullName']).strip())}")
                    print(f"    contact_number: {repr(str(row['contactNumber']).strip())}")
                    print(f"    number: {repr(number_for_customer)}")
                    print(f"    plan: {repr(plan)}")
                    print(f"    email: {repr(email)}")
                    
                    print(f"  [ATTEMPT {customer_attempt}] Calling customer_manager.create_customer()...")
                    success, customer_data, error = customer_manager.create_customer(
                        company_name=str(row['companyName']).strip(),
                        full_name=str(row['fullName']).strip(),
                        contact_number=str(row['contactNumber']).strip(),
                        number=number_for_customer,
                        address=address,
                        plan=plan,
                        email=email
                    )
                    
                    print(f"  [ATTEMPT {customer_attempt}] Result: success={success}, error={error}")

                    if success:
                        print(f"  [ATTEMPT {customer_attempt}] ✓ SUCCESS - Breaking retry loop")
                        break

                    # Check if we should retry
                    should_retry = (
                        plan == 'hardware_only'
                        and number_for_customer
                        and self._is_inflight_inventory_error(error)
                        and customer_attempt < self.CUSTOMER_INFLIGHT_RETRY_ATTEMPTS
                    )
                    print(f"  [ATTEMPT {customer_attempt}] Should retry? {should_retry}")
                    print(f"    - plan=='hardware_only': {plan == 'hardware_only'}")
                    print(f"    - has number: {bool(number_for_customer)}")
                    print(f"    - is in-flight error: {self._is_inflight_inventory_error(error)}")
                    print(f"    - not at max attempts: {customer_attempt < self.CUSTOMER_INFLIGHT_RETRY_ATTEMPTS}")
                    
                    if not should_retry:
                        print(f"  [ATTEMPT {customer_attempt}] Not retrying - Breaking retry loop")
                        break

                    print(
                        f"  [ATTEMPT {customer_attempt}] Customer number {number_for_customer} blocked by in-flight order. "
                        f"Requesting replacement ({customer_attempt}/{self.CUSTOMER_INFLIGHT_RETRY_ATTEMPTS - 1})..."
                    )
                    repl_success, replacement_number, repl_error = self._request_distinct_available_number(
                        user_manager=user_number_manager,
                        excluded_numbers=attempted_customer_numbers
                    )

                    if repl_success and replacement_number:
                        number_for_customer = replacement_number
                        attempted_customer_numbers.add(replacement_number)
                        print(f"  Retrying customer creation with replacement number {number_for_customer}")
                        continue

                    error = f"{error} | Failed to get replacement number after in-flight conflict: {repl_error}"
                    print(f"  [ATTEMPT {customer_attempt}] Failed to get replacement - Breaking retry loop")
                    break
                
                print(f"  === CUSTOMER CREATION RETRY LOOP END ===")
                print(f"  Final result: success={success}\n")
                
                if success:
                    result['success'] = True
                    result['customer_id'] = customer_data.get('customerID', customer_data.get('id', 'N/A'))
                    result['user_id'] = customer_data.get('userID', '')  # Initial portal user ID
                    result['response_data'] = customer_data
                    
                    # DEBUG: Check what email came back in the response
                    response_email = customer_data.get('email', customer_data.get('emailAddress', 'NOT_IN_RESPONSE'))
                    print(f"  DEBUG: Email in API response: {repr(response_email)}")
                    
                    # Extract key details from response
                    result['portal_id'] = customer_data.get('portalId', 'N/A')
                    result['status'] = customer_data.get('status', 'N/A')
                    result['plan'] = customer_data.get('plan', 'N/A')
                    
                    # Extract assigned number
                    if 'phoneNumberE164' in customer_data:
                        result['assigned_number'] = customer_data.get('phoneNumberE164', '')
                    elif 'number' in customer_data:
                        result['assigned_number'] = customer_data['number'].get('numberE164', 'N/A')
                    
                    # Extract SIP credentials from destination object if present
                    destination = customer_data.get('destination', {})
                    result['registration_server'] = destination.get('registrationServer', '')
                    result['sip_id'] = destination.get('sipID', '')
                    result['sip_password'] = destination.get('sipPassword', '')
                    
                    print(f"  ✓ Customer created successfully! ID: {result['customer_id']}")
                    if result['sip_id']:
                        print(f"  ✓ SIP credentials generated")
                        print(f"    Registration Server: {result['registration_server']}")
                        print(f"    SIP ID: {result['sip_id']}")
                    
                    # Create additional users using pre-enriched payload
                    if users_to_create:
                        self._create_additional_users(
                            customer_id=result['customer_id'],
                            users_to_create=users_to_create,
                            result=result
                        )
                else:
                    result['error'] = error
                    # Check if this is a timeout error that may have succeeded
                    if error and "TIMEOUT" in error and "may have been created" in error:
                        result['needs_verification'] = True
                        print(f"  ⚠ {error}")
                        
                        # Check if there are users to create despite timeout
                        if users_to_create:
                            print(f"\n  ℹ Customer may have been created successfully")
                            print(f"  DEBUG: Found {len(users_to_create)} user(s) to create")
                            retry_users = input(f"  Do you want to provide customer ID and create {len(users_to_create)} user(s)? (y/n): ").strip().lower()
                            
                            if retry_users == 'y':
                                manual_customer_id = input("  Enter customer ID from portal: ").strip()
                                if manual_customer_id:
                                    # Re-authenticate to get fresh token (previous one may have expired)
                                    print("  Re-authenticating to get fresh token...")
                                    key_id = str(row['keyID']).strip()
                                    secret = str(row['secret']).strip()
                                    auth_success, auth_error = self.authenticate(key_id, secret)
                                    if not auth_success:
                                        print(f"  ✗ Re-authentication failed: {auth_error}")
                                        result['error'] += f" | User creation skipped - re-auth failed"
                                    else:
                                        result['customer_id'] = manual_customer_id
                                        self._create_additional_users(
                                            customer_id=manual_customer_id,
                                            users_to_create=users_to_create,
                                            result=result
                                        )
                    else:
                        print(f"  ✗ Failed: {error}")
            
            except KeyError as e:
                result['error'] = f"Missing field: {str(e)}"
                print(f"  ✗ {result['error']}")
            except Exception as e:
                result['error'] = f"Unexpected error: {str(e)}"
                print(f"  ✗ {result['error']}")
            
            self.results.append(result)
            
            # Delay before next customer (allows API to prepare next available number)
            if index < len(self.df) - 1:  # Don't delay after last customer
                print(f"  Waiting {self.CYCLE_DELAY}s before next customer...")
                time.sleep(self.CYCLE_DELAY)
        
        return self.results
    
    def generate_report(self, output_file: Optional[str] = None) -> str:
        """
        Generate a report of processing results
        
        Args:
            output_file: Optional path to save report Excel file
            
        Returns:
            Path to generated report file
        """
        if not self.results:
            print("No results to report")
            return None
        
        # Create DataFrame from results, excluding the complex response_data field
        report_data = []
        for result in self.results:
            # Create a flat dictionary for Excel export
            flat_result = {
                'row': result['row'],
                'timestamp': result['timestamp'],
                # Input fields (for troubleshooting)
                'input_companyName': result.get('companyName', ''),
                'input_fullName': result.get('fullName', ''),
                'input_email': result.get('email', ''),
                'input_contactNumber': result.get('contactNumber', ''),
                'input_number': result.get('number', ''),
                'input_plan': result.get('plan', ''),
                'input_premises': result.get('premises', ''),
                'input_street': result.get('street', ''),
                'input_town': result.get('town', ''),
'input_county': result.get('county', ''),
'input_postcode': result.get('postcode', ''),
# Result fields
'success': result['success'],
'needs_verification': result.get('needs_verification', False),
'customer_id': result.get('customer_id', ''),
'user_id': result.get('user_id', ''),  # Portal user ID from customer creation
'portal_id': result.get('portal_id', ''),
'status': result.get('status', ''),
                'input_town': result.get('town', ''),
                'input_county': result.get('county', ''),
                'input_postcode': result.get('postcode', ''),
                # Result fields
                'success': result['success'],
                'needs_verification': result.get('needs_verification', False),
                'customer_id': result.get('customer_id', ''),
                'user_id': result.get('user_id', ''),
                'portal_id': result.get('portal_id', ''),
                'status': result.get('status', ''),
                'plan': result.get('plan', ''),
                'assigned_number': result.get('assigned_number', ''),
                'registration_server': result.get('registration_server', ''),
                'sip_id': result.get('sip_id', ''),
                'sip_password': result.get('sip_password', ''),
                'users_created': result.get('users_created', 0),
                'users_failed': result.get('users_failed', 0),
                'error': result.get('error', '')
            }
            report_data.append(flat_result)
            
            # Add user details as separate rows if users were created
            if result.get('user_results'):
                for user_result in result['user_results']:
                    user_row = {
                        'row': f"{result['row']}-user",
                        'timestamp': result['timestamp'],
                        # Parent customer input fields
                        'input_companyName': result.get('companyName', ''),
                        'input_fullName': user_result.get('fullName', ''),
                        'input_email': user_result.get('email', ''),
                        'input_contactNumber': '',
                        'input_number': '',
                        'input_plan': 'USER',
                        'input_premises': '',
                        'input_street': '',
                        'input_town': '',
                        'input_county': '',
                        'input_postcode': '',
                        # User result fields
                        'success': user_result.get('success', False),
                        'needs_verification': user_result.get('needs_verification', False),
                        'customer_id': result.get('customer_id', ''),
                        'user_id': user_result.get('user_id', ''),
                        'portal_id': '',
                        'status': '',
                        'plan': 'USER',
                        'assigned_number': user_result.get('assigned_number', ''),
                        'registration_server': '',
                        'sip_id': '',
                        'sip_password': '',
                        'users_created': 0,
                        'users_failed': 0,
                        'error': user_result.get('error', '')
                    }
                    report_data.append(user_row)
        
        # Create DataFrame from flattened results
        report_df = pd.DataFrame(report_data)
        
        # Generate default output filename if not provided
        if not output_file:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f"phoneline_plus_results_{timestamp}.xlsx"
        
        # Calculate summary statistics
        total = len(self.results)
        successful = sum(1 for r in self.results if r['success'])
        needs_verification = sum(1 for r in self.results if r.get('needs_verification', False))
        failed = total - successful - needs_verification
        total_users_created = sum(r.get('users_created', 0) for r in self.results)
        total_users_failed = sum(r.get('users_failed', 0) for r in self.results)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            report_df.to_excel(writer, sheet_name='Customer Results', index=False)
            
            # Add summary sheet
            summary_data = {
                'Metric': [
                    'Total Processed',
                    'Successful',
                    'Needs Verification (Timeout)',
                    'Failed',
                    'Users Created',
                    'Users Failed',
                    'Environment',
                    'Processing Date/Time'
                ],
                'Value': [
                    total,
                    successful,
                    needs_verification,
                    failed,
                    total_users_created,
                    total_users_failed,
                    self.environment.upper(),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Force phone/number columns to text format to prevent scientific notation
            customer_sheet = writer.sheets['Customer Results']
            text_columns = ['input_number', 'assigned_number']
            for column_name in text_columns:
                if column_name in report_df.columns:
                    col_idx = report_df.columns.get_loc(column_name) + 1
                    for row_idx in range(2, len(report_df) + 2):
                        cell = customer_sheet.cell(row=row_idx, column=col_idx)
                        normalized_value = self._normalize_number_for_excel_text(cell.value)
                        cell.value = normalized_value
                        cell.number_format = '@'
            
            # Auto-size columns
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Print summary
        print("\n" + "=" * 60)
        print("PROCESSING SUMMARY")
        print("=" * 60)
        total = len(self.results)
        successful = sum(1 for r in self.results if r['success'])
        needs_verification = sum(1 for r in self.results if r.get('needs_verification', False))
        failed = total - successful - needs_verification
        total_users_created = sum(r.get('users_created', 0) for r in self.results)
        total_users_failed = sum(r.get('users_failed', 0) for r in self.results)
        
        print(f"Total processed: {total}")
        print(f"Successful: {successful}")
        print(f"Needs Verification (Timeout): {needs_verification}")
        print(f"Failed: {failed}")
        if total_users_created > 0 or total_users_failed > 0:
            print(f"\nUsers:")
            print(f"  Created: {total_users_created}")
            print(f"  Failed: {total_users_failed}")
        print(f"\nReport saved to: {output_file}")
        
        # Check for users that need verification
        users_needing_verification = []
        for result in self.results:
            if result.get('user_results'):
                for user_result in result['user_results']:
                    if user_result.get('needs_verification', False):
                        users_needing_verification.append({
                            'row': result['row'],
                            'company': result.get('companyName', ''),
                            'user_name': user_result.get('fullName', ''),
                            'user_email': user_result.get('email', '')
                        })
        
        if needs_verification > 0:
            print("\n⚠ CUSTOMERS NEEDING VERIFICATION (Timeout - may have been created):")
            for result in self.results:
                if result.get('needs_verification', False):
                    print(f"  Row {result['row']}: {result['companyName']} ({result['email']})")
                    print(f"    → Please check portal to verify customer was created")
        
        if users_needing_verification:
            print("\n⚠ USERS NEEDING VERIFICATION (Success reported but no user ID returned):")
            for user_info in users_needing_verification:
                print(f"  Row {user_info['row']}: {user_info['company']} - User: {user_info['user_name']} ({user_info['user_email']})")
                print(f"    → API returned success but no user ID. Please verify in portal if user was actually created.")
        
        if failed > 0:
            print("\nFailed rows:")
            for result in self.results:
                if not result['success'] and not result.get('needs_verification', False):
                    print(f"  Row {result['row']}: {result['companyName']} - {result['error']}")
        
        return output_file


def main():
    """
    Scan input folder and process all Excel files
    """
    # Log script start time with unique ID
    import random
    run_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}-{random.randint(10000, 99999)}"
    print("\n" + "#" * 70)
    print(f"### SCRIPT RUN ID: {run_id} ###")
    print(f"### START TIME: {datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')} ###")
    print("#" * 70 + "\n")
    
    # Configuration
    SCRIPT_DIR = Path(__file__).parent
    INPUT_FOLDER = r"\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\Phoneline+ New Orders"
    OUTPUT_FOLDER = SCRIPT_DIR / "output"
    PROCESSED_FOLDER = Path(INPUT_FOLDER) / "processed"
    ENVIRONMENT = "production"  # or "production"
    
    # Ensure folders exist
    OUTPUT_FOLDER.mkdir(exist_ok=True)
    PROCESSED_FOLDER.mkdir(exist_ok=True)
    
    print("=" * 60)
    print("Phoneline+ Bulk Customer Processor")
    print("=" * 60)
    print(f"Environment: {ENVIRONMENT.upper()}")
    print(f"Input Folder: {INPUT_FOLDER}")
    print()
    
    # Process files one at a time (find, move, process) to prevent duplicate processing
    # If automation triggers script multiple times, only the first run will find files
    import glob
    files_processed = 0
    
    while True:
        # Find ONE Excel file
        excel_files = glob.glob(str(Path(INPUT_FOLDER) / "*.xlsx"))
        excel_files = [f for f in excel_files if not Path(f).name.startswith('~$')]
        
        if not excel_files:
            # No more files to process
            if files_processed == 0:
                print("✗ No Excel files found in input folder")
            else:
                print(f"\n✓ All files processed ({files_processed} total)")
            break
        
        # Take the first file from the list
        input_file = excel_files[0]
        filename = Path(input_file).name
        print(f"{'='*60}")
        print(f"Processing: {filename}")
        print(f"{'='*60}")
        
        # IMMEDIATELY move file to processing folder to prevent duplicate processing
        # if automation triggers script multiple times
        try:
            process_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            processing_folder = PROCESSED_FOLDER / "processing" / process_timestamp
            processing_folder.mkdir(parents=True, exist_ok=True)
            
            processing_path = processing_folder / filename
            shutil.move(input_file, str(processing_path))
            print(f"✓ File moved to processing folder")
            print(f"  This prevents duplicate processing if automation triggers multiple times")
            
            # Update input_file to point to new location
            input_file = str(processing_path)
        except Exception as e:
            print(f"✗ ERROR: Could not move file to processing folder: {e}")
            print(f"  Skipping this file to prevent potential duplicates")
            continue
        
        # Create processor
        processor = PhonelinePlusBulkProcessor(str(input_file), environment=ENVIRONMENT)
        
        try:
            # Load input file
            success, error = processor.load_input_file()
            if not success:
                print(f"✗ Failed to load file: {error}")
                print()
                continue
            
            # Process all customers
            results = processor.process_customers()
            
            # Generate report
            output_file = processor.generate_report()
            print(f"\n✓ Results saved to: {output_file}")
            
            # Send email with report
            print("\n" + "=" * 60)
            print("SENDING EMAIL REPORT")
            print("=" * 60)
            
            # Calculate summary statistics for email
            total = len(processor.results)
            successful = sum(1 for r in processor.results if r['success'])
            needs_verification = sum(1 for r in processor.results if r.get('needs_verification', False))
            failed = total - successful - needs_verification
            
            email_sent = send_order_report_email(
                report_file_path=output_file,
                recipient_email='psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk',
                #recipient_email='david.murphy+psmsoutput@gmail.com',
                input_filename=filename,
                total_orders=total,
                successful=successful,
                failed=failed,
                errors=needs_verification,
                report_type='Customer Creation',
                subject_prefix='Phoneline+ Customer Creation Results'
            )
            
            if email_sent:
                print("✓ Report successfully emailed to psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk")
            else:
                print("⚠ Warning: Failed to send email report")
            
            # Move from processing to completed folder
            try:
                completed_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                completed_subfolder = PROCESSED_FOLDER / "completed" / completed_timestamp
                completed_subfolder.mkdir(parents=True, exist_ok=True)

                destination_path = completed_subfolder / filename
                shutil.move(input_file, str(destination_path))

                print(f"\n✓ Input file moved to completed folder: {destination_path}")
                print("  Processing successful!")
            except Exception as e:
                print(f"\n⚠ Warning: Could not move file to completed folder: {e}")
                print(f"  File remains in processing folder: {input_file}")
                print(f"\n⚠ Warning: Could not move input file: {e}")
                print("  File may be reprocessed on next run")
        
        except Exception as e:
            print(f"\n✗ ERROR processing file: {e}")
            import traceback
            traceback.print_exc()
        
        files_processed += 1
        print()
    
    # Log script completion
    print("\n" + "#" * 70)
    print(f"### SCRIPT RUN ID: {run_id} ###")
    print(f"### END TIME: {datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')} ###")
    print(f"### SCRIPT COMPLETED NORMALLY ###")
    print("#" * 70 + "\n")


if __name__ == "__main__":
    main()


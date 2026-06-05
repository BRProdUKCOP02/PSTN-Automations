"""
Phoneline+ User Creation
Creates users for existing Phoneline+ customers
Supports both Production and UAT environments
"""

import requests
import json
import time
from typing import Dict, Optional, Tuple, List
from datetime import datetime


class PhonelinePlusUser:
    """Handles user creation for Phoneline+ customers"""
    
    # API Endpoints
    PROD_BASE_URL = "https://api-ss-gb-aws.gammaapi.net/partner/v1/customers"
    UAT_BASE_URL = "https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/customers"
    PROD_NUMBERS_BASE_URL = "https://api-ss-gb-aws.gammaapi.net/partner/v1/numbers"
    UAT_NUMBERS_BASE_URL = "https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/numbers"
    
    def __init__(self, environment: str = "uat", auth_token: Optional[str] = None):
        """
        Initialize the user creation handler
        
        Args:
            environment: Either 'production' or 'uat' (default: 'uat')
            auth_token: Optional JWT token. If not provided, must call set_token() later
        """
        self.environment = environment.lower()
        self.base_url = self.PROD_BASE_URL if self.environment == "production" else self.UAT_BASE_URL
        self.numbers_base_url = self.PROD_NUMBERS_BASE_URL if self.environment == "production" else self.UAT_NUMBERS_BASE_URL
        self.auth_token = auth_token

    @staticmethod
    def _extract_first_available_number(response_data) -> Optional[str]:
        """
        Extract first available number from varying API response formats.

        Args:
            response_data: Parsed JSON response

        Returns:
            First discovered phone number string, or None if not found
        """
        keys_to_check = (
            "phoneNumberE164",
            "numberE164",
            "phoneNumber",
            "number"
        )

        queue = [response_data]
        while queue:
            current = queue.pop(0)

            if isinstance(current, dict):
                for key in keys_to_check:
                    value = current.get(key)
                    if isinstance(value, str) and value.strip():
                        return value.strip()

                for value in current.values():
                    if isinstance(value, (dict, list)):
                        queue.append(value)

            elif isinstance(current, list):
                queue.extend(current)

            elif isinstance(current, str) and current.strip().startswith("+"):
                return current.strip()

        return None

    @staticmethod
    def _extract_available_numbers(response_data) -> List[str]:
        """
        Extract all available numbers from varying API response formats.

        Args:
            response_data: Parsed JSON response

        Returns:
            Ordered de-duplicated list of discovered phone numbers
        """
        keys_to_check = (
            "phoneNumberE164",
            "numberE164",
            "phoneNumber",
            "number"
        )

        discovered_numbers: List[str] = []
        seen_numbers = set()
        queue = [response_data]

        while queue:
            current = queue.pop(0)

            if isinstance(current, dict):
                for key in keys_to_check:
                    value = current.get(key)
                    if isinstance(value, str):
                        candidate = value.strip()
                        if candidate and candidate not in seen_numbers:
                            seen_numbers.add(candidate)
                            discovered_numbers.append(candidate)

                for value in current.values():
                    if isinstance(value, (dict, list)):
                        queue.append(value)

            elif isinstance(current, list):
                queue.extend(current)

            elif isinstance(current, str):
                candidate = current.strip()
                if candidate.startswith("+") and candidate not in seen_numbers:
                    seen_numbers.add(candidate)
                    discovered_numbers.append(candidate)

        return discovered_numbers

    def get_available_numbers(
        self,
        number_type: str = "standard_geographic",
        area_code: str = "44161",
        max_retries: int = 4,
        retry_delay_seconds: float = 1.5
    ) -> Tuple[bool, Optional[List[str]], Optional[str]]:
        """
        Request available numbers from partner API.

        Args:
            number_type: Number type query parameter (default: standard_geographic)
            area_code: Area code query parameter (default: 44161)
            max_retries: Number of attempts for transient failures
            retry_delay_seconds: Base delay between retry attempts

        Returns:
            Tuple of (success, numbers, error_message)
        """
        if not self.auth_token:
            error_msg = "No authentication token available. Set token first using set_token()"
            print(f"✗ {error_msg}")
            return False, None, error_msg

        url = f"{self.numbers_base_url}/available"
        params = {
            "numberType": number_type,
            "areaCode": area_code
        }

        headers = {
            "accept": "*/*",
            "authorization": f"Bearer {self.auth_token}",
            "cache-control": "no-cache",
            "connection": "keep-alive"
        }

        try:
            transient_status_codes = {429, 500, 502, 503, 504}
            last_error = None

            for attempt in range(1, max_retries + 1):
                response = requests.get(url, params=params, headers=headers, timeout=120)

                if response.status_code in [200, 201]:
                    try:
                        response_data = response.json()
                    except json.JSONDecodeError:
                        error_msg = f"Invalid JSON from available numbers endpoint: {response.text}"
                        print(f"✗ {error_msg}")
                        return False, None, error_msg

                    available_numbers = self._extract_available_numbers(response_data)
                    if available_numbers:
                        return True, available_numbers, None

                    error_msg = "No available numbers returned in response payload"
                    print(f"✗ {error_msg}")
                    return False, None, error_msg

                last_error = f"HTTP {response.status_code}: {response.text}"

                if response.status_code in transient_status_codes and attempt < max_retries:
                    sleep_for = retry_delay_seconds * attempt
                    print(
                        f"⚠ Available numbers request attempt {attempt}/{max_retries} failed "
                        f"({response.status_code}). Retrying in {sleep_for:.1f}s..."
                    )
                    time.sleep(sleep_for)
                    continue

                print(f"✗ Failed to fetch available numbers: {last_error}")
                return False, None, last_error

            fallback_error = last_error or "Failed to fetch available numbers"
            print(f"✗ {fallback_error}")
            return False, None, fallback_error

        except requests.exceptions.RequestException as e:
            last_exception = str(e)
            for attempt in range(2, max_retries + 1):
                sleep_for = retry_delay_seconds * (attempt - 1)
                print(
                    f"⚠ Available numbers request error on attempt {attempt - 1}/{max_retries}: {last_exception}. "
                    f"Retrying in {sleep_for:.1f}s..."
                )
                time.sleep(sleep_for)
                try:
                    response = requests.get(url, params=params, headers=headers, timeout=120)
                    if response.status_code in [200, 201]:
                        try:
                            response_data = response.json()
                        except json.JSONDecodeError:
                            error_msg = f"Invalid JSON from available numbers endpoint: {response.text}"
                            print(f"✗ {error_msg}")
                            return False, None, error_msg

                        available_numbers = self._extract_available_numbers(response_data)
                        if available_numbers:
                            return True, available_numbers, None

                        error_msg = "No available numbers returned in response payload"
                        print(f"✗ {error_msg}")
                        return False, None, error_msg

                    if response.status_code not in {429, 500, 502, 503, 504}:
                        error_msg = f"HTTP {response.status_code}: {response.text}"
                        print(f"✗ Failed to fetch available numbers: {error_msg}")
                        return False, None, error_msg

                    last_exception = f"HTTP {response.status_code}: {response.text}"
                except requests.exceptions.RequestException as retry_exc:
                    last_exception = str(retry_exc)

            error_msg = f"Request error while fetching available numbers: {last_exception}"
            print(f"✗ {error_msg}")
            return False, None, error_msg

    def get_available_number(
        self,
        number_type: str = "standard_geographic",
        area_code: str = "44161",
        max_retries: int = 4,
        retry_delay_seconds: float = 1.5
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Request an available number from partner API.

        Args:
            number_type: Number type query parameter (default: standard_geographic)
            area_code: Area code query parameter (default: 44161)
            max_retries: Number of attempts for transient failures
            retry_delay_seconds: Base delay between retry attempts

        Returns:
            Tuple of (success, number, error_message)
        """
        success, numbers, error = self.get_available_numbers(
            number_type=number_type,
            area_code=area_code,
            max_retries=max_retries,
            retry_delay_seconds=retry_delay_seconds
        )

        if success and numbers:
            return True, numbers[0], None

        return False, None, error or "Failed to fetch available number"
    
    def set_token(self, token: str):
        """
        Set the authentication token
        
        Args:
            token: JWT authentication token
        """
        self.auth_token = token
    
    def create_user(
        self,
        customer_id: str,
        full_name: str,
        email: str,
        address: Dict[str, str],
        user_type: str = "standard",
        phone_number: Optional[str] = None
    ) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Create a new user for a customer
        
        Args:
            customer_id: Customer ID to add user to
            full_name: Full name of the user
            email: Email address of the user
            address: Dictionary with keys: premises, street, town, county, postcode
            user_type: 'standard' or 'admin' (default: 'standard')
            phone_number: Phone number of the user (optional)
            
        Returns:
            Tuple of (success, user_data, error_message)
        """
        if not self.auth_token:
            error_msg = "No authentication token available. Set token first using set_token()"
            print(f"✗ {error_msg}")
            return False, None, error_msg
        
        # Build payload
        payload = {
            "type": user_type.lower(),
            "fullName": full_name,
            "email": email,
            "address": address
        }
        
        # Add phone number if provided
        if phone_number:
            payload["phoneNumber"] = phone_number
        
        headers = {
            "accept": "*/*",
            "content-type": "application/json",
            "authorization": f"Bearer {self.auth_token}",
            "cache-control": "no-cache",
            "connection": "keep-alive"
        }
        
        # Build URL with customer ID
        url = f"{self.base_url}/{customer_id}/users"
        
        try:
            print(f"Creating user for customer {customer_id} in {self.environment.upper()} environment...")
            print(f"  Full Name: {full_name}")
            print(f"  Email: {email}")
            print(f"  Type: {user_type}")
            if phone_number:
                print(f"  Phone: {phone_number}")
            print(f"  Address: {address.get('premises', '')}, {address.get('postcode', '')}")
            
            response = requests.post(url, json=payload, headers=headers, timeout=120)
            
            # Check if request was successful (200 or 201)
            if response.status_code in [200, 201]:
                user_data = response.json()
                print(f"✓ User created successfully!")
                print(f"  DEBUG - Full API response: {json.dumps(user_data, indent=2)}")
                
                # Display key user information
                if "ID" in user_data:
                    print(f"  User ID: {user_data['ID']}")
                elif "id" in user_data:
                    print(f"  User ID: {user_data['id']}")
                elif "userId" in user_data:
                    print(f"  User ID: {user_data['userId']}")
                elif "userID" in user_data:
                    print(f"  User ID: {user_data['userID']}")
                elif "user_id" in user_data:
                    print(f"  User ID: {user_data['user_id']}")
                else:
                    print(f"  ⚠ WARNING: No user ID found in response!")
                    
                if "type" in user_data:
                    print(f"  Type: {user_data['type']}")
                
                return True, user_data, None
            elif response.status_code == 409:
                # Conflict - check if it's because user already exists
                response_text = response.text.lower()
                if "email" in response_text and ("already in use" in response_text or "already exists" in response_text):
                    # User already exists - treat as success since the account is created
                    warning_msg = "User already exists with this email (409 Conflict)"
                    print(f"✓ {warning_msg}")
                    # Return minimal user data indicating it already exists
                    user_data = {
                        "email": email,
                        "fullName": full_name,
                        "alreadyExists": True
                    }
                    return True, user_data, None
                else:
                    # Other 409 conflict - treat as error
                    error_msg = f"HTTP {response.status_code}: {response.text}"
                    print(f"✗ Conflict error: {error_msg}")
                    return False, None, error_msg
            elif response.status_code == 500 and "Timeout" in response.text:
                # Timeout error - user may have been created despite the error
                error_msg = f"HTTP {response.status_code}: {response.text}"
                warning_msg = "⚠ TIMEOUT: User may have been created successfully despite timeout error. Please verify in portal."
                print(f"⚠ {warning_msg}")
                print(f"  Response: {error_msg}")
                return False, None, f"{warning_msg} | {error_msg}"
            else:
                # HTTP error occurred
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"✗ User creation failed: {error_msg}")
                return False, None, error_msg
                
        except requests.exceptions.RequestException as e:
            # Network or request error
            error_msg = f"Request error: {str(e)}"
            print(f"✗ {error_msg}")
            return False, None, error_msg
        except json.JSONDecodeError as e:
            # JSON parsing error
            error_msg = f"JSON decode error: {str(e)}"
            print(f"✗ {error_msg}")
            return False, None, error_msg
        except Exception as e:
            # Unexpected error
            error_msg = f"Unexpected error: {str(e)}"
            print(f"✗ {error_msg}")
            return False, None, error_msg
    
    def create_multiple_users(
        self,
        customer_id: str,
        users: List[Dict]
    ) -> Tuple[int, int, List[Dict]]:
        """
        Create multiple users for a customer
        
        Args:
            customer_id: Customer ID to add users to
            users: List of user dictionaries with keys: fullName, email, address, type (optional), phoneNumber (optional)
            
        Returns:
            Tuple of (successful_count, failed_count, results_list)
        """
        results = []
        successful = 0
        failed = 0
        
        print(f"\nCreating {len(users)} user(s) for customer {customer_id}...")
        
        for idx, user in enumerate(users, 1):
            print(f"\n  [{idx}/{len(users)}] Processing user: {user.get('fullName', 'Unknown')}")
            
            success, user_data, error = self.create_user(
                customer_id=customer_id,
                full_name=user['fullName'],
                email=user['email'],
                address=user['address'],
                user_type=user.get('type', 'standard'),
                phone_number=user.get('phoneNumber')
            )
            
            result = {
                'fullName': user['fullName'],
                'email': user['email'],
                'success': success,
                'user_id': (user_data.get('ID') or user_data.get('id')) if user_data else None,
                'error': error
            }
            results.append(result)
            
            if success:
                successful += 1
            else:
                failed += 1
        
        print(f"\nUser creation summary: {successful} successful, {failed} failed")
        return successful, failed, results


def main():
    """
    Example usage and testing
    """
    from phoneline_plus_jwt_auth import PhonelinePlusAuth
    
    # Example credentials
    KEY_ID = "31ef1028-2075-40fd-9acc-d40021d0c931"
    SECRET = "TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7"
    
    print("=" * 60)
    print("Phoneline+ User Creation - UAT Environment")
    print("=" * 60)
    
    # Authenticate
    print("\n[Step 1] Authenticating...")
    auth = PhonelinePlusAuth(environment="uat")
    success, token, error = auth.generate_token(KEY_ID, SECRET)
    
    if not success:
        print(f"Authentication failed: {error}")
        return
    
    # Create user manager
    user_manager = PhonelinePlusUser(environment="uat", auth_token=token)
    
    # Example: Create single user
    CUSTOMER_ID = "your-customer-id-here"  # Replace with actual customer ID
    
    print(f"\n[Step 2] Creating user for customer {CUSTOMER_ID}...")
    success, user_data, error = user_manager.create_user(
        customer_id=CUSTOMER_ID,
        full_name="John Smith",
        email="john.smith@example.com",
        address={
            "premises": "Unit 1",
            "street": "High Street",
            "town": "London",
            "county": "Greater London",
            "postcode": "SW1A1AA"
        },
        user_type="standard",
        phone_number="07700900123"
    )
    
    if success:
        print("\n[User Details]")
        print(json.dumps(user_data, indent=2))
    else:
        print(f"\nError: {error}")
    
    # Example: Create multiple users
    print("\n" + "=" * 60)
    print("[Step 3] Creating multiple users...")
    print("=" * 60)
    
    users_to_create = [
        {
            "fullName": "Admin User",
            "email": "admin@example.com",
            "type": "admin",
            "phoneNumber": "07700900456",
            "address": {
                "premises": "Building 2",
                "street": "Main Road",
                "town": "Manchester",
                "county": "Greater Manchester",
                "postcode": "M11AA"
            }
        },
        {
            "fullName": "Standard User",
            "email": "user@example.com",
            "type": "standard",
            "address": {
                "premises": "Office 3",
                "street": "Park Lane",
                "town": "Birmingham",
                "county": "West Midlands",
                "postcode": "B11AA"
            }
        }
    ]
    
    successful, failed, results = user_manager.create_multiple_users(
        customer_id=CUSTOMER_ID,
        users=users_to_create
    )
    
    print("\n[Results]")
    for result in results:
        status = "✓" if result['success'] else "✗"
        print(f"{status} {result['fullName']}: {result.get('user_id', result.get('error'))}")


if __name__ == "__main__":
    main()

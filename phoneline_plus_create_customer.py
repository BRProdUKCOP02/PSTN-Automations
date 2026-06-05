"""
Phoneline+ Customer Creation
Creates new customers in the Gamma Phoneline+ platform
Supports both Production and UAT environments
"""

import requests
import json
from typing import Dict, Optional, Tuple
from datetime import datetime
from phoneline_plus_jwt_auth import PhonelinePlusAuth


class PhonelinePlusCustomer:
    """Handles Hardware Only customer creation for Phoneline+ API"""
    
    # API Endpoints
    PROD_CUSTOMER_URL = "https://api-ss-gb-aws.gammaapi.net/partner/v1/customers"
    UAT_CUSTOMER_URL = "https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/customers"
    
    def __init__(self, environment: str = "uat", auth_token: Optional[str] = None):
        """
        Initialize the customer creation handler
        
        Args:
            environment: Either 'production' or 'uat' (default: 'uat')
            auth_token: Optional JWT token. If not provided, must call set_token() later
        """
        self.environment = environment.lower()
        self.customer_url = self.PROD_CUSTOMER_URL if self.environment == "production" else self.UAT_CUSTOMER_URL
        self.auth_token = auth_token
        self.auth = None
    
    def set_token(self, token: str):
        """
        Set the authentication token
        
        Args:
            token: JWT authentication token
        """
        self.auth_token = token
    
    def authenticate(self, key_id: str, secret: str) -> Tuple[bool, Optional[str]]:
        """
        Generate authentication token using credentials
        
        Args:
            key_id: Partner API Key ID
            secret: Partner API Secret
            
        Returns:
            Tuple of (success, error_message)
        """
        self.auth = PhonelinePlusAuth(environment=self.environment)
        success, token, error = self.auth.generate_token(key_id, secret)
        
        if success:
            self.auth_token = token
            return True, None
        else:
            return False, error
    
    def create_customer(
        self, 
        company_name: str, 
        full_name: str, 
        contact_number: str,
        number: Optional[str],
        address: Dict[str, str],
        plan: str = "hardware_only",
        email: Optional[str] = None
    ) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Create a new Hardware Only customer
        
        Args:
            company_name: Company/customer name
            full_name: Full name of primary contact
            contact_number: Contact telephone number
            number: Phone number in E164 format (required for hardware_only, optional for hardware_only_cp, e.g., '443300163504')
            address: Dictionary with keys: premises, street, town, county, postcode
            plan: Plan type - 'hardware_only' or 'hardware_only_cp' (default: 'hardware_only')
            email: Email address (optional)
            
        Returns:
            Tuple of (success, customer_data, error_message)
        """
        if not self.auth_token:
            error_msg = "No authentication token available. Call authenticate() or set_token() first."
            print(f"✗ {error_msg}")
            return False, None, error_msg
        
        # Build payload for Hardware Only account - match Postman field order
        payload = {
            "plan": plan,
            "address": address,
            "companyName": company_name,
            "fullName": full_name,
            "contactNumber": contact_number
        }
        
        # Add number if provided (required for hardware_only, optional for hardware_only_cp)
        if number:
            payload["number"] = {"numberE164": number}
        
        # Add email if provided (optional for Hardware Only)
        if email:
            payload["email"] = email
        
        headers = {
            "accept": "*/*",
            "content-type": "application/json",
            "authorization": f"Bearer {self.auth_token}",
            "cache-control": "no-cache",
            "connection": "keep-alive"
        }
        
        try:
            print(f"Creating Hardware Only customer in {self.environment.upper()} environment...")
            print(f"Company Name: {company_name}")
            print(f"Full Name: {full_name}")
            if email:
                print(f"Email: {email}")
            print(f"Contact Number: {contact_number}")
            if number:
                print(f"Number (E164): {number}")
            else:
                print(f"Number: Not provided (customer will provide later)")
            print(f"Address: {address.get('premises', '')}, {address.get('postcode', '')}")
            print(f"\nDEBUG - Request URL: {self.customer_url}")
            print(f"DEBUG - Request Payload: {json.dumps(payload, indent=2)}")
            print(f"DEBUG - Request Headers: {json.dumps({k: v for k, v in headers.items() if k != 'authorization'}, indent=2)}")
            response = requests.post(self.customer_url, json=payload, headers=headers, timeout=120)
            
            # Check if request was successful (200 or 201)
            if response.status_code in [200, 201]:
                customer_data = response.json() if response.content else {}
                print(f"✓ Customer created successfully!")
                print(f"\nDEBUG - API Response:")
                print(json.dumps(customer_data, indent=2))
                print()
                
                # Display key customer information
                # Note: API returns 'customerID' not 'id'
                if "customerID" in customer_data:
                    print(f"  Customer ID: {customer_data['customerID']}")
                    # Normalize the key to 'id' for consistency
                    customer_data['id'] = customer_data['customerID']
                elif "id" in customer_data:
                    print(f"  Customer ID: {customer_data['id']}")
                    
                if "plan" in customer_data:
                    print(f"  Plan: {customer_data['plan']}")
                
                # Display SIP credentials if present (nested in destination object)
                destination = customer_data.get('destination', {})
                if destination and 'sipID' in destination:
                    print(f"  SIP Credentials:")
                    print(f"    Registration Server: {destination.get('registrationServer', 'N/A')}")
                    print(f"    SIP ID: {destination.get('sipID', 'N/A')}")
                    print(f"    SIP Password: {destination.get('sipPassword', 'N/A')}")
                    if 'stunServer' in destination:
                        print(f"    STUN Server: {destination.get('stunServer', 'N/A')}")
                        print(f"    STUN Port: {destination.get('stunPort', 'N/A')}")
                
                return True, customer_data, None
            elif response.status_code == 500 and "Timeout" in response.text:
                # Timeout error - customer may have been created despite the error
                error_msg = f"HTTP {response.status_code}: {response.text}"
                warning_msg = "TIMEOUT: Customer may have been created successfully despite timeout error. Please verify in portal."
                print(f"⚠ {warning_msg}")
                print(f"  Response: {error_msg}")
                return False, None, f"{warning_msg} | {error_msg}"
            else:
                # HTTP error occurred
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"✗ Customer creation failed: {error_msg}")
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


def main():
    """
    Example usage and testing
    """
    # Example credentials (replace with actual values from input sheet)
    KEY_ID = "31ef1028-2075-40fd-9acc-d40021d0c931"
    SECRET = "TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7"
    
    print("=" * 60)
    print("Phoneline+ Customer Creation - UAT Environment")
    print("=" * 60)
    
    # Method 1: Initialize with authentication
    customer_manager = PhonelinePlusCustomer(environment="uat")
    
    # Authenticate
    print("\n[Step 1] Authenticating...")
    success, error = customer_manager.authenticate(KEY_ID, SECRET)
    
    if not success:
        print(f"Authentication failed: {error}")
        return
    
    # Create Hardware Only customer
    print("\n[Step 2] Creating Hardware Only customer...")
    success, customer_data, error = customer_manager.create_customer(
        company_name="Test HWO Company",
        full_name="Mike Test",
        contact_number="07944611318",
        number="443300163504",
        address={
            "premises": "Building 5",
            "street": "Main Street",
            "town": "Manchester",
            "county": "Greater Manchester",
            "postcode": "M27 5WW"
        },
        email="mike.test@example.com"
    )
    
    if success:
        print("\n[Customer Details]")
        print(json.dumps(customer_data, indent=2))
    else:
        print(f"\nError: {error}")
    
    # Alternative Method 2: Using token from separate auth
    print("\n" + "=" * 60)
    print("Alternative Method: Using Pre-generated Token")
    print("=" * 60)
    
    auth = PhonelinePlusAuth(environment="uat")
    success, token, error = auth.generate_token(KEY_ID, SECRET)
    
    if success:
        customer_manager2 = PhonelinePlusCustomer(environment="uat", auth_token=token)
        success, customer_data, error = customer_manager2.create_customer(
            company_name="Another HWO Test",
            full_name="Jane Doe",
            contact_number="07700900123",
            number="441234567890",
            address={
                "premises": "123 Test Building",
                "street": "High Street",
                "town": "London",
                "county": "Greater London",
                "postcode": "SW1A1AA"
            }
        )
        
        if success:
            print("\n[Customer Details]")
            print(f"Customer ID: {customer_data.get('id', 'N/A')}")
            print(f"Plan: {customer_data.get('plan', 'N/A')}")
            print(f"Number: {customer_data.get('number', {}).get('numberE164', 'N/A')}")


if __name__ == "__main__":
    main()

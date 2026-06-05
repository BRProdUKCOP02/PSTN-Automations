"""
Phoneline+ Number Management
Manages number allocation/deallocation for Phoneline+ customers via API

Supports:
1. Getting all numbers for a customer
2. Allocating numbers to users
3. Deallocating numbers from users
4. Identifying geographical vs non-geographical numbers
"""

import requests
from typing import Dict, List, Tuple, Optional


class PhonelinePlusNumberManager:
    """
    Handle number management for Phoneline+ customers
    """
    
    def __init__(self, environment: str = "uat", auth_token: str = None):
        """
        Initialize the number manager
        
        Args:
            environment: "uat" or "production"
            auth_token: JWT authentication token (required)
        """
        self.environment = environment.lower()
        
        # API endpoints
        if self.environment == "production":
            self.base_url = "https://api-ss-gb-aws.gammaapi.net/partner/v1"
        else:
            self.base_url = "https://api-ss-gb-aws-uat.gammaapi.net/partner/v1"
        
        self.auth_token = auth_token
    
    def get_customer_users(self, customer_id: str) -> Tuple[bool, Optional[List[Dict]], Optional[str]]:
        """
        Get all users for a customer
        
        Args:
            customer_id: Customer UUID
            
        Returns:
            Tuple of (success, list_of_users, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/users/"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'accept': '*/*',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        try:
            print(f"Fetching users for customer {customer_id}...")
            
            response = requests.get(endpoint, headers=headers, timeout=120)
            
            if response.status_code in [200, 201]:
                data = response.json()
                users = data.get('data', [])
                count = data.get('count', 0)
                print(f"  ✓ Found {count} user(s)")
                return True, users, None
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"  ✗ Failed to fetch users: {error_msg}")
                return False, None, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout (120s)"
            print(f"  ✗ Timeout: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            print(f"  ✗ Error: {error_msg}")
            return False, None, error_msg
    
    def get_customer_numbers(self, customer_id: str) -> Tuple[bool, Optional[List[Dict]], Optional[str]]:
        """
        Get all numbers for a customer
        
        Args:
            customer_id: Customer UUID
            
        Returns:
            Tuple of (success, list_of_numbers, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/numbers/"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'accept': '*/*',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        try:
            print(f"Fetching numbers for customer {customer_id}...")
            
            response = requests.get(endpoint, headers=headers, timeout=120)
            
            if response.status_code in [200, 201]:
                data = response.json()
                numbers = data.get('data', [])
                count = data.get('count', 0)
                print(f"  ✓ Found {count} number(s)")
                return True, numbers, None
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"  ✗ Failed to fetch numbers: {error_msg}")
                return False, None, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout (120s)"
            print(f"  ✗ Timeout: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            print(f"  ✗ Error: {error_msg}")
            return False, None, error_msg
    
    def allocate_number(self, customer_id: str, number_id: str, user_id: str) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Allocate a number to a user
        
        Args:
            customer_id: Customer UUID
            number_id: Number UUID
            user_id: User UUID
            
        Returns:
            Tuple of (success, number_data, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/numbers/{number_id}/allocate"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'Content-Type': 'application/json',
            'accept': '*/*',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        payload = {
            "userID": user_id
        }
        
        try:
            response = requests.put(endpoint, json=payload, headers=headers, timeout=120)
            
            if response.status_code in [200, 201]:
                number_data = response.json()
                number_e164 = number_data.get('numberE164', 'Unknown')
                print(f"  ✓ Number {number_e164} allocated to user {user_id}")
                return True, number_data, None
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"  ✗ Failed to allocate number: {error_msg}")
                return False, None, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout (120s)"
            print(f"  ✗ Timeout: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            print(f"  ✗ Error: {error_msg}")
            return False, None, error_msg
    
    def deallocate_number(self, customer_id: str, number_id: str) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Deallocate a number from a user
        
        Args:
            customer_id: Customer UUID
            number_id: Number UUID
            
        Returns:
            Tuple of (success, number_data, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/numbers/{number_id}/deallocate"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'accept': '*/*',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        try:
            response = requests.delete(endpoint, headers=headers, timeout=120)
            
            if response.status_code in [200, 201]:
                number_data = response.json()
                number_e164 = number_data.get('numberE164', 'Unknown')
                print(f"  ✓ Number {number_e164} deallocated")
                return True, number_data, None
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"  ✗ Failed to deallocate number: {error_msg}")
                return False, None, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout (120s)"
            print(f"  ✗ Timeout: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            print(f"  ✗ Error: {error_msg}")
            return False, None, error_msg


def main():
    """
    Example usage
    """
    from phoneline_plus_jwt_auth import PhonelinePlusAuth
    
    print("=" * 60)
    print("Phoneline+ Number Management Example")
    print("=" * 60)
    
    # Configuration
    ENVIRONMENT = "uat"  # or "production"
    CUSTOMER_ID = "88032dd4-fed5-40e8-abda-66a1a746971c"
    
    # Credentials
    KEY_ID = "your-key-id-here"
    SECRET = "your-secret-here"
    
    # Authenticate
    print("\nAuthenticating...")
    auth_manager = PhonelinePlusAuth(environment=ENVIRONMENT)
    success, token, error = auth_manager.generate_token(KEY_ID, SECRET)
    
    if not success:
        print(f"✗ Authentication failed: {error}")
        return
    
    print("✓ Authentication successful")
    
    # Initialize number manager
    number_manager = PhonelinePlusNumberManager(environment=ENVIRONMENT, auth_token=token)
    
    # Get all numbers for customer
    print("\n" + "=" * 60)
    print("Getting customer numbers")
    print("=" * 60)
    success, numbers, error = number_manager.get_customer_numbers(CUSTOMER_ID)
    
    if success:
        # Analyze numbers
        geo_numbers = [n for n in numbers if n.get('type') == 'standard_geographic']
        nongeo_numbers = [n for n in numbers if n.get('type') == 'standard_nongeographic']
        
        print(f"\nNumber Summary:")
        print(f"  Geographical: {len(geo_numbers)}")
        print(f"  Non-geographical: {len(nongeo_numbers)}")
        
        print(f"\nGeographical Numbers:")
        for num in geo_numbers:
            allocated_to = num.get('allocatedTo', {}).get('userID', 'Unallocated')
            print(f"  {num['numberE164']} ({num['displayNumberType']}) - {allocated_to}")
        
        print(f"\nNon-geographical Numbers:")
        for num in nongeo_numbers:
            allocated_to = num.get('allocatedTo', {}).get('userID', 'Unallocated')
            print(f"  {num['numberE164']} ({num['displayNumberType']}) - {allocated_to}")


if __name__ == "__main__":
    main()

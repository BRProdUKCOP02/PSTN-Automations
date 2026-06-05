"""
Phoneline+ Add Device
Adds hardware devices to existing Phoneline+ users via API

Supports two device types:
1. Gamma-supplied devices: Requires brand, model, macAddress
2. Customer-owned devices: Requires brand="other" and device name
"""

import requests
from typing import Dict, Tuple, Optional


class PhonelinePlusDevice:
    """
    Handle device addition for Phoneline+ users
    """
    
    def __init__(self, environment: str = "uat", auth_token: str = None):
        """
        Initialize the device manager
        
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
    
    def add_gamma_device(self, customer_id: str, user_id: str, 
                        brand: str, model: str, mac_address: str) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Add a Gamma-supplied device to a user
        
        Args:
            customer_id: Customer UUID
            user_id: User UUID
            brand: Device brand (e.g., "Yealink", "Cisco")
            model: Device model (e.g., "T46S", "SPA112")
            mac_address: MAC address of the device
            
        Returns:
            Tuple of (success, device_data, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/users/{user_id}/devices"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'Content-Type': 'application/json',
            'accept': '*/*',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        payload = {
            "brand": brand,
            "model": model,
            "macAddress": mac_address
        }
        
        try:
            print(f"Adding Gamma device to user {user_id}...")
            print(f"  Brand: {brand}, Model: {model}, MAC: {mac_address}")
            
            response = requests.post(endpoint, json=payload, headers=headers, timeout=120)
            
            if response.status_code in [200, 201]:
                device_data = response.json()
                device_id = device_data.get('ID', 'Unknown')
                print(f"  ✓ Device added successfully (ID: {device_id})")
                return True, device_data, None
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"  ✗ Failed to add device: {error_msg}")
                return False, None, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout (120s)"
            print(f"  ✗ Timeout: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            print(f"  ✗ Error: {error_msg}")
            return False, None, error_msg
    
    def add_customer_device(self, customer_id: str, user_id: str, 
                           device_name: str) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Add a customer-owned device to a user
        Returns SIP credentials for the device
        
        Args:
            customer_id: Customer UUID
            user_id: User UUID
            device_name: Name for the device
            
        Returns:
            Tuple of (success, device_data_with_sip_credentials, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/users/{user_id}/devices"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'Content-Type': 'application/json',
            'accept': '*/*',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        payload = {
            "brand": "other",
            "name": device_name
        }
        
        try:
            print(f"Adding customer-owned device to user {user_id}...")
            print(f"  Device Name: {device_name}")
            
            response = requests.post(endpoint, json=payload, headers=headers, timeout=120)
            
            if response.status_code in [200, 201]:
                device_data = response.json()
                device_id = device_data.get('ID', 'Unknown')
                sip_id = device_data.get('sipID', '')
                print(f"  ✓ Device added successfully (ID: {device_id})")
                if sip_id:
                    print(f"  ✓ SIP credentials generated")
                    print(f"    SIP ID: {sip_id}")
                    # Don't print password in logs for security
                return True, device_data, None
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"  ✗ Failed to add device: {error_msg}")
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
    print("Phoneline+ Add Device Example")
    print("=" * 60)
    
    # Configuration
    ENVIRONMENT = "uat"  # or "production"
    CUSTOMER_ID = "your-customer-uuid-here"
    USER_ID = "your-user-uuid-here"
    
    # Credentials
    KEY_ID = "ec2d3063-0079-475a-a6a3-26df72a2e8d7"
    SECRET = "frzWEWZCZrw8TYev6TvXasSV3g6W9HHiM6zqAJ2UnpMe57LrCKAolKP7ZY0NAY3z"
    
    # Authenticate
    print("\nAuthenticating...")
    auth_manager = PhonelinePlusAuth(environment=ENVIRONMENT)
    success, token, error = auth_manager.generate_token(KEY_ID, SECRET)
    
    if not success:
        print(f"✗ Authentication failed: {error}")
        return
    
    print("✓ Authentication successful")
    
    # Initialize device manager
    device_manager = PhonelinePlusDevice(environment=ENVIRONMENT, auth_token=token)
    
    # Example 1: Add Gamma-supplied device
    print("\n" + "=" * 60)
    print("Adding Gamma-supplied device")
    print("=" * 60)
    success, device_data, error = device_manager.add_gamma_device(
        customer_id=CUSTOMER_ID,
        user_id=USER_ID,
        brand="Yealink",
        model="T46S",
        mac_address="00:15:65:12:34:56"
    )
    
    if success:
        print(f"\n✓ Device ID: {device_data.get('ID')}")
        print(f"  MAC Address: {device_data.get('macAddress')}")
        print(f"  Brand: {device_data.get('metadata', {}).get('brand')}")
        print(f"  Model: {device_data.get('metadata', {}).get('model')}")
    
    # Example 2: Add customer-owned device
    print("\n" + "=" * 60)
    print("Adding customer-owned device")
    print("=" * 60)
    success, device_data, error = device_manager.add_customer_device(
        customer_id=CUSTOMER_ID,
        user_id=USER_ID,
        device_name="Office Desk Phone"
    )
    
    if success:
        print(f"\n✓ Device ID: {device_data.get('ID')}")
        print(f"  SIP ID: {device_data.get('sipID')}")
        print(f"  Registration Server: {device_data.get('registrationServer')}")
        print(f"\n⚠ Store SIP credentials securely!")


if __name__ == "__main__":
    main()

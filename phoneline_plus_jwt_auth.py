"""
Phoneline+ JWT Token Generator
Generates JWT authentication token for Gamma API access
Supports both Production and UAT environments
"""

import requests
import json
from typing import Dict, Optional, Tuple
from datetime import datetime


class PhonelinePlusAuth:
    """Handles JWT token generation for Phoneline+ API authentication"""
    
    # API Endpoints
    PROD_AUTH_URL = "https://api-ss-gb-aws.gammaapi.net/partner/v1/auth"
    UAT_AUTH_URL = "https://api-ss-gb-aws-uat.gammaapi.net/partner/v1/auth"
    
    def __init__(self, environment: str = "uat"):
        """
        Initialize the authentication handler
        
        Args:
            environment: Either 'production' or 'uat' (default: 'uat')
        """
        self.environment = environment.lower()
        self.auth_url = self.PROD_AUTH_URL if self.environment == "production" else self.UAT_AUTH_URL
        self.token = None
        self.token_timestamp = None
    
    def generate_token(self, key_id: str, secret: str) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Generate JWT token using partner credentials
        
        Args:
            key_id: Partner API Key ID
            secret: Partner API Secret
            
        Returns:
            Tuple of (success, token, error_message)
        """
        payload = {
            "keyID": key_id,
            "secret": secret
        }
        
        headers = {
            "accept": "application/json",
            "content-type": "application/json"
        }
        
        try:
            print(f"Requesting JWT token from {self.environment.upper()} environment...")
            response = requests.post(self.auth_url, json=payload, headers=headers)
            
            # Check if request was successful (200 or 201)
            if response.status_code in [200, 201]:
                response_data = response.json()
                
                # Extract token from response - handle both "token" and "access_token" fields
                if "access_token" in response_data:
                    self.token = response_data["access_token"]
                    self.token_timestamp = datetime.now()
                    print(f"✓ Token generated successfully at {self.token_timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
                    return True, self.token, None
                elif "token" in response_data:
                    self.token = response_data["token"]
                    self.token_timestamp = datetime.now()
                    print(f"✓ Token generated successfully at {self.token_timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
                    return True, self.token, None
                else:
                    # Token not in expected format
                    error_msg = f"Token not found in response: {response.text}"
                    print(f"✗ {error_msg}")
                    return False, None, error_msg
            else:
                # HTTP error occurred
                error_msg = f"HTTP {response.status_code}: {response.text}"
                print(f"✗ Authentication failed: {error_msg}")
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
    
    def get_token(self) -> Optional[str]:
        """
        Get the current token
        
        Returns:
            Current JWT token or None if not generated
        """
        return self.token
    
    def get_auth_header(self) -> Dict[str, str]:
        """
        Get authorization header with token
        
        Returns:
            Dictionary with Authorization header
        """
        if self.token:
            return {"Authorization": f"Bearer {self.token}"}
        else:
            raise ValueError("No token available. Generate token first using generate_token()")


def main():
    """
    Example usage and testing
    """
    # Example credentials (replace with actual values from input sheet)
    KEY_ID = "31ef1028-2075-40fd-9acc-d40021d0c931"
    SECRET = "TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7"
    
    # Test Production environment
    print("=" * 60)
    print("Testing PRODUCTION Environment")
    print("=" * 60)
    auth_prod = PhonelinePlusAuth(environment="production")
    success, token, error = auth_prod.generate_token(KEY_ID, SECRET)
    
    if success:
        print(f"\nToken (first 50 chars): {token[:50]}...")
        print(f"\nAuthorization Header:")
        print(auth_prod.get_auth_header())
    else:
        print(f"\nError: {error}")
    
    # Test UAT environment
    print("\n" + "=" * 60)
    print("Testing UAT Environment")
    print("=" * 60)
    auth_uat = PhonelinePlusAuth(environment="uat")
    success, token, error = auth_uat.generate_token(KEY_ID, SECRET)
    
    if success:
        print(f"\nToken (first 50 chars): {token[:50]}...")
        print(f"\nAuthorization Header:")
        print(auth_uat.get_auth_header())
    else:
        print(f"\nError: {error}")


if __name__ == "__main__":
    main()

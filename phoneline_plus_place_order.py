"""
Phoneline+ Place Order
Places hardware orders for Phoneline+ customers via API

Orders include:
- Delivery contact information
- Delivery address
- Products (with quantities)
- SKU to user mapping for device assignment
"""

import requests
import json
from typing import Dict, List, Tuple, Optional


class PhonelinePlusOrder:
    """
    Handle order placement for Phoneline+ customers
    """
    
    def __init__(self, environment: str = "uat", auth_token: str = None):
        """
        Initialize the order manager
        
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

    def place_order(self,
                   customer_id: str,
                   name: str,
                   email: str,
                   phone_number: str,
                   delivery_address: Dict[str, str],
                   products: List[Dict[str, any]],
                   sku_to_user_mapping: Optional[Dict[str, List[Dict[str, str]]]] = None,
                   tracking_email: Optional[str] = None,
                   devices: Optional[List[Dict[str, any]]] = None) -> Tuple[bool, Optional[Dict], Optional[str]]:
        """
        Place a hardware order for a customer
        
        Args:
            customer_id: Customer UUID     name: Delivery contact name
            email: Delivery contact email
            phone_number: Delivery contact phone number
            delivery_address: Dictionary with keys:
                - line1: Address line 1 (required)
                - line2: Address line 2 (optional)
                - line3: Address line 3 (optional)
                - town: Town/City (required)
                - county: County (optional)
                - country: Country (optional, defaults to "United Kingdom")
                - postcode: Postcode (required)
            products: List of product dictionaries with keys:
                - ID: Product SKU UUID (required)
                - quantity: Number of units (required)
            sku_to_user_mapping: Optional mapping of SKU to user IDs
                Dictionary with SKU as key and list of user dictionaries as value
                Example: {"sku_id": [{"userID": "user_uuid"}]}
            tracking_email: Email for shipment tracking notifications (defaults to contact email)
            devices: Device list for order payload (defaults to products list)
                
        Returns:
            Tuple of (success, order_data, error_message)
        """
        endpoint = f"{self.base_url}/customers/{customer_id}/orders"
        
        headers = {
            'Authorization': f'Bearer {self.auth_token}',
            'Content-Type': 'application/json',
            'accept': 'application/json',
            'cache-control': 'no-cache',
            'connection': 'keep-alive'
        }
        
        # Build payload in strict documentation format
        payload = {
            "deliveryAddress": {
                "line1": delivery_address.get("line1", ""),
                "town": delivery_address.get("town", ""),
                "postcode": delivery_address.get("postcode", "")
            },
            "products": products,
            "name": name,
            "email": email,
            "phoneNumber": phone_number
        }
        
        # Add optional address fields
        if delivery_address.get("line2"):
            payload["deliveryAddress"]["line2"] = delivery_address["line2"]
        if delivery_address.get("line3"):
            payload["deliveryAddress"]["line3"] = delivery_address["line3"]
        if delivery_address.get("county"):
            payload["deliveryAddress"]["county"] = delivery_address["county"]
        
        # Add country (default to UK if not specified)
        payload["deliveryAddress"]["country"] = delivery_address.get("country", "United Kingdom")
        
        # Add SKU to user mapping if provided
        if sku_to_user_mapping:
            payload["skuToUserMapping"] = sku_to_user_mapping

        # Strict docs mode: do not send trackingEmail/devices unless explicitly provided
        if tracking_email and str(tracking_email).strip():
            payload["trackingEmail"] = str(tracking_email).strip()
        if devices and len(devices) > 0:
            payload["devices"] = devices
        
        try:
            print(f"\nPlacing order for customer {customer_id}...")
            print(f"  Contact: {name} ({email})")
            print(f"  Delivery: {delivery_address.get('line1')}, {delivery_address.get('postcode')}")
            print(f"  Products: {len(products)} item(s)")
            
            # Debug output
            print(f"\nDEBUG - Request URL: {endpoint}")
            print(f"DEBUG - Request Payload: {json.dumps(payload, indent=2)}")
            print(f"DEBUG - Request Headers: {{")
            for k, v in headers.items():
                if k.lower() == 'authorization':
                    print(f'  "{k}": "Bearer <token>"')
                else:
                    print(f'  "{k}": "{v}"')
            print("}")
            
            response = requests.post(endpoint, json=payload, headers=headers, timeout=120, allow_redirects=False)

            redirect_statuses = {301, 302, 303, 307, 308}
            redirect_attempts = 0
            while response.status_code in redirect_statuses and redirect_attempts < 3:
                redirect_location = response.headers.get('Location')
                if not redirect_location:
                    break

                redirect_attempts += 1
                print(f"\nDEBUG - Redirect detected ({response.status_code}) to: {redirect_location}")
                response = requests.post(
                    redirect_location,
                    json=payload,
                    headers=headers,
                    timeout=120,
                    allow_redirects=False
                )

            print(f"\nDEBUG - Response Status: {response.status_code}")
            print(f"DEBUG - Response Body: {response.text}")

            if response.status_code in [200, 201]:
                order_data = response.json()
                order_id = order_data.get('orderID', order_data.get('id', 'Unknown'))
                print(f"✓ Order placed successfully (Order ID: {order_id})")
                return True, order_data, None

            error_msg = f"HTTP {response.status_code}: {response.text}"
            print(f"✗ Failed to place order: {error_msg}")
            return False, None, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "Request timeout (120s)"
            print(f"✗ Timeout: {error_msg}")
            return False, None, error_msg
        except Exception as e:
            error_msg = str(e)
            print(f"✗ Error: {error_msg}")
            return False, None, error_msg

"""
Gamma Broadband Ordering API Client
Version: 1.1

This module provides a Python client for interacting with the Gamma Broadband
Ordering API for provisioning and managing Broadband orders.
"""

import requests
from typing import Optional, Dict, List, Any
from datetime import datetime, timedelta
import json


class GammaOrderingAPIError(Exception):
    """Custom exception for Gamma Ordering API errors"""
    def __init__(self, message: str, status_code: int = None, response_data: Dict = None):
        self.message = message
        self.status_code = status_code
        self.response_data = response_data
        super().__init__(self.message)


class GammaAuthenticationError(GammaOrderingAPIError):
    """Exception for authentication failures"""
    pass


class GammaBroadbandOrderingAPI:
    """
    Client for Gamma Broadband Ordering API
    
    This client provides methods to:
    - View broadband orders
    - Search for orders
    - Place orders
    - Cancel/cease orders
    - Place and view regrades
    """
    
    PRODUCTION_URL = "https://api.gamma.co.uk"
    TEST_URL = "https://api-test.gamma.co.uk"
    AUTH_PRODUCTION_URL = "https://api.gamma.co.uk/auth/token"
    AUTH_TEST_URL = "https://api-test.gamma.co.uk/auth/token"
    
    def __init__(self, username: str = None, password: str = None, 
                 bearer_token: str = None, use_production: bool = False,
                 auto_refresh: bool = True):
        """
        Initialize the API client
        
        Args:
            username: Gamma API username (for JWT authentication)
            password: Gamma API password (for JWT authentication)
            bearer_token: Pre-obtained JWT bearer token (alternative to username/password)
            use_production: If True, use production environment. Default is test
            auto_refresh: If True, automatically refresh token when it expires
        """
        self.username = username
        self.password = password
        self.use_production = use_production
        self.auto_refresh = auto_refresh
        
        self.base_url = self.PRODUCTION_URL if use_production else self.TEST_URL
        self.auth_url = self.AUTH_PRODUCTION_URL if use_production else self.AUTH_TEST_URL
        
        self.access_token = bearer_token
        self.refresh_token = None
        self.token_expiry = None
        self.refresh_expiry = None
        
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/json',
            'Accept': '*/*',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive'
        })
        
        # If username/password provided, obtain token automatically
        if username and password and not bearer_token:
            self._obtain_token()
        elif bearer_token:
            self._set_token(bearer_token)
    
    def _obtain_token(self) -> Dict[str, Any]:
        """Obtain JWT token using username and password"""
        data = {
            'grant_type': 'password',
            'username': self.username,
            'password': self.password
        }
        
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        
        try:
            response = requests.post(self.auth_url, data=data, headers=headers)
            
            if response.status_code == 200:
                token_data = response.json()
                self._process_token_response(token_data)
                return token_data
            elif response.status_code == 401:
                raise GammaAuthenticationError(
                    "Authentication failed: Invalid username or password",
                    status_code=401
                )
            else:
                raise GammaAuthenticationError(
                    f"Authentication failed with status {response.status_code}",
                    status_code=response.status_code
                )
        except requests.exceptions.RequestException as e:
            raise GammaAuthenticationError(f"Authentication request failed: {str(e)}")
    
    def _refresh_access_token(self) -> Dict[str, Any]:
        """Refresh the access token using the refresh token"""
        if not self.refresh_token:
            raise GammaAuthenticationError("No refresh token available")
        
        data = {
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token
        }
        
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        
        try:
            response = requests.post(self.auth_url, data=data, headers=headers)
            
            if response.status_code == 200:
                token_data = response.json()
                self._process_token_response(token_data)
                return token_data
            else:
                if self.username and self.password:
                    return self._obtain_token()
                else:
                    raise GammaAuthenticationError(
                        f"Token refresh failed with status {response.status_code}",
                        status_code=response.status_code
                    )
        except requests.exceptions.RequestException as e:
            raise GammaAuthenticationError(f"Token refresh request failed: {str(e)}")
    
    def _process_token_response(self, token_data: Dict[str, Any]) -> None:
        """Process token response and update client state"""
        self.access_token = token_data.get('access_token')
        self.refresh_token = token_data.get('refresh_token')
        
        expires_in = token_data.get('expires_in', 600)
        refresh_expires_in = token_data.get('refresh_expires_in', 2592000)
        
        self.token_expiry = datetime.now() + timedelta(seconds=expires_in)
        self.refresh_expiry = datetime.now() + timedelta(seconds=refresh_expires_in)
        
        self._set_token(self.access_token)
    
    def _set_token(self, token: str) -> None:
        """Set the authorization token in session headers"""
        self.session.headers.update({
            'Authorization': f'Bearer {token}'
        })
    
    def _ensure_valid_token(self) -> None:
        """Ensure we have a valid token, refreshing if necessary"""
        if not self.auto_refresh:
            return
        
        if self.token_expiry:
            time_until_expiry = (self.token_expiry - datetime.now()).total_seconds()
            if time_until_expiry < 30:
                self._refresh_access_token()
    
    def _make_request(self, method: str, endpoint: str, **kwargs) -> Dict[str, Any]:
        """Make an HTTP request to the API"""
        self._ensure_valid_token()
        
        url = f"{self.base_url}{endpoint}"
        
        try:
            response = self.session.request(method, url, **kwargs)
            
            # Debug output - Enhanced for troubleshooting
            print(f"\n{'='*80}")
            print(f"DEBUG: REQUEST DETAILS")
            print(f"{'='*80}")
            print(f"Method: {method}")
            print(f"URL: {url}")
            print(f"\nRequest Headers:")
            for header, value in response.request.headers.items():
                # Mask the token for security but show it exists
                if header.lower() == 'authorization':
                    if value:
                        token_preview = value[:20] + "..." + value[-10:] if len(value) > 30 else value
                        print(f"  {header}: {token_preview}")
                else:
                    print(f"  {header}: {value}")
            
            if kwargs.get('json'):
                print(f"\nRequest Body (JSON):")
                import json
                print(json.dumps(kwargs['json'], indent=2)[:1000])
            
            print(f"\n{'='*80}")
            print(f"DEBUG: RESPONSE DETAILS")
            print(f"{'='*80}")
            print(f"Status Code: {response.status_code}")
            print(f"Response Headers:")
            for header, value in response.headers.items():
                print(f"  {header}: {value}")
            
            if response.content:
                print(f"\nResponse Body:")
                print(f"{response.text[:500]}")
            print(f"{'='*80}\n")
            
            # Success responses
            if response.status_code in [200, 201]:
                # If no content but there's a Location header (e.g., 201 Created), extract ID from it
                if not response.content and 'Location' in response.headers:
                    location = response.headers['Location']
                    # Extract order ID from location URL (e.g., .../orders/420896)
                    if '/orders/' in location:
                        order_id = location.split('/orders/')[-1].strip()
                        return {'id': order_id}
                return response.json() if response.content else {}
            
            # Client errors
            elif response.status_code == 400:
                error_data = response.json() if response.content else {}
                error_msg = error_data.get('message', 'Bad request - Invalid data or validation error')
                raise GammaOrderingAPIError(
                    error_msg,
                    status_code=400,
                    response_data=error_data
                )
            
            elif response.status_code == 401:
                raise GammaAuthenticationError(
                    "Authentication token expired or invalid - Please re-authenticate",
                    status_code=401
                )
            
            elif response.status_code == 403:
                raise GammaOrderingAPIError(
                    "Access forbidden - You don't have permission to access this resource",
                    status_code=403,
                    response_data=None
                )
            
            elif response.status_code == 404:
                raise GammaOrderingAPIError(
                    "Order not found or not accessible",
                    status_code=404,
                    response_data=None
                )
            
            elif response.status_code == 429:
                retry_after = response.headers.get('Retry-After', 'unknown')
                raise GammaOrderingAPIError(
                    f"Rate limit exceeded - Too many requests. Retry after: {retry_after} seconds",
                    status_code=429,
                    response_data=None
                )
            
            # Server errors
            elif response.status_code == 500:
                error_data = response.json() if response.content else {}
                error_msg = error_data.get('message', 'Internal server error')
                raise GammaOrderingAPIError(
                    error_msg,
                    status_code=500,
                    response_data=error_data
                )
            
            elif response.status_code == 502:
                raise GammaOrderingAPIError(
                    "Bad gateway - Upstream server error",
                    status_code=502,
                    response_data=None
                )
            
            elif response.status_code == 503:
                retry_after = response.headers.get('Retry-After', 'unknown')
                raise GammaOrderingAPIError(
                    f"Service unavailable - API is temporarily down. Retry after: {retry_after} seconds",
                    status_code=503,
                    response_data=None
                )
            
            # Any other unexpected status code
            else:
                raise GammaOrderingAPIError(
                    f"Unexpected status code: {response.status_code}",
                    status_code=response.status_code,
                    response_data=None
                )
                
        except requests.exceptions.RequestException as e:
            raise GammaOrderingAPIError(f"Request failed: {str(e)}")
    
    # ==================== Order Operations ====================
    
    def get_order(self, order_id: int) -> Dict[str, Any]:
        """
        View a broadband order
        
        Args:
            order_id: The broadband order ID
            
        Returns:
            Dictionary containing complete order details and updates
            
        Example:
            >>> api = GammaBroadbandOrderingAPI(username, password)
            >>> order = api.get_order(10000)
            >>> print(f"Status: {order['status']}")
        """
        endpoint = f"/access-ordering/v1/orders/{order_id}"
        return self._make_request('GET', endpoint)
    
    def search_orders(
        self,
        broadband_product: Optional[str] = None,
        status: Optional[str] = None,
        limit: int = 100,
        offset: int = 0,
        **kwargs
    ) -> List[Dict[str, Any]]:
        """
        Search for broadband orders
        
        Args:
            broadband_product: Filter by product name
            status: Filter by order status
            limit: Maximum number of results (default 100)
            offset: Offset for pagination (default 0)
            **kwargs: Additional filter parameters
            
        Returns:
            List of order summaries
            
        Example:
            >>> orders = api.search_orders(status="Order In Progress", limit=50)
        """
        endpoint = "/access-ordering/v1/orders"
        params = {'limit': limit, 'offset': offset}
        
        if broadband_product:
            params['broadbandProduct'] = broadband_product
        if status:
            params['status'] = status
        
        params.update(kwargs)
        
        return self._make_request('GET', endpoint, params=params)
    
    def get_regrade(self, order_id: int) -> Dict[str, Any]:
        """
        View a broadband regrade
        
        Args:
            order_id: The broadband order ID
            
        Returns:
            Dictionary containing regrade details and updates
        """
        endpoint = f"/access-ordering/v1/orders/{order_id}/regrade"
        return self._make_request('GET', endpoint)
    
    def place_order(self, order_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Place a new broadband order
        
        Args:
            order_data: Dictionary containing all order details
            
        Returns:
            Dictionary containing the order response with order ID
            
        Example:
            >>> order_data = {
            ...     "accountNumber": 132,
            ...     "broadbandProduct": "SoGEA 80:20 (1 month term)",
            ...     "careLevel": "Standard Care",
            ...     "installation": {...},
            ...     ...
            ... }
            >>> response = api.place_order(order_data)
            >>> print(f"Order ID: {response['id']}")
        """
        endpoint = "/access-ordering/v1/orders"
        return self._make_request('POST', endpoint, json=order_data)
    
    def place_regrade(self, order_id: int, regrade_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Place a regrade against an existing order
        
        Args:
            order_id: The existing broadband order ID to regrade
            regrade_data: Dictionary containing regrade details
            
        Returns:
            Dictionary containing the regrade response
            
        Example:
            >>> regrade_data = {
            ...     "broadbandProduct": "SoGEA 160:30 (1 month term)",
            ...     "customerRequiredDate": "2026-04-01",
            ...     "routerRequired": False
            ... }
            >>> response = api.place_regrade(10000, regrade_data)
            >>> print(f"Regrade Status: {response['status']}")
        """
        endpoint = f"/access-ordering/v1/orders/{order_id}/regrade"
        return self._make_request('POST', endpoint, json=regrade_data)
    
    # ==================== Helper Methods ====================
    
    def format_order_summary(self, order_data: Dict[str, Any]) -> str:
        """
        Create a human-readable summary of an order
        
        Args:
            order_data: Order data from get_order()
            
        Returns:
            Formatted string summary
        """
        lines = []
        lines.append("=" * 70)
        lines.append("GAMMA BROADBAND ORDER DETAILS")
        lines.append("=" * 70)
        
        # Basic order info
        lines.append(f"\nOrder ID: {order_data.get('id', 'N/A')}")
        lines.append(f"Status: {order_data.get('status', 'N/A')}")
        lines.append(f"Product: {order_data.get('broadbandProduct', 'N/A')}")
        lines.append(f"Account Number: {order_data.get('accountNumber', 'N/A')}")
        
        # Installation details
        if 'installation' in order_data:
            install = order_data['installation']
            lines.append(f"\nInstallation:")
            lines.append(f"  Line Type: {install.get('lineType', 'N/A')}")
            lines.append(f"  CLI: {install.get('cli', 'N/A')}")
            
            if 'site' in install:
                site = install['site']
                lines.append(f"  Company: {site.get('companyName', 'N/A')}")
                if 'address' in site:
                    addr = site['address']
                    address_str = f"{addr.get('building', '')}, {addr.get('street', '')}, {addr.get('town', '')}, {addr.get('postcode', '')}"
                    lines.append(f"  Address: {address_str}")
        
        # Dates
        if 'customerRequiredDate' in order_data.get('installation', {}):
            lines.append(f"\nCustomer Required Date: {order_data['installation']['customerRequiredDate']}")
        
        # Updates/History
        if 'updates' in order_data and order_data['updates']:
            lines.append(f"\n" + "=" * 70)
            lines.append("ORDER UPDATES")
            lines.append("=" * 70)
            
            for update in order_data['updates'][-10:]:  # Show last 10 updates
                timestamp = update.get('timestamp', 'No timestamp')
                status = update.get('status', 'No status')
                message = update.get('message', '')
                
                lines.append(f"\n[{timestamp}] {status}")
                if message:
                    lines.append(f"  {message}")
        
        lines.append("\n" + "=" * 70)
        
        return "\n".join(lines)
    
    def export_to_json(self, data: Dict[str, Any], filename: str) -> None:
        """Export order data to JSON file"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"Data exported to {filename}")


if __name__ == "__main__":
    print("Gamma Broadband Ordering API Client loaded.")
    print("Initialize with: api = GammaBroadbandOrderingAPI(username='user', password='pass')")

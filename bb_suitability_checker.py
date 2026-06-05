"""
Gamma Access Suitability Checker API Client
Version: 100.1.0

This module provides a Python client for interacting with the Gamma Broadband 
Access Suitability Checker API. It supports address lookups, portability checks,
and product suitability checks using various identifiers (CLI, postcode, NAD, ALID).
"""

import requests
from typing import Optional, Dict, List, Any
from urllib.parse import quote
import json
from datetime import datetime, timedelta


class GammaAccessAPIError(Exception):
    """Custom exception for Gamma Access API errors"""
    def __init__(self, message: str, status_code: int = None, response_data: Dict = None):
        self.message = message
        self.status_code = status_code
        self.response_data = response_data
        super().__init__(self.message)


class GammaAuthenticationError(GammaAccessAPIError):
    """Exception for authentication failures"""
    pass


class GammaAccessSuitabilityChecker:
    """
    Client for Gamma Access Suitability Checker API
    
    This client provides methods to:
    - Look up addresses by postcode
    - Check broadband product suitability
    - Check CLI portability
    - Query using CLI, postcode, NAD/DistrictID, ALID, or Order ID
    - Automatic JWT token authentication and refresh
    """
    
    PRODUCTION_URL = "https://ws.gammaoperations.com"
    TEST_URL = "https://ws-test.gammaoperations.com"
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
            use_production: If True, use production environment. Default is test/QA
            auto_refresh: If True, automatically refresh token when it expires
            
        Note:
            Either provide username+password OR bearer_token.
            If username+password provided, token will be obtained automatically.
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
            'Accept': 'application/json'
        })
        
        # If username/password provided, obtain token automatically
        if username and password and not bearer_token:
            self._obtain_token()
        elif bearer_token:
            self._set_token(bearer_token)
    
    def _obtain_token(self) -> Dict[str, Any]:
        """
        Obtain JWT token using username and password
        
        Returns:
            Token response containing access_token, refresh_token, and expiry info
            
        Raises:
            GammaAuthenticationError: If authentication fails
        """
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
            elif response.status_code == 400:
                raise GammaAuthenticationError(
                    "Authentication failed: Invalid or missing parameters",
                    status_code=400
                )
            else:
                raise GammaAuthenticationError(
                    f"Authentication failed with status {response.status_code}",
                    status_code=response.status_code
                )
        except requests.exceptions.RequestException as e:
            raise GammaAuthenticationError(f"Authentication request failed: {str(e)}")
    
    def _refresh_access_token(self) -> Dict[str, Any]:
        """
        Refresh the access token using the refresh token
        
        Returns:
            New token response
            
        Raises:
            GammaAuthenticationError: If refresh fails
        """
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
                # If refresh fails, try to re-authenticate with username/password
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
        """
        Process token response and update client state
        
        Args:
            token_data: JSON response from token endpoint
        """
        self.access_token = token_data.get('access_token')
        self.refresh_token = token_data.get('refresh_token')
        
        # Calculate expiry times
        expires_in = token_data.get('expires_in', 600)  # Default 10 minutes
        refresh_expires_in = token_data.get('refresh_expires_in', 2592000)  # Default 30 days
        
        self.token_expiry = datetime.now() + timedelta(seconds=expires_in)
        self.refresh_expiry = datetime.now() + timedelta(seconds=refresh_expires_in)
        
        # Update session header
        self._set_token(self.access_token)
    
    def _set_token(self, token: str) -> None:
        """
        Set the authorization token in session headers
        
        Args:
            token: JWT access token
        """
        self.session.headers.update({
            'Authorization': f'Bearer {token}'
        })
    
    def _ensure_valid_token(self) -> None:
        """
        Ensure we have a valid token, refreshing if necessary
        
        Raises:
            GammaAuthenticationError: If token cannot be refreshed
        """
        if not self.auto_refresh:
            return
        
        # Check if token is expired or about to expire (within 30 seconds)
        if self.token_expiry:
            time_until_expiry = (self.token_expiry - datetime.now()).total_seconds()
            if time_until_expiry < 30:
                # Token expired or about to expire, refresh it
                self._refresh_access_token()
    
    def get_token_info(self) -> Dict[str, Any]:
        """
        Get information about the current token status
        
        Returns:
            Dictionary with token expiry information
        """
        info = {
            'has_access_token': bool(self.access_token),
            'has_refresh_token': bool(self.refresh_token),
            'access_token_expiry': self.token_expiry.isoformat() if self.token_expiry else None,
            'refresh_token_expiry': self.refresh_expiry.isoformat() if self.refresh_expiry else None,
            'auto_refresh_enabled': self.auto_refresh
        }
        
        if self.token_expiry:
            seconds_until_expiry = (self.token_expiry - datetime.now()).total_seconds()
            info['seconds_until_expiry'] = max(0, int(seconds_until_expiry))
            info['token_expired'] = seconds_until_expiry < 0
        
        return info
    
    def _make_request(self, method: str, endpoint: str, **kwargs) -> Dict[str, Any]:
        """
        Make an HTTP request to the API
        
        Args:
            method: HTTP method (GET, POST, etc.)
            endpoint: API endpoint path
            **kwargs: Additional arguments to pass to requests
            
        Returns:
            JSON response as dictionary
            
        Raises:
            GammaAccessAPIError: If the request fails
        """
        # Ensure token is valid before making request
        self._ensure_valid_token()
        
        url = f"{self.base_url}{endpoint}"
        
        try:
            response = self.session.request(method, url, **kwargs)
            
            # Handle different response codes
            if response.status_code == 200:
                return response.json() if response.content else {}
            elif response.status_code == 404:
                raise GammaAccessAPIError(
                    "No data found for the provided criteria",
                    status_code=404,
                    response_data=None
                )
            elif response.status_code in [400, 500]:
                error_data = response.json() if response.content else {}
                error_msg = error_data.get('message', f'HTTP {response.status_code} error')
                raise GammaAccessAPIError(
                    error_msg,
                    status_code=response.status_code,
                    response_data=error_data
                )
            else:
                raise GammaAccessAPIError(
                    f"Unexpected status code: {response.status_code}",
                    status_code=response.status_code,
                    response_data=None
                )
                
        except requests.exceptions.RequestException as e:
            raise GammaAccessAPIError(f"Request failed: {str(e)}")
    
    # ==================== Address Lookup ====================
    
    def lookup_address(self, postcode: str) -> List[Dict[str, Any]]:
        """
        Look up addresses for a given postcode
        
        Args:
            postcode: UK postcode (e.g., "RG14 5BY")
            
        Returns:
            List of address dictionaries containing NAD, districtId, and full address details
            
        Example:
            >>> checker = GammaAccessSuitabilityChecker(token)
            >>> addresses = checker.lookup_address("RG14 5BY")
            >>> for addr in addresses:
            ...     print(f"{addr['addressString']} - NAD: {addr['nad']}")
        """
        encoded_postcode = quote(postcode)
        endpoint = f"/access/v1/address/{encoded_postcode}"
        return self._make_request('GET', endpoint)
    
    # ==================== Portability Check ====================
    
    def check_portability(self, cli: str) -> Dict[str, Any]:
        """
        Check if a CLI is portable
        
        Args:
            cli: UK geographic telephone number in E.164 format (e.g., "+441244335382")
            
        Returns:
            Dictionary with 'cli' and 'portable' fields
            portable values: "TRUE", "FALSE", or "ERROR"
            
        Example:
            >>> result = checker.check_portability("+441244335382")
            >>> if result['portable'] == "TRUE":
            ...     print(f"{result['cli']} is portable")
        """
        encoded_cli = quote(cli)
        endpoint = f"/access/v1/bb/portability/{encoded_cli}"
        return self._make_request('GET', endpoint)
    
    # ==================== Suitability Checks (V1) ====================
    
    def check_suitability_by_address_key(
        self, 
        nad: str, 
        district_id: str
    ) -> Dict[str, Any]:
        """
        Check broadband product suitability by NAD and District ID
        
        Args:
            nad: 12-character NAD key (e.g., "A15101241755")
            district_id: District code (e.g., "TH")
            
        Returns:
            Suitability outcome with products, speeds, and availability
            
        Note:
            FTTP availability is returned from address key checks
        """
        endpoint = f"/access/v1/bb/suitability/addressKey/{nad}/{district_id}"
        return self._make_request('GET', endpoint)
    
    def check_suitability_by_postcode(self, postcode: str) -> Dict[str, Any]:
        """
        Check broadband product suitability by postcode
        
        Args:
            postcode: UK postcode (e.g., "RG14 5BY")
            
        Returns:
            Suitability outcome with products, speeds, and availability
            
        Note:
            FTTC, FTTP and AnnexM availability is NOT returned from postcode checks
        """
        encoded_postcode = quote(postcode)
        endpoint = f"/access/v1/bb/suitability/postcode/{encoded_postcode}"
        return self._make_request('GET', endpoint)
    
    def check_suitability_by_alid(self, alid: str) -> Dict[str, Any]:
        """
        Check broadband product suitability by ALID
        
        Args:
            alid: Access Line Identifier (e.g., "ABCD0123456", "FBTH0026522")
            
        Returns:
            Suitability outcome with products, speeds, and availability
        """
        endpoint = f"/access/v1/bb/suitability/alid/{alid}"
        return self._make_request('GET', endpoint)
    
    def check_suitability_by_cli(self, cli: str) -> Dict[str, Any]:
        """
        Check broadband product suitability by CLI
        
        Args:
            cli: UK geographic telephone number in E.164 format (e.g., "+441244335382")
            
        Returns:
            Suitability outcome with products, speeds, and availability
        """
        encoded_cli = quote(cli)
        endpoint = f"/access/v1/bb/suitability/cli/{encoded_cli}"
        return self._make_request('GET', endpoint)
    
    # ==================== Suitability Checks (V2) ====================
    
    def check_suitability_v2(
        self,
        cli: Optional[str] = None,
        postcode: Optional[str] = None,
        nad: Optional[str] = None,
        district_id: Optional[str] = None,
        alid: Optional[str] = None,
        check_porting: bool = False
    ) -> Dict[str, Any]:
        """
        Check broadband product suitability using V2 endpoint (flexible parameters)
        
        Args:
            cli: UK geographic telephone number in E.164 format
            postcode: UK postcode
            nad: 12-character NAD key
            district_id: District code (required if nad is provided)
            alid: Access Line Identifier
            check_porting: If True, also check CLI portability
            
        Returns:
            Suitability outcome with products, speeds, and availability.
            If check_porting=True and CLI provided, includes portability info
            
        Note:
            At least one identifier (cli, postcode, nad+district_id, or alid) must be provided
        """
        endpoint = "/access/v2/bb/suitability"
        params = {}
        
        if cli:
            params['cli'] = cli
        if postcode:
            params['postcode'] = postcode
        if nad:
            params['nad'] = nad
        if district_id:
            params['districtId'] = district_id
        if alid:
            params['alid'] = alid
        if check_porting:
            params['checkPorting'] = 'true'
        
        if not params:
            raise ValueError("At least one identifier must be provided")
        
        return self._make_request('GET', endpoint, params=params)
    
    def check_suitability_by_order(
        self, 
        order_id: int, 
        check_porting: bool = False
    ) -> Dict[str, Any]:
        """
        Check broadband product suitability for regrades by Order ID
        
        Args:
            order_id: Existing order ID (must be active)
            check_porting: If True, also check CLI portability
            
        Returns:
            Suitability outcome for regrade products
        """
        endpoint = f"/access/v2/bb/suitability/order/{order_id}"
        params = {}
        
        if check_porting:
            params['checkPorting'] = 'true'
        
        return self._make_request('GET', endpoint, params=params)
    
    # ==================== Helper Methods ====================
    
    def get_available_products(self, suitability_result: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Extract only available products from a suitability check result
        
        Args:
            suitability_result: Result from any suitability check method
            
        Returns:
            List of available products
        """
        products = suitability_result.get('products', [])
        return [p for p in products if p.get('available', False)]
    
    def format_suitability_summary(self, suitability_result: Dict[str, Any]) -> str:
        """
        Create a human-readable summary of a suitability check
        
        Args:
            suitability_result: Result from any suitability check method
            
        Returns:
            Formatted string summary
        """
        lines = []
        lines.append("=" * 60)
        lines.append("GAMMA BROADBAND SUITABILITY CHECK")
        lines.append("=" * 60)
        
        # Basic info
        if 'cli' in suitability_result:
            lines.append(f"CLI: {suitability_result['cli']}")
        if 'postcode' in suitability_result:
            lines.append(f"Postcode: {suitability_result['postcode']}")
        if 'addressKey' in suitability_result:
            lines.append(f"Address Key: {suitability_result['addressKey']}")
        if 'alid' in suitability_result:
            lines.append(f"ALID: {suitability_result['alid']}")
        
        # Status
        rag = suitability_result.get('rag', 'Unknown')
        line_status = suitability_result.get('lineStatus', 'No status available')
        lines.append(f"\nRAG Status: {rag}")
        lines.append(f"Line Status: {line_status}")
        
        if 'exchange' in suitability_result:
            lines.append(f"Exchange: {suitability_result['exchange']}")
        
        # Portability
        if 'portabilityOutcome' in suitability_result:
            portable = suitability_result['portabilityOutcome'].get('portable', 'Unknown')
            lines.append(f"\nPortability: {portable}")
        
        # Products
        lines.append("\n" + "=" * 60)
        lines.append("AVAILABLE PRODUCTS")
        lines.append("=" * 60)
        
        available_products = self.get_available_products(suitability_result)
        
        if not available_products:
            lines.append("No products available")
        else:
            for i, product in enumerate(available_products, 1):
                lines.append(f"\n{i}. {product.get('name', 'Unknown Product')}")
                lines.append(f"   Family: {product.get('family', 'N/A')}")
                lines.append(f"   Technology: {product.get('technologyType', 'N/A')}")
                
                if 'estimatedDownstream' in product:
                    lines.append(f"   Est. Download: {product['estimatedDownstream']} Mbps")
                if 'estimatedUpstream' in product:
                    lines.append(f"   Est. Upload: {product['estimatedUpstream']} Mbps")
                
                if 'estimatedMinDownstream' in product and 'estimatedMaxDownstream' in product:
                    lines.append(f"   Download Range: {product['estimatedMinDownstream']}-{product['estimatedMaxDownstream']} Mbps")
                
                if 'concurrentCalls' in product:
                    lines.append(f"   Concurrent Calls: {product['concurrentCalls']}")
                
                if 'usagePolicy' in product:
                    lines.append(f"   Usage Policy: {product['usagePolicy']}")
        
        lines.append("\n" + "=" * 60)
        
        return "\n".join(lines)
    
    def export_to_json(self, data: Dict[str, Any], filename: str) -> None:
        """
        Export API response data to JSON file
        
        Args:
            data: Dictionary data to export
            filename: Output filename
        """
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"Data exported to {filename}")


# ==================== Example Usage ====================

def example_usage():
    """
    Example usage of the Gamma Access Suitability Checker client
    """
    # Option 1: Initialize with username/password (recommended - auto token management)
    USERNAME = "your_username"
    PASSWORD = "your_password"
    checker = GammaAccessSuitabilityChecker(
        username=USERNAME, 
        password=PASSWORD, 
        use_production=False
    )
    
    # Option 2: Initialize with pre-obtained bearer token
    # TOKEN = "your_bearer_token_here"
    # checker = GammaAccessSuitabilityChecker(bearer_token=TOKEN, use_production=False)
    
    try:
        # Example 1: Look up addresses by postcode
        print("Looking up addresses for postcode RG14 5BY...")
        addresses = checker.lookup_address("RG14 5BY")
        for addr in addresses:
            print(f"  {addr.get('addressString')} - NAD: {addr.get('nad')}")
        
        # Example 2: Check suitability by CLI
        print("\nChecking suitability for CLI +441244335382...")
        cli_result = checker.check_suitability_by_cli("+441244335382")
        print(checker.format_suitability_summary(cli_result))
        
        # Example 3: Check portability
        print("\nChecking portability for CLI +441244335382...")
        portability = checker.check_portability("+441244335382")
        print(f"  Portable: {portability.get('portable')}")
        
        # Example 4: Check suitability by address key (from address lookup)
        if addresses:
            first_addr = addresses[0]
            nad = first_addr.get('nad')
            district_id = first_addr.get('districtId')
            
            print(f"\nChecking suitability by address key {nad}/{district_id}...")
            addr_result = checker.check_suitability_by_address_key(nad, district_id)
            
            # Get only available products
            available = checker.get_available_products(addr_result)
            print(f"  Found {len(available)} available products")
        
        # Example 5: V2 API with multiple parameters
        print("\nUsing V2 API with CLI and portability check...")
        v2_result = checker.check_suitability_v2(
            cli="+441244335382",
            check_porting=True
        )
        print(f"  RAG Status: {v2_result.get('rag')}")
        
        # Example 6: Check token status
        token_info = checker.get_token_info()
        print(f"\nToken expires in: {token_info.get('seconds_until_expiry')} seconds")
        
        # Example 7: Export results to JSON
        checker.export_to_json(cli_result, "suitability_result.json")
        
    except GammaAuthenticationError as e:
        print(f"Authentication Error: {e.message}")
        if e.status_code:
            print(f"Status Code: {e.status_code}")
    except GammaAccessAPIError as e:
        print(f"API Error: {e.message}")
        if e.status_code:
            print(f"Status Code: {e.status_code}")
        if e.response_data:
            print(f"Details: {e.response_data}")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")


if __name__ == "__main__":
    # Uncomment to run examples
    # example_usage()
    print("Gamma Access Suitability Checker API Client loaded.")
    print("Initialize with username/password: checker = GammaAccessSuitabilityChecker(username='user', password='pass')")
    print("Or with bearer token: checker = GammaAccessSuitabilityChecker(bearer_token='token')")

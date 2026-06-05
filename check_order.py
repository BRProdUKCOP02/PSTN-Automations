"""
Check Broadband Order Status and Updates
This script retrieves and displays the status and updates for a Gamma broadband order
"""

from config import get_credentials, ENVIRONMENT
from bb_ordering_api import GammaBroadbandOrderingAPI, GammaOrderingAPIError, GammaAuthenticationError
import sys

# =============================================================================
# CONFIGURATION
# =============================================================================

# Prompt user for order ID
print("=" * 70)
print("GAMMA BROADBAND ORDER CHECKER")
print("=" * 70)

order_input = input("\nEnter the Order ID to check: ").strip()

try:
    ORDER_ID = int(order_input)
except ValueError:
    print(f"\n❌ Error: '{order_input}' is not a valid order ID")
    print("Order ID must be a number (e.g., 10000)")
    sys.exit(1)

# =============================================================================
# AUTHENTICATION
# =============================================================================

creds = get_credentials()
print(f"\nEnvironment: {creds['environment_name']}")
print(f"Username: {creds['username']}")
print(f"\nInitializing API client...")

try:
    api = GammaBroadbandOrderingAPI(
        username=creds['username'],
        password=creds['password'],
        use_production=creds['use_production'],
        auto_refresh=True
    )
    
    print("✓ Authentication successful!")
    
except GammaAuthenticationError as e:
    print(f"\n❌ AUTHENTICATION FAILED")
    print(f"Error: {e.message}")
    print(f"\nPlease check your credentials in config.py")
    print(f"Current environment: {ENVIRONMENT}")
    sys.exit(1)

# =============================================================================
# RETRIEVE ORDER DETAILS
# =============================================================================
print("\n" + "=" * 70)
print(f"RETRIEVING ORDER: {ORDER_ID}")
print("=" * 70)

try:
    print(f"\nFetching order details for Order ID {ORDER_ID}...")
    order = api.get_order(ORDER_ID)
    
    # Display formatted summary
    print("\n" + api.format_order_summary(order))
    
    # Option to export to JSON
    export = input("\nExport full order details to JSON? (y/n): ").strip().lower()
    if export == 'y':
        filename = f"order_{ORDER_ID}_{creds['environment_name'].lower()}.json"
        api.export_to_json(order, filename)
    
    print("\n" + "=" * 70)
    print("✓ ORDER CHECK COMPLETE")
    print("=" * 70)
    
except GammaOrderingAPIError as e:
    print(f"\n❌ API ERROR")
    print(f"Error: {e.message}")
    print(f"Status Code: {e.status_code}")
    
    if e.status_code == 404:
        print(f"\nOrder {ORDER_ID} not found or not accessible.")
        print("Possible reasons:")
        print("  - Order ID doesn't exist")
        print("  - Your API user is not associated with the order's billing account")
        print(f"  - Wrong environment (currently using: {creds['environment_name']})")
    elif e.status_code == 401:
        print("\nAuthentication issue - token may have expired")
    
    sys.exit(1)

except Exception as e:
    print(f"\n❌ UNEXPECTED ERROR: {e}")
    sys.exit(1)

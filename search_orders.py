"""
Search Broadband Orders
This script searches for broadband orders with various filters
"""

from config import get_credentials, ENVIRONMENT
from bb_ordering_api import GammaBroadbandOrderingAPI, GammaOrderingAPIError, GammaAuthenticationError

# =============================================================================
# SEARCH CONFIGURATION
# =============================================================================

# Configure your search filters here
SEARCH_FILTERS = {
    # 'broadband_product': 'SoGEA 80:20 (1 month term)',
    # 'status': 'Active',
    'limit': 50,  # Number of results to return
    'offset': 0   # Offset for pagination
}

# =============================================================================
# AUTHENTICATION
# =============================================================================
print("=" * 70)
print("GAMMA BROADBAND ORDER SEARCH")
print("=" * 70)

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
    exit(1)

# =============================================================================
# SEARCH ORDERS
# =============================================================================
print("\n" + "=" * 70)
print("SEARCHING ORDERS")
print("=" * 70)

# Display active filters
print("\nActive Filters:")
for key, value in SEARCH_FILTERS.items():
    if value and key not in ['limit', 'offset']:
        print(f"  {key}: {value}")

if not any(v for k, v in SEARCH_FILTERS.items() if k not in ['limit', 'offset']):
    print("  No filters - showing all orders")

print(f"\nLimit: {SEARCH_FILTERS['limit']}, Offset: {SEARCH_FILTERS['offset']}")

try:
    print("\nSearching...")
    orders = api.search_orders(**SEARCH_FILTERS)
    
    if not orders:
        print("\n⚠ No orders found matching the search criteria")
    else:
        print(f"\n✓ Found {len(orders)} orders:\n")
        print("-" * 70)
        
        for i, order in enumerate(orders, 1):
            print(f"\n{i}. Order ID: {order.get('id', 'N/A')}")
            print(f"   Status: {order.get('status', 'N/A')}")
            print(f"   Product: {order.get('broadbandProduct', 'N/A')}")
            print(f"   Account: {order.get('accountNumber', 'N/A')}")
            
            if 'installation' in order and 'cli' in order['installation']:
                print(f"   CLI: {order['installation']['cli']}")
            
            if 'installation' in order and 'site' in order['installation']:
                site = order['installation']['site']
                if 'companyName' in site:
                    print(f"   Company: {site['companyName']}")
        
        print("\n" + "-" * 70)
        print(f"\nShowing results {SEARCH_FILTERS['offset'] + 1} to {SEARCH_FILTERS['offset'] + len(orders)}")
        
        # Show pagination info
        if len(orders) == SEARCH_FILTERS['limit']:
            print(f"\nThere may be more results. To see next page:")
            print(f"  Update SEARCH_FILTERS['offset'] to {SEARCH_FILTERS['offset'] + SEARCH_FILTERS['limit']}")
    
    print("\n" + "=" * 70)
    print("✓ SEARCH COMPLETE")
    print("=" * 70)
    
    # Option to view a specific order
    if orders:
        view_order = input("\nEnter an Order ID to view details (or press Enter to skip): ").strip()
        if view_order:
            try:
                order_id = int(view_order)
                print(f"\nFetching details for Order {order_id}...")
                order_details = api.get_order(order_id)
                print("\n" + api.format_order_summary(order_details))
            except ValueError:
                print(f"Invalid Order ID: {view_order}")
            except GammaOrderingAPIError as e:
                print(f"\nError fetching order: {e.message}")
    
except GammaOrderingAPIError as e:
    print(f"\n❌ API ERROR")
    print(f"Error: {e.message}")
    print(f"Status Code: {e.status_code}")

except Exception as e:
    print(f"\n❌ UNEXPECTED ERROR: {e}")

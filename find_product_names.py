"""
Find Product Names from Existing Orders
Search your orders to find actual product names being used
"""

from config import get_credentials
from bb_ordering_api import GammaBroadbandOrderingAPI

# Initialize API
creds = get_credentials()
print("=" * 70)
print("FIND BROADBAND PRODUCT NAMES FROM ORDERS")
print("=" * 70)
print(f"\nEnvironment: {creds['environment_name']}")

api = GammaBroadbandOrderingAPI(
    username=creds['username'],
    password=creds['password'],
    use_production=creds['use_production'],
    auto_refresh=True
)

print("\nSearching recent orders...")
orders = api.search_orders(limit=50)

# Extract unique product names
products = {}
for order in orders:
    product = order.get('broadbandProduct')
    if product:
        if product not in products:
            products[product] = []
        products[product].append({
            'id': order.get('id'),
            'status': order.get('status')
        })

print(f"\n✓ Found {len(products)} unique product names:\n")
print("=" * 70)

for product in sorted(products.keys()):
    count = len(products[product])
    example_id = products[product][0]['id']
    print(f"\n• {product}")
    print(f"  Used in {count} order(s) - Example: Order {example_id}")

print("\n" + "=" * 70)
print("✓ SEARCH COMPLETE")
print("=" * 70)
print("\nUse these exact product names in your regrade template.")
print("Match the technology type (SoGEA/FTTC/FTTP) with your existing order.")

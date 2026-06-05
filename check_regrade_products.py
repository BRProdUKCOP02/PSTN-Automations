"""
Check Available Regrade Products
Query the Gamma API to see what products are available for regrading a specific order
"""

import sys
from config import get_credentials
from bb_suitability_checker import GammaAccessSuitabilityChecker

# Get order ID from user
print("=" * 70)
print("CHECK AVAILABLE REGRADE PRODUCTS")
print("=" * 70)

order_input = input("\nEnter the Order ID to check regrade options for: ").strip()

try:
    order_id = int(order_input)
except ValueError:
    print(f"\n❌ Error: '{order_input}' is not a valid order ID")
    sys.exit(1)

# Initialize API
creds = get_credentials()
print(f"\nEnvironment: {creds['environment_name']}")
print(f"Username: {creds['username']}")

try:
    checker = GammaAccessSuitabilityChecker(
        username=creds['username'],
        password=creds['password'],
        use_production=creds['use_production']
    )
    
    print(f"\nChecking regrade options for Order {order_id}...")
    print("=" * 70)
    
    # Check suitability for regrade
    result = checker.check_suitability_by_order(order_id)
    
    # Display summary
    print(f"\nRAG Status: {result.get('rag')}")
    print(f"Exchange: {result.get('exchange')}")
    print(f"Technology: {result.get('technology', 'N/A')}")
    
    # Get available products
    products = checker.get_available_products(result)
    
    if products:
        print(f"\n✓ Available Products for Regrade ({len(products)} found):")
        print("=" * 70)
        
        for i, product in enumerate(products, 1):
            print(f"\n{i}. {product['product']}")
            print(f"   Download: {product.get('downloadSpeed', 'N/A')}")
            print(f"   Upload: {product.get('uploadSpeed', 'N/A')}")
            if product.get('availability'):
                print(f"   Availability: {product['availability']}")
        
        print("\n" + "=" * 70)
        print("✓ REGRADE CHECK COMPLETE")
        print("=" * 70)
        print("\nCopy one of the product names above and use it in your")
        print("regrade template for the 'broadbandProduct' field.")
        
    else:
        print("\n⚠ No regrade products found")
        print("\nPossible reasons:")
        print("  - Order is not in a state that allows regrades")
        print("  - No alternative products available at this location")
        print("  - API limitation in test environment")

except Exception as e:
    print(f"\n❌ Error: {str(e)}")
    print("\nThis might be a limitation of the test environment.")
    print("The regrade suitability checker may only work in production.")
    sys.exit(1)

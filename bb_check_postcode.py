"""
Example 1: Address Lookup - Using Config File
This script uses the config.py file for credentials and environment selection
"""

from config import get_checker, get_credentials, ENVIRONMENT
from bb_suitability_checker import GammaAccessAPIError, GammaAuthenticationError

# =============================================================================
# INITIALIZATION
# =============================================================================
print("=" * 70)
print("GAMMA API - ADDRESS LOOKUP (USING CONFIG)")
print("=" * 70)

creds = get_credentials()
print(f"\nEnvironment: {creds['environment_name']}")
print(f"Username: {creds['username']}")
print(f"\nInitializing API client...")

try:
    checker = get_checker()
    print("✓ Authentication successful!")
    
    # Show token info
    token_info = checker.get_token_info()
    if token_info.get('seconds_until_expiry'):
        print(f"✓ Token expires in: {token_info['seconds_until_expiry']} seconds")
    
except GammaAuthenticationError as e:
    print(f"\n❌ AUTHENTICATION FAILED")
    print(f"Error: {e.message}")
    print(f"\nPlease check your credentials in config.py")
    print(f"Current environment: {ENVIRONMENT}")
    exit(1)

# =============================================================================
# EXAMPLE 1: ADDRESS LOOKUP BY POSTCODE
# =============================================================================
print("\n" + "=" * 70)
print("EXAMPLE 1: Address Lookup by Postcode")
print("=" * 70)

POSTCODE = "M17 1FG"  # Change this to test different postcodes

try:
    print(f"\nLooking up addresses for postcode: {POSTCODE}...")
    addresses = checker.lookup_address(POSTCODE)
    
    if not addresses:
        print(f"\n⚠ No addresses found for postcode '{POSTCODE}'")
    else:
        print(f"\n✓ Found {len(addresses)} addresses:\n")
        
        for i, addr in enumerate(addresses, 1):
            print(f"{i}. {addr.get('addressString')}")
            print(f"   Building Name: {addr.get('buildingName', 'N/A')}")
            print(f"   Building Number: {addr.get('buildingNumber', 'N/A')}")
            print(f"   Street: {addr.get('street', 'N/A')}")
            print(f"   Town: {addr.get('town', 'N/A')}")
            print(f"   County: {addr.get('county', 'N/A')}")
            print(f"   Postcode: {addr.get('postcode', 'N/A')}")
            print(f"   NAD: {addr.get('nad')}")
            print(f"   District ID: {addr.get('districtId')}")
            print(f"   Address Key: {addr.get('addressKey')}")
            print()
    
    print("=" * 70)
    print(f"✓ EXAMPLE 1 COMPLETED SUCCESSFULLY ({creds['environment_name']})")
    print("=" * 70)
    
except GammaAccessAPIError as e:
    print(f"\n❌ API ERROR")
    print(f"Error: {e.message}")
    print(f"Status Code: {e.status_code}")
    
    if e.status_code == 400:
        print(f"\nThe postcode '{POSTCODE}' may be invalid or incorrectly formatted")
        print("Valid UK postcode format: e.g., 'RG14 5BY', 'CB21 4LH'")
    elif e.status_code == 404:
        print(f"\nNo addresses found for postcode '{POSTCODE}'")
    elif e.status_code == 500:
        print("\nServer error - please try again later")

except Exception as e:
    print(f"\n❌ UNEXPECTED ERROR: {e}")

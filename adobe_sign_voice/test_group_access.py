"""Check which groups are accessible with current token."""
from adobe_sign_client import AdobeSignClient
import requests

client = AdobeSignClient()
token = client._get_access_token()
headers = {"Authorization": f"Bearer {token}"}

# Get user groups
url = "https://api.na1.adobesign.com/api/rest/v6/groups"
response = requests.get(url, headers=headers)

print("Groups accessible with current token:")
print("=" * 80)

if response.status_code == 200:
    groups = response.json().get("groupInfoList", [])
    for i, group in enumerate(groups, 1):
        print(f"\n{i}. {group.get('groupName')}")
        print(f"   ID: {group.get('groupId')}")
        print(f"   Created: {group.get('createdDate', 'N/A')[:10]}")
        
        # Check if this is the commercial group
        group_id = group.get('groupId')
        if group_id == "CBJCHBCAABAAtdaBCBxCah9w5QhDMT-NmGWGYVFmAGnZ":
            print(f"   ⭐ THIS IS THE COMMERCIAL GROUP (partner widget owner)")
else:
    print(f"Failed to get groups: {response.status_code}")
    print(response.text)

print("\n" + "=" * 80)
print("\nWidget groups:")
print(f"Partner widget group: CBJCHBCAABAAtdaBCBxCah9w5QhDMT-NmGWGYVFmAGnZ")
print(f"Your original group: CBJCHBCAABAAXaOhsbKiB9TvY7ZP9tm_YyR-FvMV6Yao")

"""Check widget ownership and permissions."""
from adobe_sign_client import AdobeSignClient
import requests

client = AdobeSignClient()
token = client._get_access_token()
headers = {"Authorization": f"Bearer {token}"}

partner_widget_id = "CBJCHBCAABAAHgM-E9MfQ4zVKGAAPqPXbkWE7lMnPp7T"

# Get widget details
url = f"https://api.na1.adobesign.com/api/rest/v6/widgets/{partner_widget_id}"
response = requests.get(url, headers=headers)

print("Partner Widget Details:")
print("=" * 80)

if response.status_code == 200:
    widget = response.json()
    print(f"Name: {widget.get('name')}")
    print(f"Status: {widget.get('status')}")
    print(f"Group ID: {widget.get('groupId')}")
    print(f"Created by: {widget.get('creatorEmail', widget.get('senderEmail'))}")
    print(f"Created date: {widget.get('createdDate', 'N/A')[:10]}")
    
    # Check if we're the owner
    creator_email = widget.get('creatorEmail', widget.get('senderEmail'))
    if creator_email and 'david.murphy' in creator_email.lower():
        print("\n✅ You ARE the widget owner")
    else:
        print(f"\n❌ You are NOT the owner (owned by {creator_email})")
        
    # Check security settings
    security_options = widget.get('securityOptions', {})
    if security_options:
        print(f"\nSecurity options: {security_options}")
        
    print(f"\nFull widget data:")
    import json
    print(json.dumps(widget, indent=2))
else:
    print(f"Failed to get widget: {response.status_code}")
    print(response.text)

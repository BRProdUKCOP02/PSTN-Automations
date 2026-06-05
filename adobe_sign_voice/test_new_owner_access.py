"""Test if widget ownership enables different API access patterns."""
from adobe_sign_client import AdobeSignClient
import requests

client = AdobeSignClient()
token = client._get_access_token()
headers = {"Authorization": f"Bearer {token}"}

# Check widget ownership
partner_widget_id = "CBJCHBCAABAAHgM-E9MfQ4zVKGAAPqPXbkWE7lMnPp7T"
url = f"https://api.na1.adobesign.com/api/rest/v6/widgets/{partner_widget_id}"
response = requests.get(url, headers=headers)

print("Widget Ownership Check:")
print("=" * 80)
if response.status_code == 200:
    widget = response.json()
    owner_email = widget.get('ownerEmail', widget.get('creatorEmail'))
    print(f"Widget: {widget.get('name')}")
    print(f"Owner: {owner_email}")
    
    if 'david.murphy' in owner_email.lower():
        print("✅ YOU ARE NOW THE OWNER!")
    else:
        print(f"❌ Still owned by {owner_email}")
        
print("\n" + "=" * 80)
print("Agreement Access Test:")
print("=" * 80)

# Try to get agreements
partner_agreements = client.list_widget_agreements(partner_widget_id)
print(f"Total agreements from partner widget: {len(partner_agreements)}")

if partner_agreements:
    test_agreement_id = partner_agreements[0].get('id')
    print(f"\nTesting agreement: {test_agreement_id}")
    
    # Try form data
    print("\n1. Trying /formData endpoint...")
    try:
        form_data = client.get_form_data(test_agreement_id)
        print(f"   ✅ SUCCESS - Retrieved {len(form_data)} rows of form data")
        if form_data:
            print(f"   Fields: {list(form_data[0].keys())}")
    except Exception as e:
        print(f"   ❌ FAILED - {e}")
        
        # Try alternative: Get documents and check if form fields are embedded
        print("\n2. Trying /documents endpoint...")
        try:
            docs = client.get_documents(test_agreement_id)
            print(f"   Documents found: {len(docs.get('documents', []))}")
            print(f"   Supporting docs: {len(docs.get('supportingDocuments', []))}")
        except Exception as e2:
            print(f"   ❌ FAILED - {e2}")

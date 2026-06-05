"""Test if form data access works for owned vs shared widgets."""
from adobe_sign_client import AdobeSignClient

client = AdobeSignClient()

# Test owned widget (RWP Opt-Out)
print("Testing RWP Opt-Out widget (owned directly)")
print("=" * 80)
rwp_widget_id = "CBJCHBCAABAAVqVxP0-R7wBkzV87GM2dKpTtQv0CrE-f"
rwp_agreements = client.list_widget_agreements(rwp_widget_id)
if rwp_agreements:
    test_agreement = rwp_agreements[0]
    print(f"Agreement: {test_agreement.get('id')}")
    print(f"Status: {test_agreement.get('status')}")
    try:
        form_data = client.get_form_data(test_agreement.get('id'))
        print(f"✅ SUCCESS - Form data retrieved: {len(form_data)} rows")
        if form_data:
            print(f"   Fields: {list(form_data[0].keys())}")
    except Exception as e:
        print(f"❌ FAILED - {e}")
else:
    print("No agreements found in RWP widget")

print()
print("=" * 80)
print("Testing Partner widget (shared)")
print("=" * 80)

# Test shared widget (Partner)
partner_widget_id = "CBJCHBCAABAAHgM-E9MfQ4zVKGAAPqPXbkWE7lMnPp7T"
partner_agreements = client.list_widget_agreements(partner_widget_id)
if partner_agreements:
    test_agreement = partner_agreements[0]
    print(f"Agreement: {test_agreement.get('id')}")
    print(f"Status: {test_agreement.get('status')}")
    try:
        form_data = client.get_form_data(test_agreement.get('id'))
        print(f"✅ SUCCESS - Form data retrieved: {len(form_data)} rows")
        if form_data:
            print(f"   Fields: {list(form_data[0].keys())}")
    except Exception as e:
        print(f"❌ FAILED - {e}")
else:
    print("No agreements found in partner widget")

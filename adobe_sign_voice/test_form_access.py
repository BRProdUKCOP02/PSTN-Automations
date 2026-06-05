"""Test different Adobe Sign endpoints for form data access."""
from adobe_sign_client import AdobeSignClient
import requests

client = AdobeSignClient()
agreement_id = "CBJCHBCAABAAr44d0LV32x67uO5RUv59hQQ6QFXQN-H9"
token = client._get_access_token()
headers = {"Authorization": f"Bearer {token}"}

endpoints = [
    "/formData",
    "/merge",
    "/views", 
    "/combinedDocument/fields",
    "/me/views/settings",
]

print(f"Testing agreement {agreement_id}\n")
print("=" * 80)

for ep in endpoints:
    url = f"https://api.na1.adobesign.com/api/rest/v6/agreements/{agreement_id}{ep}"
    try:
        r = requests.get(url, headers=headers, timeout=10)
        status_text = "✅ OK" if r.status_code == 200 else f"❌ {r.status_code}"
        print(f"{ep:<35} {status_text}")
        if r.status_code not in [200, 404, 403]:
            print(f"  Response: {r.text[:200]}")
        if r.status_code == 403:
            print(f"  Error: PERMISSION_DENIED")
    except Exception as e:
        print(f"{ep:<35} ❌ Exception: {e}")

print("=" * 80)

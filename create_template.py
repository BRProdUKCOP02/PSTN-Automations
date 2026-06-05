"""
Create Excel template for Phoneline+ bulk processing
Hardware Only accounts with optional users
"""
import pandas as pd

# Sample data for template - Hardware Only accounts with users
template_data = [
    {
        'keyID': '31ef1028-2075-40fd-9acc-d40021d0c931',
        'secret': 'TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7',
        'companyName': 'Example HWO Company 1',
        'fullName': 'John Smith',
        'email': 'john.smith@example.com',
        'number': '443300163504',
        'contactNumber': '07700900123',
        'premises': 'Unit 1',
        'street': 'High Street',
        'town': 'London',
        'county': 'Greater London',
        'postcode': 'SW1A 1AA',
        # User 1
        'user1_fullName': 'Admin User',
        'user1_email': 'admin@example.com',
        'user1_type': 'admin',
        'user1_phoneNumber': '07700900456',
        'user1_premises': 'Unit 1',
        'user1_street': 'High Street',
        'user1_town': 'London',
        'user1_county': 'Greater London',
        'user1_postcode': 'SW1A 1AA',
        # User 2
        'user2_fullName': 'Standard User',
        'user2_email': 'user@example.com',
        'user2_type': 'standard',
        'user2_phoneNumber': '',
        'user2_premises': 'Unit 1',
        'user2_street': 'High Street',
        'user2_town': 'London',
        'user2_county': 'Greater London',
        'user2_postcode': 'SW1A 1AA',
        # User 3 (empty)
        'user3_fullName': '',
        'user3_email': '',
        'user3_type': '',
        'user3_phoneNumber': '',
        'user3_premises': '',
        'user3_street': '',
        'user3_town': '',
        'user3_county': '',
        'user3_postcode': ''
    },
    {
        'keyID': '31ef1028-2075-40fd-9acc-d40021d0c931',
        'secret': 'TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7',
        'companyName': 'Example HWO Company 2',
        'fullName': 'Jane Doe',
        'email': 'jane.doe@example.com',
        'number': '441234567890',
        'contactNumber': '07700900456',
        'premises': 'Building 2',
        'street': 'Main Road',
        'town': 'Manchester',
        'county': 'Greater Manchester',
        'postcode': 'M1 1AA',
        # No users for this customer
        'user1_fullName': '',
        'user1_email': '',
        'user1_type': '',
        'user1_phoneNumber': '',
        'user1_premises': '',
        'user1_street': '',
        'user1_town': '',
        'user1_county': '',
        'user1_postcode': '',
        'user2_fullName': '',
        'user2_email': '',
        'user2_type': '',
        'user2_phoneNumber': '',
        'user2_premises': '',
        'user2_street': '',
        'user2_town': '',
        'user2_county': '',
        'user2_postcode': '',
        'user3_fullName': '',
        'user3_email': '',
        'user3_type': '',
        'user3_phoneNumber': '',
        'user3_premises': '',
        'user3_street': '',
        'user3_town': '',
        'user3_county': '',
        'user3_postcode': ''
    }
]

# Create DataFrame
df = pd.DataFrame(template_data)

# Save to Excel
output_file = 'c:/Users/Public/RPA/code/PSTN Migration/input/phoneline_plus_input_template.xlsx'
df.to_excel(output_file, index=False)

print(f"✓ Template created: {output_file}")
print(f"  Columns: {len(df.columns)}")
print(f"  Sample rows: {len(df)}")
print(f"\n  Account Type: Hardware Only")
print(f"    - Number (E164 format) is REQUIRED (must start with 441, 442, or 443)")
print(f"    - Email is OPTIONAL")
print(f"\n  Users: Up to 3 users per customer (optional)")
print(f"    - user1_fullName, user1_email are REQUIRED if creating user 1")
print(f"    - user1_type: 'standard' or 'admin' (defaults to 'standard')")
print(f"    - user1_phoneNumber is OPTIONAL")
print(f"    - user1_premises, user1_street, user1_town, user1_county, user1_postcode for user address")
print(f"    - Same pattern for user2_ and user3_")



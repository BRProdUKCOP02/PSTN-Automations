"""
Generate Excel template for Phoneline+ hardware orders
"""

import pandas as pd
from datetime import datetime

# Sample data for template
data = {
    'keyID': [
        '31ef1028-2075-40fd-9acc-d40021d0c931',
        '',
        ''
    ],
    'secret': [
        'TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7',
        '',
        ''
    ],
    'customer_id': [
        'customer-uuid-here',
        'customer-uuid-here',
        'another-customer-uuid'
    ],
    'name': [
        'John Smith',
        'Jane Doe',
        'Mike Johnson'
    ],
    'email': [
        'john.smith@example.com',
        'jane.doe@example.com',
        'mike.johnson@example.com'
    ],
    'phone_number': [
        '07700900123',
        '07700900456',
        '07700900789'
    ],
    'delivery_line1': [
        'Apartment 9',
        '42 High Street',
        'Unit 5, Business Park'
    ],
    'delivery_line2': [
        '6 Worsley Road',
        '',
        'Park Lane'
    ],
    'delivery_line3': [
        'Swinton',
        '',
        ''
    ],
    'delivery_town': [
        'Manchester',
        'London',
        'Birmingham'
    ],
    'delivery_county': [
        'Greater Manchester',
        '',
        'West Midlands'
    ],
    'delivery_country': [
        'United Kingdom',
        'United Kingdom',
        'United Kingdom'
    ],
    'delivery_postcode': [
        'M275WW',
        'SW1A1AA',
        'B12AB'
    ],
    'product_id': [
        'b5825399-a7ce-4806-9e86-116e05120d1d',
        'b5825399-a7ce-4806-9e86-116e05120d1d',
        'another-product-sku-uuid-here'
    ],
    'quantity': [
        1,
        2,
        1
    ],
    'user_id': [
        'd79f3004-b8d9-4091-ba3b-71ebac3e48d3',
        'user-uuid-here',
        ''
    ]
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = f'phoneline_plus_device_template_{timestamp}.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Hardware Orders', index=False)
    
    # Instructions sheet
    instructions = pd.DataFrame({
        'Field': [
            'keyID',
            'secret',
            'customer_id',
            'name',
            'email',
            'phone_number',
            'delivery_line1',
            'delivery_line2',
            'delivery_line3',
            'delivery_town',
            'delivery_county',
            'delivery_country',
            'delivery_postcode',
            'product_id',
            'quantity',
            'user_id'
        ],
        'Required': [
            'Yes (first row only)',
            'Yes (first row only)',
            'Yes',
            'Yes',
            'Yes',
            'Yes',
            'Yes',
            'No',
            'No',
            'Yes',
            'No',
            'No (defaults to UK)',
            'Yes',
            'Yes',
            'No (defaults to 1)',
            'No (for device assignment)'
        ],
        'Description': [
            'Partner API Key ID (same for all rows, only fill first row)',
            'Partner API Secret (same for all rows, only fill first row)',
            'Customer UUID from Phoneline+ system',
            'Delivery contact name',
            'Delivery contact email address',
            'Delivery contact phone number',
            'Delivery address line 1 (building/flat number, street)',
            'Delivery address line 2 (optional)',
            'Delivery address line 3 (optional)',
            'Town or city',
            'County (optional)',
            'Country (defaults to "United Kingdom" if blank)',
            'Postcode',
            'Product SKU UUID from Phoneline+ catalog',
            'Number of units to order (defaults to 1 if blank)',
            'User UUID to assign device to (optional - leave blank if not assigning)'
        ],
        'Example': [
            '31ef1028-2075-40fd-9acc-d40021d0c931',
            'TddBUXWfrOTj2Ie...',
            '123e4567-e89b-12d3-a456-426614174000',
            'John Smith',
            'john.smith@example.com',
            '07700900123',
            'Apartment 9, 6 Worsley Road',
            'Swinton',
            '',
            'Manchester',
            'Greater Manchester',
            'United Kingdom',
            'M275WW',
            'b5825399-a7ce-4806-9e86-116e05120d1d',
            '1',
            'd79f3004-b8d9-4091-ba3b-71ebac3e48d3'
        ]
    })
    instructions.to_excel(writer, sheet_name='Instructions', index=False)
    
    # Auto-size columns
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 60)
            worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"✓ Template created: {output_file}")
print("\nRequired columns:")
print("  - keyID: Partner API Key ID (first row only)")
print("  - secret: Partner API Secret (first row only)")
print("  - customer_id: Customer UUID")
print("  - name: Delivery contact name")
print("  - email: Delivery contact email")
print("  - phone_number: Delivery contact phone")
print("  - delivery_line1: Address line 1")
print("  - delivery_town: Town/City")
print("  - delivery_postcode: Postcode")
print("  - product_id: Product SKU UUID")
print("\nOptional columns:")
print("  - delivery_line2: Address line 2")
print("  - delivery_line3: Address line 3")
print("  - delivery_county: County")
print("  - delivery_country: Country (defaults to 'United Kingdom')")
print("  - quantity: Number of units (defaults to 1)")
print("  - user_id: User UUID for device assignment (leave blank if not assigning)")
print("\nNote: keyID and secret only need to be filled in the first row.")


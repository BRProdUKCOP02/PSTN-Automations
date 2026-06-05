"""
Phoneline+ Number Reallocation Template Generator
Creates an Excel template for bulk number reallocation processing
"""

import pandas as pd
from datetime import datetime


def generate_template():
    """
    Generate Excel template for number reallocation input
    """
    # Sample data
    data = {
        'keyID': [
            'your-key-id-here',
            ''
        ],
        'secret': [
            'your-secret-here',
            ''
        ],
        'customer_id': [
            'customer-uuid-here',
            'another-customer-uuid'
        ]
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'phoneline_plus_number_reallocation_template_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Number Reallocations', index=False)
        
        # Instructions sheet
        instructions = pd.DataFrame({
            'Field': [
                'keyID',
                'secret',
                'customer_id'
            ],
            'Required': [
                'Yes (Row 1 only)',
                'Yes (Row 1 only)',
                'Yes'
            ],
            'Description': [
                'Partner API Key ID - Only needed in first row, will be reused for all customers',
                'Partner API Secret - Only needed in first row, will be reused for all customers',
                'Customer UUID to process number reallocations for'
            ],
            'Example': [
                'ec2d3063-0079-475a-a6a3-26df72a2e8d7',
                'TddBUXWfrOTj2IeXCodVXshs_ZkBDVbNzpCktF1ml47cK7cfL6Di3IM8Vc9E07_7',
                '88032dd4-fed5-40e8-abda-66a1a746971c'
            ]
        })
        instructions.to_excel(writer, sheet_name='Instructions', index=False)
        
        # Process description sheet
        process_info = pd.DataFrame({
            'Step': [
                '1',
                '2',
                '3',
                '4',
                '5',
                '6'
            ],
            'Description': [
                'Script fetches all numbers for the customer',
                'Identifies geographical numbers (type: standard_geographic)',
                'Identifies non-geographical numbers (type: standard_nongeographic)',
                'For each user with a non-geographical number allocated:',
                '  - Allocates an available geographical number to that user',
                '  - Deallocates the non-geographical number from that user'
            ],
            'Notes': [
                'Uses GET /customers/{id}/numbers/',
                'These are preferred for users (e.g., Cambridge, Manchester numbers)',
                'These should be replaced (e.g., UK National numbers)',
                'Only processes users that currently have non-geo numbers',
                'Geographical numbers must be available (not already allocated)',
                'Original non-geo numbers become available for reassignment'
            ]
        })
        process_info.to_excel(writer, sheet_name='Process Information', index=False)
        
        # Auto-size columns for all sheets
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
                adjusted_width = min(max_length + 2, 80)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print("=" * 60)
    print("Phoneline+ Number Reallocation Template Generator")
    print("=" * 60)
    print(f"\n✓ Template generated: {output_file}")
    print("\nInstructions:")
    print("1. Fill in keyID and secret in the FIRST ROW only")
    print("2. Add customer UUIDs in the customer_id column")
    print("3. Save the file to the OneDrive input folder:")
    print("   Phoneline+ Number Reallocation/")
    print("4. Run: phoneline_plus_bulk_number_processor.py")
    print("\nWhat the script does:")
    print("• Fetches all numbers for each customer")
    print("• Finds users with non-geographical numbers")
    print("• Allocates geographical numbers to those users")
    print("• Deallocates the non-geographical numbers")
    print("• Generates a detailed Excel report")
    print("• Emails the report to psmanaged.delivery@gamma.co.uk")
    print("\nNote: Geographical numbers (e.g., Cambridge, Manchester)")
    print("      are preferred over non-geographical (e.g., UK National)")
    
    return output_file


if __name__ == "__main__":
    generate_template()

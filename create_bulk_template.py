"""
SoGEA Bulk Order Template Generator
Creates an Excel template for bulk ordering SoGEA broadband
"""

import pandas as pd
from datetime import datetime, timedelta

def create_bulk_order_template(filename="sogea_bulk_order_template.xlsx"):
    """
    Creates an Excel template with all required fields for SoGEA bulk ordering
    """
    
    # Define all columns
    columns = [
        # Order Type
        'orderType',
        'existingOrderId',
        
        # Account & Product
        'accountNumber',
        'broadbandProduct',
        'careLevel',
        
        # Installation Site Address
        'site_companyName',
        'site_nadKey',
        'site_building',
        'site_subPremises',
        'site_street',
        'site_town',
        'site_county',
        'site_postcode',
        
        # Site Contact
        'site_contact_name',
        'site_contact_email',
        'site_contact_phone',
        
        # Installation Details
        'lineType',
        'cli',
        'customerRequiredDate',
        'installation_type',
        'customerReference',
        'eccBand',
        'trcBand',
        
        # Reseller Contact
        'reseller_contact_name',
        'reseller_contact_email',
        'reseller_contact_phone',
        
        # IP & Voice Options
        'ipAddressOption',
        'routedIpOption',
        'voiceProduct',
        
        # Equipment (Optional)
        'routerRequired',
        'router',
        'router_delivery_building',
        'router_delivery_street',
        'router_delivery_town',
        'router_delivery_county',
        'router_delivery_postcode',
        'router_contact_name',
        'router_contact_email',
        'router_contact_phone',
        
        # Number Porting (Optional)
        'voipReference',
        
        # Notifications
        'resellerEmailNotifications'
    ]
    
    # Create example row with sample data
    example_date = (datetime.now() + timedelta(days=14)).strftime('%Y-%m-%d')
    
    example_data = {
        'orderType': 'NEW',
        'existingOrderId': '',
        
        'accountNumber': 132,
        'broadbandProduct': 'SoGEA 80:20 (1 month term)',
        'careLevel': 'Standard Care',
        
        'site_companyName': 'Example Company Ltd',
        'site_nadKey': 'A00026323803/TH',
        'site_building': 'Building 1',
        'site_subPremises': 'Flat 1',
        'site_street': 'High Street',
        'site_town': 'London',
        'site_county': 'Greater London',
        'site_postcode': 'SW1A 1AA',
        
        'site_contact_name': 'John Smith',
        'site_contact_email': 'john.smith@example.com',
        'site_contact_phone': '01234567890',
        
        'lineType': 'New',
        'cli': '',
        'customerRequiredDate': example_date,
        'installation_type': 'Managed Install',
        'customerReference': 'Example Customer Ref 001',
        'eccBand': 'Charge Band A (£0)',
        'trcBand': 'Charge Band 0 (not authorised beyond the NTE)',
        
        'reseller_contact_name': 'Your Name',
        'reseller_contact_email': 'your.email@company.com',
        'reseller_contact_phone': '01234567890',
        
        'ipAddressOption': 'Static Public IP Address',
        'routedIpOption': 'None',
        'voiceProduct': 'Gamma SIP Trunks',
        
        'routerRequired': 'TRUE',
        'router': 'Standard WiFi Router - Technicolor DGA0122/DGA4134',
        'router_delivery_building': 'Building 1',
        'router_delivery_street': 'High Street',
        'router_delivery_town': 'London',
        'router_delivery_county': 'Greater London',
        'router_delivery_postcode': 'SW1A 1AA',
        'router_contact_name': 'John Smith',
        'router_contact_email': 'john.smith@example.com',
        'router_contact_phone': '01234567890',
        
        'voipReference': '',
        
        'resellerEmailNotifications': 'TRUE'
    }
    
    # Create DataFrame with example row
    df = pd.DataFrame([example_data])
    
    # Create Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write main data sheet
        df.to_excel(writer, sheet_name='Orders', index=False)
        
        # Create instructions sheet
        instructions = pd.DataFrame({
            'INSTRUCTIONS': [
                'SoGEA BULK ORDER TEMPLATE',
                '',
                'HOW TO USE:',
                '1. Fill in one row per order in the "Orders" sheet',
                '2. Delete the example row before processing',
                '3. Save the file',
                '4. Run: python sogea_bulk_ordering.py',
                '',
                'ORDER TYPES:',
                '- orderType: "NEW" (or leave blank) for new orders, "REGRADE" for upgrades/downgrades',
                '- existingOrderId: Required if orderType is "REGRADE" (the Order ID to regrade)',
                '',
                'FOR NEW ORDERS:',
                '- orderType = "NEW" or leave blank',
                '- Fill in all standard fields below',
                '',
                'FOR REGRADE ORDERS:',
                '- orderType = "REGRADE"',
                '- existingOrderId = Order ID to regrade',
                '- Only fill in: broadbandProduct, customerRequiredDate, routerRequired, router details',
                '- Other fields inherited from existing order',
                '',
                'REQUIRED FIELDS (must not be empty):',
                '- accountNumber: Your Gamma billing account number',
                '- broadbandProduct: Exact product name from Suitability Checker',
                '- careLevel: e.g., "Standard Care"',
                '- site_nadKey: NAD/DistrictID from address lookup (format: A00026323803/TH)',
                '- site_companyName, site_building, site_street, site_town, site_postcode',
                '- site_contact_name, site_contact_email, site_contact_phone',
                '- lineType: "Existing" or "New"',
                '- cli: Phone number (required if lineType is "Existing")',
                '- customerRequiredDate: Format YYYY-MM-DD (e.g., 2026-02-15)',
                '- reseller_contact_name, reseller_contact_email, reseller_contact_phone',
                '- ipAddressOption: e.g., "Static Public IP Address"',
                '- routedIpOption: e.g., "None"',
                '- voiceProduct: e.g., "Gamma SIP Trunks" or "None"',
                '',
                'OPTIONAL FIELDS:',
                '- site_subPremises: e.g., "Flat 1" (optional sub-address)',
                '- installation_type: e.g., "Managed Install" (installation type)',
                '- customerReference: Your internal reference/order number',
                '- eccBand: e.g., "Charge Band A (£0)" (Excess Construction Charge band)',
                '- trcBand: e.g., "Charge Band 0 (not authorised beyond the NTE)" (Tie Cable band)',
                '- routerRequired: TRUE or FALSE',
                '- router: Router model (if routerRequired is TRUE)',
                '- router_delivery_*: Delivery address (if router required)',
                '- router_contact_*: Delivery contact (if router required)',
                '- voipReference: For number porting',
                '- resellerEmailNotifications: TRUE or FALSE',
                '',
                'IMPORTANT NOTES:',
                '- NAD Key must be obtained from Suitability Checker API first',
                '- Product names must match exactly as returned from API',
                '- Date must be at least 2 weeks in the future',
                '- Router fields required if routerRequired is TRUE',
                '- CLI required if lineType is "Existing"',
                '',
                'For valid product names and options, use:',
                '  python run_example1.py  (to get NAD keys)',
                '',
                'CARE LEVEL OPTIONS:',
                '- Standard Care',
                '- Enhanced Care',
                '- Premium Care',
                '',
                'LINE TYPE OPTIONS:',
                '- Existing (has existing phone line)',
                '- New (new installation)',
                '',
                'BOOLEAN VALUES:',
                '- Use TRUE or FALSE (all caps)',
                ''
            ]
        })
        
        instructions.to_excel(writer, sheet_name='Instructions', index=False)
        
        # Format the sheets
        workbook = writer.book
        
        # Format Orders sheet
        worksheet = writer.sheets['Orders']
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        # Format Instructions sheet
        inst_worksheet = writer.sheets['Instructions']
        inst_worksheet.column_dimensions['A'].width = 100
    
    print(f"✓ Template created: {filename}")
    print(f"\nNext steps:")
    print(f"1. Open {filename}")
    print(f"2. Fill in your order details in the 'Orders' sheet")
    print(f"3. Delete the example row")
    print(f"4. Save the file")
    print(f"5. Run: python sogea_bulk_ordering.py")
    
    return filename


if __name__ == "__main__":
    create_bulk_order_template()

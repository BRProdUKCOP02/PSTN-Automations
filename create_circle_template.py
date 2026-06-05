"""
Create SoGEA Bulk Order Template with Circle Cloud Example
Generates standard Excel template populated with Circle Cloud order data
"""

import pandas as pd
from datetime import datetime
import os

def create_circle_template():
    """
    Creates standard SoGEA bulk order template with Circle Cloud order data
    """
    
    # Define all columns (standard template format)
    columns = [
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
    
    # Circle Cloud order data from JSON
    order_data = {
        'accountNumber': '000169',
        'broadbandProduct': 'SoGEA 80:20 (1 month term)',
        'careLevel': 'Standard Care',
        
        'site_companyName': 'circle test',
        'site_nadKey': 'A14975671306/SW',
        'site_building': '1',
        'site_subPremises': 'Flat 1',
        'site_street': 'Goldsland Walk',
        'site_town': 'Cardiff',
        'site_county': 'Wenvoe',
        'site_postcode': 'CF5 6FD',
        
        'site_contact_name': 'Molly test',
        'site_contact_email': 'provisioning@circle.cloud',
        'site_contact_phone': '02382549891',
        
        'lineType': 'New',
        'cli': '',
        'customerRequiredDate': '2026-02-13',
        'installation_type': 'Managed Install',
        'customerReference': 'circle test',
        'eccBand': 'Charge Band A (£0)',
        'trcBand': 'Charge Band 0 (not authorised beyond the NTE)',
        
        'reseller_contact_name': 'Tyler Comley',
        'reseller_contact_email': 'provisioning@circle.cloud',
        'reseller_contact_phone': '+443330436600',
        
        'ipAddressOption': 'Static Public IP Address',
        'routedIpOption': 'None',
        'voiceProduct': 'Gamma SIP Trunks',
        
        'routerRequired': 'FALSE',
        'router': '',
        'router_delivery_building': '',
        'router_delivery_street': '',
        'router_delivery_town': '',
        'router_delivery_county': '',
        'router_delivery_postcode': '',
        'router_contact_name': '',
        'router_contact_email': '',
        'router_contact_phone': '',
        
        'voipReference': '',
        
        'resellerEmailNotifications': 'TRUE'
    }
    
    # Create DataFrame with the order
    df = pd.DataFrame([order_data])
    
    # Generate filename
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = os.path.join(script_dir, f"circle_cloud_order_{timestamp}.xlsx")
    
    # Create Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write main data sheet
        df.to_excel(writer, sheet_name='Orders', index=False)
        
        # Create instructions sheet
        instructions = pd.DataFrame({
            'INSTRUCTIONS': [
                'SoGEA BULK ORDER TEMPLATE - CIRCLE CLOUD TEST ORDER',
                '',
                'This template is pre-filled with Circle Cloud test order data.',
                '',
                'HOW TO USE:',
                '1. Review the order in the "Orders" sheet',
                '2. Modify values if needed or add more rows',
                '3. Save the file',
                '4. Run: python sogea_bulk_ordering.py',
                '',
                'REQUIRED FIELDS (must not be empty):',
                '- accountNumber: Your Gamma billing account number',
                '- broadbandProduct: Exact product name from Suitability Checker',
                '- careLevel: e.g., "Standard Care"',
                '- site_nadKey: NAD/DistrictID from address lookup',
                '- site_companyName, site_building, site_street, site_town, site_postcode',
                '- site_contact_name, site_contact_email, site_contact_phone',
                '- lineType: "Existing" or "New"',
                '- cli: Phone number (required if lineType is "Existing")',
                '- customerRequiredDate: Format YYYY-MM-DD',
                '- reseller_contact_name, reseller_contact_email, reseller_contact_phone',
                '- ipAddressOption, routedIpOption, voiceProduct',
                '',
                'OPTIONAL FIELDS:',
                '- site_subPremises: Sub-address like "Flat 1"',
                '- installation_type: e.g., "Managed Install"',
                '- customerReference: Your internal reference',
                '- eccBand: Excess Construction Charge band',
                '- trcBand: Tie Cable band',
                '- routerRequired: TRUE or FALSE',
                '- router: Router model (if routerRequired is TRUE)',
                '- router_delivery_*: Delivery address',
                '- router_contact_*: Delivery contact',
                '- voipReference: For number porting',
                '- resellerEmailNotifications: TRUE or FALSE',
                '',
                'IMPORTANT NOTES:',
                '- NAD Key must be from Gamma Suitability Checker',
                '- Product names must match exactly from API',
                '- Date must be at least 2 weeks in future',
                '- Boolean values: TRUE or FALSE (all caps)',
                '- CLI required if lineType is "Existing"',
                '',
                'LINE TYPE OPTIONS:',
                '- Existing (has existing phone line)',
                '- New (new installation)',
                ''
            ]
        })
        
        instructions.to_excel(writer, sheet_name='Instructions', index=False)
        
        # Format the sheets
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
    
    print(f"✓ Template created: {os.path.basename(filename)}")
    print("\nPre-filled with Circle Cloud test order:")
    print("  Account: 000169")
    print("  Product: SoGEA 80:20 (1 month term)")
    print("  Site: circle test, Cardiff")
    print("  NAD Key: A14975671306/SW")
    print("\nReady to use with: python sogea_bulk_ordering.py")

if __name__ == "__main__":
    create_circle_template()

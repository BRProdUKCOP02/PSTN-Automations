"""
Bulk Check Broadband Orders
This script reads a list of order IDs from a spreadsheet and retrieves detailed status
information for each order, outputting a comprehensive Excel report.
"""

from config import get_credentials, ENVIRONMENT
from bb_ordering_api import GammaBroadbandOrderingAPI, GammaOrderingAPIError, GammaAuthenticationError
from graph_mailbox_check import send_order_report_email
import pandas as pd
import sys
import os
from datetime import datetime
import time
import glob
import shutil

# =============================================================================
# CONFIGURATION
# =============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Input folder for MANUAL status checks from SharePoint/OneDrive
# (Status checks are run separately, 24+ hours after order placement)
INPUT_FOLDER = r"\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\SoGEA Bulk Status Checker"

# Output folder for results
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, "order_check_output")

# Processed folder for completed checks
PROCESSED_FOLDER = os.path.join(INPUT_FOLDER, "processed")

CHECK_DELAY = 0.5  # seconds between each API call

# Ensure folders exist
os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def flatten_order_data(order_data):
    """
    Flatten nested order JSON into a single-level dictionary for Excel export
    
    Args:
        order_data: Order data from API
        
    Returns:
        Flattened dictionary with all order details
    """
    flat = {}
    
    # Basic order info
    flat['Order_ID'] = order_data.get('id', '')
    flat['Status'] = order_data.get('status', '')
    flat['Account_Number'] = order_data.get('accountNumber', '')
    flat['Broadband_Product'] = order_data.get('broadbandProduct', '')
    flat['Voice_Product'] = order_data.get('voiceProduct', '')
    flat['Care_Level'] = order_data.get('careLevel', '')
    flat['Routed_IP_Option'] = order_data.get('routedIpOption', '')
    flat['IP_Address_Option'] = order_data.get('ipAddressOption', '')
    flat['Reseller_Email_Notifications'] = order_data.get('resellerEmailNotifications', '')
    flat['Under_Regrade'] = order_data.get('underRegrade', '')
    
    # Installation details
    if 'installation' in order_data:
        install = order_data['installation']
        flat['Line_Type'] = install.get('lineType', '')
        flat['CLI'] = install.get('cli', '')
        flat['Customer_Required_Date'] = install.get('customerRequiredDate', '')
        flat['Customer_Reference'] = install.get('customerReference', '')
        flat['Installation_Type'] = install.get('type', '')
        flat['Supplier_Promised_Date'] = install.get('supplierPromisedDate', '')
        
        # Site details
        if 'site' in install:
            site = install['site']
            flat['Site_Company_Name'] = site.get('companyName', '')
            flat['Site_NAD_Key'] = site.get('nadKey', '')
            
            # Site contact
            if 'contact' in site:
                contact = site['contact']
                flat['Site_Contact_Name'] = contact.get('name', '')
                flat['Site_Contact_Email'] = contact.get('emailAddress', '')
                flat['Site_Contact_Phone'] = contact.get('telephoneNumber', '')
            
            # Site address
            if 'address' in site:
                addr = site['address']
                flat['Site_Building'] = addr.get('building', '')
                flat['Site_SubPremises'] = addr.get('subPremises', '')
                flat['Site_Street'] = addr.get('street', '')
                flat['Site_Town'] = addr.get('town', '')
                flat['Site_County'] = addr.get('county', '')
                flat['Site_Postcode'] = addr.get('postcode', '')
        
        # Engineer appointment
        if 'engineerAppointment' in install:
            appt = install['engineerAppointment']
            flat['Engineer_Appointment_Date'] = appt.get('date', '')
            flat['Engineer_Appointment_Timeslot'] = appt.get('timeslot', '')
            flat['Engineer_Appointment_Start'] = appt.get('timeslotStart', '')
            flat['Engineer_Appointment_End'] = appt.get('timeslotEnd', '')
    
    # Reseller contact
    if 'resellerContact' in order_data:
        reseller = order_data['resellerContact']
        flat['Reseller_Contact_Name'] = reseller.get('name', '')
        flat['Reseller_Contact_Email'] = reseller.get('emailAddress', '')
        flat['Reseller_Contact_Phone'] = reseller.get('telephoneNumber', '')
    
    # Equipment details
    if 'equipment' in order_data:
        equipment = order_data['equipment']
        flat['Router_Required'] = equipment.get('routerRequired', '')
        flat['Router_Model'] = equipment.get('router', '')
        flat['Equipment_Company_Name'] = equipment.get('companyName', '')
        
        # Delivery contact
        if 'deliveryContact' in equipment:
            delivery = equipment['deliveryContact']
            flat['Delivery_Contact_Name'] = delivery.get('name', '')
            flat['Delivery_Contact_Email'] = delivery.get('emailAddress', '')
            flat['Delivery_Contact_Phone'] = delivery.get('telephoneNumber', '')
        
        # Delivery address
        if 'deliveryAddress' in equipment:
            del_addr = equipment['deliveryAddress']
            flat['Delivery_Building'] = del_addr.get('building', '')
            flat['Delivery_Street'] = del_addr.get('street', '')
            flat['Delivery_Town'] = del_addr.get('town', '')
            flat['Delivery_County'] = del_addr.get('county', '')
            flat['Delivery_Postcode'] = del_addr.get('postcode', '')
        
        # Router configuration
        if 'routerConfiguration' in equipment:
            router_config = equipment['routerConfiguration']
            flat['Router_DSL_Username'] = router_config.get('dslUsername', '')
            flat['Router_WAN_IP_Address'] = router_config.get('wanIpAddress', '')
    
    # Number port details
    if 'numberPort' in order_data:
        port = order_data['numberPort']
        flat['Number_Port_VOIP_Reference'] = port.get('voipReference', '')
        flat['Number_Port_Order_ID'] = port.get('portOrderId', '')
        flat['Number_Port_Status'] = port.get('portOrderStatus', '')
    
    # Updates summary (last update only for main sheet)
    if 'updates' in order_data and order_data['updates']:
        last_update = order_data['updates'][-1]
        flat['Last_Update_Timestamp'] = last_update.get('timestamp', '')
        flat['Last_Update_Status'] = last_update.get('status', '')
        flat['Last_Update_Message'] = last_update.get('message', '')
        flat['Total_Updates'] = len(order_data['updates'])
    else:
        flat['Last_Update_Timestamp'] = ''
        flat['Last_Update_Status'] = ''
        flat['Last_Update_Message'] = ''
        flat['Total_Updates'] = 0
    
    return flat

def extract_updates(order_data, order_id):
    """
    Extract all updates from an order into a list of dictionaries
    
    Args:
        order_data: Order data from API
        order_id: Order ID
        
    Returns:
        List of update dictionaries
    """
    updates = []
    
    if 'updates' in order_data and order_data['updates']:
        for idx, update in enumerate(order_data['updates'], 1):
            updates.append({
                'Order_ID': order_id,
                'Update_Number': idx,
                'Timestamp': update.get('timestamp', ''),
                'Status': update.get('status', ''),
                'Message': update.get('message', '')
            })
    
    return updates

def get_input_files():
    """
    Get all Excel files from the SharePoint input folder
    
    Returns:
        List of tuples: (full_path, filename_without_extension)
    """
    excel_patterns = ['*.xlsx', '*.xls']
    input_files = []
    
    for pattern in excel_patterns:
        files = glob.glob(os.path.join(INPUT_FOLDER, pattern))
        for file_path in files:
            # Skip temporary Excel files that start with ~$
            filename = os.path.basename(file_path)
            if not filename.startswith('~$'):
                filename_without_ext = os.path.splitext(filename)[0]
                input_files.append((file_path, filename_without_ext))
    
    return sorted(input_files)  # Sort for consistent processing order

def move_to_processed(file_path, processed_folder):
    """
    Move a processed file to the processed folder
    
    Args:
        file_path: Full path to the file to move
        processed_folder: Destination folder path
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        filename = os.path.basename(file_path)
        destination = os.path.join(processed_folder, filename)
        
        # If file already exists in destination, append timestamp
        if os.path.exists(destination):
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{name}_moved_{timestamp}{ext}"
            destination = os.path.join(processed_folder, filename)
        
        shutil.move(file_path, destination)
        print(f"  ✓ Moved input file to: {destination}")
        return True
    except Exception as e:
        print(f"  ⚠ Warning: Could not move file to processed folder: {e}")
        return False

# =============================================================================
# MAIN PROCESS
# =============================================================================

def process_single_file(input_file_path, input_filename, api, creds):
    """
    Process a single input file
    
    Args:
        input_file_path: Full path to the input Excel file
        input_filename: Filename without extension (for output naming)
        api: Authenticated GammaBroadbandOrderingAPI instance
        creds: Credentials dictionary for environment info
    
    Returns:
        bool: True if successful, False otherwise
    """
    print("=" * 80)
    print(f"PROCESSING: {input_filename}")
    print("=" * 80)
    
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(OUTPUT_FOLDER, f"{input_filename}_results_{timestamp}.xlsx")
    
    # Load input file
    print(f"\nLoading order IDs from: {os.path.basename(input_file_path)}")
    
    if not os.path.exists(input_file_path):
        print(f"\n❌ Error: Input file not found: {input_file_path}")
        print("\nPlease create an Excel file with a column named 'Order_ID'")
        return False
    
    try:
        df = pd.read_excel(input_file_path)
    except Exception as e:
        print(f"\n❌ Error reading file: {e}")
        return False
    
    # Validate the file has Order_ID column
    if 'Order_ID' not in df.columns:
        print(f"\n❌ Error: Input file must have a column named 'Order_ID'")
        print(f"\nFound columns: {', '.join(df.columns)}")
        return False
    
    # Remove any empty rows
    df = df.dropna(subset=['Order_ID'])
    
    # Convert to integers
    try:
        df['Order_ID'] = df['Order_ID'].astype(int)
    except ValueError as e:
        print(f"\n❌ Error: Order_ID column must contain only numbers")
        print(f"Error: {e}")
        return False
    
    order_ids = df['Order_ID'].tolist()
    print(f"✓ Loaded {len(order_ids)} order IDs to check")
    
    # Process orders
    print("\n" + "=" * 80)
    print("CHECKING ORDERS")
    print("=" * 80)
    
    results = []
    all_updates = []
    
    for idx, order_id in enumerate(order_ids, 1):
        print(f"\nChecking order {idx} of {len(order_ids)}: Order ID {order_id}...")
        
        try:
            # Fetch order details
            order_data = api.get_order(order_id)
            
            status = order_data.get('status', 'Unknown')
            product = order_data.get('broadbandProduct', 'Unknown')
            
            print(f"  ✓ SUCCESS - Status: {status}, Product: {product}")
            
            # Flatten order data
            flat_order = flatten_order_data(order_data)
            flat_order['Check_Status'] = 'SUCCESS'
            flat_order['Check_Error'] = ''
            
            results.append(flat_order)
            
            # Extract updates
            updates = extract_updates(order_data, order_id)
            all_updates.extend(updates)
            
        except GammaOrderingAPIError as e:
            print(f"  ❌ FAILED - {e.message}")
            
            # Add minimal row with error
            results.append({
                'Order_ID': order_id,
                'Check_Status': 'FAILED',
                'Check_Error': f"HTTP {e.status_code}: {e.message}",
                'Status': '',
                'Broadband_Product': ''
            })
        
        except Exception as e:
            print(f"  ❌ ERROR - {str(e)}")
            
            results.append({
                'Order_ID': order_id,
                'Check_Status': 'ERROR',
                'Check_Error': str(e),
                'Status': '',
                'Broadband_Product': ''
            })
        
        # Delay before next request
        if idx < len(order_ids):
            time.sleep(CHECK_DELAY)
    
    # Create output report
    print("\n" + "=" * 80)
    print("GENERATING REPORT")
    print("=" * 80)
    
    results_df = pd.DataFrame(results)
    updates_df = pd.DataFrame(all_updates) if all_updates else pd.DataFrame()
    
    # Save to Excel with multiple sheets
    print(f"\nSaving results to: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Order Details', index=False)
        
        if not updates_df.empty:
            updates_df.to_excel(writer, sheet_name='Order Updates', index=False)
        
        # Add summary sheet
        summary_data = {
            'Metric': [
                'Total Orders Checked',
                'Successful Checks',
                'Failed Checks',
                'Errors',
                'Environment',
                'Check Date/Time'
            ],
            'Value': [
                len(order_ids),
                len([r for r in results if r.get('Check_Status') == 'SUCCESS']),
                len([r for r in results if r.get('Check_Status') == 'FAILED']),
                len([r for r in results if r.get('Check_Status') == 'ERROR']),
                creds['environment_name'],
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Auto-size columns
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"✓ Report saved successfully!")
    
    # Summary
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"\nTotal orders checked: {len(order_ids)}")
    
    successful_count = len([r for r in results if r.get('Check_Status') == 'SUCCESS'])
    failed_count = len([r for r in results if r.get('Check_Status') == 'FAILED'])
    error_count = len([r for r in results if r.get('Check_Status') == 'ERROR'])
    
    print(f"Successful: {successful_count}")
    print(f"Failed: {failed_count}")
    print(f"Errors: {error_count}")
    print(f"\nResults saved to: {output_file}")
    
    print("\n" + "=" * 80)
    print(f"✓ FILE PROCESSING COMPLETE: {input_filename}")
    print("=" * 80)
    
    # Move the processed input file to the processed folder
    print("\nMoving input file to processed folder...")
    move_to_processed(input_file_path, PROCESSED_FOLDER)
    
    # Send email report with status check results
    print("\n" + "=" * 80)
    print("SENDING EMAIL REPORT")
    print("=" * 80)
    
    email_sent = send_order_report_email(
        report_file_path=output_file,
        recipient_email='psmanaged.delivery@gamma.co.uk',
        input_filename=input_filename,
        total_orders=len(order_ids),
        successful=successful_count,
        failed=failed_count,
        errors=error_count,
        report_type='Order Status Check',
        subject_prefix='SoGEA Order Status Update'
    )
    
    if email_sent:
        print("✓ Report successfully emailed to psmanaged.delivery@gamma.co.uk")
    else:
        print("⚠ Warning: Failed to send email report")
    
    return True

def main():
    """
    Main entry point - finds and processes all Excel files in the input folder
    """
    print("=" * 80)
    print("BULK ORDER CHECK")
    print("=" * 80)
    
    # Get all input files
    input_files = get_input_files()
    
    if not input_files:
        print(f"\n❌ No Excel files found in: {INPUT_FOLDER}")
        print("\nPlease place your order ID spreadsheets in the input folder.")
        print(f"Input folder location: {INPUT_FOLDER}")
        print("\nExpected format: Excel file with 'Order_ID' column")
        sys.exit(1)
    
    print(f"\n✓ Found {len(input_files)} file(s) to process:")
    for _, filename in input_files:
        print(f"  • {filename}")
    
    # Authenticate once for all files
    print("\n" + "=" * 80)
    print("AUTHENTICATION")
    print("=" * 80)
    
    creds = get_credentials()
    print(f"\nEnvironment: {creds['environment_name']}")
    print(f"Username: {creds['username']}")
    print(f"\nAuthenticating...")
    
    try:
        api = GammaBroadbandOrderingAPI(
            username=creds['username'],
            password=creds['password'],
            use_production=creds['use_production'],
            auto_refresh=True
        )
        print("✓ Authentication successful!")
    except GammaAuthenticationError as e:
        print(f"\n❌ Authentication failed: {e.message}")
        print(f"\nPlease check credentials in config.py")
        sys.exit(1)
    
    # Process each file
    print("\n" + "=" * 80)
    print("PROCESSING FILES")
    print("=" * 80)
    
    processed_count = 0
    failed_count = 0
    
    for file_path, filename in input_files:
        print(f"\n\n")
        success = process_single_file(file_path, filename, api, creds)
        if success:
            processed_count += 1
        else:
            failed_count += 1
    
    # Final summary
    print("\n\n" + "=" * 80)
    print("ALL FILES PROCESSED")
    print("=" * 80)
    print(f"\nTotal files: {len(input_files)}")
    print(f"✓ Successfully processed: {processed_count}")
    print(f"❌ Failed: {failed_count}")
    print(f"\n✓ Output files saved to: {OUTPUT_FOLDER}")
    print("\n" + "=" * 80)

if __name__ == "__main__":
    main()

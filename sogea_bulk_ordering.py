"""
SoGEA Bulk Ordering Script
Reads orders from an Excel spreadsheet and places the order via Gamma API

Order Support:
- NEW orders: Place new broadband orders
- REGRADE orders: Upgrade/downgrade existing orders (e.g., migrate from existing service) **Regrade functionality is currently broken server-side, will be re-enabled once fixed by Gamma**

For NEW orders:
- Use orderType = "NEW" (or leave blank, defaults to NEW)
- Fill in all standard fields

For REGRADE orders:
- Use orderType = "REGRADE"
- Add existingOrderId column with the ID of the order to regrade
- Only requires: broadbandProduct, customerRequiredDate, routerRequired, router details (if router needed)
- Does not require: accountNumber, careLevel, site details, reseller contact (uses existing order values)
"""

import pandas as pd
from config import get_credentials, ENVIRONMENT
from bb_ordering_api import GammaBroadbandOrderingAPI, GammaOrderingAPIError, GammaAuthenticationError
from graph_mailbox_check import send_order_report_email
from datetime import datetime
import sys
import time
import shutil

# =============================================================================
# CONFIGURATION
# =============================================================================

import os
import glob

# Get script directory and build paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = r"\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\SoGEA New orders & Regrades"
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, "output")
CHECK_INPUT_FOLDER = os.path.join(SCRIPT_DIR, "order_check_input")
PROCESSED_FOLDER = r"\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\SoGEA New orders & Regrades\processed"

# Delay between orders (in seconds) to avoid rate limiting
ORDER_DELAY = 2

# Ensure folders exist
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(CHECK_INPUT_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def convert_to_bool(value):
    """Convert string TRUE/FALSE to boolean"""
    if pd.isna(value) or value == '':
        return None
    if isinstance(value, bool):
        return value
    return str(value).upper() == 'TRUE'

def convert_to_date_string(value):
    """Convert Excel date or string to YYYY-MM-DD format"""
    if pd.isna(value) or value == '':
        return None
    
    # If it's already a pandas Timestamp or datetime object
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime('%Y-%m-%d')
    
    # If it's a string, check if it's already in correct format
    date_str = str(value).strip()
    
    # Remove time component if present (e.g., "2026-01-15 00:00:00")
    if ' ' in date_str:
        date_str = date_str.split(' ')[0]
    
    return date_str

def format_phone_number(value):
    """
    Format phone number, preserving leading zeros that Excel may have stripped
    
    Args:
        value: Phone number from Excel (may be numeric or string)
        
    Returns:
        Properly formatted phone number string with leading zeros
    """
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and remove any decimals Excel may have added
    phone_str = str(value).replace('.0', '').strip()
    
    # If it's a UK mobile number without leading 0, add it
    if phone_str.isdigit() and len(phone_str) == 10 and phone_str.startswith('7'):
        return '0' + phone_str
    
    # If it's a UK landline without leading 0, add it
    if phone_str.isdigit() and len(phone_str) == 10 and phone_str[0] in '1238':
        return '0' + phone_str
    
    return phone_str

def build_order_payload(row):
    """
    Build the API order payload from a spreadsheet row
    
    Args:
        row: Pandas Series containing order data
        
    Returns:
        Dictionary formatted for API request
    """
    
    # Basic order structure
    order = {
        "accountNumber": int(row['accountNumber']),
        "broadbandProduct": str(row['broadbandProduct']),
        "careLevel": str(row['careLevel']),
        "resellerEmailNotifications": convert_to_bool(row.get('resellerEmailNotifications', True)),
        "ipAddressOption": str(row['ipAddressOption']),
        "voiceProduct": str(row['voiceProduct']) if pd.notna(row.get('voiceProduct')) else 'None'
    }
    
    # Add routedIpOption only if not "none"
    if pd.notna(row.get('routedIpOption')) and str(row['routedIpOption']).strip().lower() != 'none':
        order["routedIpOption"] = str(row['routedIpOption'])
    
    # Installation details
    installation = {
        "lineType": str(row['lineType']),
        "customerRequiredDate": convert_to_date_string(row['customerRequiredDate']),
        "site": {
            "companyName": str(row['site_companyName']),
            "nadKey": str(row['site_nadKey']),
            "address": {
                "building": str(row['site_building']),
                "street": str(row['site_street']),
                "town": str(row['site_town']),
                "county": str(row.get('site_county', '')),
                "postcode": str(row['site_postcode'])
            },
            "contact": {
                "name": str(row['site_contact_name']),
                "emailAddress": str(row['site_contact_email']),
                "telephoneNumber": format_phone_number(row['site_contact_phone'])
            }
        }
    }
    
    # Add optional installation fields if provided
    if not pd.isna(row.get('installation_type')) and str(row.get('installation_type', '')).strip():
        installation["type"] = str(row['installation_type'])
    
    if not pd.isna(row.get('customerReference')) and str(row.get('customerReference', '')).strip():
        installation["customerReference"] = str(row['customerReference'])
    
    if not pd.isna(row.get('eccBand')) and str(row.get('eccBand', '')).strip():
        installation["eccBand"] = str(row['eccBand'])
    
    if not pd.isna(row.get('trcBand')) and str(row.get('trcBand', '')).strip():
        installation["trcBand"] = str(row['trcBand'])
    
    # Add subPremises to address if provided
    if not pd.isna(row.get('site_subPremises')) and str(row.get('site_subPremises', '')).strip():
        installation["site"]["address"]["subPremises"] = str(row['site_subPremises'])

    # Add CLI if existing line
    if row['lineType'].upper() == 'EXISTING' and not pd.isna(row.get('cli')):
        installation["cli"] = format_phone_number(row['cli'])
    
    order["installation"] = installation
    
    # Reseller contact
    order["resellerContact"] = {
        "name": str(row['reseller_contact_name']),
        "emailAddress": str(row['reseller_contact_email']),
        "telephoneNumber": format_phone_number(row['reseller_contact_phone'])
    }
    
    # Equipment (REQUIRED by API - even if no router needed)
    router_required = convert_to_bool(row.get('routerRequired', False))
    equipment = {
        "routerRequired": router_required
    }
    
    if router_required:
        equipment["router"] = str(row.get('router', ''))
        
        # Delivery address
        if not pd.isna(row.get('router_delivery_postcode')):
            equipment["deliveryAddress"] = {
                "building": str(row.get('router_delivery_building', '')),
                "street": str(row.get('router_delivery_street', '')),
                "town": str(row.get('router_delivery_town', '')),
                "county": str(row.get('router_delivery_county', '')),
                "postcode": str(row.get('router_delivery_postcode', ''))
            }
        
        # Delivery contact
        if not pd.isna(row.get('router_contact_email')):
            equipment["deliveryContact"] = {
                "name": str(row.get('router_contact_name', '')),
                "emailAddress": str(row.get('router_contact_email', '')),
                "telephoneNumber": format_phone_number(row.get('router_contact_phone', ''))
            }
    
    order["equipment"] = equipment
    
    # Number porting (if provided)
    if not pd.isna(row.get('voipReference')) and row.get('voipReference', '').strip():
        order["numberPort"] = {
            "voipReference": str(row['voipReference'])
        }
    
    return order

def build_regrade_payload(row):
    """
    Build the API regrade payload from a spreadsheet row
    
    Args:
        row: Pandas Series containing regrade data
        
    Returns:
        Dictionary formatted for regrade API request
    """
    
    # Required fields
    regrade = {
        "broadbandProduct": str(row['broadbandProduct']),
        "customerRequiredDate": convert_to_date_string(row['customerRequiredDate']),
        "routerRequired": convert_to_bool(row.get('routerRequired', False))
    }
    
    # Optional: Voice product
    if not pd.isna(row.get('voiceProduct')) and str(row.get('voiceProduct', '')).strip():
        regrade["voiceProduct"] = str(row['voiceProduct'])
    
    # Optional: Routed IP option (skip if 'none')
    if not pd.isna(row.get('routedIpOption')) and str(row.get('routedIpOption', '')).strip() and str(row['routedIpOption']).strip().lower() != 'none':
        regrade["routedIpOption"] = str(row['routedIpOption'])
    
    # Optional: Care level
    if not pd.isna(row.get('careLevel')) and str(row.get('careLevel', '')).strip():
        regrade["careLevel"] = str(row['careLevel'])
    
    # Optional: Install type (required when regrading from non-FTTC to FTTC)
    if not pd.isna(row.get('installType')) and str(row.get('installType', '')).strip():
        regrade["installType"] = str(row['installType'])
    
    # Optional: Router details (if router is required)
    if regrade["routerRequired"]:
        if not pd.isna(row.get('router')) and str(row.get('router', '')).strip():
            regrade["router"] = str(row['router'])
        
        # Optional: Company name
        if not pd.isna(row.get('router_companyName')) and str(row.get('router_companyName', '')).strip():
            regrade["companyName"] = str(row['router_companyName'])
        
        # Optional: Delivery address for router
        if not pd.isna(row.get('router_delivery_postcode')):
            regrade["deliveryAddress"] = {
                "building": str(row.get('router_delivery_building', '')),
                "street": str(row.get('router_delivery_street', '')),
                "town": str(row.get('router_delivery_town', '')),
                "county": str(row.get('router_delivery_county', '')),
                "postcode": str(row.get('router_delivery_postcode', ''))
            }
        
        # Optional: Delivery contact for router
        if not pd.isna(row.get('router_contact_email')):
            regrade["deliveryContact"] = {
                "name": str(row.get('router_contact_name', '')),
                "emailAddress": str(row.get('router_contact_email', '')),
                "telephoneNumber": format_phone_number(row.get('router_contact_phone', ''))
            }
    
    return regrade

def validate_row(row, row_num):
    """
    Validate that a row has all required fields
    
    Returns:
        Tuple: (is_valid, error_message)
    """
    errors = []
    
    # Determine order type (defaults to NEW if not specified)
    order_type = str(row.get('orderType', 'NEW')).upper()
    
    if order_type == 'REGRADE':
        # REGRADE validation - only validate fields needed for regrade
        regrade_required_fields = [
            'broadbandProduct',
            'customerRequiredDate',
            'routerRequired'
        ]
        
        # Check existingOrderId is provided
        if pd.isna(row.get('existingOrderId')) or str(row.get('existingOrderId', '')).strip() == '':
            errors.append("Missing required field: existingOrderId")
        
        # Check regrade-specific required fields
        for field in regrade_required_fields:
            if pd.isna(row.get(field)) or str(row.get(field, '')).strip() == '':
                errors.append(f"Missing required field: {field}")
    
    else:
        # NEW order validation - validate all standard fields
        required_fields = [
            'accountNumber',
            'broadbandProduct',
            'careLevel',
            'site_companyName',
            'site_nadKey',
            'site_building',
            'site_street',
            'site_town',
            'site_postcode',
            'site_contact_name',
            'site_contact_email',
            'site_contact_phone',
            'lineType',
            'customerRequiredDate',
            'reseller_contact_name',
            'reseller_contact_email',
            'reseller_contact_phone',
            'ipAddressOption'
        ]
        
        for field in required_fields:
            if pd.isna(row.get(field)) or str(row.get(field, '')).strip() == '':
                errors.append(f"Missing required field: {field}")
    
        # Validate lineType
        if not pd.isna(row.get('lineType')):
            line_type = str(row['lineType']).upper()
            if line_type not in ['EXISTING', 'NEW']:
                errors.append(f"lineType must be 'Existing' or 'New', got: {row['lineType']}")
    
    # Validate NAD Key format (DISABLED - uncomment to re-enable)
    # if not pd.isna(row.get('site_nadKey')):
    #     nad_key = str(row['site_nadKey'])
    #     if not (nad_key.startswith('A') and len(nad_key) >= 11):
    #         errors.append(f"Invalid NAD Key format: {nad_key} (should be like 'A00010584545')")
    
    # Validate date format
    if not pd.isna(row.get('customerRequiredDate')):
        try:
            date_str = convert_to_date_string(row['customerRequiredDate'])
            if date_str:
                datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            errors.append(f"Invalid date format: {row['customerRequiredDate']} (should be YYYY-MM-DD)")
    
    # Validate router fields if router required
    router_required = convert_to_bool(row.get('routerRequired', False))
    if router_required:
        if pd.isna(row.get('router')) or str(row.get('router', '')).strip() == '':
            errors.append("Router model required when routerRequired is TRUE")
    
    if errors:
        return False, f"Row {row_num}: " + "; ".join(errors)
    
    return True, None

def get_input_files():
    """
    Get all Excel files from the input folder
    
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

# =============================================================================
# MAIN PROCESS
# =============================================================================

def process_single_file(input_file_path, input_filename, api):
    """
    Process a single input file
    
    Args:
        input_file_path: Full path to the input Excel file
        input_filename: Filename without extension (for output naming)
        api: Authenticated GammaBroadbandOrderingAPI instance
    """
    print("=" * 80)
    print(f"PROCESSING: {input_filename}")
    print("=" * 80)
    
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(OUTPUT_FOLDER, f"{input_filename}_results_{timestamp}.xlsx")
    check_file = os.path.join(CHECK_INPUT_FOLDER, f"{input_filename}_order_ids_{timestamp}.xlsx")
    
    # Check if input file exists and load it
    try:
        df = pd.read_excel(input_file_path, sheet_name='Orders')
    except FileNotFoundError:
        print(f"\n❌ Error: Input file '{input_file_path}' not found")
        return False
    except Exception as e:
        print(f"\n❌ Error reading file: {e}")
        return False
    
    print(f"\n✓ Loaded {len(df)} orders from {os.path.basename(input_file_path)}")
    
    # Validate all rows first
    print("\nValidating orders...")
    validation_errors = []
    
    for idx, row in df.iterrows():
        is_valid, error_msg = validate_row(row, idx + 2)  # +2 for Excel row (1-indexed + header)
        if not is_valid:
            validation_errors.append(error_msg)
    
    if validation_errors:
        print(f"\n❌ Validation failed! Found {len(validation_errors)} errors:\n")
        for error in validation_errors:
            print(f"  • {error}")
        print("\nPlease fix the errors in the spreadsheet and try again.")
        return False
    
    print("✓ All orders validated successfully")
    
    # Process orders
    print("\n" + "=" * 80)
    print("PLACING ORDERS")
    print("=" * 80)
    
    results = []
    
    for idx, row in df.iterrows():
        row_num = idx + 2
        print(f"\nProcessing order {idx + 1} of {len(df)} (Row {row_num})...")
        
        # Determine order type
        order_type = str(row.get('orderType', 'NEW')).upper()
        
        # Print appropriate info based on order type
        if order_type == 'REGRADE':
            existing_order_id = str(row.get('existingOrderId', 'Unknown'))
            print(f"  Type: REGRADE (Order {existing_order_id})")
        else:
            company = row.get('site_companyName', 'Unknown')
            print(f"  Company: {company}")
        
        print(f"  Product: {row['broadbandProduct']}")
        
        # Start with all original row data
        result_row = row.to_dict()
        result_row['Original_Row_Number'] = row_num
        
        try:
            if order_type == 'REGRADE':
                # Regrade existing order
                # Get existing order ID
                if pd.isna(row.get('existingOrderId')):
                    raise ValueError("existingOrderId is required for REGRADE orders")
                
                existing_order_id = str(row['existingOrderId'])
                
                # Build regrade payload
                regrade_data = build_regrade_payload(row)
                
                # Place regrade
                response = api.place_regrade(existing_order_id, regrade_data)
                
                order_id = response.get('id', existing_order_id)
                status = response.get('status', 'Unknown')
                
                print(f"  ✓ SUCCESS - Regrade Order ID: {order_id}, Status: {status}")
                
                # Add outcome columns
                result_row['Outcome_Status'] = 'SUCCESS'
                result_row['Outcome_Order_ID'] = order_id
                result_row['Outcome_Order_Status'] = status
                result_row['Outcome_Error_Code'] = ''
                result_row['Outcome_Error_Message'] = ''
                
            else:
                # New order
                print(f"  Type: NEW")
                
                # Build order payload
                order_data = build_order_payload(row)
                
                # Place order
                response = api.place_order(order_data)
                
                order_id = response.get('id', 'Unknown')
                status = response.get('status', 'Unknown')
                
                print(f"  ✓ SUCCESS - Order ID: {order_id}, Status: {status}")
                
                # Add outcome columns
                result_row['Outcome_Status'] = 'SUCCESS'
                result_row['Outcome_Order_ID'] = order_id
                result_row['Outcome_Order_Status'] = status
                result_row['Outcome_Error_Code'] = ''
                result_row['Outcome_Error_Message'] = ''
            
            results.append(result_row)
            
            # Delay before next order
            if idx < len(df) - 1:
                time.sleep(ORDER_DELAY)
            
        except GammaOrderingAPIError as e:
            print(f"  ❌ FAILED - {e.message}")
            
            # Extract error code if present
            error_code = ''
            if e.response_data and 'code' in e.response_data:
                error_code = e.response_data['code']
            elif e.status_code:
                error_code = f"HTTP_{e.status_code}"
            
            # Add outcome columns
            result_row['Outcome_Status'] = 'FAILED'
            result_row['Outcome_Order_ID'] = ''
            result_row['Outcome_Order_Status'] = ''
            result_row['Outcome_Error_Code'] = error_code
            result_row['Outcome_Error_Message'] = e.message
            
            results.append(result_row)
        
        except Exception as e:
            print(f"  ❌ ERROR - {str(e)}")
            
            # Add outcome columns
            result_row['Outcome_Status'] = 'ERROR'
            result_row['Outcome_Order_ID'] = ''
            result_row['Outcome_Order_Status'] = ''
            result_row['Outcome_Error_Code'] = 'SYSTEM_ERROR'
            result_row['Outcome_Error_Message'] = str(e)
            
            results.append(result_row)
    
    # Save results
    print("\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    
    results_df = pd.DataFrame(results)
    
    # Create error code legend
    error_legend = pd.DataFrame([
        {
            'Error Code': 'HTTP_400',
            'Description': 'Bad Request - Invalid data or parameters',
            'Resolution': 'Check all required fields are present and correctly formatted. Review API documentation.'
        },
        {
            'Error Code': 'HTTP_401',
            'Description': 'Unauthorized - Authentication failed',
            'Resolution': 'Verify credentials in config.py are correct. Check token has not expired.'
        },
        {
            'Error Code': 'HTTP_403',
            'Description': 'Forbidden - Insufficient permissions',
            'Resolution': 'Verify account has permission to place orders. Contact Gamma support.'
        },
        {
            'Error Code': 'HTTP_404',
            'Description': 'Not Found - Resource does not exist',
            'Resolution': 'Check NAD key, account number, or other identifiers are valid.'
        },
        {
            'Error Code': 'HTTP_409',
            'Description': 'Conflict - Duplicate order or resource conflict',
            'Resolution': 'Order may already exist. Check for duplicate submissions.'
        },
        {
            'Error Code': 'HTTP_422',
            'Description': 'Unprocessable Entity - Validation error',
            'Resolution': 'Data is syntactically correct but fails business rules. Check product availability at address.'
        },
        {
            'Error Code': 'HTTP_500',
            'Description': 'Internal Server Error - Gamma API issue',
            'Resolution': 'Temporary API problem. Wait and retry. Contact Gamma if persists.'
        },
        {
            'Error Code': 'HTTP_503',
            'Description': 'Service Unavailable - API temporarily down',
            'Resolution': 'Gamma API maintenance or outage. Wait and retry later.'
        },
        {
            'Error Code': 'INVALID_NAD',
            'Description': 'NAD key is invalid or not found',
            'Resolution': 'Run suitability checker to verify NAD key. May need new address lookup.'
        },
        {
            'Error Code': 'PRODUCT_UNAVAILABLE',
            'Description': 'Requested product not available at address',
            'Resolution': 'Run suitability checker to see available products for this address.'
        },
        {
            'Error Code': 'INVALID_ACCOUNT',
            'Description': 'Account number is invalid or inactive',
            'Resolution': 'Verify account number with Gamma. Ensure account is active.'
        },
        {
            'Error Code': 'INVALID_CLI',
            'Description': 'CLI format invalid or not portable',
            'Resolution': 'Verify CLI format (E.164: +44...). Run portability check.'
        },
        {
            'Error Code': 'INVALID_DATE',
            'Description': 'Customer required date is invalid',
            'Resolution': 'Ensure date is in future and format is YYYY-MM-DD. Check minimum lead time.'
        },
        {
            'Error Code': 'SYSTEM_ERROR',
            'Description': 'Unexpected system error in script',
            'Resolution': 'Check Python error message. May indicate data type issue or script bug.'
        },
        {
            'Error Code': 'VALIDATION_ERROR',
            'Description': 'Pre-submission validation failed',
            'Resolution': 'Review validation error message. Fix data in spreadsheet.'
        }
    ])
    
    # Save to Excel with multiple sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name='Order Results', index=False)
        error_legend.to_excel(writer, sheet_name='Error Code Legend', index=False)
        
        # Auto-size columns for readability
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
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    success_count = len([r for r in results if r['Outcome_Status'] == 'SUCCESS'])
    failed_count = len([r for r in results if r['Outcome_Status'] in ['FAILED', 'ERROR']])
    
    print(f"\nTotal orders processed: {len(results)}")
    print(f"✓ Successful: {success_count}")
    print(f"❌ Failed: {failed_count}")
    
    print(f"\n✓ Results saved to: {output_file}")
    print(f"  - Sheet 'Order Results': Complete order data with outcomes")
    print(f"  - Sheet 'Error Code Legend': Reference guide for error codes")
    
    # Create order check input file for successful orders
    if success_count > 0:
        print("\n" + "=" * 80)
        print("CREATING ORDER CHECK INPUT FILE")
        print("=" * 80)
        
        successful_order_ids = [
            r['Outcome_Order_ID'] for r in results 
            if r['Outcome_Status'] == 'SUCCESS' and r['Outcome_Order_ID']
        ]
        
        if successful_order_ids:
            check_df = pd.DataFrame({'Order_ID': successful_order_ids})
            check_df.to_excel(check_file, index=False)
            
            print(f"\n✓ Created order check input file: {check_file}")
            print(f"  - Contains {len(successful_order_ids)} successfully placed order IDs")
            print(f"  - Ready to use with: python bulk_check_orders.py")
    
    if failed_count > 0:
        print("\n" + "=" * 80)
        print("FAILED ORDERS")
        print("=" * 80)
        for result in results:
            if result['Outcome_Status'] in ['FAILED', 'ERROR']:
                error_code = result.get('Outcome_Error_Code', '')
                error_msg = result.get('Outcome_Error_Message', '')
                print(f"  Row {result['Original_Row_Number']}: {result['site_companyName']}")
                if error_code:
                    print(f"    Error Code: {error_code}")
                print(f"    Message: {error_msg}")
    
    print("\n" + "=" * 80)
    print(f"✓ FILE PROCESSING COMPLETE: {input_filename}")
    print("=" * 80)
    
    # Move processed file to processed folder
    try:
        # Create timestamped subfolder in processed directory
        process_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        processed_subfolder = os.path.join(PROCESSED_FOLDER, process_timestamp)
        os.makedirs(processed_subfolder, exist_ok=True)
        
        # Move the file
        destination_path = os.path.join(processed_subfolder, os.path.basename(input_file_path))
        shutil.move(input_file_path, destination_path)
        
        print(f"\n✓ Input file moved to: {destination_path}")
        print("  (File will not be reprocessed on next run)")
        
    except Exception as e:
        print(f"\n⚠ Warning: Could not move input file: {e}")
        print("  File may be reprocessed on next run")
    
    # 
    # Send email with report
    print("\n" + "=" * 80)
    print("SENDING EMAIL REPORT")
    print("=" * 80)
    
    email_sent = send_order_report_email(
        report_file_path=output_file,
        recipient_email='psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk',
        #recipient_email='david.murphy+psmsoutput@gamma.co.uk',
        input_filename=input_filename,
        total_orders=len(results),
        successful=success_count,
        failed=failed_count,
        errors=0,
        report_type='Order Placement',
        subject_prefix='SoGEA Bulk Order Results'
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
    print("SOGEA BULK ORDER PROCESSOR")
    print("=" * 80)
    
    # Get all input files
    input_files = get_input_files()
    
    if not input_files:
        print(f"\n❌ No Excel files found in: {INPUT_FOLDER}")
        print("\nPlease place your order spreadsheets in the input folder.")
        print(f"Input folder location: {INPUT_FOLDER}")
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
    
    # Confirm before processing
    print("\n" + "=" * 80)
    print("READY TO PROCESS FILES")
    print("=" * 80)
    print(f"\nEnvironment: {creds['environment_name']}")
    print(f"Number of files: {len(input_files)}")
    
    """confirm = input("\nProceed with processing all files? (yes/no): ").strip().lower()
    if confirm != 'yes':
        print("\n❌ Cancelled by user")
        sys.exit(0)"""
    
    # Process each file
    print("\n" + "=" * 80)
    print("PROCESSING FILES")
    print("=" * 80)
    
    processed_count = 0
    failed_count = 0
    
    for file_path, filename in input_files:
        print(f"\n\n")
        success = process_single_file(file_path, filename, api)
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
    print(f"✓ Order check files saved to: {CHECK_INPUT_FOLDER}")
    print("\n" + "=" * 80)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Process cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        sys.exit(1)

"""
Phoneline+ Bulk Order Processor
Reads hardware orders from Excel and places orders via API

Supports:
- Multiple products per order
- Delivery address specification
- SKU to user mapping for device assignment
- Bulk processing from OneDrive folder
"""

import pandas as pd
import shutil
import time
import json
from typing import Dict, List, Optional
from datetime import datetime
from pathlib import Path
from phoneline_plus_jwt_auth import PhonelinePlusAuth
from phoneline_plus_place_order import PhonelinePlusOrder
from graph_mailbox_check import send_order_report_email


class PhonelinePlusBulkOrderProcessor:
    """
    Process bulk hardware orders from Excel input
    """
    
    # Delay between order processing cycles (seconds)
    # Allows API time to process the order
    CYCLE_DELAY = 2.0

    @staticmethod
    def _format_phone_number(value) -> str:
        """Normalize phone values from Excel and preserve UK leading zero."""
        if pd.isna(value) or value == '':
            return ''

        phone_str = str(value).replace('.0', '').strip()
        if phone_str.isdigit() and len(phone_str) == 10 and phone_str.startswith(('7', '1', '2', '3', '8')):
            return '0' + phone_str
        return phone_str

    COLUMN_ALIASES = {
        'key id': 'keyID',
        'keyid': 'keyID',
        'secret key': 'secret',
        'customer id': 'customer_id',
        'customerid': 'customer_id',
        'full name': 'name',
        'contact name': 'name',
        'e-mail': 'email',
        'email address': 'email',
        'mail': 'email',
        'tracking email': 'tracking_email',
        'trackingemail': 'tracking_email',
        'devices json': 'devices_json',
        'devices_json': 'devices_json',
        'sku mapping json': 'sku_mapping_json',
        'sku_mapping_json': 'sku_mapping_json',
        'phone': 'phone_number',
        'phone number': 'phone_number',
        'delivery line1': 'delivery_line1',
        'delivery line 1': 'delivery_line1',
        'delivery line2': 'delivery_line2',
        'delivery line 2': 'delivery_line2',
        'delivery line3': 'delivery_line3',
        'delivery line 3': 'delivery_line3',
        'delivery postcode': 'delivery_postcode',
        'postcode': 'delivery_postcode',
        'product sku': 'product_id',
        'sku': 'sku_code',
        'sku code': 'sku_code',
        'product model': 'sku_code',
        'model': 'sku_code',
        'qty': 'quantity',
        'user uuid': 'user_id',
        'userid': 'user_id',
        'user id': 'user_id'
    }
    
    def __init__(self, input_file: str, environment: str = "production"):
        """
        Initialize the bulk order processor
        
        Args:
            input_file: Path to Excel file with orders
            environment: "uat" or "production"
        """
        self.input_file = input_file
        self.environment = environment
        self.auth_token = None
        self.results = []
        self.df = None
    
    def authenticate(self, key_id: str, secret: str) -> bool:
        """
        Authenticate and get JWT token
        
        Args:
            key_id: Partner API Key ID
            secret: Partner API Secret
        
        Returns:
            bool: True if successful
        """
        print("Authenticating...")
        auth_manager = PhonelinePlusAuth(environment=self.environment)
        success, token, error = auth_manager.generate_token(key_id, secret)
        
        if success:
            self.auth_token = token
            print("✓ Authentication successful")
            return True
        else:
            print(f"✗ Authentication failed: {error}")
            return False
    
    def load_input_file(self) -> bool:
        """
        Load and validate the Excel input file
        
        Expected columns:
        - keyID: Partner API Key ID (required, first row only)
        - secret: Partner API Secret (required, first row only)
        - customer_id: Customer UUID (required)
        - name: Delivery contact name (required)
        - email: Delivery contact email (required)
        - phone_number: Delivery contact phone (required)
        - delivery_line1: Delivery address line 1 (required)
        - delivery_line2: Delivery address line 2 (optional)
        - delivery_line3: Delivery address line 3 (optional)
        - delivery_town: Town/City (required)
        - delivery_county: County (optional)
        - delivery_country: Country (optional, defaults to "United Kingdom")
        - delivery_postcode: Postcode (required)
        - product_id: Product SKU UUID (required)
        - sku_code: Product SKU code like "W73P" (required for device assignment)
        - quantity: Number of units (required, defaults to 1)
        - user_id: User UUID for device assignment (optional)
        
        Returns:
            bool: True if loaded successfully
        """
        try:
            print(f"\nLoading input file: {Path(self.input_file).name}")
            print(f"Input file path: {self.input_file}")
            self.df = pd.read_excel(self.input_file)

            # Normalise column names from Excel (trim spaces, normalise case, map aliases)
            normalized_columns = []
            for original_col in self.df.columns:
                if not isinstance(original_col, str):
                    normalized_columns.append(original_col)
                    continue

                cleaned = original_col.strip()
                alias_key = cleaned.lower().replace('_', ' ')
                canonical = self.COLUMN_ALIASES.get(alias_key, cleaned)
                normalized_columns.append(canonical)

            self.df.columns = normalized_columns
            
            # Validate required columns
            required_columns = [
                'keyID', 'secret', 'customer_id', 'name', 'email', 'phone_number',
                'delivery_line1', 'delivery_town', 'delivery_postcode', 'product_id', 'sku_code'
            ]
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                print(f"✗ Missing required columns: {', '.join(missing_columns)}")
                return False
            
            # Remove empty rows
            self.df = self.df.dropna(subset=['customer_id', 'product_id'])
            
            # Default quantity to 1 if not provided
            if 'quantity' not in self.df.columns:
                self.df['quantity'] = 1
            else:
                self.df['quantity'] = self.df['quantity'].fillna(1)
            
            print(f"✓ Loaded {len(self.df)} order(s)")
            return True
            
        except Exception as e:
            print(f"✗ Error loading file: {e}")
            return False
    
    def validate_row(self, row: pd.Series, row_num: int) -> tuple[bool, Optional[str]]:
        """
        Validate a single row
        
        Args:
            row: DataFrame row
            row_num: Row number for error messages
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Check required fields
        if self._get_str(row, 'customer_id') == '':
            return False, f"Row {row_num}: customer_id is required"
        
        if self._get_str(row, 'name') == '':
            return False, f"Row {row_num}: name is required"
        
        if self._get_str(row, 'email') == '':
            return False, f"Row {row_num}: email is required"
        
        if self._get_str(row, 'phone_number') == '':
            return False, f"Row {row_num}: phone_number is required"
        
        if self._get_str(row, 'product_id') == '':
            return False, f"Row {row_num}: product_id is required"
        
        if self._get_str(row, 'sku_code') == '':
            return False, f"Row {row_num}: sku_code is required"
        
        return True, None

    @staticmethod
    def _get_str(row: pd.Series, key: str, default: str = '') -> str:
        """Safely get a string field from a pandas row."""
        value = row.get(key, default)
        if pd.isna(value):
            return ''
        return str(value).strip()

    @staticmethod
    def _get_json_dict(row: pd.Series, key: str) -> Optional[Dict]:
        """Parse optional JSON object from sheet cell."""
        value = row.get(key)
        if value is None or pd.isna(value):
            return None

        if isinstance(value, dict):
            return value

        if isinstance(value, str):
            raw = value.strip()
            if not raw:
                return None
            try:
                parsed = json.loads(raw)
                return parsed if isinstance(parsed, dict) else None
            except json.JSONDecodeError:
                return None

        return None

    @staticmethod
    def _get_json_list(row: pd.Series, key: str) -> Optional[List[Dict]]:
        """Parse optional JSON array from sheet cell."""
        value = row.get(key)
        if value is None or pd.isna(value):
            return None

        if isinstance(value, list):
            return value

        if isinstance(value, str):
            raw = value.strip()
            if not raw:
                return None
            try:
                parsed = json.loads(raw)
                return parsed if isinstance(parsed, list) else None
            except json.JSONDecodeError:
                return None

        return None
    
    def process_orders(self) -> List[Dict]:
        """
        Process all orders in the input file
        
        Returns:
            List of result dictionaries
        """
        if self.df is None or len(self.df) == 0:
            print("No data to process")
            return []
        
        # Get credentials from first row
        first_row = self.df.iloc[0]
        key_id = str(first_row['keyID']).strip()
        secret = str(first_row['secret']).strip()
        
        # Authenticate
        print("\nAuthenticating...")
        if not self.authenticate(key_id, secret):
            print("✗ Authentication failed - cannot process orders")
            return []
        
        order_manager = PhonelinePlusOrder(environment=self.environment, auth_token=self.auth_token)
        
        print("\n" + "=" * 60)
        print("PROCESSING ORDERS")
        print("=" * 60)
        
        for idx, row in self.df.iterrows():
            row_num = idx + 2  # Excel row number (1-indexed + header)
            
            print(f"\nProcessing row {row_num}...")
            
            # Validate row
            is_valid, error_msg = self.validate_row(row, row_num)
            if not is_valid:
                print(f"  ✗ Validation failed: {error_msg}")
                self.results.append({
                    'row': row_num,
                    'timestamp': datetime.now().isoformat(),
                    'customer_id': self._get_str(row, 'customer_id'),
                    'name': self._get_str(row, 'name'),
                    'email': self._get_str(row, 'email'),
                    'phone_number': self._get_str(row, 'phone_number'),
                    'delivery_line1': self._get_str(row, 'delivery_line1'),
                    'delivery_line2': self._get_str(row, 'delivery_line2'),
                    'delivery_line3': self._get_str(row, 'delivery_line3'),
                    'delivery_town': self._get_str(row, 'delivery_town'),
                    'delivery_county': self._get_str(row, 'delivery_county'),
                    'delivery_country': self._get_str(row, 'delivery_country'),
                    'delivery_postcode': self._get_str(row, 'delivery_postcode'),
                    'product_id': self._get_str(row, 'product_id'),
                    'sku_code': self._get_str(row, 'sku_code'),
                    'quantity': row.get('quantity', 1),
                    'user_id': self._get_str(row, 'user_id'),
                    'success': False,
                    'order_number': '',
                    'order_created': '',
                    'product_category': '',
                    'product_make': '',
                    'product_name': '',
                    'product_unit_cost': '',
                    'error': error_msg
                })
                continue
            
            # Get row data
            customer_id = self._get_str(row, 'customer_id')
            name = self._get_str(row, 'name')
            email = self._get_str(row, 'email')
            tracking_email = self._get_str(row, 'tracking_email') or email
            phone_number = self._format_phone_number(row.get('phone_number'))
            product_id = self._get_str(row, 'product_id')
            sku_code = self._get_str(row, 'sku_code')
            quantity = int(row.get('quantity', 1))
            user_id = self._get_str(row, 'user_id') or None
            
            # Build delivery address
            delivery_address = {
                "line1": self._get_str(row, 'delivery_line1'),
                "town": self._get_str(row, 'delivery_town'),
                "postcode": self._get_str(row, 'delivery_postcode')
            }
            
            # Add optional address fields
            if self._get_str(row, 'delivery_line2'):
                delivery_address["line2"] = self._get_str(row, 'delivery_line2')
            if self._get_str(row, 'delivery_line3'):
                delivery_address["line3"] = self._get_str(row, 'delivery_line3')
            if self._get_str(row, 'delivery_county'):
                delivery_address["county"] = self._get_str(row, 'delivery_county')
            if self._get_str(row, 'delivery_country'):
                delivery_address["country"] = self._get_str(row, 'delivery_country')
            
            # Build products list
            products = [
                {
                    "ID": product_id,
                    "quantity": quantity
                }
            ]

            # Build devices list - allow JSON override from sheet
            devices = self._get_json_list(row, 'devices_json')
            if not devices:
                devices = [
                    {
                        "ID": product_id,
                        "quantity": quantity
                    }
                ]
            
            # Build SKU to user mapping if user_id provided
            # Note: API expects SKU code (like "W73P") as key, not product UUID
            sku_to_user_mapping = None
            sku_mapping_override = self._get_json_dict(row, 'sku_mapping_json')
            if sku_mapping_override:
                sku_to_user_mapping = sku_mapping_override
            elif user_id:
                sku_to_user_mapping = {
                    sku_code: [
                        {
                            "userID": user_id
                        }
                    ]
                }
            
            # Place the order
            success, order_data, error = order_manager.place_order(
                customer_id=customer_id,
                name=name,
                email=email,
                phone_number=phone_number,
                delivery_address=delivery_address,
                products=products,
                sku_to_user_mapping=sku_to_user_mapping,
                tracking_email=tracking_email if tracking_email else None,
                devices=devices
            )
            
            result = {
                'row': row_num,
                'timestamp': datetime.now().isoformat(),
                'customer_id': customer_id,
                'name': name,
                'email': email,
                'tracking_email': tracking_email,
                'phone_number': phone_number,
                'delivery_line1': delivery_address.get('line1', ''),
                'delivery_line2': delivery_address.get('line2', ''),
                'delivery_line3': delivery_address.get('line3', ''),
                'delivery_town': delivery_address.get('town', ''),
                'delivery_county': delivery_address.get('county', ''),
                'delivery_country': delivery_address.get('country', 'United Kingdom'),
                'delivery_postcode': delivery_address.get('postcode', ''),
                'product_id': product_id,
                'sku_code': sku_code,
                'quantity': quantity,
                'user_id': user_id if user_id else '',
                'success': success,
                'order_number': order_data.get('orderNumber', '') if order_data else '',
                'order_created': order_data.get('created', '') if order_data else '',
                'product_category': order_data.get('products', [{}])[0].get('category', '') if order_data and order_data.get('products') else '',
                'product_make': order_data.get('products', [{}])[0].get('make', '') if order_data and order_data.get('products') else '',
                'product_name': order_data.get('products', [{}])[0].get('name', '') if order_data and order_data.get('products') else '',
                'product_unit_cost': order_data.get('products', [{}])[0].get('unitCost', '') if order_data and order_data.get('products') else '',
                'error': error if not success else ''
            }
            
            self.results.append(result)
            
            # Delay before next order (allows API to process the order)
            if idx < len(self.df) - 1:  # Don't delay after last order
                print(f"  Waiting {self.CYCLE_DELAY}s before next order...")
                time.sleep(self.CYCLE_DELAY)
        
        return self.results
    
    def generate_report(self, output_file: Optional[str] = None) -> str:
        """
        Generate Excel report of order results
        
        Args:
            output_file: Optional path to save report
            
        Returns:
            Path to generated report file
        """
        if not self.results:
            print("No results to report")
            return None
        
        # Create DataFrame from results with clear input/output field separation
        report_data = []
        for result in self.results:
            flat_result = {
                'row': result['row'],
                'timestamp': result['timestamp'],
                # Input fields (for troubleshooting)
                'input_customer_id': result['customer_id'],
                'input_name': result['name'],
                'input_email': result['email'],
                'input_tracking_email': result.get('tracking_email', ''),
                'input_phone_number': result['phone_number'],
                'input_delivery_line1': result.get('delivery_line1', ''),
                'input_delivery_line2': result.get('delivery_line2', ''),
                'input_delivery_line3': result.get('delivery_line3', ''),
                'input_delivery_town': result.get('delivery_town', ''),
                'input_delivery_county': result.get('delivery_county', ''),
                'input_delivery_country': result.get('delivery_country', ''),
                'input_delivery_postcode': result.get('delivery_postcode', ''),
                'input_product_id': result['product_id'],                'input_sku_code': result.get('sku_code', ''),                'input_quantity': result['quantity'],
                'input_user_id': result.get('user_id', ''),
                # Result fields
                'success': result['success'],
                'order_number': result.get('order_number', ''),
                'order_created': result.get('order_created', ''),
                'product_category': result.get('product_category', ''),
                'product_make': result.get('product_make', ''),
                'product_name': result.get('product_name', ''),
                'product_unit_cost': result.get('product_unit_cost', ''),
                'error': result.get('error', '')
            }
            report_data.append(flat_result)
        
        report_df = pd.DataFrame(report_data)
        
        # Generate default output filename if not provided
        if not output_file:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f"phoneline_plus_order_results_{timestamp}.xlsx"
        
        # Calculate summary statistics
        total = len(self.results)
        successful = sum(1 for r in self.results if r['success'])
        failed = total - successful
        total_quantity = sum(r.get('quantity', 1) for r in self.results)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            report_df.to_excel(writer, sheet_name='Order Results', index=False)
            
            # Add summary sheet
            summary_data = {
                'Metric': [
                    'Total Orders Processed',
                    'Successful',
                    'Failed',
                    'Total Devices Ordered',
                    'Environment',
                    'Processing Date/Time'
                ],
                'Value': [
                    total,
                    successful,
                    failed,
                    total_quantity,
                    self.environment.upper(),
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
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Print summary
        print("\n" + "=" * 60)
        print("PROCESSING SUMMARY")
        print("=" * 60)
        print(f"Total processed: {total}")
        print(f"Successful: {successful}")
        print(f"Failed: {failed}")
        print(f"Total devices ordered: {total_quantity}")
        print(f"\nReport saved to: {output_file}")
        
        if failed > 0:
            print("\nFailed orders:")
            for result in self.results:
                if not result['success']:
                    print(f"  Row {result['row']}: {result['error']}")
        
        return output_file


def main():
    """
    Scan input folder and process all Excel files
    """
    # Configuration
    SCRIPT_DIR = Path(__file__).parent
    INPUT_FOLDER = r"\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\Phoneline+ Device Assignments"
    OUTPUT_FOLDER = SCRIPT_DIR / "output"
    PROCESSED_FOLDER = Path(INPUT_FOLDER) / "processed"
    ENVIRONMENT = "production"  # or "production"
    
    # Ensure folders exist
    OUTPUT_FOLDER.mkdir(exist_ok=True)
    
    # Check if input folder exists, create if needed
    if not Path(INPUT_FOLDER).exists():
        print(f"⚠ Warning: Input folder does not exist: {INPUT_FOLDER}")
        print("  Please create this folder or update the INPUT_FOLDER path in the script.")
        return
    
    # Create processed folder
    try:
        PROCESSED_FOLDER.mkdir(exist_ok=True)
    except Exception as e:
        print(f"⚠ Warning: Could not create processed folder: {e}")
        print("  Files will not be moved after processing.")
        PROCESSED_FOLDER = None
    
    print("=" * 60)
    print("Phoneline+ Bulk Order Processor")
    print("=" * 60)
    print(f"Environment: {ENVIRONMENT.upper()}")
    print(f"Input Folder: {INPUT_FOLDER}")
    print()
    
    # Find all Excel files in input folder
    import glob
    excel_files = glob.glob(str(Path(INPUT_FOLDER) / "*.xlsx"))
    excel_files = [f for f in excel_files if not Path(f).name.startswith('~$')]  # Exclude temp files
    
    if not excel_files:
        print("✗ No Excel files found in input folder")
        return
    
    print(f"✓ Found {len(excel_files)} file(s) to process")
    print()
    
    # Process each file
    for input_file in excel_files:
        filename = Path(input_file).name
        print(f"{'='*60}")
        print(f"Processing: {filename}")
        print(f"File path: {input_file}")
        print(f"{'='*60}")
        
        # Create processor
        processor = PhonelinePlusBulkOrderProcessor(str(input_file), environment=ENVIRONMENT)
        
        # Load input file (authentication happens inside process_orders)
        if not processor.load_input_file():
            print("Skipping file due to load failure")
            print()
            continue
        
        # Process all orders
        results = processor.process_orders()
        
        # Generate report
        output_file = processor.generate_report()
        print(f"\n✓ Results saved to: {output_file}")
        
        # Send email with report
        print("\n" + "=" * 60)
        print("SENDING EMAIL REPORT")
        print("=" * 60)
        
        # Calculate summary statistics for email
        total = len(processor.results)
        successful = sum(1 for r in processor.results if r['success'])
        failed = total - successful
        
        email_sent = send_order_report_email(
            report_file_path=output_file,
            recipient_email='psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk',
            #recipient_email='david.murphy+plp-orders@gamma.co.uk',
            input_filename=filename,
            total_orders=total,
            successful=successful,
            failed=failed,
            errors=0,
            report_type='Hardware Order',
            subject_prefix='Phoneline+ Hardware Order Results'
        )
        
        if email_sent:
            print("✓ Report successfully emailed to psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk")
        else:
            print("⚠ Warning: Failed to send email report")
        
        # Move processed file to processed folder
        if PROCESSED_FOLDER:
            try:
                process_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                processed_subfolder = PROCESSED_FOLDER / process_timestamp
                processed_subfolder.mkdir(exist_ok=True)
                
                destination_path = processed_subfolder / filename
                shutil.move(input_file, str(destination_path))
                
                print(f"\n✓ Input file moved to: {destination_path}")
                print("  (File will not be reprocessed on next run)")
            except Exception as e:
                print(f"\n⚠ Warning: Could not move input file: {e}")
                print("  File may be reprocessed on next run")
        else:
            print("\n⚠ Warning: Processed folder not available")
            print("  File was not moved and may be reprocessed on next run")
        
        print()


if __name__ == "__main__":
    main()

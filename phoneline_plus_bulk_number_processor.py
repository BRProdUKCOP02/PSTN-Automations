"""
Phoneline+ Bulk Number Processor
Reallocates geographical numbers to users and assigns numbers to users without numbers

Process:
1. Get all numbers and users for customer
2. Identify users with assigned numbers
3. Identify numbers available on the account
4. Swap assigned user numbers with available account numbers
5. Prioritize users with non-geographical numbers when available numbers are limited
6. Deallocate replaced numbers after successful reassignment
7. Generate report with results
"""

import pandas as pd
import shutil
import time
from typing import Dict, List, Optional
from datetime import datetime
from pathlib import Path
import glob
from phoneline_plus_jwt_auth import PhonelinePlusAuth
from phoneline_plus_number_management import PhonelinePlusNumberManager
from graph_mailbox_check import send_order_report_email


class PhonelinePlusBulkNumberProcessor:
    """
    Process bulk number reallocation from Excel input
    """
    
    # Delay between customer processing cycles (seconds)
    # Allows API time to prepare next available number
    CYCLE_DELAY = 2.0
    
    def __init__(self, input_file: str, environment: str = "uat"):
        """
        Initialize the bulk processor
        
        Args:
            input_file: Path to Excel file with customer IDs
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
            Success boolean
        """
        print("Authenticating...")
        auth = PhonelinePlusAuth(environment=self.environment)
        success, token, error = auth.generate_token(key_id, secret)
        
        if success:
            self.auth_token = token
            print("✓ Authentication successful\n")
            return True
        else:
            print(f"✗ Authentication failed: {error}\n")
            return False
    
    def load_input_file(self) -> bool:
        """
        Load and validate the input Excel file
        
        Returns:
            Success boolean
        """
        try:
            print(f"Loading input file: {Path(self.input_file).name}")
            
            # Read Excel file
            self.df = pd.read_excel(self.input_file)
            
            # Required columns
            required_columns = ['keyID', 'secret', 'customer_id']
            
            # Check for missing columns
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                print(f"✗ Missing required columns: {', '.join(missing_columns)}")
                return False
            
            print(f"✓ Loaded {len(self.df)} customer(s) to process\n")
            return True
            
        except Exception as e:
            print(f"✗ Error loading file: {e}")
            return False
    
    def process_customer(self, row: pd.Series, row_num: int) -> Dict:
        """
        Process number reallocation for one customer
        
        Args:
            row: DataFrame row with customer data
            row_num: Row number for reporting
            
        Returns:
            Result dictionary
        """
        result = {
            'customer_id': str(row['customer_id']).strip(),
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'success': False,
            'total_numbers': 0,
            'geo_numbers': 0,
            'nongeo_numbers': 0,
            'total_users': 0,
            'users_with_numbers': 0,
            'users_without_numbers': 0,
            'reallocations': 0,
            'new_allocations': 0,
            'deallocations': 0,
            'failed_operations': 0,
            'error': None,
            'details': []
        }
        
        customer_id = result['customer_id']
        
        print(f"[Row {row_num}] Processing customer: {customer_id}")
        
        # Authenticate if needed
        if not self.auth_token:
            key_id = str(row['keyID']).strip()
            secret = str(row['secret']).strip()
            if not self.authenticate(key_id, secret):
                result['error'] = "Authentication failed"
                return result
        
        # Initialize number manager
        number_manager = PhonelinePlusNumberManager(
            environment=self.environment,
            auth_token=self.auth_token
        )
        
        # Get all numbers for customer
        success, numbers, error = number_manager.get_customer_numbers(customer_id)
        
        if not success:
            result['error'] = f"Failed to fetch numbers: {error}"
            print(f"  ✗ {result['error']}\n")
            return result
        
        # Get all users for customer
        success_users, users, error_users = number_manager.get_customer_users(customer_id)
        
        if not success_users:
            result['error'] = f"Failed to fetch users: {error_users}"
            print(f"  ✗ {result['error']}\n")
            return result
        
        # Analyze numbers
        result['total_numbers'] = len(numbers)
        result['total_users'] = len(users)
        
        geo_numbers = [n for n in numbers if n.get('type') == 'standard_geographic']
        nongeo_numbers = [n for n in numbers if n.get('type') == 'standard_nongeographic']
        
        result['geo_numbers'] = len(geo_numbers)
        result['nongeo_numbers'] = len(nongeo_numbers)
        
        print(f"  Number breakdown:")
        print(f"    Total: {len(numbers)}")
        print(f"    Geographical: {len(geo_numbers)}")
        print(f"    Non-geographical: {len(nongeo_numbers)}")
        print(f"  User breakdown:")
        print(f"    Total users: {len(users)}")
        
        # Build map of which users have which numbers
        user_numbers = {}  # user_id -> list of numbers
        for num in numbers:
            if 'allocatedTo' in num and 'userID' in num['allocatedTo']:
                user_id = num['allocatedTo']['userID']
                if user_id not in user_numbers:
                    user_numbers[user_id] = []
                user_numbers[user_id].append(num)
        
        result['users_with_numbers'] = len(user_numbers)
        result['users_without_numbers'] = len(users) - len(user_numbers)
        
        print(f"    Users with numbers: {len(user_numbers)}")
        print(f"    Users without numbers: {len(users) - len(user_numbers)}")
        
        # Build users that currently have assigned numbers and need swap
        users_with_assigned_numbers = []
        users_with_nongeo_count = 0

        for user in users:
            user_id = user.get('ID')
            if not user_id or user_id not in user_numbers:
                continue

            current_numbers = user_numbers[user_id]
            if not current_numbers:
                continue

            current_number = current_numbers[0]
            current_type = current_number.get('type', '')
            is_nongeo = current_type == 'standard_nongeographic'

            if is_nongeo:
                users_with_nongeo_count += 1

            users_with_assigned_numbers.append({
                'user_id': user_id,
                'user_name': user.get('name', 'Unknown'),
                'old_number_id': current_number.get('ID', ''),
                'old_number_e164': current_number.get('numberE164', ''),
                'old_type': current_type,
                'old_display_type': current_number.get('displayNumberType', 'Unknown'),
                'priority': 0 if is_nongeo else 1
            })

        print(f"    Users with non-geo numbers: {users_with_nongeo_count}")
        print(f"    Users with assigned numbers to swap: {len(users_with_assigned_numbers)}")

        if not users_with_assigned_numbers:
            result['success'] = True
            result['error'] = "No operations needed - no assigned user numbers to swap"
            print(f"  ℹ {result['error']}\n")
            return result

        # Find available numbers on account (prioritize geographical first)
        available_numbers = [
            n for n in numbers
            if 'allocatedTo' not in n or 'userID' not in n.get('allocatedTo', {})
        ]

        available_geo_numbers = [n for n in available_numbers if n.get('type') == 'standard_geographic']
        available_other_numbers = [n for n in available_numbers if n.get('type') != 'standard_geographic']
        available_numbers_ordered = available_geo_numbers + available_other_numbers

        print(f"    Available numbers on account: {len(available_numbers_ordered)}")

        if not available_numbers_ordered:
            result['error'] = "No available numbers on account to use for swapping"
            print(f"  ✗ {result['error']}\n")
            return result

        # Prioritize non-geographical users first when numbers are limited
        users_with_assigned_numbers.sort(key=lambda user_info: user_info['priority'])

        swaps_possible = min(len(users_with_assigned_numbers), len(available_numbers_ordered))
        print(f"    Planned swaps: {swaps_possible} of {len(users_with_assigned_numbers)}")

        if swaps_possible < len(users_with_assigned_numbers):
            result['error'] = (
                f"Only {len(available_numbers_ordered)} available numbers for "
                f"{len(users_with_assigned_numbers)} users. Prioritized non-geographical users first."
            )
            print(f"  ⚠ {result['error']}")

        print(f"\n  Processing number swaps (assigned → available)...")
        for swap_index in range(swaps_possible):
            user_info = users_with_assigned_numbers[swap_index]
            target_number = available_numbers_ordered[swap_index]

            user_id = user_info['user_id']
            old_number_e164 = user_info['old_number_e164']
            new_number_e164 = target_number.get('numberE164', '')

            operation_detail = {
                'operation_type': 'swap',
                'user_id': user_id,
                'old_number': old_number_e164,
                'old_type': user_info['old_display_type'],
                'new_number': new_number_e164,
                'new_type': target_number.get('displayNumberType', 'Unknown'),
                'allocation_success': False,
                'deallocation_success': False,
                'error': None
            }

            print(f"\n    User {user_id} ({user_info['user_name']}):")
            print(f"      Replacing: {old_number_e164} ({user_info['old_display_type']})")
            print(f"      With: {new_number_e164} ({target_number.get('displayNumberType', 'Unknown')})")

            # Allocate available number to user first
            alloc_success, alloc_data, alloc_error = number_manager.allocate_number(
                customer_id, target_number['ID'], user_id
            )

            operation_detail['allocation_success'] = alloc_success

            if not alloc_success:
                operation_detail['error'] = f"Allocation failed: {alloc_error}"
                result['failed_operations'] += 1
                result['details'].append(operation_detail)
                continue

            result['reallocations'] += 1

            # Deallocate previous assigned number
            if user_info['old_number_id']:
                dealloc_success, dealloc_data, dealloc_error = number_manager.deallocate_number(
                    customer_id, user_info['old_number_id']
                )
            else:
                dealloc_success, dealloc_data, dealloc_error = False, None, "Missing old number ID"

            operation_detail['deallocation_success'] = dealloc_success

            if not dealloc_success:
                operation_detail['error'] = f"Deallocation failed: {dealloc_error}"
                result['failed_operations'] += 1
            else:
                result['deallocations'] += 1

            result['details'].append(operation_detail)
        
        # Mark as success if any operations completed
        if result['reallocations'] > 0:
            result['success'] = True
        
        print(f"\n  Summary:")
        print(f"    Swaps completed: {result['reallocations']}")
        print(f"    Deallocations: {result['deallocations']}")
        print(f"    Failed operations: {result['failed_operations']}")
        print()
        
        return result
    
    def process_customers(self):
        """
        Process all customers from input file
        """
        if self.df is None:
            print("✗ No data loaded")
            return
        
        print("=" * 60)
        print("PROCESSING NUMBER ASSIGNMENTS")
        print("=" * 60)
        print()
        
        for index, row in self.df.iterrows():
            row_num = index + 2  # +2 for Excel 1-indexed + header
            result = self.process_customer(row, row_num)
            self.results.append(result)
            
            # Delay before next customer (allows API to prepare next available number)
            if index < len(self.df) - 1:  # Don't delay after last customer
                print(f"  Waiting {self.CYCLE_DELAY}s before next customer...")
                time.sleep(self.CYCLE_DELAY)
    
    def generate_report(self) -> str:
        """
        Generate Excel report of results
        
        Returns:
            Path to generated report file
        """
        if not self.results:
            print("No results to report")
            return None
        
        # Create combined results with operation details
        combined_data = []
        for result in self.results:
            # If there are operation details, create one row per operation
            if result['details']:
                for detail in result['details']:
                    combined_data.append({
                        'timestamp': result['timestamp'],
                        'input_customer_id': result['customer_id'],
                        'success': result['success'],
                        'total_numbers': result['total_numbers'],
                        'geo_numbers': result['geo_numbers'],
                        'nongeo_numbers': result['nongeo_numbers'],
                        'total_users': result['total_users'],
                        'users_with_numbers': result['users_with_numbers'],
                        'users_without_numbers': result['users_without_numbers'],
                        'reallocations': result['reallocations'],
                        'new_allocations': result['new_allocations'],
                        'deallocations': result['deallocations'],
                        'failed_operations': result['failed_operations'],
                        'customer_error': result.get('error', ''),
                        # Operation details
                        'operation_type': detail.get('operation_type', ''),
                        'user_id': detail['user_id'],
                        'old_number': detail['old_number'],
                        'old_type': detail['old_type'],
                        'new_number': detail['new_number'],
                        'new_type': detail['new_type'],
                        'allocation_success': detail['allocation_success'],
                        'deallocation_success': detail['deallocation_success'],
                        'operation_error': detail.get('error', '')
                    })
            else:
                # If no operation details, create a summary row
                combined_data.append({
                    'timestamp': result['timestamp'],
                    'input_customer_id': result['customer_id'],
                    'success': result['success'],
                    'total_numbers': result['total_numbers'],
                    'geo_numbers': result['geo_numbers'],
                    'nongeo_numbers': result['nongeo_numbers'],
                    'total_users': result['total_users'],
                    'users_with_numbers': result['users_with_numbers'],
                    'users_without_numbers': result['users_without_numbers'],
                    'reallocations': result['reallocations'],
                    'new_allocations': result['new_allocations'],
                    'deallocations': result['deallocations'],
                    'failed_operations': result['failed_operations'],
                    'customer_error': result.get('error', ''),
                    # Empty operation details
                    'operation_type': '',
                    'user_id': '',
                    'old_number': '',
                    'old_type': '',
                    'new_number': '',
                    'new_type': '',
                    'allocation_success': '',
                    'deallocation_success': '',
                    'operation_error': ''
                })
        
        combined_df = pd.DataFrame(combined_data)
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f"phoneline_plus_number_results_{timestamp}.xlsx"
        
        # Calculate summary
        total = len(self.results)
        successful = sum(1 for r in self.results if r['success'])
        failed = total - successful
        total_reallocations = sum(r['reallocations'] for r in self.results)
        total_new_allocations = sum(r['new_allocations'] for r in self.results)
        total_deallocations = sum(r['deallocations'] for r in self.results)
        
        # Save to Excel - Summary first, then Results
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Summary sheet first
            summary_data = {
                'Metric': [
                    'Total Customers Processed',
                    'Successful',
                    'Failed',
                    'Total Reallocations (Non-geo → Geo)',
                    'Total New Allocations (No Number → Geo)',
                    'Total Deallocations',
                    'Environment',
                    'Processing Date/Time'
                ],
                'Value': [
                    total,
                    successful,
                    failed,
                    total_reallocations,
                    total_new_allocations,
                    total_deallocations,
                    self.environment.upper(),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Combined results sheet
            combined_df.to_excel(writer, sheet_name='Results', index=False)
            
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
        
        print("\n" + "=" * 60)
        print("PROCESSING SUMMARY")
        print("=" * 60)
        print(f"Total processed: {total}")
        print(f"Successful: {successful}")
        print(f"Failed: {failed}")
        print(f"Total reallocations: {total_reallocations}")
        print(f"Total new allocations: {total_new_allocations}")
        print(f"Total deallocations: {total_deallocations}")
        print(f"\nReport saved to: {output_file}")
        
        if failed > 0:
            print("\nFailed customers:")
            for result in self.results:
                if not result['success']:
                    print(f"  {result['customer_id']} - {result['error']}")
        
        return output_file


def main():
    """
    Main entry point - scan input folder and process all Excel files
    """
    # Configuration
    INPUT_FOLDER = r"\\localhost\c$\Users\dmurphy\OneDrive - Gamma Telecom Ltd\Automation Solutions - Power Automate\PSTN Switch off\Phoneline+ Number Assignment & Removal"
    PROCESSED_FOLDER = Path(INPUT_FOLDER) / "processed"
    ENVIRONMENT = "production"  # or "production"
    
    # Ensure processed folder exists
    PROCESSED_FOLDER.mkdir(exist_ok=True)
    
    print("=" * 60)
    print("Phoneline+ Bulk Number Processor")
    print("=" * 60)
    print(f"Environment: {ENVIRONMENT.upper()}")
    print(f"Input Folder: {INPUT_FOLDER}")
    print()
    
    # Find all Excel files
    excel_files = glob.glob(str(Path(INPUT_FOLDER) / "*.xlsx"))
    excel_files = [f for f in excel_files if not Path(f).name.startswith('~$')]
    
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
        print(f"{'='*60}")
        print()
        
        # Create processor
        processor = PhonelinePlusBulkNumberProcessor(str(input_file), environment=ENVIRONMENT)
        
        # Load input file
        if not processor.load_input_file():
            continue
        
        # Process all customers
        processor.process_customers()
        
        # Generate report
        output_file = processor.generate_report()
        
        if output_file:
            print(f"\n✓ Results saved to: {output_file}")
            
            # Send email
            print("\n" + "=" * 60)
            print("SENDING EMAIL REPORT")
            print("=" * 60)
            
            total = len(processor.results)
            successful = sum(1 for r in processor.results if r['success'])
            failed = total - successful
            
            email_sent = send_order_report_email(
                report_file_path=output_file,
                recipient_email='psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk',
                #recipient_email='david.murphy+plp-numbers@gamma.co.uk',
                input_filename=filename,
                total_orders=total,
                successful=successful,
                failed=failed,
                errors=0,
                report_type='Number Reallocation',
                subject_prefix='Phoneline+ Number Reallocation Results'
            )
            
            if email_sent:
                print("✓ Report successfully emailed to psmanaged.delivery@gamma.co.uk, Dave.Siddall@gamma.co.uk")
            else:
                print("⚠ Warning: Failed to send email report")
            
            # Move processed file to processed folder
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
        
        print()


if __name__ == "__main__":
    main()

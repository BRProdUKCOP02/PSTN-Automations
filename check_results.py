import pandas as pd

# Read the most recent results file
file = r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_results_20260219_173754.xlsx'

# Read customer results
df = pd.read_excel(file, sheet_name='Customer Results')

print('=' * 60)
print('MOST RECENT TEST RESULTS (17:37:54)')
print('=' * 60)
print(f'Total rows in results: {len(df)}')
print()

# Show all rows with their key details
for idx, row in df.iterrows():
    print(f"Row {row['row']}: {row['input_companyName']}")
    print(f"  Customer ID: {row['customer_id']}")
    print(f"  Success: {row['success']}")
    print(f"  Plan: {row['input_plan']}")
    if pd.notna(row['error']) and str(row['error']).strip():
        print(f"  Error: {row['error']}")
    print()

# Count customers vs users
customer_rows = df[df['input_plan'] != 'USER']
user_rows = df[df['input_plan'] == 'USER']

print(f'Customers created: {len(customer_rows)}')
print(f'Users created: {len(user_rows)}')
print()

# Check for duplicate customer IDs
customer_ids = customer_rows['customer_id'].tolist()
print(f'DUPLICATE CHECK:')
print(f'Customer IDs: {customer_ids}')
if len(customer_ids) != len(set(customer_ids)):
    print('⚠ WARNING: DUPLICATE CUSTOMER IDs FOUND!')
    from collections import Counter
    duplicates = [id for id, count in Counter(customer_ids).items() if count > 1]
    print(f'Duplicates: {duplicates}')

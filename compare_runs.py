import pandas as pd

df1 = pd.read_excel(r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_results_20260219_173729.xlsx')
df2 = pd.read_excel(r'c:\Users\Public\RPA\code\PSTN Migration\phoneline_plus_results_20260219_173754.xlsx')

print('=' * 60)
print('COMPARING TWO CONSECUTIVE RUNS (25 seconds apart)')
print('=' * 60)
print()
print('FILE 1 (17:37:29):')
print(f"  Total rows: {len(df1)}")
print(f"  Company: {df1.iloc[0]['input_companyName']}")
print(f"  Customer ID: {df1.iloc[0]['customer_id']}")
print()

print('FILE 2 (17:37:54):')
print(f"  Total rows: {len(df2)}")
print(f"  Company: {df2.iloc[0]['input_companyName']}")
print(f"  Customer ID: {df2.iloc[0]['customer_id']}")
print()

if df1.iloc[0]['customer_id'] == df2.iloc[0]['customer_id']:
    print('✓ SAME CUSTOMER ID - Duplicate detection working!')
else:
    print('✗ DIFFERENT CUSTOMER IDs - This is the duplicate creation bug!')
    print()
    print('This means the script was invoked TWICE by automation,')
    print('and each time it created a NEW customer from the same Excel row.')

#!/usr/bin/env python3
"""
Debug script to reproduce the Excel handler error
"""
import pandas as pd
import numpy as np

# Simulate the problematic data
print("Testing the error scenario...")

# Create a test DataFrame with mixed types in email column (like the real Excel data)
data = {
    'col1': ['A', 'B', 'C'],
    'col2': [1, 2, 3],
    'col3': ['Name1', 'Name2', 'Name3'],
    'col4': ['First1', 'First2', 'First3'], 
    'col5': ['Last1', 'Last2', 'Last3'],
    'col6': ['Dept1', 'Dept2', 'Dept3'],
    'col7': ['Data1', 'Data2', 'Data3'],
    'col8': ['email1@test.com;manager1@test.com', np.nan, None]  # This is the problematic column
}

df = pd.DataFrame(data)
print("DataFrame:")
print(df)
print("\nColumn 8 (email_col) data types:")
print(df.iloc[:, 7].dtype)
print("Values:", df.iloc[:, 7].values)

# Try to reproduce the error
try:
    email_col = df.iloc[:, 7]  # Column H (8th column) - this is column index 7
    print(f"\nTrying .astype(str)...")
    email_col_str = email_col.astype(str)
    print("Success! String values:", email_col_str.values)
    
    print(f"\nTrying .str operations on astype(str)...")
    result = email_col.astype(str).str.split(';').str[0].str.strip()
    print("Success! Split result:", result.values)
    
except Exception as e:
    print(f"ERROR: {e}")

# Test the fix
print("\n" + "="*50)
print("TESTING THE FIX:")

try:
    email_col = df.iloc[:, 7]
    
    # Proper way: handle NaN/None values first, then convert to string
    # Replace NaN/None with empty string, then convert to string
    email_col_clean = email_col.fillna('').astype(str)
    print("Clean email column:", email_col_clean.values)
    
    cardholder_emails = email_col_clean.str.split(';').str[0].str.strip()
    manager_emails = email_col_clean.str.split(';').str[1].str.strip()
    
    print("Cardholder emails:", cardholder_emails.values)
    print("Manager emails:", manager_emails.values)
    
except Exception as e:
    print(f"ERROR in fix: {e}")
#!/usr/bin/env python3
"""
Test script to validate the Excel handler fixes
"""
import sys
import os
import pandas as pd
import numpy as np
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def create_test_excel_data():
    """Create test data that mimics the problematic Excel structure"""
    
    # Simulate the Excel file structure with various data quality issues
    data = {
        # Column 1-4: Various metadata
        'Card-Number-(From-Sheet-1-with-formula)': ['12345', '67890', '', None, 'TEMP123'],
        'Card-Number-(Without-formula)': ['12345', '67890', '', None, 'TEMP123'],
        'Card-Number-Last-5-Digits': ['*2345', '*7890', '', None, '*P123'],
        'Title': ['Mr', 'Ms', None, '', 'Dr'],
        
        # Column 5-6: Names (positional columns 4-5 in 0-indexed)
        'First-Name': ['John', 'Jane', '', None, 'Bob-Smith'],
        'Last-Name': ['Doe', 'Smith', 'NoFirst', None, ''],
        
        # Column 7: Location
        'Location': ['Office1', 'Office2', '', None, 'Home'],
        
        # Column 8: Email data (this is the problematic column - 0-indexed position 7)
        'E-mail-Address': [
            'john.doe@company.com;manager1@company.com',  # Valid email pair
            'jane.smith@company.com',  # Single email
            '',  # Empty string
            None,  # None value
            np.nan  # NaN value
        ],
        
        # Additional columns
        'Line-Manager': ['Manager1', 'Manager2', '', None, 'Manager3'],
        'CARD-STATUS': ['Active', 'Active', 'Inactive', None, 'Active'],
        
        # Statement columns
        'June-2024---Statement-Amount': [100.0, 200.0, 0.0, None, 150.0],
        'July-2024---Statement-Amount': [150.0, None, 0.0, 50.0, np.nan]
    }
    
    return pd.DataFrame(data)

def test_excel_handler():
    """Test the Excel handler with the fixed code"""
    
    # Import after setting path
    from src.modules.excel_handler import ExcelHandler
    from src.core.config import DashboardConfig
    
    print("Creating test configuration...")
    config = DashboardConfig()
    
    print("Creating Excel handler...")
    handler = ExcelHandler(config)
    
    print("Creating test data...")
    test_df = create_test_excel_data()
    
    print("\nOriginal test data:")
    print(test_df.head())
    print(f"\nData types:\n{test_df.dtypes}")
    print(f"\nColumn 8 (email) values: {test_df.iloc[:, 7].values}")
    
    print("\n" + "="*60)
    print("TESTING FIXED _clean_cardholder_data method...")
    
    try:
        cleaned_df = handler._clean_cardholder_data(test_df)
        
        print(f"\nCleaned data shape: {cleaned_df.shape}")
        print("\nCleaned data columns:", list(cleaned_df.columns))
        
        if len(cleaned_df) > 0:
            print("\nFirst few rows of cleaned data:")
            # Only show relevant columns
            cols_to_show = ['name', 'email', 'manager_email', 'cardholder_email', 'card_number']
            available_cols = [col for col in cols_to_show if col in cleaned_df.columns]
            print(cleaned_df[available_cols].head())
        else:
            print("WARNING: No records in cleaned data!")
            
        print(f"\nSuccess! Processed {len(cleaned_df)} records without errors.")
        return True
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_excel_handler()
    if success:
        print("\n✅ All tests passed!")
    else:
        print("\n❌ Tests failed!")
        sys.exit(1)
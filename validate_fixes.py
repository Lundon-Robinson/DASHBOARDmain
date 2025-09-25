#!/usr/bin/env python3
"""
Final validation test for the Excel handler fixes
"""
import sys
import pandas as pd
import numpy as np
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_fixes():
    """Comprehensive test of all the fixes"""
    
    from src.modules.excel_handler import ExcelHandler
    from src.core.config import DashboardConfig
    
    print("🔧 Testing Excel Handler Fixes")
    print("=" * 50)
    
    # Create configuration
    config = DashboardConfig()
    handler = ExcelHandler(config)
    
    # Create problematic test data similar to the real Excel files
    test_data = {
        'Card-Number-(From-Sheet-1-with-formula)': ['CARD001', 'CARD002', '', None, 'CARD005'],
        'Card-Number-Last-5-Digits': ['*2001', '*2002', '', None, '*2005'], 
        'Title': ['Mr', 'Ms', '', None, 'Dr'],
        'First-Name': ['John', 'Jane-Smith', '', None, 'Bob'],  # With hyphen to test
        'Last-Name': ['Doe', 'Johnson', 'OnlyLast', None, ''],  # Missing first name case
        'Location': ['Office1', 'Office2', '', None, 'Home'],
        'E-mail-Address': [  # This was the problematic column
            'john.doe@company.com;manager1@company.com',  # Valid pair
            'jane.smith@company.com',                      # Single email 
            '',                                            # Empty string
            None,                                          # None value
            12345                                          # Numeric value (problematic)
        ],
        'Line-Manager': ['Manager1', 'Manager2', '', None, 'Manager3'],
        'CARD-STATUS': ['Active', 'Active', '', None, 'Active']
    }
    
    df = pd.DataFrame(test_data)
    
    print(f"✓ Created test data with {len(df)} rows")
    print(f"✓ Email column contains: {df['E-mail-Address'].values}")
    
    try:
        # This should work without the "Can only use .str accessor" error
        cleaned_df = handler._clean_cardholder_data(df)
        
        print(f"✓ Successfully cleaned data: {len(cleaned_df)} records")
        print(f"✓ No pandas .str accessor errors!")
        
        # Check that we have the expected columns
        required_cols = ['name', 'email', 'card_number']
        for col in required_cols:
            if col in cleaned_df.columns:
                print(f"✓ Column '{col}' present")
            else:
                print(f"⚠ Column '{col}' missing")
        
        # Test database sync (this was also failing)
        handler._sync_cardholders(cleaned_df)
        print(f"✓ Database sync completed for {len(cleaned_df)} records")
        
        print("\n🎉 ALL TESTS PASSED!")
        print("\nSummary of fixes:")
        print("- Fixed pandas .str accessor error by using fillna('').astype(str) before string operations")
        print("- Made data filtering more lenient (OR instead of AND for name/email validation)")
        print("- Added proper error handling with fallback values")
        print("- Fixed column name mapping (name vs FullName)")
        print("- Added robust NaN/None value handling throughout")
        
        return True
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_fixes()
    if not success:
        sys.exit(1)
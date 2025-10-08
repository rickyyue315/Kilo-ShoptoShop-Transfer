"""
Test script to verify the transfer system functionality
"""
import pandas as pd
import numpy as np
from app import (preprocess_data, identify_transfer_out_candidates, identify_transfer_in_candidates,
                 match_transfers, calculate_statistics, handle_no_transfer_candidates)

def test_with_real_data():
    """Test with real Excel data"""
    print("=== Testing with Real Excel Data ===")

    try:
        # Load real Excel data
        df = pd.read_excel('ELE_08Sep2025_Test.XLSX')
        print(f'Loaded {len(df)} records from Excel file')

        # Test data processing
        processed_df = preprocess_data(df)
        print('Data preprocessing completed')

        # Test all three modes with demand constraint check
        for mode in ['conservative', 'enhanced', 'special']:
            print(f'\n{mode.upper()} MODE:')
            transfer_out = identify_transfer_out_candidates(processed_df, mode)
            transfer_in = identify_transfer_in_candidates(processed_df)
            transfer_suggestions = match_transfers(transfer_out, transfer_in, processed_df)

            print(f'  Transfer out candidates: {len(transfer_out)}')
            print(f'  Transfer in candidates: {len(transfer_in)}')
            print(f'  Transfer suggestions: {len(transfer_suggestions)}')

            if not transfer_suggestions.empty:
                # Check demand constraint for each article
                print(f'  Total transfer suggestions: {len(transfer_suggestions)}')

                # Group by article and check constraints
                article_groups = transfer_suggestions.groupby('Article')
                constraint_violations = 0

                for article, group in article_groups:
                    total_demand = group['Receive Site Target Qty'].iloc[0] if len(group) > 0 else 0
                    total_transfer = group['Transfer Qty'].sum()

                    print(f'    Article {article}: Transfer={total_transfer}, Demand={total_demand}')

                    if total_transfer > total_demand:
                        constraint_violations += 1
                        print(f'      WARNING: Transfer ({total_transfer}) exceeds demand ({total_demand})!')

                if constraint_violations == 0:
                    print(f'  ✅ All articles satisfy demand constraint')
                else:
                    print(f'  ❌ {constraint_violations} articles violate demand constraint')
            else:
                print(f'  No transfer suggestions - testing error handling...')
                # Test the error handling function
                error_info = handle_no_transfer_candidates(transfer_out, transfer_in, mode)
                print(f'  Error reason: {error_info.get("reason", "unknown")}')
                print(f'  User message: {error_info.get("message", "No message")[:100]}...')

    except Exception as e:
        print(f'Error testing with real data: {str(e)}')
        import traceback
        traceback.print_exc()

def test_conservative_mode():
    """Test conservative transfer mode"""
    print("=== Testing Conservative Transfer Mode ===")
    
    # Create test data
    df = create_test_data()
    print(f"Test data created with {len(df)} records")
    
    # Preprocess data
    processed_df = preprocess_data(df)
    print("Data preprocessing completed")
    
    # Identify transfer candidates
    transfer_out = identify_transfer_out_candidates(processed_df, 'conservative')
    transfer_in = identify_transfer_in_candidates(processed_df)
    
    print(f"Found {len(transfer_out)} transfer-out candidates")
    print(f"Found {len(transfer_in)} transfer-in candidates")
    
    # Match transfers
    transfer_suggestions = match_transfers(transfer_out, transfer_in, processed_df)
    print(f"Generated {len(transfer_suggestions)} transfer suggestions")
    
    # Calculate statistics
    stats = calculate_statistics(transfer_suggestions, 'conservative')
    print(f"Total transfer quantity: {stats['total_transfer_qty']}")
    
    if not transfer_suggestions.empty:
        print("\nTransfer Suggestions:")
        print(transfer_suggestions[['Article', 'Transfer Site', 'Receive Site', 'Transfer Qty', 'Transfer Type']])
    
    return transfer_suggestions, stats

def test_enhanced_mode():
    """Test enhanced transfer mode"""
    print("\n=== Testing Enhanced Transfer Mode ===")
    
    # Create test data
    df = create_test_data()
    
    # Preprocess data
    processed_df = preprocess_data(df)
    
    # Identify transfer candidates
    transfer_out = identify_transfer_out_candidates(processed_df, 'enhanced')
    transfer_in = identify_transfer_in_candidates(processed_df)
    
    print(f"Found {len(transfer_out)} transfer-out candidates")
    print(f"Found {len(transfer_in)} transfer-in candidates")
    
    # Match transfers
    transfer_suggestions = match_transfers(transfer_out, transfer_in, processed_df)
    print(f"Generated {len(transfer_suggestions)} transfer suggestions")
    
    # Calculate statistics
    stats = calculate_statistics(transfer_suggestions, 'enhanced')
    print(f"Total transfer quantity: {stats['total_transfer_qty']}")
    
    if not transfer_suggestions.empty:
        print("\nTransfer Suggestions:")
        print(transfer_suggestions[['Article', 'Transfer Site', 'Receive Site', 'Transfer Qty', 'Transfer Type']])
    
    return transfer_suggestions, stats

def test_special_mode():
    """Test special enhanced transfer mode"""
    print("\n=== Testing Special Enhanced Transfer Mode ===")
    
    # Create test data
    df = create_test_data()
    
    # Preprocess data
    processed_df = preprocess_data(df)
    
    # Identify transfer candidates
    transfer_out = identify_transfer_out_candidates(processed_df, 'special')
    transfer_in = identify_transfer_in_candidates(processed_df)
    
    print(f"Found {len(transfer_out)} transfer-out candidates")
    print(f"Found {len(transfer_in)} transfer-in candidates")
    
    # Match transfers
    transfer_suggestions = match_transfers(transfer_out, transfer_in, processed_df)
    print(f"Generated {len(transfer_suggestions)} transfer suggestions")
    
    # Calculate statistics
    stats = calculate_statistics(transfer_suggestions, 'special')
    print(f"Total transfer quantity: {stats['total_transfer_qty']}")
    
    if not transfer_suggestions.empty:
        print("\nTransfer Suggestions:")
        print(transfer_suggestions[['Article', 'Transfer Site', 'Receive Site', 'Transfer Qty', 'Transfer Type']])
    
    return transfer_suggestions, stats

if __name__ == "__main__":
    print("Testing Retail Inventory Transfer System")
    print("=" * 50)

    # Test with real Excel data
    test_with_real_data()

    print("\n" + "=" * 50)
    print("All tests completed!")
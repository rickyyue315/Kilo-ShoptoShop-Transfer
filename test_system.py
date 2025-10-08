"""
Test script to verify the transfer system functionality
"""
import pandas as pd
import numpy as np
from app import preprocess_data, identify_transfer_out_candidates, identify_transfer_in_candidates, match_transfers, calculate_statistics

def create_test_data():
    """Create sample test data similar to the Excel structure"""
    test_data = {
        'Article': ['108433502001', '108433502001', '108433502001', '108433502001'],
        'Article Description': ['PRESSED POWDER PUFF', 'PRESSED POWDER PUFF', 'PRESSED POWDER PUFF', 'PRESSED POWDER PUFF'],
        'RP Type': ['ND', 'RF', 'RF', 'RF'],
        'Site': ['HA40', 'HB38', 'HC68', 'HC25'],
        'OM': ['Candy', 'Candy', 'Candy', 'Candy'],
        'MOQ': [6, 6, 6, 6],
        'SaSa Net Stock': [16, 9, 6, 0],
        'Target': [0, 0, 30, 0],
        'Pending Received': [0, 0, 0, 0],
        'Safety Stock': [0, 8, 8, 8],
        'Last Month Sold Qty': [0, 1, 0, 3],
        'MTD Sold Qty': [0, 0, 1, 0]
    }
    return pd.DataFrame(test_data)

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
    
    # Test conservative mode
    conservative_results, conservative_stats = test_conservative_mode()
    
    # Test enhanced mode
    enhanced_results, enhanced_stats = test_enhanced_mode()
    
    # Test special mode
    special_results, special_stats = test_special_mode()
    
    print("\n=== Test Summary ===")
    print(f"Conservative mode: {conservative_stats['total_transfer_qty']} units transferred")
    print(f"Enhanced mode: {enhanced_stats['total_transfer_qty']} units transferred")
    print(f"Special mode: {special_stats['total_transfer_qty']} units transferred")
    print("All tests completed successfully!")
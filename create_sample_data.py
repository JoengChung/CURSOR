#!/usr/bin/env python3
"""
Create sample master products Excel file for testing the Product Order Generator
"""

import pandas as pd
import os

def create_sample_data():
    """Create sample product data"""
    sample_data = {
        'Product Number': ['PROD001', 'PROD002', 'PROD003', 'PROD004', 'PROD005', 'PROD006', 'PROD007', 'PROD008', 'PROD009', 'PROD010'],
        'Model': ['Model-A1', 'Model-B2', 'Model-C3', 'Model-D4', 'Model-E5', 'Model-F6', 'Model-G7', 'Model-H8', 'Model-I9', 'Model-J10'],
        'Flavor Name': ['Vanilla', 'Chocolate', 'Strawberry', 'Mint', 'Caramel', 'Coffee', 'Lemon', 'Orange', 'Berry', 'Coconut'],
        'Capacity': ['500ml', '1L', '750ml', '500ml', '1.5L', '750ml', '500ml', '1L', '2L', '500ml'],
        'Case Size': [24, 12, 18, 24, 6, 18, 24, 12, 4, 24],
        'Volume (L)': [0.5, 1.0, 0.75, 0.5, 1.5, 0.75, 0.5, 1.0, 2.0, 0.5],
        'Cubic Meter': [0.012, 0.018, 0.015, 0.012, 0.025, 0.015, 0.012, 0.018, 0.032, 0.012],
        'Price per Unit': [2.50, 4.00, 3.25, 2.50, 6.00, 3.25, 2.50, 4.00, 8.00, 2.50],
        'Category': ['Beverage', 'Beverage', 'Beverage', 'Beverage', 'Beverage', 'Beverage', 'Beverage', 'Beverage', 'Beverage', 'Beverage'],
        'Brand': ['Brand X', 'Brand Y', 'Brand X', 'Brand Z', 'Brand Y', 'Brand X', 'Brand Z', 'Brand Y', 'Brand X', 'Brand Z'],
        'Weight (kg)': [0.52, 1.05, 0.78, 0.52, 1.58, 0.78, 0.52, 1.05, 2.1, 0.52],
        'Barcode': ['1234567890123', '2345678901234', '3456789012345', '4567890123456', '5678901234567', '6789012345678', '7890123456789', '8901234567890', '9012345678901', '0123456789012']
    }
    
    # Create DataFrame
    df = pd.DataFrame(sample_data)
    
    # Save to Excel
    filename = 'sample_master_products.xlsx'
    df.to_excel(filename, index=False)
    print(f"Sample master products file created: {filename}")
    print(f"Contains {len(df)} sample products")
    print(f"Columns: {', '.join(df.columns)}")
    
    return filename

if __name__ == "__main__":
    create_sample_data()
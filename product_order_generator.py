#!/usr/bin/env python3
"""
Product Order Generator

A tool to automatically generate customer-specific spreadsheets from a master product database
by simply entering product numbers.

Usage:
    python product_order_generator.py
"""

import pandas as pd
import os
import sys
from pathlib import Path
from datetime import datetime
import argparse


class ProductOrderGenerator:
    def __init__(self, master_file_path):
        """
        Initialize the Product Order Generator
        
        Args:
            master_file_path (str): Path to the master Excel file containing all products
        """
        self.master_file_path = master_file_path
        self.master_data = None
        self.load_master_data()
    
    def load_master_data(self):
        """Load the master product database from Excel file"""
        try:
            print(f"Loading master product database from: {self.master_file_path}")
            self.master_data = pd.read_excel(self.master_file_path)
            print(f"Successfully loaded {len(self.master_data)} products")
            
            # Display available columns
            print("\nAvailable columns in master database:")
            for i, col in enumerate(self.master_data.columns, 1):
                print(f"  {i}. {col}")
                
        except FileNotFoundError:
            print(f"Error: Master file not found at {self.master_file_path}")
            print("Please ensure the master Excel file exists at the specified location.")
            sys.exit(1)
        except Exception as e:
            print(f"Error loading master file: {str(e)}")
            sys.exit(1)
    
    def find_product_by_number(self, product_number):
        """
        Find a product by its product number
        
        Args:
            product_number (str): The product number to search for
            
        Returns:
            pandas.Series or None: Product data if found, None otherwise
        """
        # Try to find the product number column (common variations)
        product_number_cols = ['product_number', 'Product Number', 'ProductNumber', 
                              'product_id', 'Product ID', 'ProductID', 'SKU', 'sku',
                              'model', 'Model', 'Item Number', 'item_number']
        
        product_col = None
        for col in product_number_cols:
            if col in self.master_data.columns:
                product_col = col
                break
        
        if product_col is None:
            print("Warning: Could not automatically detect product number column.")
            print("Available columns:", list(self.master_data.columns))
            return None
        
        # Search for the product
        mask = self.master_data[product_col].astype(str).str.strip() == str(product_number).strip()
        matching_products = self.master_data[mask]
        
        if len(matching_products) == 0:
            return None
        elif len(matching_products) == 1:
            return matching_products.iloc[0]
        else:
            print(f"Warning: Multiple products found for number {product_number}")
            return matching_products.iloc[0]  # Return first match
    
    def generate_customer_spreadsheet(self, product_numbers, customer_name, output_dir="output"):
        """
        Generate a customer-specific spreadsheet with the requested products
        
        Args:
            product_numbers (list): List of product numbers to include
            customer_name (str): Name of the customer for the filename
            output_dir (str): Directory to save the output file
            
        Returns:
            str: Path to the generated file
        """
        # Create output directory if it doesn't exist
        Path(output_dir).mkdir(exist_ok=True)
        
        # Find all requested products
        found_products = []
        missing_products = []
        
        print(f"\nSearching for {len(product_numbers)} products...")
        
        for product_num in product_numbers:
            product = self.find_product_by_number(product_num)
            if product is not None:
                found_products.append(product)
                print(f"✓ Found: {product_num}")
            else:
                missing_products.append(product_num)
                print(f"✗ Not found: {product_num}")
        
        if not found_products:
            print("Error: No products were found!")
            return None
        
        # Create DataFrame with found products
        customer_df = pd.DataFrame(found_products)
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_customer_name = "".join(c for c in customer_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        filename = f"{safe_customer_name}_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, filename)
        
        # Save to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Main products sheet
            customer_df.to_excel(writer, sheet_name='Products', index=False)
            
            # Summary sheet
            summary_data = {
                'Customer': [customer_name],
                'Generated On': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                'Total Products Requested': [len(product_numbers)],
                'Products Found': [len(found_products)],
                'Products Missing': [len(missing_products)]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Missing products sheet (if any)
            if missing_products:
                missing_df = pd.DataFrame({'Missing Product Numbers': missing_products})
                missing_df.to_excel(writer, sheet_name='Missing Products', index=False)
        
        print(f"\n✓ Customer spreadsheet generated: {output_path}")
        print(f"  - Products found: {len(found_products)}")
        if missing_products:
            print(f"  - Products missing: {len(missing_products)}")
            print(f"  - Missing products listed in 'Missing Products' sheet")
        
        return output_path
    
    def interactive_mode(self):
        """Run the generator in interactive mode"""
        print("\n" + "="*60)
        print("         PRODUCT ORDER GENERATOR")
        print("="*60)
        
        while True:
            print("\nOptions:")
            print("1. Generate customer spreadsheet")
            print("2. Search for a specific product")
            print("3. View master database info")
            print("4. Exit")
            
            choice = input("\nSelect an option (1-4): ").strip()
            
            if choice == '1':
                self.generate_order_interactive()
            elif choice == '2':
                self.search_product_interactive()
            elif choice == '3':
                self.show_database_info()
            elif choice == '4':
                print("Goodbye!")
                break
            else:
                print("Invalid choice. Please select 1-4.")
    
    def generate_order_interactive(self):
        """Interactive mode for generating customer orders"""
        print("\n" + "-"*40)
        print("GENERATE CUSTOMER SPREADSHEET")
        print("-"*40)
        
        # Get customer name
        customer_name = input("Enter customer name: ").strip()
        if not customer_name:
            print("Customer name is required.")
            return
        
        # Get product numbers
        print("\nEnter product numbers (one per line, empty line to finish):")
        product_numbers = []
        while True:
            product_num = input("Product number: ").strip()
            if not product_num:
                break
            product_numbers.append(product_num)
        
        if not product_numbers:
            print("No product numbers entered.")
            return
        
        # Generate the spreadsheet
        output_path = self.generate_customer_spreadsheet(product_numbers, customer_name)
        
        if output_path:
            print(f"\nSpreadsheet saved to: {output_path}")
        
    def search_product_interactive(self):
        """Interactive mode for searching products"""
        print("\n" + "-"*40)
        print("SEARCH PRODUCT")
        print("-"*40)
        
        product_num = input("Enter product number to search: ").strip()
        if not product_num:
            return
        
        product = self.find_product_by_number(product_num)
        if product is not None:
            print(f"\n✓ Product found:")
            for col, value in product.items():
                print(f"  {col}: {value}")
        else:
            print(f"✗ Product {product_num} not found in database.")
    
    def show_database_info(self):
        """Show information about the master database"""
        print("\n" + "-"*40)
        print("MASTER DATABASE INFO")
        print("-"*40)
        print(f"File: {self.master_file_path}")
        print(f"Total products: {len(self.master_data)}")
        print(f"Columns: {len(self.master_data.columns)}")
        print("\nColumn names:")
        for i, col in enumerate(self.master_data.columns, 1):
            print(f"  {i:2d}. {col}")


def main():
    parser = argparse.ArgumentParser(description='Product Order Generator')
    parser.add_argument('--master-file', '-m', 
                       default='master_products.xlsx',
                       help='Path to master Excel file (default: master_products.xlsx)')
    parser.add_argument('--customer', '-c',
                       help='Customer name')
    parser.add_argument('--products', '-p', nargs='+',
                       help='List of product numbers')
    parser.add_argument('--output-dir', '-o',
                       default='output',
                       help='Output directory (default: output)')
    
    args = parser.parse_args()
    
    # Initialize the generator
    try:
        generator = ProductOrderGenerator(args.master_file)
    except Exception as e:
        print(f"Failed to initialize: {e}")
        return 1
    
    # If customer and products provided via command line, generate directly
    if args.customer and args.products:
        output_path = generator.generate_customer_spreadsheet(
            args.products, args.customer, args.output_dir
        )
        return 0 if output_path else 1
    
    # Otherwise run in interactive mode
    generator.interactive_mode()
    return 0


if __name__ == "__main__":
    sys.exit(main())
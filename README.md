# Product Order Generator

A Python-based solution that automatically generates customer-specific Excel spreadsheets from a master product database by simply entering product numbers. This eliminates the manual process of searching, copying, and pasting data from master spreadsheets for each customer order.

## Features

- **Automatic Product Lookup**: Find products by product number from your master Excel database
- **Batch Processing**: Handle multiple product numbers at once
- **Customer-Specific Spreadsheets**: Generate organized Excel files for each customer
- **Missing Product Tracking**: Identify and report products that couldn't be found
- **Multiple Interfaces**: Both command-line and graphical user interfaces
- **Flexible Column Detection**: Automatically detects common product number column names
- **Summary Reports**: Includes summary sheets with order statistics

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Setup Instructions

1. **Clone or download the project files**
   ```bash
   git clone <repository-url>
   cd product-order-generator
   ```

2. **Install required dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Prepare your master Excel file**
   - Ensure your master product database is in Excel format (.xlsx or .xls)
   - The file should have a column containing product numbers (common names like "Product Number", "SKU", "Model", etc. are automatically detected)

## Usage

### Method 1: Graphical User Interface (Recommended for beginners)

Run the GUI application:
```bash
python gui_interface.py
```

1. **Load Master Database**: Click "Browse" to select your master Excel file, then click "Load Master Database"
2. **Enter Customer Information**: Fill in the customer name and select output directory
3. **Add Product Numbers**: Enter product numbers in the text area (one per line or comma-separated)
4. **Generate**: Click "Generate Customer Spreadsheet"

### Method 2: Command Line Interface

#### Interactive Mode
```bash
python product_order_generator.py
```

Follow the on-screen prompts to:
- Load your master database
- Generate customer spreadsheets
- Search for specific products
- View database information

#### Direct Command Line Usage
```bash
# Generate spreadsheet directly
python product_order_generator.py --master-file "master_products.xlsx" --customer "ABC Company" --products "PROD001" "PROD002" "PROD003"

# Specify custom output directory
python product_order_generator.py -m "master_products.xlsx" -c "XYZ Corp" -p "SKU123" "SKU456" -o "customer_orders"
```

### Command Line Options
- `--master-file` or `-m`: Path to master Excel file (default: master_products.xlsx)
- `--customer` or `-c`: Customer name
- `--products` or `-p`: List of product numbers
- `--output-dir` or `-o`: Output directory (default: output)

## File Structure

After running the application, you'll have:

```
project-directory/
├── product_order_generator.py    # Main CLI application
├── gui_interface.py             # GUI application
├── requirements.txt             # Python dependencies
├── README.md                   # This file
├── master_products.xlsx        # Your master database (you provide this)
└── output/                     # Generated customer spreadsheets
    ├── Customer_Name_20240101_123456.xlsx
    └── ...
```

## Generated Spreadsheet Structure

Each customer spreadsheet contains:

1. **Products Sheet**: All found products with complete information from master database
2. **Summary Sheet**: Order statistics including:
   - Customer name
   - Generation timestamp
   - Total products requested
   - Products found vs. missing
3. **Missing Products Sheet** (if applicable): List of product numbers that couldn't be found

## Master Database Requirements

Your master Excel file should:
- Be in .xlsx or .xls format
- Have a header row with column names
- Include a column with product numbers/SKUs/model numbers
- Contain all product information you want to include in customer spreadsheets

### Supported Product Number Column Names
The system automatically detects columns with these names:
- `product_number`, `Product Number`, `ProductNumber`
- `product_id`, `Product ID`, `ProductID`
- `SKU`, `sku`
- `model`, `Model`
- `Item Number`, `item_number`

## Examples

### Example 1: Basic Usage
```bash
python product_order_generator.py -m "inventory.xlsx" -c "Acme Corp" -p "A001" "B002" "C003"
```

### Example 2: Large Order
```bash
python product_order_generator.py -m "products.xlsx" -c "BigBox Retail" -p "SKU001" "SKU002" "SKU003" "SKU004" "SKU005"
```

### Example 3: Custom Output Location
```bash
python product_order_generator.py -m "master.xlsx" -c "Local Store" -p "ITEM123" -o "/path/to/customer/orders"
```

## Troubleshooting

### Common Issues

1. **"Master file not found" error**
   - Ensure the file path is correct
   - Check that the file exists and is accessible

2. **"Could not detect product number column" warning**
   - Check your master file has a column with product numbers
   - Rename the column to one of the supported names listed above

3. **Products not found**
   - Verify product numbers match exactly (case-sensitive)
   - Check for extra spaces or special characters
   - Use the search function to verify product numbers exist

4. **Permission errors when saving**
   - Ensure the output directory exists and is writable
   - Close any Excel files that might be open in the output directory

### Getting Help

1. **View database information**: Use the "Database Info" feature to see all available columns
2. **Search individual products**: Use the search function to verify product numbers
3. **Check the log/status area**: Error messages provide specific details about issues

## Technical Details

### Dependencies
- **pandas**: Excel file reading and data manipulation
- **openpyxl**: Excel file writing (.xlsx format)
- **xlrd**: Excel file reading (.xls format support)
- **tkinter**: GUI interface (included with Python)

### Performance
- Handles thousands of products efficiently
- Memory usage scales with master database size
- Fast lookup using pandas indexing

### File Formats
- **Input**: Excel files (.xlsx, .xls)
- **Output**: Excel files (.xlsx) with multiple sheets

## License

This project is provided as-is for internal business use. Modify and distribute as needed for your organization.

## Support

For technical support or feature requests, contact your IT department or the system administrator.
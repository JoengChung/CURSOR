#!/usr/bin/env python3
"""
GUI Interface for Product Order Generator

A simple graphical user interface for the Product Order Generator using tkinter.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
from pathlib import Path

# Import the main generator class
from product_order_generator import ProductOrderGenerator


class ProductOrderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Product Order Generator")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)
        
        # Initialize variables
        self.generator = None
        self.master_file_path = tk.StringVar()
        self.customer_name = tk.StringVar()
        self.output_dir = tk.StringVar(value="output")
        
        self.setup_gui()
        
    def setup_gui(self):
        """Set up the GUI layout"""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Product Order Generator", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Master file selection
        ttk.Label(main_frame, text="Master Excel File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        master_file_frame = ttk.Frame(main_frame)
        master_file_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        master_file_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(master_file_frame, textvariable=self.master_file_path, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(master_file_frame, text="Browse", 
                  command=self.browse_master_file).grid(row=0, column=1)
        
        # Load button
        ttk.Button(main_frame, text="Load Master Database", 
                  command=self.load_master_database).grid(row=2, column=0, columnspan=3, pady=10)
        
        # Customer information frame
        customer_frame = ttk.LabelFrame(main_frame, text="Customer Order", padding="10")
        customer_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        customer_frame.columnconfigure(1, weight=1)
        
        # Customer name
        ttk.Label(customer_frame, text="Customer Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(customer_frame, textvariable=self.customer_name, width=30).grid(
            row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0)
        )
        
        # Output directory
        ttk.Label(customer_frame, text="Output Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        output_frame = ttk.Frame(customer_frame)
        output_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        output_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_frame, textvariable=self.output_dir, width=30).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(output_frame, text="Browse", 
                  command=self.browse_output_dir).grid(row=0, column=1)
        
        # Product numbers input
        ttk.Label(customer_frame, text="Product Numbers:").grid(
            row=2, column=0, sticky=(tk.W, tk.N), pady=5
        )
        
        # Text area for product numbers
        text_frame = ttk.Frame(customer_frame)
        text_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=(5, 0))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.product_numbers_text = scrolledtext.ScrolledText(
            text_frame, height=6, width=40,
            wrap=tk.WORD
        )
        self.product_numbers_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Instructions
        instructions = ttk.Label(
            customer_frame, 
            text="Enter product numbers (one per line or separated by commas)",
            font=('Arial', 8),
            foreground='gray'
        )
        instructions.grid(row=3, column=1, sticky=tk.W, pady=(0, 5), padx=(5, 0))
        
        # Generate button
        ttk.Button(customer_frame, text="Generate Customer Spreadsheet", 
                  command=self.generate_spreadsheet).grid(
            row=4, column=0, columnspan=2, pady=10
        )
        
        # Status/Log area
        log_frame = ttk.LabelFrame(main_frame, text="Status", padding="5")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=8, width=80,
            wrap=tk.WORD, state=tk.DISABLED
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Additional buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Clear Log", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Search Product", 
                  command=self.search_product_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Database Info", 
                  command=self.show_database_info).pack(side=tk.LEFT, padx=5)
        
    def log_message(self, message):
        """Add a message to the log area"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
        
    def clear_log(self):
        """Clear the log area"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
    def browse_master_file(self):
        """Browse for master Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Master Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.master_file_path.set(filename)
            
    def browse_output_dir(self):
        """Browse for output directory"""
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir.set(dirname)
            
    def load_master_database(self):
        """Load the master database"""
        if not self.master_file_path.get():
            messagebox.showerror("Error", "Please select a master Excel file first.")
            return
            
        try:
            self.clear_log()
            self.log_message("Loading master database...")
            
            # Redirect stdout to capture print statements
            import io
            from contextlib import redirect_stdout
            
            f = io.StringIO()
            with redirect_stdout(f):
                self.generator = ProductOrderGenerator(self.master_file_path.get())
            
            # Display the captured output
            output = f.getvalue()
            for line in output.strip().split('\n'):
                if line.strip():
                    self.log_message(line)
                    
            self.log_message("✓ Master database loaded successfully!")
            messagebox.showinfo("Success", "Master database loaded successfully!")
            
        except Exception as e:
            error_msg = f"Error loading master database: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)
            
    def generate_spreadsheet(self):
        """Generate customer spreadsheet"""
        if not self.generator:
            messagebox.showerror("Error", "Please load the master database first.")
            return
            
        customer_name = self.customer_name.get().strip()
        if not customer_name:
            messagebox.showerror("Error", "Please enter a customer name.")
            return
            
        # Parse product numbers
        product_text = self.product_numbers_text.get(1.0, tk.END).strip()
        if not product_text:
            messagebox.showerror("Error", "Please enter at least one product number.")
            return
            
        # Split by lines and commas, then clean up
        product_numbers = []
        for line in product_text.split('\n'):
            for item in line.split(','):
                item = item.strip()
                if item:
                    product_numbers.append(item)
                    
        if not product_numbers:
            messagebox.showerror("Error", "No valid product numbers found.")
            return
            
        try:
            self.log_message(f"\nGenerating spreadsheet for customer: {customer_name}")
            self.log_message(f"Product numbers: {', '.join(product_numbers)}")
            
            # Redirect stdout to capture print statements
            import io
            from contextlib import redirect_stdout
            
            f = io.StringIO()
            with redirect_stdout(f):
                output_path = self.generator.generate_customer_spreadsheet(
                    product_numbers, customer_name, self.output_dir.get()
                )
            
            # Display the captured output
            output = f.getvalue()
            for line in output.strip().split('\n'):
                if line.strip():
                    self.log_message(line)
                    
            if output_path:
                self.log_message(f"\n✓ Spreadsheet generated successfully!")
                messagebox.showinfo("Success", 
                    f"Customer spreadsheet generated successfully!\n\nSaved to: {output_path}")
            else:
                messagebox.showerror("Error", "Failed to generate spreadsheet.")
                
        except Exception as e:
            error_msg = f"Error generating spreadsheet: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)
            
    def search_product_dialog(self):
        """Open dialog to search for a specific product"""
        if not self.generator:
            messagebox.showerror("Error", "Please load the master database first.")
            return
            
        product_num = tk.simpledialog.askstring("Search Product", "Enter product number:")
        if not product_num:
            return
            
        try:
            product = self.generator.find_product_by_number(product_num)
            if product is not None:
                # Create a new window to display product details
                detail_window = tk.Toplevel(self.root)
                detail_window.title(f"Product Details - {product_num}")
                detail_window.geometry("500x400")
                
                # Create text widget to display product info
                text_widget = scrolledtext.ScrolledText(detail_window, wrap=tk.WORD)
                text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                text_widget.insert(tk.END, f"Product Number: {product_num}\n")
                text_widget.insert(tk.END, "="*50 + "\n\n")
                
                for col, value in product.items():
                    text_widget.insert(tk.END, f"{col}: {value}\n")
                    
                text_widget.config(state=tk.DISABLED)
                
            else:
                messagebox.showinfo("Not Found", f"Product {product_num} not found in database.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error searching for product: {str(e)}")
            
    def show_database_info(self):
        """Show database information"""
        if not self.generator:
            messagebox.showerror("Error", "Please load the master database first.")
            return
            
        try:
            # Create info window
            info_window = tk.Toplevel(self.root)
            info_window.title("Database Information")
            info_window.geometry("400x500")
            
            # Create text widget
            text_widget = scrolledtext.ScrolledText(info_window, wrap=tk.WORD)
            text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Add database info
            text_widget.insert(tk.END, "MASTER DATABASE INFORMATION\n")
            text_widget.insert(tk.END, "="*50 + "\n\n")
            text_widget.insert(tk.END, f"File: {self.generator.master_file_path}\n")
            text_widget.insert(tk.END, f"Total products: {len(self.generator.master_data)}\n")
            text_widget.insert(tk.END, f"Columns: {len(self.generator.master_data.columns)}\n\n")
            text_widget.insert(tk.END, "Column names:\n")
            text_widget.insert(tk.END, "-"*20 + "\n")
            
            for i, col in enumerate(self.generator.master_data.columns, 1):
                text_widget.insert(tk.END, f"{i:2d}. {col}\n")
                
            text_widget.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error displaying database info: {str(e)}")


def main():
    # Import tkinter.simpledialog for the search dialog
    import tkinter.simpledialog
    tk.simpledialog = tkinter.simpledialog
    
    root = tk.Tk()
    app = ProductOrderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
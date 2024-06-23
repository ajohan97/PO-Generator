import pandas as pd
import tkinter as tk
from tkinter import filedialog
import curses
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
import os

def read_spreadsheet(file_path):

    try:
        # Read the spreadsheet into a pandas ExcelFile object
        xl = pd.ExcelFile(file_path)

        # Read the main sheet (first sheet)
        df = xl.parse(xl.sheet_names[0])

        # Check if the DataFrame has columns named "SKU", "Title", and "Barcode"
        required_columns = ['SKU', 'Title', 'Barcode']
        missing_columns = [col for col in required_columns if col not in df.columns]
            
        if missing_columns:
            raise ValueError(f"The spreadsheet is missing the following columns: {', '.join(missing_columns)}.")

        # Convert the DataFrame into a list of dictionaries, where each dictionary represents a row
        data_set = df.to_dict(orient='records')

                # Initialize meta_data dictionary for additional categories
        meta_data = {}

        # Read data from the second sheet if it exists
        if len(xl.sheet_names) > 1:
            df_meta = xl.parse(xl.sheet_names[1])
                
            # Define categories to read from the second sheet
            meta_categories = [
                'po_number', 'po_date', 'ship_to_name', 'ship_to_address_1', 'ship_to_address_2',
                'ship_to_address_3', 'comments_1', 'comments_2', 'comments_3', 'comments_4',
                'comments_5', 'comments_6', 'comments_7', 'comments_8', 'comments_9', 'comments_10',
                'company_name', 'company_address_1', 'company_address_2', 'company_country',
                'company_logo', 'sender_name', 'sender_company_name', 'sender_company_address_1',
                'sender_company_address_2', 'sender_company_address_3', 'sender_company_country'
            ]

            # Extract values for each meta category if it exists in the second sheet
            for category in meta_categories:
                if category in df_meta.columns:
                    meta_data[category] = df_meta[category].iloc[0]
                else:
                    meta_data[category] = ""  # or provide default value if needed

        return data_set, meta_data
    
    except ValueError as ve:
        raise ve
    except Exception as e:
        raise ValueError("Error reading spreadsheet.") from e

def open_file_dialog():
    # Initialize Tkinter root
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open file dialog
    file_path = filedialog.askopenfilename(
        title="Select a Spreadsheet",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    if file_path:
        return file_path
    else:
        print("No file selected")
        return None

def get_user_input(screen, skus):
    selected_index = 0
    page_size = curses.LINES - 2  # Reserve space for header and navigation
    current_page = 0

    while True:
        screen.clear()
        screen.addstr(0, 0, "Select SKU using arrow keys and press Enter:")
        start_index = current_page * page_size
        end_index = start_index + page_size
        page_skus = skus[start_index:end_index]

        for idx, sku in enumerate(page_skus):
            actual_index = start_index + idx
            if actual_index == selected_index:
                screen.addstr(idx + 1, 0, f"> {sku}", curses.A_REVERSE)
            else:
                screen.addstr(idx + 1, 0, f"  {sku}")
        
        screen.addstr(curses.LINES - 1, 0, f"Page {current_page + 1} of {((len(skus) - 1) // page_size) + 1}")
        key = screen.getch()

        if key == curses.KEY_UP and selected_index > 0:
            selected_index -= 1
            if selected_index < start_index:
                current_page -= 1
        elif key == curses.KEY_DOWN and selected_index < len(skus) - 1:
            selected_index += 1
            if selected_index >= end_index:
                current_page += 1
        elif key == curses.KEY_ENTER or key in [10, 13]:
            return skus[selected_index]

def generate_pdf(selected_products, meta_data):
    # Create a PDF document
    pdf_file = "selected_products_report.pdf"
    c = canvas.Canvas(pdf_file, pagesize=letter)
    
    # Load and position company logo if available
    company_logo_path = "logo.jpg"
    if company_logo_path:
        try:
            company_logo = ImageReader(company_logo_path)

            #Resize the image
            logo_width, logo_height = company_logo.getSize()
            logo_width = logo_width / 7
            logo_height = logo_height / 7

            #Center the image on the page
            page_width, page_height = letter
            x = (page_width - logo_width) / 2
            y = (page_height - logo_height)
    
            # Draw the image on the canvas
            c.drawImage(company_logo, x, y, width=logo_width, height=logo_height)
    
        except Exception as e:
            print(f"Failed to load company logo: {e}")

    # Set the font and size for the title text
    c.setFont("Helvetica-Bold", 12)
    
    # Define the title text and its position
    title_text = "PURCHASE ORDER"
    title_x = 500  # Adjust this value for horizontal position
    title_y = 750  # Adjust this value for vertical position
    
    # Draw the title text on the canvas
    c.drawString(title_x, title_y, title_text)
    
    # Set the font and size for the smaller text
    c.setFont("Helvetica", 10)
    
    # Get the PO details from meta_data
    po_number = meta_data.get("po_number", "N/A")
    po_date = meta_data.get("po_date", "N/A")

    # Define the smaller text and its position
    po_number_text = f"PO Number: {po_number}"
    po_date_text = f"PO Date: {po_date}"
    
    po_number_x = 500  # Same horizontal position as the title
    po_number_y = title_y - 15  # Slightly below the title
    
    po_date_x = 500  # Same horizontal position as the title
    po_date_y = po_number_y - 15  # Slightly below the PO number
    
    # Draw the smaller text on the canvas
    c.drawString(po_number_x, po_number_y, po_number_text)
    c.drawString(po_date_x, po_date_y, po_date_text)


    # Set up the table headers
    table_headers = ["SKU", "Barcode", "Quantity"]
    col_widths = [150, 150, 150]
    row_height = 30
    y_start = 750  # Start position of the first row
    x_start = 50   # Starting x position for the table


    # Draw table headers
    for i, header in enumerate(table_headers):
        c.drawString(x_start + sum(col_widths[:i]), y_start, header)
    
    # Draw a line under the headers
    c.line(x_start, y_start-5, sum(col_widths) + x_start, y_start-5)

    # Print each selected SKU, barcode, and quantity in the PDF
    for idx, (sku, barcode, qty) in enumerate(selected_products, start=1):
        y_pos = y_start - idx * row_height
        c.drawString(x_start, y_pos, sku)
        c.drawString(x_start + col_widths[0], y_pos, str(barcode))
        c.drawString(x_start + col_widths[0] + col_widths[1], y_pos, qty)
    
    # Print meta data
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_start, y_pos - 50, "Metadata:")
    c.setFont("Helvetica", 10)
    y_pos -= 20
    for key, value in meta_data.items():
        c.drawString(x_start, y_pos, f"{value}")
        y_pos -= 15
    
    # Save the PDF file
    c.save()
    print(f"PDF report generated: {pdf_file}")

def main():
    #file_path = open_file_dialog()
    file_path = "C:\\Users\Alex - User\\Documents\\Lee Products Ltd\\Product\\PO\\PO Generator\\Sample Spreadsheet.xlsx"
    if file_path:
        try:
            data_set, meta_data = read_spreadsheet(file_path)
            
            selected_products = []

            # Initialize curses
            while True:

                selected_product = curses.wrapper(get_user_input, data_set)
                selected_product_sku = selected_product["SKU"]
                selected_product_barcode = selected_product["Barcode"]
                # Get quantity input from user
                selected_product_quantity = input(f"Enter Quantity for SKU {selected_product_sku}: ")
                
                # Store the SKU, barcode, and quantity in a list
                selected_products.append((selected_product_sku, selected_product_barcode, selected_product_quantity))
                
                # Ask user if they want to add another SKU or finish
                choice = input("Enter 1 to add another SKU or 2 to finish: ")
                if choice == '2':
                    break
        
            # Print all selected SKUs, quantities, and barcodes
            print("Selected SKUs, Barcodes, and Quantities:")
            for idx, (sku, barcode, qty) in enumerate(selected_products, start=1):
                print(f"{idx}. SKU: {sku}, Barcode: {barcode}, Quantity: {qty}")
            
            # Generate PDF report
            generate_pdf(selected_products, meta_data)

        except ValueError as e:
            print(e)
    else:
        print("No file selected, exiting.")

# Entry point
if __name__ == "__main__":
    main()

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import curses
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PIL import Image
import shutil

import os

def get_string_height(text, font_name, font_size):
    # Get the ascent and descent of the font
    ascent = pdfmetrics.getAscent(font_name)
    descent = pdfmetrics.getDescent(font_name)
    
    # Calculate the height
    height = (ascent - descent) / 1000 * font_size
    return height

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
                'po_number', 'po_date', 'subtotal',	'shipping',	'transaction_fee',	'total', 'ship_to_name', 'ship_to_address_1', 'ship_to_address_2',
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
    po_number = f"{meta_data.get('po_number', 'N/A')}"
    pdf_file = f"{po_number}.pdf"
    c = canvas.Canvas(pdf_file, pagesize=letter)

    # Define margins
    margin_width = 0.5 * 72  # 0.75 inches converted to points (72 points per inch)
    margin_height = 0.5 * 72

    # Define text sizes
    Title1_size = 18
    Title2_size = 14
    Text_size = 10

    # Define line spacing
    line_spacing = 2

    def add_new_page():
        c.showPage()
        c.setFont("Helvetica", Text_size)

    #Center the image on the page
    page_width, page_height = letter

    # Get the available width and height inside the margins
    available_width = page_width - 2 * margin_width
    available_height = page_height -2 * margin_height

    # Load and position company logo if available
    company_logo_path = "assets\\logo.jpg"
    if company_logo_path:
        try:
            company_logo = ImageReader(company_logo_path)

            #Resize the image
            logo_width, logo_height = company_logo.getSize()
            logo_width = logo_width / 9
            logo_height = logo_height / 9

            #Center the image on the page
            page_width, page_height = letter

            x = (available_width - logo_width) / 2 + margin_width
            y = (page_height - logo_height)
    
            # Draw the image on the canvas
            c.drawImage(company_logo, x, y, width=logo_width, height=logo_height)
    
        except Exception as e:
            print(f"Failed to load company logo: {e}")

    # Set the font and size for the company name text
    c.setFont("Helvetica-Bold", Title2_size)
    
    # Get the company details from meta_data and draw it on the PDF
    company_name = f"{meta_data.get('company_name', 'N/A')}"
    company_name_height = get_string_height(company_name, "Helvetica-Bold", Title2_size)
    company_name_y = page_height - margin_height - company_name_height
    c.drawString(margin_width, company_name_y, company_name)

    # Set the font size for the company address text
    c.setFont("Helvetica", Text_size)

    # Get the company details from meta_data and draw it on the PDF
    company_address_1 = f"{meta_data.get('company_address_1', 'N/A')}"
    company_address_2 = f"{meta_data.get('company_address_2', 'N/A')}"
    company_country = f"{meta_data.get('company_country', 'N/A')}"
    
    # Align the company details vertically
    company_address_1_y = company_name_y - 15
    company_address_2_y = company_address_1_y - 15
    company_country_y = company_address_2_y - 15

    # Draw the smaller text on the canvas
    c.drawString(margin_width, company_address_1_y, company_address_1)
    c.drawString(margin_width, company_address_2_y, company_address_2)
    c.drawString(margin_width, company_country_y, company_country)

    # Set the font and size for the title text
    c.setFont("Helvetica-Bold", Title1_size)
    
    # Define the title text and its position
    title_text = "PURCHASE ORDER"
    title_width = c.stringWidth(title_text, "Helvetica-Bold", Title1_size)
    title_x = page_width - margin_width - title_width  # Align to the right
    title_y_height = get_string_height(title_text, "Helvetica-Bold", Title1_size)
    title_y = page_height - margin_height - title_y_height # Adjust this value for vertical position
    
    # Draw the title text on the canvas
    c.drawString(title_x, title_y, title_text)
    
    # Set the font and size for the smaller text
    c.setFont("Helvetica", Text_size)
    
    # Get the PO details from meta_data
    po_number = f"{meta_data.get('po_number', 'N/A')}"
    po_date = f"Date: {meta_data.get('po_number', 'N/A')}"

    po_number_width = c.stringWidth(po_number, "Helvetica", Text_size)
    po_date_width = c.stringWidth(po_date, "Helvetica", Text_size)

    # Align the smaller text to the right
    po_number_x = page_width - margin_width - po_number_width
    po_date_x = page_width - margin_width - po_date_width
    
    po_number_y = title_y - 15  # Slightly below the title
    po_date_y = po_number_y - 15  # Slightly below the PO number
    
    # Draw the smaller text on the canvas
    c.drawString(po_number_x, po_number_y, po_number)
    c.drawString(po_date_x, po_date_y, po_date)

    # Separate lines for sender and ship to data
    sender_data = [
        f"SENDER:",
        f"{meta_data.get('sender_name', 'N/A')}",
        f"{meta_data.get('sender_company_name', 'N/A')}",
        f"{meta_data.get('sender_company_address_1', 'N/A')}",
        f"{meta_data.get('sender_company_address_2', 'N/A')}",
        f"{meta_data.get('sender_company_address_3', 'N/A')}",
        f"{meta_data.get('sender_company_country', 'N/A')}"
    ]

    ship_to_data = [
        f"SHIP TO:",
        f"{meta_data.get('ship_to_name', 'N/A')}",
        f"{meta_data.get('ship_to_address_1', 'N/A')}",
        f"{meta_data.get('ship_to_address_2', 'N/A')}",
        f"{meta_data.get('ship_to_address_3', 'N/A')}"
    ]

    # Combine sender and ship-to data into pairs
    max_length = max(len(sender_data), len(ship_to_data))
    table_data = [
        [sender_data[i] if i < len(sender_data) else '', 
        ship_to_data[i] if i < len(ship_to_data) else '']
        for i in range(max_length)
    ]

    # Calculate column widths dynamically
    sender_width = max(c.stringWidth(line, "Helvetica", 12) for line in sender_data)
    ship_to_width = max(c.stringWidth(line, "Helvetica", 12) for line in ship_to_data)

    total_width = sender_width + ship_to_width
    scaling_factor = available_width / total_width

    sender_width *= scaling_factor
    ship_to_width *= scaling_factor

    # Define the table style
    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),  # Background color for the first row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),   # Text color for the first row
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold font for the first row
        ('FONTSIZE', (0, 0), (-1, 0), 12),  # Font size for the first row
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),  # No bottom padding for all rows
        ('TOPPADDING', (0, 0), (-1, -1), 0),     # No top padding for all rows
        ('LEFTPADDING', (0, 0), (-1, -1), 0),    # No left padding for all rows
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),   # No right padding for all rows
    ])

    table = Table(table_data, colWidths=[sender_width, ship_to_width])
    table.setStyle(table_style)

    # Calculate the position for the table
    table_width, table_height = table.wrap(available_width, available_height)
    table_x = margin_width
    table_y = po_date_y - table_height - 40  # Adjust this value to control spacing
    
    # Draw the table on the canvas
    table.wrapOn(c, available_width, available_height)
    table.drawOn(c, table_x, table_y)

    # ------- Draw the comments---------
    c.setFont("Helvetica-Bold", Text_size)
    comments_y = table_y - 30
    c.drawString(margin_width, comments_y, "COMMENTS OR SPECIAL INSTRUCTIONS")
    #comments_y -= 5

    # List of comments from meta_data
    comments = [
        meta_data.get('comments_1', ''),
        meta_data.get('comments_2', ''),
        meta_data.get('comments_3', ''),
        meta_data.get('comments_4', ''),
        meta_data.get('comments_5', ''),
        meta_data.get('comments_6', ''),
        meta_data.get('comments_7', ''),
        meta_data.get('comments_8', ''),
        meta_data.get('comments_9', ''),
        meta_data.get('comments_10', '')
    ]

    # Convert all comments to strings and filter out empty or 'nan' comments
    filtered_comments = [str(comment) for comment in comments if str(comment).strip() and str(comment).lower() != 'nan']

    # Create a style for the comments
    styles = getSampleStyleSheet()
    bullet_style = ParagraphStyle(
        'Bullet',
        parent=styles['BodyText'],
        leftIndent=20,       # Offset for all lines
        firstLineIndent=-6,  # Negative offset to bring bullet back
        spaceBefore=0,       # No space before the paragraph
        spaceAfter=2,        # No space after the paragraph
        leading=12           # Line height
    )

    # Draw the bullet points for each comment
    for comment in filtered_comments:
        para = Paragraph(f"• {comment}", bullet_style)
        width, height = para.wrap(available_width, available_height)  # Measure the width and height of the paragraph
        comments_y -= (height + bullet_style.spaceAfter)  # Adjust the position for the next comment
        para.drawOn(c, margin_width, comments_y)

    #-----------SKU TABLE-----------

    c.setFont("Helvetica-Bold", Text_size)

    # Set up the table headers
    table_headers = ["ITEM #", "QTY", "Description", "Sample", "Barcode", "Total (USD)"]

    item_no_col_width = 40
    qty_col_width = 30
    description_col_width = 260
    sample_image_col_width = 75
    barcode_col_width = 75
    total_col_width = 60

    col_widths = [item_no_col_width, qty_col_width, description_col_width, sample_image_col_width, barcode_col_width, total_col_width]
    header_height = 25
    row_height = 50
    font_height = get_string_height("This is a test", "Helvetica", Text_size)  # Example usage, replace with actual font and size
    y_start = comments_y - 50  # Start position of the first row
    x_start = margin_width  # Starting x position for the table

    # Draw table headers
    for i, header in enumerate(table_headers):
        # Calculate center position for each header
        header_width = c.stringWidth(header)  # Get the width of the header string
        header_x = x_start + sum(col_widths[:i]) + (col_widths[i] - header_width) / 2
        c.drawString(header_x, y_start + (header_height - font_height) / 2, header)
    
    c.setFont("Helvetica", Text_size)

    # Draw a line above the headers
    c.line(x_start, y_start + header_height, sum(col_widths) + x_start, y_start + header_height)

    # Draw a line under the headers
    c.line(x_start, y_start, sum(col_widths) + x_start, y_start)
    line_spacing = 5
    page_count = 1
    relative_height = 1

    # Print each selected SKU, barcode, and quantity in the PDF
    for idx, (qty, sku, title, barcode) in enumerate(selected_products, start=1):
    
        y_pos = y_start - relative_height * row_height   # Center vertically

        #If the table would extend into the margin of the existing pages, add a new page
        if y_pos < (margin_height):

            # Draw vertical lines for the grid
            for i in range(len(col_widths) + 1):
                
                x_pos = x_start + sum(col_widths[:i])
                c.line(x_pos, y_start + header_height, x_pos, y_start - row_height * (relative_height - 1))
            
            # Print page number on previous page
            c.drawString((page_width - 2 * margin_width), margin_height / 2 + get_string_height("Page", "Helvetica", Text_size) / 2, f"Page {page_count}")

            add_new_page()
            page_count = page_count +1
            relative_height = 0
            y_start = page_height - margin_height - 1.5 * row_height
            y_pos = y_start

            #Draw vertical line at the top of the next page
            c.line(x_start, y_pos + row_height, sum(col_widths) + x_start, y_pos + row_height)


        # Description split into two parts
        description_part1 = f'{title}.'
        description_part2 = f'See attachment "{sku}.png"'

        # Calculate positions for each line of the description with extra space
        description_y_pos1 = y_pos + (row_height - 2 * font_height - line_spacing) / 2 + font_height + line_spacing/2  # First line
        description_y_pos2 = y_pos + (row_height - 2 * font_height - line_spacing) / 2 # Second line

         # Item Number
        c.drawString(x_start + (col_widths[0] - c.stringWidth(str(idx))) / 2, y_pos + (row_height - font_height) / 2 + font_height / 4, str(idx))
        
        # Quantity
        c.drawString(x_start + col_widths[0] + (col_widths[1] - c.stringWidth(str(qty))) / 2, y_pos + (row_height - font_height) / 2 + font_height / 4, str(qty))
        
        # Description - First Line (Title) (Not centered horizontally)
        c.drawString(x_start + col_widths[0] + col_widths[1] + 5, description_y_pos1 + font_height / 4, description_part1)

        # Description - Second Line (See attachment {sku}.png) (Not centered horizontally)
        c.drawString(x_start + col_widths[0] + col_widths[1] + 5, description_y_pos2 + font_height / 4, description_part2)

        # Sample Image
        image_path = f"assets/{sku}.jpg"
        try:
            with Image.open(image_path) as img:
                img_width, img_height = img.size
                
                # Calculate scaling factor
                scale_factor = min(sample_image_col_width / img_width, row_height / img_height)
                
                # Calculate new dimensions
                new_width = img_width * scale_factor
                new_height = img_height * scale_factor
                
                # Calculate positions to center the image
                image_x = x_start + col_widths[0] + col_widths[1] + col_widths[2] + (col_widths[3] - new_width) / 2
                image_y = y_pos + (row_height - new_height) / 2
                
                # Draw the image
                c.drawImage(image_path, image_x, image_y, width=new_width, height=new_height, preserveAspectRatio=True)
        except IOError:
            # Handle the case where the image does not exist
            c.drawString(x_start + col_widths[0] + col_widths[1] + col_widths[2] + (col_widths[3] - c.stringWidth("No Image")) / 2, y_pos + (row_height - font_height) / 2, "No Image")

        # Barcode
        #Save the barcode x position for later when printing the summary data
        barcode_pos_x = x_start + col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3] + col_widths[4] 
        c.drawString(x_start + col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3] + (col_widths[4] - c.stringWidth(str(barcode))) / 2, y_pos + (row_height - font_height) / 2 + font_height / 4, str(barcode))

        # Total
        # Save the x position of this for later
        total_x_pos = x_start + col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3] + col_widths[4] + (col_widths[5] - c.stringWidth(str(idx))) / 2
        # Leave the total empty for now
        # c.drawString(x_start + col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3] + col_widths[4] + (col_widths[5] - c.stringWidth(str(idx))) / 2, y_pos + (row_height - font_height) / 2 + font_height / 4, str(idx))

        # Draw horizontal line for each row
        c.line(x_start, y_pos, sum(col_widths) + x_start, y_pos)

        # Increment relative_height
        relative_height = relative_height + 1

    # Draw vertical lines for the grid
    for i in range(len(col_widths) + 1):
        
        x_pos = x_start + sum(col_widths[:i])

        # If we are not on the first page, don't use the header_height and
        # len(selected_products) to determine where to draw the vertical lines
        if page_count == 1:
            c.line(x_pos, y_start + header_height, x_pos, y_start - row_height * len(selected_products))
        else:
            c.line(x_pos, y_start + row_height, x_pos, y_start - row_height * (relative_height - 1))
        

    #----------- Additional Items -----------
    final_row_height = 15
    additional_items = [
        "SUBTOTAL",
        "SALES TAX",
        "WARNING STICKERS, BARCODE STICKERS, AND OPAQUE POLY BAG + LABOR",
        "SHIPPING",
        "TRANSACTION FEE",
        "TOTAL"
    ]

    # Calculate starting y position for the additional items
    items_start_y = y_pos #- final_row_height # Adjust for spacing below the table

    relative_index = 1

    # Draw the additional items
    for idx, item in enumerate(additional_items):
        # Calculate y position for each item
        item_y_pos = items_start_y - relative_index * final_row_height  

        #If the table would extend into the margin of the existing pages, add a new page
        if item_y_pos < (margin_height):

            # Draw lines to the left and right of the lines we just drew to make boxes
            c.line(barcode_pos_x, y_pos, barcode_pos_x, item_y_pos + final_row_height)
            c.line(margin_width + available_width, y_pos, margin_width + available_width, item_y_pos + final_row_height)

            # Print page number on previous page
            c.drawString((page_width - 2 * margin_width), margin_height / 2 + get_string_height("Page", "Helvetica", Text_size) / 2, f"Page {page_count}")

            add_new_page()
            page_count = page_count +1
            relative_index = 0

            #Need to reset items_start_y and item_y_pos because the former is used to calculate the latter
            items_start_y = page_height - margin_height - 1.5 * final_row_height
            item_y_pos = items_start_y

        buffer = 5

        # Calculate the x position to align the item to the right side of the page
        item_width = c.stringWidth(item)
        item_x_pos = barcode_pos_x - item_width - buffer # Add 5 buffer

        # Draw the item name
        c.drawString(item_x_pos, item_y_pos + (final_row_height - font_height) / 2, item)

        if item == 'SUBTOTAL':
            subtotal = meta_data.get('subtotal', 'N/A')
            subtotal = f"{subtotal:.2f}"
            c.drawString(total_x_pos - c.stringWidth(str(subtotal))/2 + buffer, item_y_pos + (final_row_height - font_height) / 2, str(subtotal))
        elif item == 'SALES TAX':
            sales_tax = "EXEMPT"
            c.drawString(total_x_pos - c.stringWidth(str(sales_tax))/2 + buffer, item_y_pos + (final_row_height - font_height) / 2, str(sales_tax))
        elif item == 'WARNING STICKERS, BARCODE STICKERS, AND OPAQUE POLY BAG + LABOR':
            sticks_and_labor = "INCL"
            c.drawString(total_x_pos - c.stringWidth(str(sticks_and_labor))/2 + buffer, item_y_pos + (final_row_height - font_height) / 2, str(sticks_and_labor))
        elif item == 'SHIPPING':
            shipping = meta_data.get('shipping', 'N/A')
            shipping = f"{shipping:.2f}"
            c.drawString(total_x_pos - c.stringWidth(str(shipping))/2 + buffer, item_y_pos + (final_row_height - font_height) / 2, str(shipping))
        elif item == 'TRANSACTION FEE':
            transaction_fee = meta_data.get('transaction_fee', 'N/A')
            transaction_fee = f"{transaction_fee:.2f}"
            c.drawString(total_x_pos - c.stringWidth(str(transaction_fee))/2 + buffer, item_y_pos + (final_row_height - font_height) / 2, str(transaction_fee))
        elif item == 'TOTAL':
            total = meta_data.get('total', 'N/A')
            total = f"{total:.2f}"
            c.drawString(total_x_pos - c.stringWidth(str(total))/2 + buffer, item_y_pos + (final_row_height - font_height) / 2, str(total))

        # Draw lines to separate items
        c.line(barcode_pos_x, item_y_pos, margin_width + available_width, item_y_pos)

        relative_index = relative_index + 1


    if page_count == 1:
        # Draw lines to the left and right of the lines we just drew to make boxes
        c.line(barcode_pos_x, y_pos, barcode_pos_x, item_y_pos)
        c.line(margin_width + available_width, y_pos, margin_width + available_width, item_y_pos)
    else:
        # Draw line across the top box
        c.line(barcode_pos_x, items_start_y + final_row_height, margin_width + available_width, items_start_y + final_row_height)

        # Draw lines to the left and right of the lines we just drew to make boxes
        c.line(barcode_pos_x, items_start_y + final_row_height, barcode_pos_x, item_y_pos)
        c.line(margin_width + available_width, items_start_y + final_row_height, margin_width + available_width, item_y_pos)

    # Draw a final line below the last additional item
    # c.line(availabe_width, item_y_pos - final_row_height, sum(col_widths) + x_start, item_y_pos - final_row_height)

    # Draw the page number
    c.drawString((page_width - 2 * margin_width), margin_height / 2 + get_string_height("Page", "Helvetica", Text_size) / 2, f"Page {page_count}")

    # Save the PDF file
    c.save()
    print(f"PDF report generated: {pdf_file}")

def create_output_folder(selected_products, meta_data):

    source_folder = "assets"
    destination_folder = "outputs"

    po_number = f"{meta_data.get('po_number', 'N/A')}"


    try:
        # Check if the source folder exists
        if not os.path.exists(source_folder):
            print(f"Source folder '{source_folder}' does not exist.")
            return
        
        # Create a new folder with the PO number within the destination folder
        new_folder_name = po_number
        new_folder_path = os.path.join(destination_folder, new_folder_name)
        os.makedirs(new_folder_path, exist_ok=True)
        
        # Get list of files in the source folder
        files = os.listdir(source_folder)
        
        for product, (qty, sku, title, barcode) in enumerate(selected_products, start=1):
            # Copy the SKU design
            source_file = os.path.join(source_folder, f"{sku}.jpg")
            destination_file = os.path.join(new_folder_path,  f"{sku}.jpg")
            shutil.copy2(source_file, destination_file)  # Copy the file

            # Copy the barcode
            source_file = os.path.join(source_folder, f"{barcode}.png")
            destination_file = os.path.join(new_folder_path,  f"{barcode}.png")
            shutil.copy2(source_file, destination_file)  # Copy the file
        
        # Move the new PDF we generated
        destination_file = os.path.join(new_folder_path,  f"{po_number}.pdf")
        shutil.move(f"{po_number}.pdf", destination_file)  # Move the file

        print(f"Files copied from '{source_folder}' to '{new_folder_path}' successfully.")
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")


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
                selected_product_title = selected_product["Title"]
                selected_product_barcode = selected_product["Barcode"]
                # Get quantity input from user
                selected_product_quantity = input(f"Enter Quantity for SKU {selected_product_sku}: ")
                
                # Store the SKU, barcode, and quantity in a list
                selected_products.append((selected_product_quantity, selected_product_sku, selected_product_title, selected_product_barcode))
                
                # Ask user if they want to add another SKU or finish
                choice = input("Enter 1 to add another SKU or 2 to finish: ")
                if choice == '2':
                    break
        
            # Print all selected SKUs, quantities, and barcodes
            print("Selected SKUs, Barcodes, and Quantities:")
            for idx, (sku, title, barcode, qty) in enumerate(selected_products, start=1):
                print(f"{idx}. SKU: {sku}, Barcode: {barcode}, Quantity: {qty}")
            
            # Generate PDF report
            generate_pdf(selected_products, meta_data)

            
            create_output_folder(selected_products, meta_data)

        except ValueError as e:
            print(e)
    else:
        print("No file selected, exiting.")

# Entry point
main()

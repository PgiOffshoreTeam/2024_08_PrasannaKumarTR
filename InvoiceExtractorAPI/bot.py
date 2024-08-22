import re
from PyPDF2 import PdfReader
from botcity.plugins.googlesheets import BotGoogleSheetsPlugin
from botcity.plugins.googledrive import BotGoogleDrivePlugin

class Bot:
    def action(self, execution=None):
        # List of PDF files (Google Drive file IDs) to process
        pdf_file_ids = [
            "1mWMDyUk-5dJ55sxprJnElgefyf9aqFYo","19MdU-MHvX9DDvkn2k77kszrDQAzfj8M2" # Replace with actual Google Drive file IDs
            # Add more Google Drive file IDs here
        ]

        # Initialize Google Drive and Sheets plugins
        client_secret_path = r"C:\Users\Prasanna\Downloads\GoogleAPI.json"
        spreadsheet_id = "15xdUK0XJ2Cu3HlIK2Bn7b4G5z3A8bXkWknShyiXuuvE"  # Replace with your Google Sheet ID
        drive_plugin = BotGoogleDrivePlugin(client_secret_path)
        sheets_plugin = BotGoogleSheetsPlugin(client_secret_path, spreadsheet_id)

        # Function to extract text from PDF
        def extract_text_from_pdf(pdf_path):
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                return text

        # Function to map data
        def map_invoice_data(text):
            invoice_data = {}

            # Adjusted regex patterns
            invoice_number_match = re.search(r'INV-\d+', text)
            order_number_match = re.search(r'Order  Number\s+(\d+)', text)
            invoice_date_match = re.search(r'Invoice\s+Date\s+([A-Za-z]+) (\d{1,2}), (\d{4})', text)
            due_date_match = re.search(r'Due\s+Date\s+([A-Za-z]+) (\d{1,2}), (\d{4})', text)
            from_details_match = re.search(r'From:\s+([\s\S]+?)\s+To:', text)
            to_details_match = re.search(r'To:\s+([\s\S]+?)\s+(?:Hrs/Qty|$)', text)

            # Extracting email addresses
            from_email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', from_details_match.group(1)) if from_details_match else None
            to_email_match = re.search(r'([a-zA-Z0-9._%+-]+@\s+[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', to_details_match.group(1)) if to_details_match else None

            # Remove emails from 'From' and 'To' addresses
            from_address = re.sub(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', '', from_details_match.group(1)).strip() if from_details_match else 'N/A'
            to_address = re.sub(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', '', to_details_match.group(1)).strip() if to_details_match else 'N/A'

            # Extracting service details
            service_details_pattern = re.compile(r'(\d+\.\d{2})\s+([\w\s]+?)\s+(\d+\.\d{2})\s+(\d+\.\d{2}%)\s+\$(\d+\.\d{2})')
            service_details_matches = service_details_pattern.findall(text)

            service_entries = [{'Quantity/Hours': match[0], 'Description': match[1].strip(),
                                'Unit Price': match[2], 'Tax Rate': match[3], 'Amount': match[4]}
                               for match in service_details_matches]

            # Extracting financial details
            sub_total_match = re.search(r'Sub Total\s*\$([\d,]+\.\d+)', text)
            tax_match = re.search(r'Tax\s*\$([\d,]+\.\d+)', text)
            total_due_match = re.search(r'Total\s+Due\s+\$([\d,]+\.\d+)', text)

            # Storing extracted data
            invoice_data['Invoice Number'] = invoice_number_match.group(0) if invoice_number_match else 'N/A'
            invoice_data['Order Number'] = order_number_match.group(1) if order_number_match else 'N/A'
            invoice_data['Invoice Date'] = f"{invoice_date_match.group(1)} {invoice_date_match.group(2)}, {invoice_date_match.group(3)}" if invoice_date_match else 'N/A'
            invoice_data['Due Date'] = f"{due_date_match.group(1)} {due_date_match.group(2)}, {due_date_match.group(3)}" if due_date_match else 'N/A'
            invoice_data['From'] = from_address
            invoice_data['From Email'] = from_email_match.group(0) if from_email_match else 'N/A'
            invoice_data['To'] = to_address
            invoice_data['To Email'] = to_email_match.group(0) if to_email_match else 'N/A'
            invoice_data['Services'] = service_entries
            invoice_data['Sub Total'] = sub_total_match.group(1) if sub_total_match else 'N/A'
            invoice_data['Tax'] = tax_match.group(1) if tax_match else 'N/A'
            invoice_data['Total Due'] = total_due_match.group(1) if total_due_match else 'N/A'
            
            return invoice_data

        # Function to save mapped data to Google Sheets using BotCity Google plugin
        def save_mapped_data_to_google_sheets(data):
            # Define the headers for the sheet
            headers = ['Invoice Number', 'Order Number', 'Invoice Date', 'Due Date', 
                       'From', 'From Email', 'To', 'To Email', 
                       'Service 1 - Quantity/Hours', 'Service 1 - Description', 
                       'Service 1 - Unit Price', 'Service 1 - Tax Rate', 'Service 1 - Amount',
                       'Service 2 - Quantity/Hours', 'Service 2 - Description', 
                       'Service 2 - Unit Price', 'Service 2 - Tax Rate', 'Service 2 - Amount',
                       'Service 3 - Quantity/Hours', 'Service 3 - Description', 
                       'Service 3 - Unit Price', 'Service 3 - Tax Rate', 'Service 3 - Amount',
                       'Sub Total', 'Tax', 'Total Due']

            # Check if headers are present
            existing_data = sheets_plugin.get_range('A1:A1')  # Get a single cell to check if headers are there
            if not existing_data or existing_data[0][0] != headers[0]:
                # If headers are not present, add them
                sheets_plugin.add_row(headers)

            # Prepare the row to append
            row = [
                data['Invoice Number'],
                data['Order Number'],
                data['Invoice Date'],
                data['Due Date'],
                data['From'],
                data['From Email'],
                data['To'],
                data['To Email']
            ]
            
            # Append service entries to the row
            for service_entry in data['Services']:
                row.extend([
                    service_entry['Quantity/Hours'],
                    service_entry['Description'],
                    service_entry['Unit Price'],
                    service_entry['Tax Rate'],
                    service_entry['Amount']
                ])
            
            # Fill remaining service columns with empty strings if less than 3 services
            for _ in range(3 - len(data['Services'])):
                row.extend(['', '', '', '', ''])
            
            # Add sub total, tax, and total due
            row.extend([data['Sub Total'], data['Tax'], data['Total Due']])
            
            # Append the row to the Google Sheets
            try:
                sheets_plugin.add_row(row)
            except Exception as e:
                print(f"Error appending row to Google Sheets: {e}")

        # Process each PDF file
        for file_id in pdf_file_ids:
            # Define the local path to save the downloaded PDF
            pdf_path = f"./{file_id}.pdf"

            # Download or export the file from Google Drive
            try:
                drive_plugin.download_file(file_id, pdf_path)
            except Exception as e:
                if "fileNotDownloadable" in str(e):
                    # Export Google Docs files as PDFs
                    drive_plugin.export_file(file_id, pdf_path, mime_type='application/pdf')
                else:
                    raise

            # Extract text from the PDF
            extracted_text = extract_text_from_pdf(pdf_path)

            # Map the data
            mapped_data = map_invoice_data(extracted_text)

            # Save the mapped data to Google Sheets
            save_mapped_data_to_google_sheets(mapped_data)

        print("Extracted and mapped invoice data has been saved to Google Sheets.")

    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == "__main__":
    Bot().action()
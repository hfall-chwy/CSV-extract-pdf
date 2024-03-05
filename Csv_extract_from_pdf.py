import os
import pdfplumber
import csv
from pathlib import Path


pdf_folder_path = r'C:\Users\hfall\Downloads\testing_contract'
output_folder = r'C:\Users\hfall\Downloads\testing_contract'

def Job_Complete_Notification(message):
    print(message)

def extract_table_from_pdf(pdf_file):
    try:
        table_data = []

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        table_data.append(row)

        return table_data
    except Exception as e:
        raise Exception(f"Error extracting table from {pdf_file}: {str(e)}")

def save_as_csv(table_data, csv_file):
    try:
        with open(csv_file, 'w', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)
            for row in table_data:
                csv_writer.writerow(row)
    except Exception as e:
        raise Exception(f"Error saving table as CSV: {str(e)}")

def PDF_to_CSV(pdf_folder, output_folder):
    try:
        for pdf_file in Path(pdf_folder).glob("*.pdf"):
            csv_file = os.path.join(output_folder, os.path.splitext(pdf_file.name)[0] + ".csv")
            table_data = extract_table_from_pdf(pdf_file)
            save_as_csv(table_data, csv_file)

        status_message = "Files Finished Converting."
    except Exception as e:
        status_message = f"Error encountered during conversion: {str(e)}"

    Job_Complete_Notification(status_message)

PDF_to_CSV(pdf_folder_path, output_folder)

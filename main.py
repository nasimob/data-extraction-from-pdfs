
import re
import PyPDF2
import pdfplumber
import os
import shutil
from docx import Document



def extract_data_from_pdf(file_path):
    extracted_data = ""
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)

        for page in pdf_reader.pages:
            extracted_data+= page.extract_text()

    return extracted_data


# Extract text from the specified coordinates using pdfplumber
def extract_data_from_pdf_plumber(file_path, coordinates):
    extracted_data = []
    with pdfplumber.open(file_path) as pdf:
        for coord in coordinates:
            page_number = coord["page"]
            x1, x2, y1, y2 = coord["x1"], coord["x2"], coord["y1"], coord["y2"]
            # Check if the page number is valid
            if page_number <= len(pdf.pages):
                # Extract text from the specified coordinates using pdfplumber
                page = pdf.pages[page_number - 1]
                extracted_text = page.crop((x1, y1, x2, y2)).extract_text()

                # Append the extracted text to the results
                extracted_data.append(extracted_text)
            else:
                extracted_data.append('')


    return extracted_data



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    document = Document('orders_table.docx')
    table = document.tables[0]

    coord = [ {"page":1,"x1":165.97491373233032,"x2":221.73479883364868,"y1":100.73039268493653,"y2":121.65953269958497}]
    folder_path = 'PDFS'
    processed_folder = 'Done_pdfs'
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        # Skip if the file is not a PDF
        if not file_name.endswith('.pdf'):
            continue

        # Skip if the file has already been processed
        if os.path.exists(os.path.join(processed_folder, file_name)):
            continue

        extracted_text = extract_data_from_pdf(file_path)
        order_id = extract_data_from_pdf_plumber(file_path,coord)
        #id_matches = re.findall(r'(.+)מספר הזמנה', extracted_text , re.UNICODE)
        order_price = re.findall(r'(.+)סה"כ להזמנה', extracted_text , re.UNICODE)
        order_pay = re.findall(r'(.+)סה"כ לתשלום',extracted_text,re.UNICODE)
        order_spec = re.findall(r'(.+)כתובת החיבור:',extracted_text,re.UNICODE)

        float_price = None  # Initialize order_price
        float_pay = None  # Initialize order_pay

        for price in order_price:
            # Extract numeric part of the order price
            numeric_part = ''.join(filter(str.isdigit, price))
            if numeric_part:
                float_price = float(numeric_part) / 100

        for pay in order_pay:
            # Remove parentheses and extra spaces
            cleaned_pay = pay.replace(',', '').replace('(', '').replace(')', '').strip()
            if cleaned_pay:
                float_pay = float(cleaned_pay)

        for spec in order_spec:
            # Remove the list brackets
            cleaned_spec = spec.strip("[]")


        for id in order_id:
            cleaned_id = id.strip("[]")
        if float_price == None or float_pay == None :
            progress = 'לא התקבל חשבון'
        elif float_price * 0.16 > float_pay :
            progress = 'שלב א'
        elif float_price * 0.16 <= float_pay:
            progress = 'שלב ב'

        row = table.add_row().cells
        row[0].text = str(cleaned_spec)
        row[1].text = progress
        row[2].text = str(cleaned_id)
        row[3].text = ""
        document.save('orders_table.docx')
        print(cleaned_id)
        print(float_price)
        print(float_pay)
        print(cleaned_spec+ '\n')
        shutil.copy(file_path, os.path.join(processed_folder, file_name))



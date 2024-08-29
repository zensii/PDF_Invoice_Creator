import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob('xls_files\\*.xlsx')

for file in filepaths:
    df = pd.read_excel(file)
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=10, txt=f'Invoice Number: {file.split('\\')[1].split('-')[0]}', ln=1)
    pdf.cell(w=50, h=10, txt=f'Date: {file.split('\\')[1].split('-')[1].replace('.xlsx', '')}', ln=1)
    pdf.cell(w=45, h=10, txt="", ln=1)

    pdf.cell(w=45, h=10, txt="Product ID", border=1)
    pdf.cell(w=90, h=10, txt="Product Name", border=1)
    pdf.cell(w=45, h=10, txt="Amount", border=1)
    pdf.cell(w=45, h=10, txt="Price per Unit", border=1)
    pdf.cell(w=45, h=10, txt="Total Price", border=1, ln=1)

    pdf.set_font(family='Times', style='', size=16)

    for index, row in df.iterrows():
        pdf.cell(w=45, h=10, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=90, h=10, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=45, h=10, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=45, h=10, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=45, h=10, txt=f"{row['total_price']}", border=1, ln=1)

    pdf.cell(w=45, h=10, txt="", border=1)
    pdf.cell(w=90, h=10, txt=f"", border=1)
    pdf.cell(w=45, h=10, txt=f"", border=1)
    pdf.cell(w=45, h=10, txt=f"", border=1)
    pdf.cell(w=45, h=10, txt=f"{df['total_price'].sum()}", border=1, ln=1)

    pdf.cell(w=45, h=10, txt="", ln=1)
    pdf.cell(w=45, h=10, txt="", ln=1)

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=0, h=10, txt=f'The total due amount is {df['total_price'].sum()} Euros.')

    pdf.output('PDFs\\' + file.split('\\')[1].replace('.xlsx', '.pdf'))

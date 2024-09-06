import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob(r"C:\Users\mrokas\PycharmProjects\pdf_from_excel\*.xlsx") #everything on the filepath that ends on .xlsx

for filepath in filepaths:

    pdf=FPDF(orientation='P', unit='mm',format='A4')
    pdf.add_page()
    filename=Path(filepath).stem #you get the last entry of the excel so e.g 10001-2023.1.18
    invoice_nr,date =filename.split('-')

    pdf.set_font(family="Times", size=16,style='B')
    pdf.cell(w=50,h=8,txt=f"Invoice nr.{invoice_nr}",ln=1) #ln=1 creates a new line

    pdf.set_font(family="Times", size=16,style='B')
    pdf.cell(w=50,h=8,txt=f"Date.{date}",ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    columns=list(df.columns)
    pdf.set_font(family="Times", style="B",size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0].replace('_',' ').capitalize(), border=1)
    pdf.cell(w=70, h=8, txt=columns[1].replace('_',' ').capitalize(), border=1)
    pdf.cell(w=30, h=8, txt=columns[2].replace('_',' ').capitalize(), border=1)
    pdf.cell(w=30, h=8, txt=columns[3].replace('_',' ').capitalize(), border=1)
    pdf.cell(w=30, h=8, txt=columns[4].replace('_',' ').capitalize(), border=1, ln=1)
    for index,row in df.iterrows():
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]),border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]),border=1,ln=1)
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=70, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt=str(df['total_price'].sum()), border=1, ln=1)

    #add total amount
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt=f"The total amount due is {str(df['total_price'].sum())} euros",ln=1)

    #add company logo
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt=f"Company Name")
    pdf.image(r"C:\Users\mrokas\PycharmProjects\pdf_from_excel\pythonhow.png",w=10)


    pdf.output(f"PDFs/{filename}.pdf")



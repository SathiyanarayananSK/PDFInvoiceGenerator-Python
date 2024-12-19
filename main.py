import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No: {invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)


    # Add a header
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = list(df.columns)
    columns_names = [name.replace("_", " ") for name in columns]
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(columns_names[0].title()), border=1)
    pdf.cell(w=60, h=8, txt=str(columns_names[1].title()), border=1)
    pdf.cell(w=40, h=8, txt=str(columns_names[2].title()), border=1)
    pdf.cell(w=30, h=8, txt=str(columns_names[3].title()), border=1)
    pdf.cell(w=30, h=8, txt=str(columns_names[4].title()), border=1, ln=1)

    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row[columns[0]]), border=1)
        pdf.cell(w=60, h=8, txt=str(row[columns[1]]), border=1)
        pdf.cell(w=40, h=8, txt=str(row[columns[2]]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[columns[3]]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[columns[4]]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=28, h=8, txt="PythonHow")
    pdf.image("images/pythonhow.png", w=10)



    pdf.output(f"PDFs/{filename}.pdf")






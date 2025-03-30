import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
from datetime import date

#Get a list of all Excel files in the "invoices" folder
file_paths = glob.glob("invoices/*.xlsx")

#print(files)
for file_path in file_paths:
    #load each Excel file
    df = pd.read_excel(file_path, sheet_name="Sheet 1")
    print(df)
    #Create pdf file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    #get filename from the file
    filename = Path(file_path).stem
    #splits the invoice name where there is - in 2 parts
    # and [0] gives the 1st part of name
    invoice_nr, invoice_date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}")
    pdf.output(f"PDFs/{filename}.pdf")

#print(df)
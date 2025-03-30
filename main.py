import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
from datetime import date

#Get a list of all Excel files in the "invoices" folder
file_paths = glob.glob("invoices/*.xlsx")

#print(files)
for file_path in file_paths:

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
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)
    pdf.ln(15)

    # load each Excel file
    df = pd.read_excel(file_path, sheet_name="Sheet 1")

    #Add header to the table
    columns= df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=45, h=8, txt=columns[2], border=1)
    pdf.cell(w=40, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    #Add row items to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=45, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln = 1)

    #Total sum of all items calculated
    total_sum = df["total_price"].sum()

    # Total sum added to the table
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=45, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.ln(15)

    #Add total sum sentence
    pdf.set_font(family="Times", size=17)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=f"The total due amount is {total_sum}  Euros", ln=1)
    pdf.ln(2)

    #Add logo of the company
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=35, h=8, txt=f"BoringCompany")
    pdf.image("pythonhow.png", x=62, y=99, w=10)

    pdf.output(f"PDFs/{filename}.pdf")
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=12, txt=f"Invoice nr. {invoice_nr}", align="L",
             ln=1, border=0)
    pdf.cell(w=0, h=12, txt=f"Date: {date}", align="L",
             ln=1, border=0)
    pdf.cell(w=0, h=12, txt="", align="L",
             ln=1, border=0)

    headers = df.columns.values.tolist()

    for header in headers:
        width=30
        if header=="product_name":
            width=70
        if header=="amount_purchased":
            header="Amount"
        better_header = header.replace("_", " ").title()
        if header == headers[len(headers) - 1]:
            pdf.cell(w=width, h=8, txt=better_header, border=1, ln=1)
            break
        pdf.set_font(family="Times", size=10, style="B")
        pdf.set_text_color(10,10,10)
        pdf.cell(w=width, h=8, txt=better_header, border=1)


    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        for header in headers:
            width = 30
            if len(str(row[header])) > 10:
                width=70
            if header == headers[len(headers)-1]:
                pdf.cell(w=width, h=8, txt=str(row[header]), border=1, ln=1)
                break
            pdf.cell(w=width, h=8, txt=str(row[header]), border=1)

    for header in headers:
        width = 30
        if header == "product_name":
            width = 70
        if header == headers[len(headers) - 1]:
            pdf.cell(w=width, h=8, txt=str(df["total_price"].sum()), border=1, ln=1)
            break
        pdf.cell(w=width, h=8, txt="", border=1)

    pdf.output(f"PDFs/{filename}.pdf")
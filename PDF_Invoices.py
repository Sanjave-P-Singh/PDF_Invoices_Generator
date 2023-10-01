import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

# Iterate through a list of file paths
for filepath in filepaths:

    # Create a PDF object with specified orientation, unit, and format
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # Add a new page to the PDF document
    pdf.add_page()

    # Extract the filename without extension from the file path
    filename = Path(filepath).stem
    # Extract invoice number from the filename by splitting at "-"
    invoice_nr, date = filename.split("-")
    # date = filename.split("-")[1] <-- Can also be used

    # Set font properties for the PDF content
    pdf.set_font(family="Times", size=16, style="B")
    # Add a cell to the PDF with the formatted invoice number
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_nr}", ln=1)

    # Set font properties for the PDF content
    pdf.set_font(family="Times", size=14, style="B")
    # Add a cell to the PDF with the formatted date
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Read data from an Excel file into a DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1, align='C')  # Center aligned text
    pdf.cell(w=70, h=8, txt=columns[1], border=1, align='C')  # Center aligned text
    pdf.cell(w=30, h=8, txt=columns[2], border=1, align='C')  # Center aligned text
    pdf.cell(w=30, h=8, txt=columns[3], border=1, align='C')  # Center aligned text
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1, align='C')  # Center aligned text

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1, align='C')  # Center aligned text
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1, align='C')  # Center aligned text
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1, align='C')  # Center aligned text
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1, align='C')  # Center aligned text
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1, align='C')  # Center aligned text

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1, align='C')  # Center aligned text
    pdf.cell(w=70, h=8, txt="", border=1, align='C')  # Center aligned text
    pdf.cell(w=30, h=8, txt="", border=1, align='C')  # Center aligned text
    pdf.cell(w=30, h=8, txt="", border=1, align='C')  # Center aligned text
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1, align='C')  # Center aligned text

    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is: {total_sum}",ln=1)

    # Add Company Name and Logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    # Output the PDF file to the specified directory with the formatted filename
    pdf.output(f"PDFs/{filename}.pdf")

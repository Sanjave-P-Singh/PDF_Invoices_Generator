import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

# Iterate through a list of file paths
for filepath in filepaths:
    # Read data from an Excel file into a DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

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
    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    # Output the PDF file to the specified directory with the formatted filename
    pdf.output(f"PDFs/{filename}.pdf")


from fileinput import filename
import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

# Get list of files
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    #print(df)
    # Create 1 PDF Document per excel file
    pdf = FPDF(orientation="P", unit="mm", format="A4") # 'P' stands for Portrait, 'L' for Landscape
    pdf.add_page()
    ## Get filename without the extension
    filename = Path(filepath).stem
    invoice_num = filename.split("-")[0]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, text=f"Invoice #: {invoice_num}")
    pdf.output(f"PDFs/{filename}.pdf")

    for index, row in df.iterrows():
        #print(index)
        print(row)
        #print(df.shape)
        # Define the number of rows and columns
        rows_per_page = int(index) + 1
        cols_per_page = len(df.columns)
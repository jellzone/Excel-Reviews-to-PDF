import pandas as pd
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import html

# Load the Excel file
excel_file = 'example.xlsx'
sheet_name = 'Sheet1'
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Set up the paragraph style
style = getSampleStyleSheet()['BodyText']
style.leading = 24

# Constants
rows_per_page = 8  # Adjust this value based on your specific requirements
max_pages_per_pdf = 1000
rows_per_pdf = rows_per_page * max_pages_per_pdf
total_rows = len(df.index)
num_pdfs = (total_rows // rows_per_pdf) + (1 if total_rows % rows_per_pdf > 0 else 0)

# Function to create a PDF with given rows
def create_pdf(pdf_name, start_row, end_row):
    doc = SimpleDocTemplate(pdf_name, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
    doc.pagesize = landscape(letter)
    elements = []

    for index, row in df.iloc[start_row:end_row].iterrows():
        text = f"{index + 1}. {html.escape(str(row['review']))}"
        elements.append(Paragraph(text, style))
        elements.append(Spacer(1, 12))

    # Build the PDF
    doc.build(elements)

# Create PDFs
for i in range(num_pdfs):
    pdf_name = f"reviews_part{i + 1}.pdf"
    start_row = i * rows_per_pdf
    end_row = min((i + 1) * rows_per_pdf, total_rows)
    create_pdf(pdf_name, start_row, end_row)

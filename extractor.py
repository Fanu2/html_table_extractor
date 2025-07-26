import pandas as pd
from io import BytesIO
from docx import Document

def extract_all_tables(html):
    """Extract all tables from HTML and return list of DataFrames."""
    try:
        tables = pd.read_html(html)  # Parses all tables from the HTML
        return tables
    except Exception:
        return []

def df_to_docx(df):
    """Convert a DataFrame into a Word (.docx) file and return as BytesIO."""
    doc = Document()

    # Add table with column headers
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells

    for i, col in enumerate(df.columns):
        hdr_cells[i].text = str(col)

    # Add rows of data
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, item in enumerate(row):
            row_cells[i].text = str(item)

    # Save to BytesIO
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

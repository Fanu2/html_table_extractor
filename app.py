import pandas as pd
from bs4 import BeautifulSoup
from io import BytesIO
from docx import Document

def extract_all_tables(html):
    """Extract all tables from HTML and return list of DataFrames."""
    try:
        tables = pd.read_html(html)
        return tables
    except Exception:
        return []

def df_to_docx(df):
    doc = Document()
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'

    # Add headers
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(df.columns):
        hdr_cells[i].text = str(col)

    # Add data rows
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, item in enumerate(row):
            row_cells[i].text = str(item)

    # Output as BytesIO
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

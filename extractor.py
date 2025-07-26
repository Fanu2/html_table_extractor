import pandas as pd
from bs4 import BeautifulSoup
from io import BytesIO
from docx import Document

def html_table_to_df(html_content, max_columns=7):
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')
    if not table:
        return None

    rows = table.find_all('tr')
    data = [[col.get_text(strip=True) for col in row.find_all(['td', 'th'])] for row in rows]
    df = pd.DataFrame(data).iloc[:, :max_columns]
    return df

def df_to_docx(df):
    doc = Document()
    table = doc.add_table(rows=1, cols=len(df.columns))
    for i, col_name in enumerate(df.iloc[0]):
        table.rows[0].cells[i].text = str(col_name)

    for _, row in df.iloc[1:].iterrows():
        row_cells = table.add_row().cells
        for i, item in enumerate(row):
            row_cells[i].text = str(item)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
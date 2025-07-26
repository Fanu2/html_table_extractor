import streamlit as st
from extractor import extract_all_tables, df_to_docx

st.title("HTML Table to Word Converter")

uploaded_file = st.file_uploader("Upload HTML File", type=["html", "htm"])

if uploaded_file:
    try:
        content = uploaded_file.read().decode('utf-8')
        tables = extract_all_tables(content)

        if not tables:
            st.error("No tables found in the uploaded HTML file.")
        else:
            st.success(f"Found {len(tables)} table(s)")

            table_index = st.selectbox("Select Table", range(len(tables)), format_func=lambda i: f"Table {i+1}")
            max_cols = len(tables[table_index].columns)
            max_rows = len(tables[table_index])

            skip_rows = st.number_input("Rows to skip (from top)", min_value=0, max_value=max_rows - 1, value=1)
            num_columns = st.number_input("Number of columns to import", min_value=1, max_value=max_cols, value=8)

            # Slice the selected DataFrame
            df = tables[table_index].iloc[skip_rows:, :num_columns]

            st.write("📄 Extracted Table Preview:")
            st.dataframe(df)

            docx_file = df_to_docx(df)
            st.download_button(
                label="⬇ Download as Word (.docx)",
                data=docx_file,
                file_name="extracted_table.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    except Exception as e:
        st.error(f"❌ An error occurred: {e}")

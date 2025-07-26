import streamlit as st
from extractor import extract_all_tables, df_to_docx

st.title("HTML Table to Word Converter")

uploaded_file = st.file_uploader("Upload HTML File", type=["html", "htm"])

if uploaded_file:
    try:
        content = uploaded_file.read().decode('utf-8')
        tables = extract_all_tables(content)

        # Filter out empty or invalid tables
        tables = [t for t in tables if not t.empty and t.shape[1] >= 5]

        if not tables:
            st.error("No valid tables found in the uploaded HTML file.")
        else:
            # Use the largest table by number of rows
            table = max(tables, key=len)
            max_cols = len(table.columns)
            max_rows = len(table)

            st.success(f"Found {len(tables)} table(s). Using the largest one with {max_rows} rows and {max_cols} columns.")

            if max_rows <= 1:
                skip_rows = 0
                st.info("Table has 1 row or less — skipping rows is not needed.")
            else:
                skip_rows = st.number_input(
                    "Rows to skip (from top)",
                    min_value=0,
                    max_value=max(max_rows - 1, 0),
                    value=1
                )

            num_columns = st.number_input(
                "Number of columns to import",
                min_value=1,
                max_value=max_cols,
                value=min(8, max_cols)
            )

            df = table.iloc[skip_rows:, :num_columns]

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

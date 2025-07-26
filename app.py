import streamlit as st
from extractor import html_table_to_df, df_to_docx

st.title("HTML Table to Word Converter")

uploaded_file = st.file_uploader("Upload HTML File", type=["html", "htm"])

if uploaded_file:
    try:
        content = uploaded_file.read().decode('utf-8')
        df = html_table_to_df(content)

        if df is not None:
            st.write("Extracted Table (First 7 columns):")
            st.dataframe(df)

            docx_file = df_to_docx(df)
            st.download_button(
                label="Download as Word (.docx)",
                data=docx_file,
                file_name="extracted_table.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("No table found in the uploaded HTML file.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
import streamlit as st
import pandas as pd
import base64
import io

def process_excel(file):
    df = pd.read_excel(file)
    if 'Name' in df.columns and 'Email' in df.columns:
        return df[['Name', 'Email']]
    else:
        st.error("Excel file must contain 'Name' and 'Email' columns.")

def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def main():
    st.set_page_config(page_title="Mentor Matching App", page_icon="IUF_Marketing_Lockup_V_WEB_REV.png") 
    logo_image = 'IUF_Marketing_Lockup_V_WEB_REV.png'

    st.image(logo_image, use_column_width=True)
    st.title("Mentor Matching")
    st.subheader("Upload an Excel file and extract Names and Emails")
    

    file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'], accept_multiple_files=False)
    
    if file is not None:
        st.write("Uploaded file:", file.name)
        
        df = process_excel(file)
        if df is not None:
            st.write("Extracted Names and Emails:")
            st.write(df)

            csv = df.to_csv(index=False)
            b64 = base64.b64encode(csv.encode()).decode()
            href = f'<a href="data:file/csv;base64,{b64}" download="output.csv">Download CSV File</a>'
            st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd

# Function to process Excel file and extract names and emails
def process_excel(file):
    df = pd.read_excel(file)
    if 'Name' in df.columns and 'Email' in df.columns:
        return df[['Name', 'Email']]
    else:
        st.error("Excel file must contain 'Name' and 'Email' columns.")

def main():
    st.title("Mentor Matching")  # Adding a title
    st.subheader("Upload an Excel file and extract Names and Emails")
    

    file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'], accept_multiple_files=False)
    
    if file is not None:
        st.write("Uploaded file:", file.name)
        
        df = process_excel(file)
        if df is not None:
            st.write("Extracted Names and Emails:")
            st.write(df)

if __name__ == "__main__":
    main()

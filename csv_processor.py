
import streamlit as st
import pandas as pd


def process_excel(file):
    df = pd.read_excel(file)
    if 'Name' in df.columns and 'Email' in df.columns:
        return df[['Name', 'Email']]
    else:
        st.error("Excel file must contain 'Name' and 'Email' columns.")

def main():
    st.title("Mentor Matching Email Processer")
    st.subheader("Upload a CSV file and extract Names and Emails")


    file = st.file_uploader("Upload CSV file", type=['csv'], accept_multiple_files=False)

    if file is not None:
        st.write("Uploaded file:", file.name)
        
        # Process the uploaded file
        df = process_excel(file)
        if df is not None:
            # Display the extracted names and emails in a table
            st.write("Extracted Names and Emails:")
            st.write(df)

if __name__ == "__main__":
    main()

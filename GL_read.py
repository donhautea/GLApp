import streamlit as st
import pandas as pd

# Define the column names for columns A to Z
COLUMN_NAMES = [
    "schemeid", "Scheme Short Name", "Journal Entry No", "Document Type", "Document Date",
    "Posting Date", "Company Code", "Reference", "Document Header Text", "Amount", "Currency",
    "Posting Key", "Cross Company", "GL Account", "Vendor ID", "Customer ID", "Tax Code",
    "Withholding Tax Type", "Withholding Tax Code", "Withholding Tax Base Amount", "Cost Center",
    "Profit Center", "Order", "Fund Center", "Assignment", "Text"
]

def read_and_process_file(file):
    # Read the Excel file with the specified columns and start at row 2
    df = pd.read_excel(file, usecols="A:Z", header=1, names=COLUMN_NAMES)
    
    # Adjust Amount if Posting Key is 50
    df.loc[df['Posting Key'] == 50, 'Amount'] *= -1
    
    return df

def main():
    st.title("Excel File Consolidator for SAP GL")

    # Sidebar for file selection
    st.sidebar.header("File Selection")
    uploaded_files = st.sidebar.file_uploader("Choose Excel files", accept_multiple_files=True, type="xlsx")

    if uploaded_files:
        combined_data = pd.DataFrame()
        
        for file in uploaded_files:
            df = read_and_process_file(file)
            combined_data = pd.concat([combined_data, df], ignore_index=True)

        # Sum Amount if Scheme ID, GL Account, and Document Type are the same
        combined_data = combined_data.groupby(['schemeid', 'GL Account', 'Document Type'], as_index=False).agg({
            'Scheme Short Name': 'first',
            'Journal Entry No': 'first',
            'Document Date': 'first',
            'Posting Date': 'first',
            'Company Code': 'first',
            'Reference': 'first',
            'Document Header Text': 'first',
            'Amount': 'sum',
            'Currency': 'first',
            'Posting Key': 'first',
            'Cross Company': 'first',
            'Vendor ID': 'first',
            'Customer ID': 'first',
            'Tax Code': 'first',
            'Withholding Tax Type': 'first',
            'Withholding Tax Code': 'first',
            'Withholding Tax Base Amount': 'first',
            'Cost Center': 'first',
            'Profit Center': 'first',
            'Order': 'first',
            'Fund Center': 'first',
            'Assignment': 'first',
            'Text': 'first'
        })

        # Make Amount an absolute number
        combined_data['Amount'] = combined_data['Amount'].abs()

        st.header("Combined Data")
        st.dataframe(combined_data)
        
        # Option to download the combined data
        st.sidebar.header("Download Combined Data")
        if st.sidebar.button("Download"):
            output_file = st.sidebar.text_input("Enter the output file name (without extension):", "combined_data")
            combined_data.to_excel(f"{output_file}.xlsx", index=False)
            st.sidebar.markdown(f"[Download {output_file}.xlsx](./{output_file}.xlsx)")

if __name__ == "__main__":
    main()

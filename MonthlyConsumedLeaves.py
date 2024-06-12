import streamlit as st
import pandas as pd
import base64
import io
import openpyxl

# # Function to create download link for DataFrame
def get_download_link(df, label):
    # Reset index if MultiIndex columns are present
    if isinstance(df.columns, pd.MultiIndex):
        df = df.reset_index()
    # Convert DataFrame to XLSX file in memory
    output = io.BytesIO()
    df.to_excel(output, engine='xlsxwriter')
    xlsx_data = output.getvalue()
    # Encode Excel data as base64 string
    b64 = base64.b64encode(xlsx_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="report_{label}.xlsx">{label}</a>'
    return href

# Main Streamlit app
def main():
    st.set_page_config(page_title="Monthly Consumed Leaves")
    st.markdown("<h1 style='text-align: center; font-size: 45px;'>All Types Report</h1>", unsafe_allow_html=True)
    st.title("Upload Nazih Leaves")
  
    uploaded_file = st.file_uploader("Choose a file", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("Uploaded DataFrame:")
            st.write(df)
            
            # Data processing
            df = df[~df['Category meaning'].str.contains('-', na=False)]
            df = df.dropna(subset=['Full Name'])
            df = df[df['Person Employmnt Type'] != 'Ex-employee']
            df['Sector'] = df.apply(lambda x: x['Central Department'] if '-' in str(x['Sector']) else x['Sector'], axis=1)
            df['Required Date Starting'] = pd.to_datetime(df['Required Date Starting'])
            df['Required Date Ending'] = pd.to_datetime(df['Required Date Ending'])

            # Create a DataFrame with a row for each day in the range for each employee
            date_ranges = []
            for _, row in df.iterrows():
                sector = row['Sector']
                start_date = row['Required Date Starting']
                end_date = row['Required Date Ending']
                absence_type = row['Absense type']
                for single_date in pd.date_range(start_date, end_date):
                    date_ranges.append({'Sector': sector, 'Absense type': absence_type, 'date': single_date})

            expanded_df = pd.DataFrame(date_ranges)
            expanded_df['year_month'] = expanded_df['date'].dt.strftime('%m %B')

            # Group by Sector, Absense type, and year_month to count days
            test = expanded_df[(expanded_df['Absense type'] != 'أجازة عارضة') & (expanded_df['Absense type'] != 'أجازة اعتيادية')]

            # result = expanded_df.groupby(['Sector', 'Absense type', 'year_month']).size().reset_index(name='days_count')

            pivot_table1 = pd.pivot_table(test, index=['Sector'], columns=['year_month', 'Absense type'], aggfunc='size', fill_value=0)
            # Suppose pivot_table is your DataFrame with multi-index columns
            pivot_table1.reset_index(inplace=True)

            pivot_table2 = pd.pivot_table(expanded_df, index=['Sector', 'Absense type'], columns='year_month', aggfunc='size', fill_value=0)

            pivot_table3 = pd.pivot_table(expanded_df, index=['year_month', 'Absense type'], columns='Sector', aggfunc='size', fill_value=0)

            # Add labels for each download link
            download_links = {
                'Download Report - Sectors as rows': pivot_table1,
                'Download Report - Sectors and Absense type as rows': pivot_table2,
                'Download Report - Month and Absense type as rows': pivot_table3

                # Add more download links as needed
            }
            # Download button
            # Display download links
            for label, data in download_links.items():
                download_link = get_download_link(data, label)
                st.markdown(download_link, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

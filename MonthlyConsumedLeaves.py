import streamlit as st
import pandas as pd
import base64
import io
import openpyxl
import os

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
    st.markdown("<h1 style='text-align: left; font-size: 35px;'>Monthly Consumed Leaves Report</h1>", unsafe_allow_html=True)

  
    uploaded_file = st.file_uploader("Upload Monthly Report Leaves Data as .xlsx from discoverer", type=["xlsx"])

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

            pivot_table1 = pd.pivot_table(test, index=['Sector'], columns=['year_month', 'Absense type'], aggfunc='size', fill_value=0)

            pivot_table2 = pd.pivot_table(test, index=['Sector', 'Absense type'], columns='year_month', aggfunc='size', fill_value=0)

            pivot_table3 = pd.pivot_table(test, index=['year_month', 'Absense type'], columns='Sector', aggfunc='size', fill_value=0)

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


    uploaded_file2 = st.file_uploader("Upload Annual leaves Report from discoverer", type=["xlsx"])

    if uploaded_file2 is not None:
        try:
            df2 = pd.read_excel(uploaded_file2)
            st.write("Uploaded DataFrame:")
            st.write(df2)
            
            # Data processing
            df2 = df2[~df2['Category meaning'].str.contains('-', na=False)]
            df2 = df2.dropna(subset=['Full Name'])
            df2 = df2[df2['Person Employmnt Type'] != 'Ex-employee']
            df2['Sector'] = df2.apply(lambda x: x['Central Department'] if '-' in str(x['Sector']) else x['Sector'], axis=1)
            df2['Required Date Starting'] = pd.to_datetime(df2['Required Date Starting'])
            df2['Required Date Ending'] = pd.to_datetime(df2['Required Date Ending'])


            # Define the list of specific dates to exclude
            exclude_dates = [
                pd.Timestamp('2024-01-01'),
                pd.Timestamp('2024-01-07'),
                pd.Timestamp('2024-01-25'),
                pd.Timestamp('2024-04-09'),
                pd.Timestamp('2024-04-10'),
                pd.Timestamp('2024-04-11'),
                pd.Timestamp('2024-04-14'),
                pd.Timestamp('2024-04-25'),
                pd.Timestamp('2024-06-16'),
                pd.Timestamp('2024-06-17'),
                pd.Timestamp('2024-06-18'),
                pd.Timestamp('2024-06-19'),
                pd.Timestamp('2024-06-20'),
                pd.Timestamp('2024-05-05'),# Add specific dates here
                pd.Timestamp('2024-05-06')
            ]
            # Create a DataFrame with a row for each day in the range for each employee
            date_ranges2 = []
            for _, row2 in df2.iterrows():
                sector2 = row2['Sector']
                start_date2 = row2['Required Date Starting']
                end_date2 = row2['Required Date Ending']
                absence_type2 = row2['Absense type']
                for single_date2 in pd.date_range(start_date2, end_date2):
                    if single_date2.weekday() not in [4, 5] and single_date2 not in exclude_dates:
                        date_ranges2.append({'Sector': sector2, 'Absense type': absence_type2, 'date': single_date2})



                    # date_ranges2.append({'Sector': sector2, 'Absense type': absence_type2, 'date': single_date2})

            expanded_df2 = pd.DataFrame(date_ranges2)
            expanded_df2['year_month'] = expanded_df2['date'].dt.strftime('%m %B')

            # Group by Sector, Absense type, and year_month to count days
            test2 = expanded_df2[(expanded_df2['Absense type'] == 'أجازة عارضة') | (expanded_df2['Absense type'] == 'أجازة اعتيادية')]
            result = test2.groupby(['Sector', 'year_month'], as_index=False).size()

            pivot_table12 = pd.pivot_table(result, index=['Sector'], columns=['year_month'], aggfunc='sum', fill_value=0)


            # Add labels for each download link
            download_links2 = {
                'Download Report - Sectors Annual leaves consumption': pivot_table12,

                # Add more download links as needed
            }
            # Download button
            # Display download links
            for label2, data2 in download_links2.items():
                download_link2 = get_download_link(data2, label2)
                st.markdown(download_link2, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

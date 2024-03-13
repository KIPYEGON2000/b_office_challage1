import streamlit as st
import xlwings as xw
import pandas as pd
import io


def write_to_excel(data):
    # Connect to Excel
    wb = xw.Book()

    # Write data to Excel
    wb.sheets['Sheet1'].range('A1').value = data
    wb.sheets['Sheet1'].range('M10').value = data

    # Format the ranges
    format_range(wb.sheets['Sheet1'].range('B1:J6'))  
    format_range(wb.sheets['Sheet1'].range('N10:V15'))  

    # Format the header ranges
    format_header(wb.sheets['Sheet1'].range('A1:J1')) 
    format_header(wb.sheets['Sheet1'].range('M10:V10')) 

    # Set Century Gothic font for all cells with font size 10
    set_font(wb.sheets['Sheet1'].range('A1:V15'), 'Century Gothic', font_size=10)

    # Add borders
    add_borders(wb.sheets['Sheet1'].range('A1:J6'))
    add_borders(wb.sheets['Sheet1'].range('M10:V15'))

    # Auto fit columns and remove gridlines
    auto_fit_and_remove_gridlines(wb.sheets['Sheet1'])

    # Save the Excel file
    wb.save('output.xlsx')

def format_range(range_to_format):
    range_to_format.api.Font.Bold = False
    range_to_format.api.Font.Color = 5321497  
    range_to_format.api.Interior.Color = 16314344 

def format_header(header_to_format):
    header_to_format.api.Font.Bold = True
    header_to_format.api.Font.Color = 16777215
    header_to_format.api.Interior.Color = 5321497 

def set_font(range_to_format, font_name, font_size):
    range_to_format.api.Font.Name = font_name
    range_to_format.api.Font.Size = font_size

def add_borders(range_to_format):
    border_line_style = -4142  # xlContinuous
    border_color = 15389115 
    for i in range(1, 5):
        border = range_to_format.api.Borders(i)
        border.LineStyle = border_line_style
        border.Color = border_color

def auto_fit_and_remove_gridlines(sheet):
    sheet.api.Columns.AutoFit()
    sheet.api.Rows.AutoFit()
    sheet.api.Application.ActiveWindow.DisplayGridlines = False

def main():
    st.title('Excel Writer App')

    # Get user input
    data = pd.read_csv(r"E:\KPLC_backoffice\complete_advance\diff_feeders_suplicates.csv").reset_index(drop=True)
   

    # if st.button('Write to Excel'):
    #     write_to_excel(data.head())
    #     st.success('Data written to Excel successfully!')


    def download_excel():
        # Create an in-memory Excel file
        excel_buffer = io.BytesIO()
        # Write the DataFrame to the Excel file
        data.to_excel(excel_buffer, index=False)
        # Set the cursor to the beginning of the file
        excel_buffer.seek(0)
        # Return the Excel file as bytes
        return excel_buffer.getvalue()

    # Create a download button
    if st.button('Download Excel'):
        excel_data = download_excel()
        st.download_button(
            label="Download",
            data=excel_data,
            file_name='data.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()
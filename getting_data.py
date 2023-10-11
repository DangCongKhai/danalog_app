# Import all necessary library
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from icecream import ic  # Use this module for debugging
import pandas as pd
from io import BytesIO


pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)

with open('spread_sheetId.txt', mode='r') as file:
    SAMPLE_SPREADSHEET_ID = file.read()
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Define the scope for editing and reading
SERVICE_ACCOUNT_FILE = 'keys.json'
SAMPLE_RANGE_NAME = 'DataTaiXe!A1:K'


def setup_credentials():
    """This function is used to set up the credentials"""
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return creds


def connectToSheet(creds):
    """This function is used to connect to Google Sheet"""
    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    except HttpError as e:
        ic(e)
    else:
        return result


def create_data_frame(result) -> pd.DataFrame:
    """ This function creates a data frame from the result of google sheet"""
    values = result.get('values', [])
    if not values:
        ic('No value is found. We cannot create a data frame')
    else:
        # Find the maximum number of columns in any row
        max_columns = max([len(row) for row in values])
        # Pad rows with fewer columns with None
        values = [row + [None] * (max_columns - len(row)) for row in values]
        # Create a data frame
        df = pd.DataFrame(data=values[1:], columns=values[0])
        df = df.drop(columns=['Dấu thời gian'], axis=1)
        # df['Ngày'] = pd.to_datetime(df['Ngày'])
        # Convert the 'Ngày' column to datetime format with the correct format
        df['Ngày'] = pd.to_datetime(df['Ngày'], format="%d/%m/%Y",errors='coerce')

        # Format the 'Ngày' column as "%d-%m-%Y" and assign it back to the column

        columns = df.columns.to_list()
        columns.remove('Ngày')
        new_columns = ['Ngày']+columns
        df = df[new_columns]
        # Set new columns
        df.columns = ['Ngày', 'Tên Tài Xế', 'Biển số xe', 'Container No.', 'Size ', 'S/C', 'N/X', 'Tuyến đường',
                      'Lưu đêm', 'Ghi chú']
        column1 = 'Tổng doanh thu (v/c + nâng hạ + đường biển,…)'
        column2 = 'Dthu vận chuyển (chưa VAT)'
        df.insert(0, column='STT', value=None)
        df.insert(8,column=column1, value=None)
        df.insert(9,column=column2, value=None)
        list_number_type=[column1, column2, 'Size ', 'S/C', 'Lưu đêm']
        df[list_number_type]=df[list_number_type].astype(float)
        return df


def data_extracting(df, *arg) -> pd.DataFrame:
    """This function is used to extract the data we want"""
    name = arg[0]
    start_date = arg[1]
    end_date = arg[2]
    # df['Ngày'] = pd.to_datetime(df['Ngày'], format="%d/%m/%Y",errors='coerce')  # Convert to date format like 10/5/2005
    new_df = df[(df['Tên Tài Xế'] == name) & (df['Ngày'].dt.date >= start_date) & (df['Ngày'].dt.date <= end_date)]
    # new_df['Ngày'] = new_df['Ngày'].dt.strftime('%d/%m/%Y')
    return new_df


def to_excel(data):
    """This function is used """
    output = BytesIO()  # Create binary stream
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    data.to_excel(writer,startrow=2, index=False, sheet_name='Tên Biển Số')
    workbook = writer.book
    worksheet = writer.sheets['Tên Biển Số']

    # Loop through each column to adjust the width
    for col_num, col_data in enumerate(data.columns):
        column_len = max(data[col_data].astype('str').apply(len).max(), len(col_data))
        worksheet.set_column(col_num, col_num, column_len)

    format1 = workbook.add_format({'num_format': '0'})
    worksheet.set_column('A:A', None, format1)
    writer.close() # Save is deprecated
    processed_data = output.getvalue()
    return processed_data







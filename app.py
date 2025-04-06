# Import all necessary library
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from icecream import ic  # Use this module for debugging
import pandas as pd
from io import BytesIO
import streamlit as st
import gspread
import streamlit_authenticator as stauth
import pickle # Use this library for serialization of the object
from pathlib import Path
import threading
#  Set display
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)

with open('spread_sheetId.txt', mode='r') as file:
    SAMPLE_SPREADSHEET_ID = file.read()
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Define the scope for editing and reading
SERVICE_ACCOUNT_FILE = 'credential.json'
SAMPLE_RANGE_NAME = 'Data!A1:AB'

def main():
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
            st.write('Bạn không thể kết nối đến Google Sheet. Vui lòng thử lại hoặc liên lạc với người làm app')
        else:
            return result


    def create_data_frame(result) -> pd.DataFrame:
        """ This function creates a data frame from the result of google sheet"""
        values = result.get('values', [])
        if not values:
            st.write('Không tìm thấy dữ liệu trên google sheet. Không thể tạo bảng')
        else:
            # Find the maximum number of columns in any row
            max_columns = max([len(row) for row in values])
            # Pad rows with fewer columns with None
            values = [row + [None] * (max_columns - len(row)) for row in values]
            # Create a data frame
            df = pd.DataFrame(data=values[1:], columns=values[0])
            # Drop unnecessary columns on google sheet
            deleted_columns = ['Dấu thời gian', 'Nhập Mật Khẩu Danalog Cấp', 'Nơi đến']
            df = df.drop(columns=deleted_columns, axis=1)

            # Convert the 'Ngày' column to datetime format with the correct format
            df['Ngày'] = pd.to_datetime(df['Ngày'], format="%d/%m/%Y", errors='coerce')
            column1 = 'Tổng doanh thu (v/c + nâng hạ + đường biển,…)'
            column2 = 'Dthu vận chuyển (chưa VAT)'
            dem_position = df.columns.get_loc('Lưu đêm')
            df.insert(0, column='STT', value=None)
            df.insert(dem_position - 2, column=column1, value=None)
            df.insert(dem_position - 1, column=column2, value=None)
            list_number_type = [column1, column2, 'Doanh thu', 'Số chuyến', 'Lưu đêm']

            df[list_number_type] = df[list_number_type].astype(float)
            df['Size'] = pd.to_numeric(df['Size'], errors='coerce')
            # df[df['Size']!='']['Size']=df[df['Size']!='']['Size'].astype(int)

            df = df[df['Nơi Đến'] != '#N/A']
            return df


    def data_extracting(df, *arg) -> pd.DataFrame:
        """This function is used to extract the data we want based on the condition given"""
        name = arg[0]
        start_date = arg[1]
        end_date = arg[2]
        # df['Ngày'] = pd.to_datetime(df['Ngày'], format="%d/%m/%Y",errors='coerce')  # Convert to date format like 10/5/2005

        true_df = df['Tên Tài Xế'].isin(name)
        name_df = df[true_df]
        new_df = name_df[(name_df['Ngày'].dt.date >= start_date) & (name_df['Ngày'].dt.date <= end_date)]

        # filter_df = df[(df['Ngày'].dt.date >= start_date) & (df['Ngày'].dt.date <= end_date)]
        return new_df


    def to_excel(data):
        """This function is used to write data to the buffer and store it """
        output = BytesIO()  # Create binary stream
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        data.T.reset_index().T.to_excel(writer, startrow=2, index=False, header=None, sheet_name='Tên Biển Số')
        workbook = writer.book
        worksheet = writer.sheets['Tên Biển Số']

        # Loop through each column to adjust the width
        for col_num, col_data in enumerate(data.columns):
            column_len = max(data[col_data].astype('str').apply(len).max(), len(col_data))
            worksheet.set_column(col_num, col_num, column_len)

        format1 = workbook.add_format({'num_format': '0'})
        worksheet.set_column('A:A', None, format1)
        writer.close()  # Save is deprecated
        processed_data = output.getvalue()
        return processed_data

    def update_data(df1, df2, df3, wrksheet, column_name, target_column, line_start=2):
        """ This function is used to update the data on Google Sheet"""
        result = df1[column_name].equals(df2[column_name])  # Check if
        if not result:
            data_ = df3[column_name].to_dict()
            value_to_update = []
            for index, value in data_.items():
                position = int(index) + line_start
                value_to_update.append({'range': f'{target_column}{position}', 'values': [[value]]})
            wrksheet.batch_update(value_to_update)

    # Connect to worksheet using gspread package
    gc = gspread.service_account('credential.json')
    sheet = gc.open('DATA Tài Xế DNL')
    wrksheet = sheet.worksheet('Data')

    # Set up credentials
    creds = setup_credentials()

    # Get the result
    result = connectToSheet(creds)

    # Create data frame
    df = create_data_frame(result)
    disabled_columns = ['STT', 'Tên Tài Xế', 'Ngày', 'Tổng doanh thu (v/c + nâng hạ + đường biển,…)',
                        'Dthu vận chuyển (chưa VAT)', 'Size',
                        'Số chuyến', 'Lưu đêm', 'Ghi chú của tài xế (nếu có)']

    # Read road sheet
    road_table = pd.read_excel('road.xlsx')
    # Final result
    final_result = pd.merge(df, road_table, on='Nơi đi', how='left')


    #  USER-AUTHENTICATION INTERFACE
    names = ['Danalog']
    user_name = ['DNL']
    # Find current directory
    file_path = Path(__file__).parent/'hashed_wd.pkl'

    with file_path.open(mode='rb') as file:
        hashed_passwords = pickle.load(file)

    authenticator = stauth.Authenticate(names=names, usernames=user_name, passwords=hashed_passwords,cookie_name='danalog',key='dana',cookie_expiry_days=30)

    name, authentication_status, user_name = authenticator.login('Đăng nhập tại đây','main')
    if authentication_status == False:
        st.error('Bạn đã nhập sai mật khẩu')
    if authentication_status == None:
        st.warning('Vui lòng nhập thông tin để đăng nhập')
    if authentication_status:

        # Create a Streamlit app
        st.title("DanaLog Webapp")

        driver_name = df['Tên Tài Xế'].unique()  # Return as a list
        authenticator.logout(button_name='Đăng xuất tài khoản',location='sidebar')
        container = st.sidebar.container()
        all = st.sidebar.checkbox("Chọn tất cả tài xế")

        if all:
            selected_name = container.multiselect("Select one or more options:",
                                                  driver_name, default=driver_name)
        else:
            selected_name = container.multiselect("Select one or more options:",
                                                  driver_name)
        # Create a dropdown to select a name
        # selected_name = st.multiselect("Chọn tên tài xế", driver_name, default=driver_name)

        # Create date input widgets for start and end dates
        start_date = st.sidebar.date_input("Chọn ngày bắt đầu", value=df["Ngày"].median(), min_value=df["Ngày"].min(),
                                   max_value=df['Ngày'].max(), format="DD/MM/YYYY")
        end_date = st.sidebar.date_input("Chọn ngày kết thúc", value=df["Ngày"].median(), min_value=df['Ngày'].min(),
                                 max_value=df['Ngày'].max(), format="DD/MM/YYYY")


        # Display the selected date range
        if start_date > end_date:
            st.error("Ngày kết thúc phải lớn hơn hoặc bằng ngày bắt đầu")

        # Filter the DataFrame based on the selected date range
        filtered_df = data_extracting(final_result, selected_name, start_date, end_date)
        option = st.sidebar.selectbox(label='Chọn phòng ban', options=['Kế toán', 'CS'])
        if option == 'Kế toán':
            filtered_df.drop(columns='Tuyến đường',inplace=True)
        modified_df = st.data_editor(filtered_df, disabled=disabled_columns)
        filtered_df['number'] = 1
        modified_df['number'] = 2

        # Only get the rows that were modified
        changed_df = pd.concat([modified_df, filtered_df]).drop_duplicates(keep=False)
        changed_df = changed_df.loc[changed_df['number'] == 2]


        save_button = st.button(label='Cập nhập dữ liệu mới')
        if save_button:
            if modified_df is not None:
                t1=threading.Thread(target=update_data, args=(modified_df, filtered_df, changed_df, wrksheet, 'Doanh thu', 'AB'))
                t2=threading.Thread(target=update_data, args=(modified_df, filtered_df, changed_df, wrksheet, 'Nơi Đến', 'V'))
                t3=threading.Thread(target=update_data, args=(modified_df, filtered_df, changed_df, wrksheet, 'Nơi đi', 'M'))
                t4=threading.Thread(target=update_data, args=(modified_df, filtered_df, changed_df, wrksheet, 'Dịch Vụ/ Container No.', 'W'))
                t1.start()
                t2.start()
                t3.start()
                t4.start()
        filtered_df = filtered_df.drop(columns=['number'])

        # Download the excel file
        st.download_button(label='📥 Tải file excel tại đây',
                           data=to_excel(data=filtered_df),
                           file_name='Danalog.xlsx')
if __name__ == "__main__":
    main()


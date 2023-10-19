# Import all necessary library

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from icecream import ic  # Use this module for debugging
import pandas as pd
from io import BytesIO
import streamlit as st



#  Set display
pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)

with open('spread_sheetId.txt', mode='r') as file:
    SAMPLE_SPREADSHEET_ID = file.read()
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # Define the scope for editing and reading
SERVICE_ACCOUNT_FILE = 'credential.json'
SAMPLE_RANGE_NAME = 'TestData!A1:M'


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

        # Convert the 'Ngày' column to datetime format with the correct format
        df['Ngày'] = pd.to_datetime(df['Ngày'], format="%d/%m/%Y",errors='coerce')
        column1 = 'Tổng doanh thu (v/c + nâng hạ + đường biển,…)'
        column2 = 'Dthu vận chuyển (chưa VAT)'
        df.insert(0, column='STT', value=None)
        df.insert(8,column=column1, value=None)
        df.insert(9,column=column2, value=None)
        list_number_type=[column1, column2]
        int_type = ['Size', 'Số chuyến', 'Lưu đêm']
        df[list_number_type]=df[list_number_type].astype(float)
        df[int_type] = df[int_type].astype(int)
        return df


def data_extracting(df, *arg) -> pd.DataFrame:
    """This function is used to extract the data we want"""
    name = arg[0]
    start_date = arg[1]
    end_date = arg[2]
    # df['Ngày'] = pd.to_datetime(df['Ngày'], format="%d/%m/%Y",errors='coerce')  # Convert to date format like 10/5/2005
    if name:
        true_df = df['Tên Tài Xế'].isin(name)
        name_df = df[true_df]
        new_df = name_df[(name_df['Ngày'].dt.date >= start_date) & (name_df['Ngày'].dt.date <= end_date)]
        return new_df
    # new_df['Ngày'] = new_df['Ngày'].dt.strftime('%d/%m/%Y')
    else:
        filter_df = df[(df['Ngày'].dt.date >= start_date) & (df['Ngày'].dt.date <= end_date)]
        return filter_df


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


def update_data(new_df):

    pass






# Set up credentials
creds = setup_credentials()
# Get the result
result = connectToSheet(creds)
# Create data frame
df = create_data_frame(result)
# Read dien giai sheet
dien_giai = pd.read_excel('diengiai.xlsx')
# Create ID column based on Dien Giai 1
result1 = pd.merge(df, dien_giai, on='Diễn giải 1', how='inner')
# Read road sheet
road_table = pd.read_excel('road.xlsx')
# Final result
final_result = pd.merge(result1, road_table, on='ID', how='inner').drop('ID',axis=1)
columns = ['STT', 'Ngày', 'Tên Tài Xế', 'Biển Số Xe', 'Container No.', 'Size', 'Số chuyến', 'Nhập Xuất', 'Tổng doanh thu (v/c + nâng hạ + đường biển,…)', 'Dthu vận chuyển (chưa VAT)', 'Diễn giải 1', 'Diễn giải 2','Tuyến đường', 'Lưu đêm', 'Ghi chú của tài xế ( nếu có)']
final_result = final_result[columns]









# Create a Streamlit app
st.title("DanaLog Webapp")

driver_name = df['Tên Tài Xế'].drop_duplicates().tolist()


# Create a dropdown to select a name
selected_name = st.multiselect("Chọn tên tài xế", driver_name, default=driver_name)

# Create date input widgets for start and end dates
start_date = st.date_input("Chọn ngày bắt đầu",value=df["Ngày"].median(), min_value=df["Ngày"].min(), max_value=df['Ngày'].max(),format="DD/MM/YYYY")
end_date = st.date_input("Chọn ngày kết thúc",value=df["Ngày"].median(), min_value=df['Ngày'].min(), max_value=df['Ngày'].max(),format="DD/MM/YYYY")



# Display the selected date range
if start_date <= end_date:
    st.write(f"Tên của tài xế là {selected_name}")
    st.write(f"Chọn dữ liệu từ: {start_date.strftime('%d-%m-%Y')} tới {end_date.strftime('%d-%m-%Y')}")
else:
    st.error("Ngày kết thúc phải lớn hơn hoặc bằng ngày bắt đầu")

# Filter the DataFrame based on the selected date range
filtered_df = data_extracting(final_result, selected_name, start_date, end_date)
option = st.selectbox('Chọn vị trí của bạn',('Kế toán','CS'))
if 'changes' not in st.session_state:
    st.session_state['changes'] = {}
if option == 'Kế toán':
    columns = ['Container No.', 'Tuyến đường']
    filtered_df = filtered_df.drop(columns=['Diễn giải 1','Diễn giải 2'])
else:
    columns = ['Container No.', 'Diễn giải 1', 'Diễn giải 2']
    filtered_df = filtered_df.drop(columns=['Tuyến đường'])
st.dataframe(filtered_df)

try:
    index =st.number_input('Nhap vao hang muon doi:', min_value=df.index.min(), max_value=df.index.max(), step=1)
    if index in filtered_df.index:
        with st.form("Edit Data Form"):
            # Create input fields for each cell in the DataFrame
            for column in columns:
                new_value = st.text_input(f"Edit {column} for row {index}", filtered_df.loc[index, column])
                if new_value != filtered_df.loc[index, column]:
                    # If the user made changes, store them in a dictionary
                    st.session_state['changes'][(index, column)] = new_value

            # When the "Save Changes" button is clicked
            if st.form_submit_button("Save Changes"):
                # Apply the changes to the DataFrame
                for (index, column), new_value in st.session_state.changes.items():
                    if column == 'Container No.':
                        new_value = str(new_value) if new_value else None  # Convert the data for compatible data type
                    elif column == 'Tuyến đường':
                        new_value = str(new_value) if new_value else None
                    filtered_df.at[index, column] = new_value
except ValueError:
    st.error('Nhập dữ liệu sai')
except KeyError:
    st.error('Dòng cần sửa không có trong bản')
st.dataframe(filtered_df)

stringified_changes = {f"Chỉnh sửa dòng:{key[0]}, cột: {key[1]}": value for key, value in st.session_state.changes.items()}
st.write(stringified_changes)


st.download_button(label='📥 Tải file excel tại đây',
               data=to_excel(data=filtered_df),
               file_name='Danalog.xlsx')
import streamlit as st
from getting_data import *

# Set up credentials
creds = setup_credentials()
# Get the result
result = connectToSheet(creds)
# Create data frame
df = create_data_frame(result)

# Create a Streamlit app
st.title("DanaLog Webapp")

# with st.form("Data Filter Form"):
# Create a dropdown to select a name
selected_name = st.selectbox("Chọn tên tài xế", df['Tên Tài Xế'].drop_duplicates())

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
filtered_df = data_extracting(df, selected_name, start_date, end_date)
st.write(filtered_df)

# Download excel file button
st.download_button(label='📥 Tải file excel tại đây',
                                data=to_excel(data=filtered_df),
                                file_name= 'Danalog.xlsx')


import streamlit as st
import pandas as pd
import os
import io
from io import BytesIO
from io import StringIO

def retype(x):
    if isinstance(x, str):
        return 0
    return x

# def process_excel_file(file):

#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(file, sheet_name="Câu trả lời biểu mẫu 1", engine='openpyxl')

#     # Convert all string values in n_empl_df to lowercase
#     df = df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
#     df.columns = df.columns.str.lower()

#     # retype str to int
#     for i in range(3,8):
#         df.iloc[:, i] = df.iloc[:, i].apply(retype)

#     # copute sum of all columns from 3 to 7
#     df['tổng cộng'] = df.iloc[:, 3:7].sum(axis=1) 
    
#     # set group 1 & 2
#     group_2 = ['chuyển giao', 'phòng mẫu', 'mặt giày 2', 'kho đế', 'gia công đế', 'xưởng c', 'qc 2']
#     df['group'] = df['xưởng'].copy()
#     for i,item in enumerate(df['group']):
#         if item.lower() in group_2:
#             df.loc[i, 'group'] = 2
#         else:
#             df.loc[i, 'group'] = 1

#     df = df.sort_values(by = ['group', 'xưởng'])
#     df.to_csv('sorted_qr_lunch.csv', index=False, encoding='utf-8')
    
#     return df



def process_excel_file(file):

    # Save to a buffer
    buffer = io.StringIO()
    file.to_csv(buffer, index=False)
    buffer.seek(0)

    # Use the buffer as input
    new_df = pd.read_csv(buffer)

    # Convert all string values in n_empl_df to lowercase
    new_df = new_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
    new_df.columns = new_df.columns.str.lower()

    # retype str to int
    for i in range(3,8):
        new_df.iloc[:, i] = new_df.iloc[:, i].apply(retype)

    # copute sum of all columns from 3 to 7
    new_df['tổng cộng总共'] = new_df.iloc[:, 3:7].sum(axis=1) 
    
    # set group 1 & 2
    group_2 = ['chuyển giao', 'phòng mẫu', 'mặt giày 2', 'kho đế', 'gia công đế', 'xưởng c', 'qc 2']
    new_df['group'] = new_df['xưởng厂别'].copy()
    for i,item in enumerate(new_df['group']):
        if item.lower() in group_2:
            new_df.loc[i, 'group'] = 2
        else:
            new_df.loc[i, 'group'] = 1

    new_df = new_df.sort_values(by = ['Dấu thời gian', 'group', 'xưởng厂别'])
    new_df.to_csv('sorted_qr_lunch.csv', index=False, encoding='utf-8')
    
    return new_df

# Giao diện Streamlit
st.title("Excel Processor")
st.header("Sorted 'BÁO CƠM'")

uploaded_file = st.file_uploader("Upload file .xlsx", type=["xlsx"])

if uploaded_file is not None:
    file_name = uploaded_file.name
    st.write(f"**Tên file:** {file_name}")
    
    try:
        # Đọc file vào DataFrame
        df = pd.read_excel(uploaded_file)
        # st.write(df.head())

        # Gọi hàm xử lý
        processed_df = process_excel_file(df) # uploaded_file

        # Ghi kết quả ra file Excel mới (trong RAM)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            processed_df.to_excel(writer, index=False)
        output.seek(0)

        # Nút tải về
        st.download_button(
            label="Tải xuống file đã xử lý",
            data=output,
            file_name="processed_" + file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {e}")

st.header("Merged 'BÁO BIỂU'")

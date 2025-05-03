import streamlit as st
import pandas as pd
import os
from io import BytesIO

def retype(x):
    if isinstance(x, str):
        return 0
    return x

def process_excel_file(file):

    # Read the Excel file into a DataFrame
    df = pd.read_excel(file)

    # Convert all string values in n_empl_df to lowercase
    df = df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
    df.columns = df.columns.str.lower()

    # retype str to int
    for i in range(3,8):
        df.iloc[:, i] = df.iloc[:, i].apply(retype)

    # copute sum of all columns from 3 to 7
    df['tổng cộng'] = df.iloc[:, 3:7].sum(axis=1) 
    
    # set group 1 & 2
    group_2 = ['chuyển giao', 'phòng mẫu', 'mặt giày 2', 'kho đế', 'gia công đế', 'xưởng c', 'qc 2']
    df['group'] = df['xưởng'].copy()
    for i,item in enumerate(df['group']):
        if item.lower() in group_2:
            df.loc[i, 'group'] = 2
        else:
            df.loc[i, 'group'] = 1

    df = df.sort_values(by = ['group', 'xưởng'])
    df.to_csv('sorted_qr_lunch.csv', index=False, encoding='utf-8')
    
    return df

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
        processed_df = process_excel_file(uploaded_file)

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

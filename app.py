import streamlit as st
import pandas as pd
import os
from io import BytesIO

def process_excel_file(file):
    df = pd.read_excel(file)
    group_1 = ['văn phòng', 'tổng vụ', 'cơ điện', 'kho dán hộp', 'kho vật tư', 'xưởng a', 'mặt giày 1', 'xưởng b', 'qc 1']
    df['Group'] = df['Xưởng'].copy()
    for i,item in enumerate(df['Group']):
        if item.lower() in group_1:
            df.loc[i, 'Group'] = 1
        else:
            df.loc[i, 'Group'] = 2

    df = df.sort_values(by = ['Group', 'Xưởng'])
    df.to_csv('sorted_qr_lunch.csv', index=False, encoding='utf-8')
    
    return df

# Giao diện Streamlit
st.title("Excel Processor")

uploaded_file = st.file_uploader("Upload file .xlsx", type=["xlsx"])

if uploaded_file is not None:
    file_name = uploaded_file.name
    st.write(f"**Tên file:** {file_name}")
    
    try:
        # Đọc file vào DataFrame
        df = pd.read_excel(uploaded_file)
        st.write(df.head())

        # Gọi hàm xử lý
        processed_df = process_excel_file(file_name)

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

import streamlit as st
import pandas as pd
import os
import re
import io
from io import BytesIO
from io import StringIO


# def remove_non_vietnamese_characters(text):
#     # Giữ lại: chữ cái Latin có dấu tiếng Việt, số, dấu cách, dấu chấm câu cơ bản
#     text = re.sub(r'[^a-zA-Z0-9\sàáảãạăắằẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởợùúủũụưừứửữựýỳỷỹỵ]', ' ', text)
#     text = re.sub(' +', ' ',text)
#     text = text.strip()
#     return text

# def retype(x):
#     if isinstance(x, str):
#         return 0
#     return x

# def fix_output_type(df):

#     # Rename content in 'xưởng': mặt giày 1 ab -> mặt giày 1; mặt giày 2 c -> mặt giày 2; 
#     # qc 1 ab -> qc 1; qc 2 cs -> qc2; vp doreen 1 -> vp doreen; vp una 2 -> vp una
#     rename_map = {
#         "mặt giày 1 ab": "mặt giày 1",
#         "mặt giày 2 c": "mặt giày 2",
#         "qc 1 ab": "qc 1",
#         "qc 2 cs": "qc 2",
#         "vp doreen 1": "vp doreen",
#         "vp una 2": "vp una"
#     }
#     df['xưởng'] = df['xưởng'].replace(rename_map)

#     # Rename content in 'ca ăn', 'đợt ăn'
#     # if 'ca ăn' = 1, 'đợt ăn' = 1 -> 'đợt ăn' : '11h30'; 'ca ăn' = 1, 'đợt ăn' = 2 --> 'đợt ăn': '12h00'...
#     def rename_dot_an(row):
#         if row['ca ăn (trưa, chiều)'] == 1 and row['đợt ăn 1,2'] == 1:
#             return '11h30'
#         elif row['ca ăn (trưa, chiều)'] == 1 and row['đợt ăn 1,2'] == 2:
#             return '12h00'
#         elif row['ca ăn (trưa, chiều)'] == 2 and row['đợt ăn 1,2'] == 1:
#             return '16h30'
#         return '17h00'

#     df['đợt ăn 1,2'] = df.apply(rename_dot_an, axis=1)
#     # if "ca ăn (trưa, chiều)" = 1 -> trưa else chiều
#     df['ca ăn (trưa, chiều)'] = df['ca ăn (trưa, chiều)'].replace({1: 'trưa', 2: 'chiều'})

#     # Convert all string values in qr_lunch_df to uppercase
#     df = df.map(lambda x: x.title() if isinstance(x, str) else x)
#     df.columns = df.columns.str.title()

#     return df

# def process_excel_file(file):

#     # Save to a buffer
#     # buffer = StringIO()
#     # file.to_csv(buffer, index=False)
#     # buffer.seek(0)

#     # # Use the buffer as input
#     # new_df = pd.read_csv(buffer)
#     new_df = file

#     # Convert all string values in n_empl_df to lowercase
#     new_df = new_df.map(lambda x: x.lower() if isinstance(x, str) else x)
#     new_df.columns = new_df.columns.str.lower()

#     # retype str to int
#     for i in range(3,8):
#         new_df.iloc[:, i] = new_df.iloc[:, i].apply(retype)
    
#     # Apply the remove_non_vietnamese_characters function to all columns in the DataFrame
#     new_df.columns = new_df.columns.map(remove_non_vietnamese_characters)
#     columns_to_process = [col for col in new_df.columns if col not in ["dấu thời gian", "ghi chú"]]
#     new_df[columns_to_process] = new_df[columns_to_process].map(lambda x: remove_non_vietnamese_characters(x) if isinstance(x, str) else x)
    

#     # copute sum of all columns from 3 to 7
#     new_df['tổng cộng'] = new_df.iloc[:, 3:7].sum(axis=1) 

#     # fill na by "" 
#     new_df = new_df.fillna("")
    
#     # set 'đợt ăn' 1 & 2
#     group_2 = ['chuyển giao', 'phòng mẫu', 'mặt giày 2', 'kho đế', 'gia công đế', 'xưởng c', 'qc 2', 'mặt giày 2c', 'qc 2cs']
#     new_df['đợt ăn 1,2'] = new_df['xưởng'].copy()
#     for i,item in enumerate(new_df['đợt ăn 1,2']):
#         if item.lower() in group_2:
#             new_df.loc[i, 'đợt ăn 1,2'] = 2 # dot 2
#         else:
#             new_df.loc[i, 'đợt ăn 1,2'] = 1 # dot 1
  
#     # tạo cột 'giờ' để phân loại        
#     # new_df['dấu thời gian'] = new_df['dấu thời gian'].astype('object') # streamlit view it as timestamp
#     new_df['dấu thời gian'] = pd.to_datetime(new_df['dấu thời gian']) 

#     # new_df['giờ'] = new_df['dấu thời gian'].apply(lambda x: x.split()[1])
#     # new_df['giờ'] = new_df['giờ'].apply(lambda x: x.split(':')[0] + ':' + x.split(':')[1])
#     new_df['giờ'] = new_df['dấu thời gian'].dt.strftime('%H:%M') 
#     new_df['ngày'] = new_df['dấu thời gian'].dt.strftime('%Y-%m-%d') 
#     # new_df['dấu thời gian'] = pd.to_datetime(new_df['dấu thời gian'], format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
    
#     # set 'ca ăn' trưa & chiều
#     new_df['ca ăn (trưa, chiều)'] = new_df['dấu thời gian'] .copy()
#     for i, item in enumerate(new_df['ca ăn (trưa, chiều)']):
#         time = pd.to_datetime(item, format='%H:%M', errors='coerce').time()
#         if pd.to_datetime('12:00', format='%H:%M').time() < time < pd.to_datetime('16:00', format='%H:%M').time():
#             new_df.loc[i, 'ca ăn (trưa, chiều)'] = 2
#         else:
#             new_df.loc[i, 'ca ăn (trưa, chiều)'] = 1

#     new_df = new_df.sort_values(by = ['ngày', 'ca ăn (trưa, chiều)', 'đợt ăn 1,2', 'xưởng', 'bộ phận'])
#     new_df = new_df[[
#         'ca ăn (trưa, chiều)', 'đợt ăn 1,2', 'xưởng', 'bộ phận','món mặn đánh số vd 25',
#         'món chay đánh số vd 25', 'món nước đánh số vd 25',
#         'phiếu đổi đánh số vd 25', 'tổng cộng', 'ghi chú', 'dấu thời gian']]
    
#     new_df.reset_index(drop=True, inplace=True)
#     new_df = fix_output_type(new_df)
    
#     # new_df.to_csv('sorted_qr_lunch.csv', index=False, encoding='utf-8')
    
#     return new_df


def remove_non_vietnamese_characters(text):
    # Giữ lại: chữ cái Latin có dấu tiếng Việt, số, dấu cách, dấu chấm câu cơ bản
    text = re.sub(r'[^a-zA-Z0-9\sàáảãạăắằẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởợùúủũụưừứửữựýỳỷỹỵ]', ' ', text)
    text = re.sub(' +', ' ',text)
    text = text.strip()
    return text

def retype(x):
    if isinstance(x, str):
        return 0
    return x


def clean_old_cols(df):

    # Convert all string values in n_empl_df to lowercase
    df = df.map(lambda x: x.lower() if isinstance(x, str) else x)
    df.columns = df.columns.str.lower()

    # retype str to int
    for i in range(3,8):
        df.iloc[:, i] = df.iloc[:, i].apply(retype)
    
    # Apply the remove_non_vietnamese_characters function to all columns in the DataFrame
    df.columns = df.columns.map(remove_non_vietnamese_characters)
    columns_to_process = [col for col in df.columns if col not in ["dấu thời gian", "ghi chú"]]
    df[columns_to_process] = df[columns_to_process].map(lambda x: remove_non_vietnamese_characters(x) if isinstance(x, str) else x)
    

    # copute sum of all columns from 3 to 7
    df['tổng cộng'] = df.iloc[:, 3:7].sum(axis=1) 

    # fill na by "" 
    df = df.fillna("")

    return df

def create_new_cols(df):
    # Creat new columns: 'ngày', 'giờ', 'đợt ăn (1,2)', 'ca ăn (trưa, chiều)

    # set 'đợt ăn' 1 & 2
    group_2 = ['chuyển giao', 'phòng mẫu', 'mặt giày 2', 'kho đế', 'gia công đế', 'xưởng c', 'qc 2', 'mặt giày 2 c', 'qc 2 cs']
    df['đợt ăn (1,2)'] = df['xưởng'].copy()
    for i,item in enumerate(df['đợt ăn (1,2)']):
        if item.lower() in group_2:
            df.loc[i, 'đợt ăn (1,2)'] = 2 # dot 2
        else:
            df.loc[i, 'đợt ăn (1,2)'] = 1 # dot 1
  
    # tạo cột 'ngày', 'giờ' để phân loại        
    df['dấu thời gian'] = pd.to_datetime(df['dấu thời gian']) 
    df['giờ'] = df['dấu thời gian'].dt.strftime('%H:%M') 
    df['ngày'] = df['dấu thời gian'].dt.strftime('%Y-%m-%d') 

    # set 'ca ăn' trưa & chiều
    df['ca ăn (trưa, chiều)'] = df['dấu thời gian'] .copy()
    for i, item in enumerate(df['ca ăn (trưa, chiều)']):
        time = pd.to_datetime(item, format='%H:%M', errors='coerce').time()
        if pd.to_datetime('12:00', format='%H:%M').time() < time < pd.to_datetime('16:00', format='%H:%M').time():
            df.loc[i, 'ca ăn (trưa, chiều)'] = 2
        else:
            df.loc[i, 'ca ăn (trưa, chiều)'] = 1

    # choosing needed columns
    df = df.sort_values(by = ['ngày', 'ca ăn (trưa, chiều)', 'đợt ăn (1,2)', 'xưởng', 'bộ phận'])
    df = df[[
        'ca ăn (trưa, chiều)', 'đợt ăn (1,2)', 'xưởng', 'bộ phận','món mặn đánh số vd 25',
        'món chay đánh số vd 25', 'món nước đánh số vd 25',
        'phiếu đổi đánh số vd 25', 'tổng cộng', 'ghi chú', 'dấu thời gian']]
    
    return df

def fix_output_type(df):

    # Rename content in 'xưởng': mặt giày 1 ab -> mặt giày 1; mặt giày 2 c -> mặt giày 2; 
    # qc 1 ab -> qc 1; qc 2 cs -> qc2; vp doreen 1 -> vp doreen; vp una 2 -> vp una
    rename_map = {
        "mặt giày 1 ab": "mặt giày 1",
        "mặt giày 2 c": "mặt giày 2",
        "qc 1 ab": "qc 1",
        "qc 2 cs": "qc 2",
        "vp doreen 1": "vp doreen",
        "vp una 2": "vp una"
    }
    df['xưởng'] = df['xưởng'].replace(rename_map)

    # Rename content in 'ca ăn', 'đợt ăn'
    # if 'ca ăn' = 1, 'đợt ăn' = 1 -> 'đợt ăn' : '11h30'; 'ca ăn' = 1, 'đợt ăn' = 2 --> 'đợt ăn': '12h00'...
    def rename_dot_an(row):
        if row['ca ăn (trưa, chiều)'] == 1 and row['đợt ăn (1,2)'] == 1:
            return '11h30'
        elif row['ca ăn (trưa, chiều)'] == 1 and row['đợt ăn (1,2)'] == 2:
            return '12h00'
        elif row['ca ăn (trưa, chiều)'] == 2 and row['đợt ăn (1,2)'] == 1:
            return '16h30'
        return '17h00'

    df['đợt ăn (1,2)'] = df.apply(rename_dot_an, axis=1)
    # if "ca ăn (trưa, chiều)" = 1 -> trưa else chiều
    df['ca ăn (trưa, chiều)'] = df['ca ăn (trưa, chiều)'].replace({1: 'trưa', 2: 'chiều'})

    # Convert all string values in qr_lunch_df to uppercase
    df = df.map(lambda x: x.title() if isinstance(x, str) else x)
    df.columns = df.columns.str.title()

    return df

def process_excel_file(file):

    new_df = file

    # Clean dataframe
    new_df = clean_old_cols(new_df)

    # Creat new columns: 'ngày', 'giờ', 'đợt ăn (1,2)', 'ca ăn (trưa, chiều)
    new_df = create_new_cols(new_df)
    
    # change output type
    new_df.reset_index(drop=True, inplace=True)
    new_df = fix_output_type(new_df)
    
    # new_df.to_csv('sorted_qr_lunch.csv', index=False, encoding='utf-8')
    
    return new_df

# === Streamlit Interface ===

st.title("Excel Processor")
st.header("1. Sorted 'BÁO CƠM'")

uploaded_file = st.file_uploader("Chọn tệp Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.sidebar.write("Available sheets:")
        sheet = st.sidebar.selectbox("Chọn sheet", excel_file.sheet_names)

        df = pd.read_excel(excel_file, sheet_name=sheet)
        st.write(f"### 📄 Sheet: {sheet}")
        st.dataframe(df)

        df = pd.read_excel(excel_file, sheet_name='Câu trả lời biểu mẫu 1')

        processed_df = process_excel_file(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            processed_df.to_excel(writer, index=False)
        output.seek(0)

        st.write(f"### 📄 File sau khi xử lý")
        st.dataframe(processed_df)

        st.download_button(
            label="📥 Tải xuống file đã xử lý",
            data=output,
            file_name=f"processed_{sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Lỗi khi xử lý file Excel: {e}")
# st.header("Merged 'BÁO BIỂU'")

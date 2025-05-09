import pandas as pd
import numpy as np
import re
pd.set_option('future.no_silent_downcasting', True)

# 1. Preprocess the data
def process_bao_bieu(df):

    # Convert all string values in file 'BAO BIEU' to lowercase
    df = df.map(lambda x: x.lower() if isinstance(x, str) else x)
    df.columns = df.columns.str.lower()

    # get row from 5 to 128
    df = df[4:128]
    # get col B, C, O (index: 1,2,14)
    df = df[df.columns[[1,2,14]].values]
    
    # rename header
    df.columns = ['department_name_chinese', 'id_code', 'num_employee']
    
    return df

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
    df['ca ăn (trưa, chiều)'] = df['ca ăn (trưa, chiều)'].astype(object)
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

    return df

def process_bao_com(file):

    new_df = file

    # Clean dataframe
    new_df = clean_old_cols(new_df)

    # Creat new columns: 'ngày', 'giờ', 'đợt ăn (1,2)', 'ca ăn (trưa, chiều)
    new_df = create_new_cols(new_df)
    
    # change output type
    new_df.reset_index(drop=True, inplace=True)
    new_df = fix_output_type(new_df)
    
    return new_df

def process_dept_id(df):
    
    ''' df: department_id dataframe (df_dept_id)'''

    # Convert all string values in df to lowercase
    df = df.map(lambda x: x.lower() if isinstance(x, str) else x)
    df.columns = df.columns.str.lower()
    df.fillna("", inplace=True)

    return df



# Merged
# df_work
def create_work_df(df_bao_bieu, df_dept_id):

    df_work = pd.merge(df_bao_bieu, df_dept_id, on = ['department_name_chinese', 'id_code'], how='outer')
    df_work.drop_duplicates( keep='first', inplace=True)
    df_work = df_work.fillna(0)
    df_work = df_work.groupby(['id', 'xưởng', 'bộ phận'])['num_employee'].sum().reset_index()
    df_work = df_work[df_work['num_employee'] != 0]
    df_work.drop_duplicates(keep='first', inplace=True)
    df_work.reset_index(drop=True, inplace=True)
    df_work['bộ phận 2'] = np.where(df_work['bộ phận']=="", df_work['xưởng'], df_work['bộ phận'])

    return df_work 

# df_lunch
def create_lunch_df(df_bao_com, df_dept_id):

    # create column 'bộ phận 2' on df_bao_com & df_dept_id & df_work
    df_bao_com['bộ phận 2'] = np.where(df_bao_com['bộ phận']=="", df_bao_com['xưởng'], df_bao_com['bộ phận'])
    df_dept_id['bộ phận 2'] = np.where(df_dept_id['bộ phận']=="", df_dept_id['xưởng'], df_dept_id['bộ phận'])

    # replace content in df_bao_com['xưởng'] to fit with df_dept_id['xưởng]
    rename_map = {
            "vp thu mua 1": "vp thu mua",
            "vp thu mua 2": "vp thu mua",
            "vp doreen": "vp nghiệp vụ",
            "vp una": "vp nghiệp vụ"
        }
    df_bao_com['xưởng'] = df_bao_com['xưởng'].replace(rename_map)

    df_lunch = pd.merge(df_bao_com, df_dept_id, on = ['xưởng', 'bộ phận 2'], how='left')

    df_lunch = df_lunch[['ca ăn (trưa, chiều)', 'đợt ăn (1,2)', 'department_name_ver2', 'xưởng', 'bộ phận_x',
    'món mặn đánh số vd 25', 'món chay đánh số vd 25',
    'món nước đánh số vd 25', 'phiếu đổi đánh số vd 25', 'tổng cộng',
    'ghi chú', 'dấu thời gian', 'id', 'bộ phận 2']]
    df_lunch.rename({'bộ phận_x': 'bộ phận'})
    df_lunch.drop_duplicates(inplace=True, keep='first')
    df_lunch.reset_index(drop=True, inplace=True)

    return df_lunch

# merge_df
def create_merged_df(df_work, df_lunch):

    df_merge4 = pd.merge(df_work, df_lunch, on=['xưởng', 'bộ phận 2'], how='left')

    df_merge4[['ca ăn (trưa, chiều)', 'đợt ăn (1,2)', 'tổng cộng']] = df_merge4[['ca ăn (trưa, chiều)', 'đợt ăn (1,2)', 'tổng cộng']].fillna("")
    df_merge4['tổng cộng'] = np.where(df_merge4['tổng cộng'] == "", df_merge4['num_employee'], df_merge4['tổng cộng'])
    df_merge4['tổng cộng'] = df_merge4['tổng cộng'].replace('', 0)

    # Filter rows where 'bộ phận' matches the specified values
    target_rows = df_merge4[df_merge4['xưởng'].isin(['vp nhân sự', 'vp tài vụ', 'vp xuất nhập khẩu', 'văn phòng'])]

    # Create duplicates with 'ca ăn (trưa, chiều)' = 1 and 'đợt ăn (1,2)' = 1 
    dup1 = target_rows.copy()
    dup1['ca ăn (trưa, chiều)'] = 1
    dup1['đợt ăn (1,2)'] = 1

    # Create duplicates with 'ca ăn (trưa, chiều)' = 2 and 'đợt ăn (1,2)' = 1 
    dup2 = target_rows.copy()
    dup2['ca ăn (trưa, chiều)'] = 2
    dup2['đợt ăn (1,2)'] = 1

    # concatenate the original dataframe with the duplicates and drop the target rows
    df_merge4 = pd.concat([df_merge4, dup1, dup2], ignore_index=True)
    df_merge4.drop(target_rows.index, inplace=True)
    df_merge4.reset_index(drop=True, inplace=True)

    # df_merge4.sort_values(by=['ca ăn (trưa, chiều)', 'đợt ăn (1,2)', 'id_x', 'xưởng', 'bộ phận'], inplace=True)


    df_merge4 = df_merge4[['id_x', 'xưởng', 'bộ phận', 'department_name_ver2', 'num_employee',  'tổng cộng', 'ghi chú', 'ca ăn (trưa, chiều)', 'đợt ăn (1,2)']]
    df_merge4['department_name_ver2'] = df_merge4['department_name_ver2'].fillna('辦公室')

    # pivot table columns=['ca ăn (trưa, chiều)']
    df_merge4 = pd.pivot_table(df_merge4, columns=['ca ăn (trưa, chiều)'], index=['id_x', 'xưởng', 'bộ phận', 'department_name_ver2', 'num_employee', 'đợt ăn (1,2)'], values='tổng cộng', aggfunc='sum').reset_index()
    df_merge4.columns.name = None
    df_merge4.rename(columns={1: 'báo cơm trưa', 2: 'báo cơm chiều'}, inplace=True)
    df_merge4[['báo cơm trưa', 'báo cơm chiều']] = df_merge4[['báo cơm trưa', 'báo cơm chiều']].fillna(0)

    # combine bộ phận == 'vp thu mua 1' and 'vp thu mua 2' into 'vp thu mua'; 'vp doreen' and 'vp una' into 'vp nghiệp vụ'
    df_merge4['bộ phận'] = df_merge4['bộ phận'].replace(['vp thu mua 1', 'vp thu mua 2'], 'vp thu mua')
    df_merge4['bộ phận'] = df_merge4['bộ phận'].replace(['vp doreen', 'vp una'], 'vp nghiệp vụ')

    # groupby 'id_x', 'xưởng', 'bộ phận', 'department_name_ver2', 'num_employee', 'đợt ăn (1,2)' and sum the values
    df_merge4 = df_merge4.groupby(by=['id_x', 'xưởng', 'bộ phận', 'department_name_ver2', 'num_employee', 'đợt ăn (1,2)'])[['báo cơm trưa', 'báo cơm chiều']].sum().reset_index()

    # groupby 'id_x', 'xưởng', 'bộ phận', 'department_name_ver2', 'num_employee' and sum the values
    df_merge4 = df_merge4.groupby(['id_x', 'department_name_ver2', 'đợt ăn (1,2)'])[['num_employee', 'báo cơm trưa', 'báo cơm chiều']].sum().reset_index()
    df_merge4 = df_merge4[['đợt ăn (1,2)', 'id_x', 'department_name_ver2', 'num_employee', 'báo cơm trưa', 'báo cơm chiều']]

    df_merge4.sort_values(by=['đợt ăn (1,2)', 'id_x', 'department_name_ver2'], inplace=True)
    df_merge4.reset_index(drop=True, inplace=True)

    df_merge4['NO'] = df_merge4['id_x'].str[0]
    replace_map_df4 = {
        'a': 'A  厰',
        'b': 'B  厰',
        'c': 'C  厰',
        's': ''
    }
    df_merge4['NO'] = df_merge4['NO'].replace(replace_map_df4)

    # output type
    df_merge4 = df_merge4[['NO', 'department_name_ver2', 'num_employee', 'báo cơm trưa', 'báo cơm chiều']]
    df_merge4.columns = ['NO', '單位', '人員出勤實到', '中午餐人數', '下午餐人數']
    df_merge4.sort_values(by=['NO', '單位'], inplace=True)
    df_merge4.reset_index(drop=True, inplace=True)

    return df_merge4


# 2. Read the excel files
# # file path
# path_bao_bieu = r"C:\LAN DIEP\HR_TEAM\data\BAO BIEU MOI NGAY 08.05.2025.xlsx"
# path_bao_com = r"C:\LAN DIEP\HR_TEAM\data\BAO COM MOI NGAY 08.05.2025.xlsx"
# path_dept_id = r"C:\LAN DIEP\HR_TEAM\data\department_id.xlsx"

# # read excel file
# df_bao_bieu = pd.read_excel(path_bao_bieu)
# df_bao_com = pd.read_excel(path_bao_com, sheet_name="Câu trả lời biểu mẫu 1")
# df_dept_id = pd.read_excel(path_dept_id)

# # preprocessing
# df_bao_bieu = process_bao_bieu(df_bao_bieu)
# df_bao_com = process_bao_com(df_bao_com)
# df_dept_id = process_dept_id(df_dept_id)

# # 3. Create the final merged dataframe
# df_work = create_work_df(df_bao_bieu, df_dept_id)
# df_lunch = create_lunch_df(df_bao_com, df_dept_id)
# df_merge4 = create_merged_df(df_work, df_lunch)
# print(df_merge4.head())

import streamlit as st
import pandas as pd
from io import BytesIO
from sorted import process_excel_file
from merged import process_bao_com, process_bao_bieu, process_dept_id, create_work_df, create_lunch_df, create_merged_df

st.set_page_config(page_title="Báo cơm xử lý", layout="wide")
st.title("🍱 Ứng dụng xử lý 'BÁO CƠM' và 'BÁO BIỂU'")

# --- Lựa chọn chức năng ---
mode = st.radio("🔧 Chọn chế độ xử lý:", ["📄 Sorted BÁO CƠM", "📊 Merged BÁO CƠM + BÁO BIỂU"])

# --- Sorted Mode ---
if mode == "📄 Sorted BÁO CƠM":
    st.header("1️⃣ Sorted 'BÁO CƠM'")

    import streamlit as st

    container = st.container(border=True)
    with container:
         uploaded_file = st.file_uploader("📁 Chọn file Excel báo cơm", type=["xlsx", "xls"], key="bao_com", accept_multiple_files=False)

    # uploaded_file = st.file_uploader("📁 Chọn file Excel báo cơm", type=["xlsx", "xls"], key="bao_com", accept_multiple_files=False)
    print('upload done!')
    
    if uploaded_file is not None:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet = st.sidebar.selectbox("Chọn sheet", excel_file.sheet_names)
            df = pd.read_excel(excel_file, sheet_name=sheet)
            st.write(f"### 📄 Sheet: {sheet}")
            st.dataframe(df)

            processed_df = process_excel_file(df)

            st.write("### ✅ File sau khi xử lý:")
            st.dataframe(processed_df)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                processed_df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="📥 Tải xuống file đã xử lý",
                data=output,
                file_name=f"processed_{sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            # st.error(f"❌ Lỗi khi xử lý file Excel: {e}")
            st.exception(e)

# --- Merged Mode ---
elif mode == "📊 Merged BÁO CƠM + BÁO BIỂU":
    st.header("2️⃣ Gộp 'BÁO CƠM' + 'BÁO BIỂU'")
    uploaded_bao_com = st.file_uploader("📁 Chọn file báo cơm", type=["xlsx"], key="com")
    uploaded_bao_bieu = st.file_uploader("📁 Chọn file báo biểu", type=["xlsx"], key="bieu")
    # uploaded_dept = st.file_uploader("📁 Chọn file department_id", type=["xlsx"], key="dept")

    if uploaded_bao_com and uploaded_bao_bieu:
        try:
            df_bao_com = pd.read_excel(uploaded_bao_com, sheet_name="Câu trả lời biểu mẫu 1")
            df_bao_bieu = pd.read_excel(uploaded_bao_bieu)
            # df_dept_id = pd.read_excel(uploaded_dept)
            df_dept_id = pd.read_excel('department_id.xlsx')

            # Xử lý
            df_bao_com = process_bao_com(df_bao_com)
            df_bao_bieu = process_bao_bieu(df_bao_bieu)
            df_dept_id = process_dept_id(df_dept_id)

            df_work = create_work_df(df_bao_bieu, df_dept_id)
            df_lunch = create_lunch_df(df_bao_com, df_dept_id)
            df_merged = create_merged_df(df_work, df_lunch)

            st.success("✅ Đã xử lý xong dữ liệu gộp.")
            st.dataframe(df_merged)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_merged.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="📥 Tải xuống kết quả đã gộp",
                data=output,
                file_name="merged_baocom_baobieu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Lỗi xử lý: {e}")












import streamlit as st
import pandas as pd
from openpyxl import Workbook
from datetime import datetime, timedelta
import tempfile
import re
import os

st.set_page_config(
    page_title="Attendance Converter",
    layout="wide"
)

st.title("Attendance Converter Tool")
st.write("حول ملفات البصمة إلى ملف جاهز للاستيراد")

uploaded_file = st.file_uploader(
    "ارفع ملف البصمة",
    type=["xls", "xlsx"]
)

if uploaded_file:

    try:

        # حفظ الملف مؤقتًا
        suffix = os.path.splitext(uploaded_file.name)[1]

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.read())
            temp_path = tmp.name

        # قراءة ملف الإكسل
        engine = "openpyxl"

        if suffix == ".xls":
            engine = "xlrd"

        excel_file = pd.ExcelFile(
            temp_path,
            engine=engine
        )

        # اختيار الشيت
        sheet_name = "Att.log report"

        if sheet_name not in excel_file.sheet_names:
            sheet_name = excel_file.sheet_names[0]

        df = pd.read_excel(
            temp_path,
            sheet_name=sheet_name,
            header=None,
            engine=engine
        )

        # استخراج تاريخ البداية
        date_text = str(df.iloc[2, 2])

        # مثال:
        # 2026-05-01 ~ 2026-05-11

        start_date_str = date_text.split("~")[0].strip()

        start_date = datetime.strptime(
            start_date_str,
            "%Y-%m-%d"
        )

        # إنشاء ملف الإخراج
        out_wb = Workbook()
        out_ws = out_wb.active

        out_ws.title = "Attendance"

        out_ws.append([
            "ref_no",
            "date",
            "check_in_at",
            "check_out_at"
        ])

        # تحديد أعمدة الأيام
        day_columns = {}

        for col in range(df.shape[1]):

            value = df.iloc[3, col]

            if pd.notna(value):

                try:
                    day_num = int(value)

                    current_date = start_date + timedelta(days=day_num - 1)

                    day_columns[col] = current_date.strftime("%Y-%m-%d")

                except:
                    pass

        current_id = None

        # قراءة البيانات
        for row_idx in range(df.shape[0]):

            row = df.iloc[row_idx]

            values = row.tolist()

            # استخراج ID
            if len(values) > 2 and values[0] == "ID:":

                if pd.notna(values[2]):
                    current_id = str(values[2]).strip()

                continue

            # استخراج البصمات
            if current_id:

                has_time = any(
                    isinstance(v, str) and ":" in v
                    for v in values
                )

                if has_time:

                    for col_idx, date_str in day_columns.items():

                        val = df.iloc[row_idx, col_idx]

                        if pd.notna(val):

                            val = str(val)

                            if ":" in val:

                                # استخراج كل الأوقات
                                times = re.findall(
                                    r"\d{2}:\d{2}",
                                    val
                                )

                                if times:

                                    check_in = times[0]
                                    check_out = times[-1]

                                    out_ws.append([
                                        current_id,
                                        date_str,
                                        check_in,
                                        check_out
                                    ])

                    current_id = None

        # حفظ الملف النهائي
        output_file = "attendance_ready.xlsx"

        out_wb.save(output_file)

        st.success("تم تحويل الملف بنجاح")

        with open(output_file, "rb") as file:

            st.download_button(
                label="Download Ready File",
                data=file,
                file_name="attendance_ready.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error(f"حدث خطأ: {str(e)}")
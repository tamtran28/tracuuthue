import streamlit as st
import pandas as pd
import requests
import time
import re
import os
from pathlib import Path
from io import BytesIO
from datetime import datetime

# ================= CONFIG =================
OUTPUT_DIR = "output_files"
os.makedirs(OUTPUT_DIR, exist_ok=True)

TIMEOUT = 20

MAX_RETRY_PER_MST = 3
SLEEP_OK = 0.6
SLEEP_RETRY = 30
SLEEP_RATE_LIMIT = 90

# ================= FUNCTION =================

def clean_mst(mst):
    """Chuẩn hóa MST, chỉ giữ số."""
    return re.sub(r"[^0-9]", "", str(mst))


def normalize_status(text):
    """Chuẩn hóa trạng thái doanh nghiệp."""
    if not text:
        return ""

    t = str(text).lower()

    if "đang hoạt động" in t:
        return "Đang hoạt động"
    if "ngừng hoạt động" in t:
        return "Ngừng hoạt động"
    if "chấm dứt" in t:
        return "Chấm dứt hiệu lực MST"

    return str(text).strip()


def query_vietqr(mst):
    """Gọi API VietQR tra cứu MST."""
    url = f"https://api.vietqr.io/v2/business/{mst}"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json"
    }

    try:
        r = requests.get(url, headers=headers, timeout=TIMEOUT)

        if "Too many requests" in r.text:
            return "RATE_LIMIT", None

        js = r.json()

        if js.get("code") != "00":
            return "NOT_FOUND", None

        data = js.get("data", {})

        return "OK", {
            "MST": mst,
            "Tên DN": data.get("name", ""),
            "Địa chỉ": data.get("address", ""),
            "Trạng thái": normalize_status(data.get("status", "")),
            "Nguồn": "vietqr.io",
            "Thời gian tra cứu": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    except Exception as e:
        return "ERROR", str(e)


def tra_cuu_1_mst(mst):
    """Tra cứu 1 MST có retry."""
    attempt = 0
    final_row = None

    while attempt < MAX_RETRY_PER_MST:
        status, data = query_vietqr(mst)

        if status == "OK":
            final_row = data
            time.sleep(SLEEP_OK)
            break

        elif status == "NOT_FOUND":
            final_row = {
                "MST": mst,
                "Tên DN": "",
                "Địa chỉ": "",
                "Trạng thái": "Không tồn tại MST",
                "Nguồn": "vietqr.io",
                "Thời gian tra cứu": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            time.sleep(0.5)
            break

        elif status == "ERROR":
            attempt += 1
            time.sleep(SLEEP_RETRY)

        elif status == "RATE_LIMIT":
            attempt += 1
            time.sleep(SLEEP_RATE_LIMIT)

    if final_row is None:
        final_row = {
            "MST": mst,
            "Tên DN": "",
            "Địa chỉ": "",
            "Trạng thái": "Không tra được",
            "Nguồn": "vietqr.io",
            "Thời gian tra cứu": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    return final_row


def convert_df_to_excel(df):
    """Xuất DataFrame ra Excel để tải xuống."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Ket_qua_tra_cuu")

        workbook = writer.book
        worksheet = writer.sheets["Ket_qua_tra_cuu"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1
        })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 25)

    output.seek(0)
    return output


# ================= STREAMLIT UI =================

st.set_page_config(
    page_title="Tra cứu MST doanh nghiệp",
    page_icon="🔎",
    layout="wide"
)

st.title("🔎 Web nội bộ tra cứu MST doanh nghiệp")
st.caption("Upload file Excel có cột MST để tra cứu thông tin doanh nghiệp từ VietQR.")

with st.sidebar:
    st.header("⚙️ Cấu hình")

    max_retry = st.number_input(
        "Số lần thử lại mỗi MST",
        min_value=1,
        max_value=10,
        value=MAX_RETRY_PER_MST
    )

    sleep_ok = st.number_input(
        "Thời gian nghỉ khi tra thành công",
        min_value=0.1,
        max_value=10.0,
        value=SLEEP_OK,
        step=0.1
    )

    sleep_retry = st.number_input(
        "Thời gian nghỉ khi lỗi tạm",
        min_value=1,
        max_value=120,
        value=SLEEP_RETRY
    )

    sleep_rate_limit = st.number_input(
        "Thời gian nghỉ khi bị rate limit",
        min_value=10,
        max_value=300,
        value=SLEEP_RATE_LIMIT
    )

    st.warning(
        "Không nên để thời gian nghỉ quá thấp vì có thể bị chặn do gọi API quá nhanh."
    )

uploaded_files = st.file_uploader(
    "📂 Upload một hoặc nhiều file Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"Đã upload {len(uploaded_files)} file.")

    for uploaded_file in uploaded_files:
        st.divider()
        st.subheader(f"📄 File: {uploaded_file.name}")

        try:
            df = pd.read_excel(uploaded_file, dtype=str)
            df.columns = df.columns.str.strip()

            st.write("### Xem trước dữ liệu")
            st.dataframe(df.head(10), use_container_width=True)

            column_options = list(df.columns)

            default_index = column_options.index("MST") if "MST" in column_options else 0

            mst_col = st.selectbox(
                f"Chọn cột chứa MST cho file {uploaded_file.name}",
                column_options,
                index=default_index,
                key=f"mst_col_{uploaded_file.name}"
            )

            if st.button(f"🚀 Bắt đầu tra cứu file {uploaded_file.name}", key=f"run_{uploaded_file.name}"):

                # Cập nhật config từ sidebar
                MAX_RETRY_PER_MST = int(max_retry)
                SLEEP_OK = float(sleep_ok)
                SLEEP_RETRY = int(sleep_retry)
                SLEEP_RATE_LIMIT = int(sleep_rate_limit)

                work_df = df.copy()

                work_df["MST_CLEAN"] = work_df[mst_col].apply(clean_mst)
                work_df = work_df[work_df["MST_CLEAN"] != ""]

                if work_df.empty:
                    st.error("Không có MST hợp lệ để tra cứu.")
                    continue

                unique_mst = work_df["MST_CLEAN"].drop_duplicates().tolist()

                st.info(f"Tổng số dòng hợp lệ: {len(work_df)}")
                st.info(f"Số MST không trùng cần tra cứu: {len(unique_mst)}")

                progress_bar = st.progress(0)
                status_text = st.empty()
                result_area = st.empty()

                results = []
                result_map = {}

                for idx, mst in enumerate(unique_mst, start=1):
                    status_text.write(f"Đang tra cứu {idx}/{len(unique_mst)}: MST {mst}")

                    row = tra_cuu_1_mst(mst)
                    results.append(row)
                    result_map[mst] = row

                    progress_bar.progress(idx / len(unique_mst))

                    if len(results) % 10 == 0:
                        temp_df = pd.DataFrame(results)
                        result_area.dataframe(temp_df.tail(10), use_container_width=True)

                result_df = pd.DataFrame(results)

                # Ghép kết quả tra cứu về file gốc
                result_lookup = result_df.set_index("MST")

                work_df["Tên DN"] = work_df["MST_CLEAN"].map(result_lookup["Tên DN"])
                work_df["Địa chỉ"] = work_df["MST_CLEAN"].map(result_lookup["Địa chỉ"])
                work_df["Trạng thái"] = work_df["MST_CLEAN"].map(result_lookup["Trạng thái"])
                work_df["Nguồn"] = work_df["MST_CLEAN"].map(result_lookup["Nguồn"])
                work_df["Thời gian tra cứu"] = work_df["MST_CLEAN"].map(result_lookup["Thời gian tra cứu"])

                work_df = work_df.drop(columns=["MST_CLEAN"])

                st.success("✅ Hoàn tất tra cứu!")

                st.write("### Kết quả tra cứu")
                st.dataframe(work_df, use_container_width=True)

                output_excel = convert_df_to_excel(work_df)

                output_name = uploaded_file.name.replace(".xlsx", "_ket_qua_tra_cuu.xlsx")

                st.download_button(
                    label="📥 Tải file kết quả Excel",
                    data=output_excel,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Lỗi khi xử lý file {uploaded_file.name}: {e}")

else:
    st.info("Vui lòng upload file Excel để bắt đầu.")

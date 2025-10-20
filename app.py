import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import re
from copy import copy
import logging
import io
import zipfile
import tempfile

# --- CẤU HÌNH LOGGING ---
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- CẤU HÌNH CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ ---
TOOL1_COLUMN_MAPPING = {
    'A': 'T', 'B': 'U', 'C': 'Y', 'D': 'C', 'E': 'H',
    'F': 'I', 'G': 'X', 'I': 'K', 'N': 'AY'
}
TOOL1_START_ROW_DESTINATION = 7

# --- CẤU HÌNH CÔNG CỤ 2 & 3: LÀM SẠCH & TÁCH FILE ---
STEP1_CHECK_COLS = ["D", "E", "F", "I", "J", "L", "M", "R", "S", "T", "U"]
STEP1_START_ROW = 5
STEP1_YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
STEP1_EMPTY_FILL = PatternFill(fill_type=None)
STEP2_TARGET_COL = "G"
STEP2_START_ROW = 5
STEP2_EMPTY_FILL = PatternFill(fill_type=None)


# --- CÁC HÀM HELPER CHUNG ---

def helper_copy_cell_format(src_cell, tgt_cell):
    if src_cell.has_style:
        tgt_cell.font = copy(src_cell.font)
        tgt_cell.border = copy(src_cell.border)
        tgt_cell.fill = copy(src_cell.fill)
        tgt_cell.number_format = copy(src_cell.number_format)
        tgt_cell.protection = copy(src_cell.protection)
        tgt_cell.alignment = copy(src_cell.alignment)

def helper_normalize_value(val):
    if pd.isna(val) or val is None:
        return np.nan
    str_val = str(val).strip()
    str_val = re.sub(r'\s+', ' ', str_val)
    return str_val.lower() if str_val else np.nan

def helper_calculate_column_width(ws):
    for col in range(1, ws.max_column + 1):
        max_length = 0
        column_letter = get_column_letter(col)
        for cell in ws[column_letter]:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max(max_length + 2, 8), 60)
        ws.column_dimensions[column_letter].width = adjusted_width

def helper_cell_has_bg(c):
    try:
        if c.fill and c.fill.fgColor and c.fill.fgColor.rgb:
            rgb_val = str(c.fill.fgColor.rgb).upper()
            return rgb_val not in ('00000000', 'FFFFFFFF')
        return False
    except:
        return False
        
# --- CÁC HÀM CHO CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ ---

def tool1_excel_col_to_index(col_letter):
    index = 0
    for char in col_letter.upper():
        index = index * 26 + (ord(char) - ord('A')) + 1
    return index - 1

def tool1_get_sheet_names_from_buffer(file_buffer):
    try:
        wb = load_workbook(file_buffer, read_only=True)
        return wb.sheetnames
    except Exception as e:
        st.error(f"Không thể đọc sheet từ file: {e}")
        return []

def tool1_transform_and_copy(source_buffer, source_sheet, dest_buffer, dest_sheet, progress_bar, status_label):
    try:
        # 1. Đọc dữ liệu nguồn
        status_label.info("Đang đọc dữ liệu từ file nguồn...")
        source_cols_letters = list(TOOL1_COLUMN_MAPPING.keys())
        source_cols_indices = [tool1_excel_col_to_index(col) for col in source_cols_letters]
        df_source = pd.read_excel(source_buffer, sheet_name=source_sheet, header=None, skiprows=2, usecols=source_cols_indices, engine='openpyxl')
        df_source.columns = source_cols_letters
        df_source_renamed = df_source.rename(columns=TOOL1_COLUMN_MAPPING)
        progress_bar.progress(20)

        # 2. Mở workbook đích
        status_label.info("Đang mở file đích để ghi dữ liệu...")
        wb_dest = load_workbook(dest_buffer)
        if dest_sheet not in wb_dest.sheetnames:
            st.error(f"Lỗi: Không tìm thấy sheet '{dest_sheet}' trong file đích.")
            return None
        ws_dest = wb_dest[dest_sheet]
        progress_bar.progress(40)

        # 3. Ghi dữ liệu
        status_label.info("Đang sao chép dữ liệu...")
        dest_cols = list(TOOL1_COLUMN_MAPPING.values())
        total_rows = len(df_source_renamed)
        for i, dest_col in enumerate(dest_cols):
            col_index_dest = tool1_excel_col_to_index(dest_col)
            for j, value in enumerate(df_source_renamed[dest_col], start=TOOL1_START_ROW_DESTINATION):
                cell_value = value if pd.notna(value) else None
                ws_dest.cell(row=j, column=col_index_dest + 1, value=cell_value)
            
            progress_bar.progress(40 + int((i + 1) / len(dest_cols) * 40))

        # 4. Kẻ viền
        status_label.info("Đang kẻ viền cho vùng dữ liệu mới...")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        end_row_border = TOOL1_START_ROW_DESTINATION + total_rows - 1
        for row in ws_dest.iter_rows(min_row=TOOL1_START_ROW_DESTINATION, max_row=end_row_border, min_col=1, max_col=50): # A -> AX
            for cell in row:
                cell.border = thin_border
        progress_bar.progress(95)

        # 5. Lưu kết quả vào buffer
        status_label.info("Đang lưu kết quả...")
        output_buffer = io.BytesIO()
        wb_dest.save(output_buffer)
        output_buffer.seek(0)
        progress_bar.progress(100)
        return output_buffer

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
        logging.error(f"Lỗi Công cụ 1: {e}")
        return None

# --- CÁC HÀM CHO CÔNG CỤ 2 & 3: LÀM SẠCH, PHÂN LOẠI & TÁCH FILE ---

def run_step_1_process(wb, sheet_name, progress_bar, status_label, base_percent, step_budget):
    # (Giữ nguyên code gốc của bạn)
    # ...
    try:
        def update_progress(local_percent, step_text=""):
            status_label.info(f"Bước 1: {step_text}...")
            master_percent = base_percent + (local_percent / 100) * step_budget
            progress_bar.progress(int(master_percent))

        update_progress(0, "Kiểm tra sheet")
        if sheet_name not in wb.sheetnames:
            st.error(f"Lỗi Bước 1: Không tìm thấy sheet '{sheet_name}'.")
            return None
        ws = wb[sheet_name]
        last_row = ws.max_row
        
        update_progress(10, "Tìm hàng thiếu dữ liệu")
        rows_to_color = set()
        for row_idx in range(STEP1_START_ROW, last_row + 1):
            for col in STEP1_CHECK_COLS:
                cell_value = ws[f"{col}{row_idx}"].value
                if cell_value is None or str(cell_value).strip() == "":
                    rows_to_color.add(row_idx)
                    break
        
        update_progress(30, "Xóa màu nền cũ")
        for row in ws.iter_rows(min_row=1, max_row=last_row):
            for cell in row:
                cell.fill = STEP1_EMPTY_FILL

        update_progress(40, "Tô màu vàng các hàng thiếu dữ liệu")
        for row_idx in rows_to_color:
            for cell in ws[row_idx]:
                cell.fill = STEP1_YELLOW_FILL

        update_progress(50, "Chuẩn bị tách sheet")
        def copy_to_new_sheet(title, condition_fn):
            if title in wb.sheetnames:
                wb.remove(wb[title])
            ws_dst = wb.create_sheet(title)
            # Copy header
            for r in range(1, 5):
                for c in range(1, ws.max_column + 1):
                    src = ws.cell(row=r, column=c)
                    dst = ws_dst.cell(row=r, column=c)
                    dst.value = src.value
                    if src.has_style:
                        helper_copy_cell_format(src, dst)
            # Copy data rows
            next_row = 5
            for r in range(5, last_row + 1):
                if condition_fn(r):
                    for c in range(1, ws.max_column + 1):
                        src = ws.cell(row=r, column=c)
                        dst = ws_dst.cell(row=next_row, column=c)
                        dst.value = src.value
                        if src.has_style:
                            helper_copy_cell_format(src, dst)
                    next_row += 1
            helper_calculate_column_width(ws_dst)
        
        update_progress(60, "Tạo sheet 'Nhóm 1' (đủ dữ liệu)")
        copy_to_new_sheet("Nhóm 1", lambda r_idx: r_idx not in rows_to_color)
        
        update_progress(80, "Tạo sheet 'Nhóm 2' (thiếu dữ liệu)")
        copy_to_new_sheet("Nhóm 2", lambda r_idx: r_idx in rows_to_color)
        
        update_progress(100, "Hoàn tất Bước 1")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 1): {e}")
        logging.error(f"Lỗi Bước 1: {e}")
        return None

def run_step_2_clear_fill(wb, progress_bar, status_label, base_percent, step_budget):
    # (Giữ nguyên code gốc của bạn)
    # ...
    try:
        TARGET_SHEET = "Nhóm 2"
        def update_progress(local_percent, step_text=""):
            status_label.info(f"Bước 2: {step_text}...")
            master_percent = base_percent + (local_percent / 100) * step_budget
            progress_bar.progress(int(master_percent))

        update_progress(0, f"Kiểm tra sheet '{TARGET_SHEET}'")
        if TARGET_SHEET not in wb.sheetnames:
            st.warning(f"Cảnh báo (Bước 2): Không tìm thấy sheet '{TARGET_SHEET}', bỏ qua.")
            update_progress(100, "Bỏ qua")
            return wb
        ws = wb[TARGET_SHEET]
        
        update_progress(20, "Xóa màu theo điều kiện cột G")
        for row_idx in range(STEP2_START_ROW, ws.max_row + 1):
            cell_g_val = ws[f"{STEP2_TARGET_COL}{row_idx}"].value
            if cell_g_val is not None and str(cell_g_val).strip() != "":
                for cell_in_row in ws[row_idx]:
                    cell_in_row.fill = STEP2_EMPTY_FILL
        
        update_progress(100, "Hoàn tất Bước 2")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 2): {e}")
        logging.error(f"Lỗi Bước 2: {e}")
        return None

def run_step_3_split_by_color(wb, progress_bar, status_label, base_percent, step_budget):
    # (Giữ nguyên code gốc của bạn)
    # ...
    try:
        TARGET_SHEET = "Nhóm 2"
        def update_progress(local_percent, step_text=""):
            status_label.info(f"Bước 3: {step_text}...")
            master_percent = base_percent + (local_percent / 100) * step_budget
            progress_bar.progress(int(master_percent))

        update_progress(0, f"Kiểm tra sheet '{TARGET_SHEET}'")
        if TARGET_SHEET not in wb.sheetnames:
            st.warning(f"Cảnh báo (Bước 3): Không tìm thấy sheet '{TARGET_SHEET}', bỏ qua.")
            update_progress(100, "Bỏ qua")
            return wb
        ws_src = wb[TARGET_SHEET]

        def copy_to_new_sheet(title, condition_fn):
            if title in wb.sheetnames:
                wb.remove(wb[title])
            ws_dst = wb.create_sheet(title)
            # Copy header
            for r in range(1, 5):
                for c in range(1, ws_src.max_column + 1):
                    src = ws_src.cell(row=r, column=c)
                    dst = ws_dst.cell(row=r, column=c)
                    dst.value = src.value
                    if src.has_style:
                        helper_copy_cell_format(src, dst)
            # Copy data
            next_row = 5
            for r in range(5, ws_src.max_row + 1):
                if condition_fn(ws_src.cell(row=r, column=1)):
                    for c in range(1, ws_src.max_column + 1):
                        src = ws_src.cell(row=r, column=c)
                        dst = ws_dst.cell(row=next_row, column=c)
                        dst.value = src.value
                        if src.has_style:
                            helper_copy_cell_format(src, dst)
                    next_row += 1
            helper_calculate_column_width(ws_dst)

        update_progress(25, "Tạo sheet 'Nhóm 2_TC' (không màu)")
        copy_to_new_sheet("Nhóm 2_TC", lambda c: not helper_cell_has_bg(c))
        
        update_progress(75, "Tạo sheet 'Nhóm 2_GDC' (còn màu)")
        copy_to_new_sheet("Nhóm 2_GDC", lambda c: helper_cell_has_bg(c))
        
        update_progress(100, "Hoàn tất Bước 3")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 3): {e}")
        logging.error(f"Lỗi Bước 3: {e}")
        return None

def run_step_4_split_files(data_buffer, progress_bar, status_label, base_percent, step_budget):
    # (Giữ nguyên code gốc của bạn, chỉ điều chỉnh lại một chút cho rõ ràng)
    # ...
    DATA_SHEET = "Nhóm 2_GDC"
    TEMPLATE_SHEET = "TongHop"
    FILTER_COLUMN = "T"
    START_ROW = 5
    
    try:
        def update_progress(local_percent, step_text=""):
            status_label.info(f"Bước 4: {step_text}...")
            master_percent = base_percent + (local_percent / 100) * step_budget
            progress_bar.progress(int(master_percent))

        update_progress(0, "Đọc dữ liệu từ bộ nhớ")
        wb_main = load_workbook(data_buffer, data_only=True)
        
        if TEMPLATE_SHEET not in wb_main.sheetnames or DATA_SHEET not in wb_main.sheetnames:
            st.error(f"Lỗi: File đầu vào phải chứa sheet '{TEMPLATE_SHEET}' và '{DATA_SHEET}'.")
            return None

        template_ws = wb_main[TEMPLATE_SHEET]
        data_buffer.seek(0)
        df = pd.read_excel(data_buffer, sheet_name=DATA_SHEET, header=None)

        update_progress(10, f"Lọc giá trị duy nhất từ cột '{FILTER_COLUMN}'")
        col_index = column_index_from_string(FILTER_COLUMN) - 1
        data_col = df.iloc[START_ROW - 1:, col_index].apply(helper_normalize_value)
        unique_values = data_col.dropna().unique().tolist()
        if data_col.isnull().any():
            unique_values.append("BLANK")

        if not unique_values:
            st.warning("Không tìm thấy giá trị nào để tách file.")
            return None

        update_progress(20, f"Chuẩn bị tách thành {len(unique_values)} file con")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f, tempfile.TemporaryDirectory() as tmpdir:
            total = len(unique_values)
            for i, value in enumerate(unique_values, start=1):
                mask = data_col.isnull() if value == "BLANK" else (data_col == value)
                filtered_df = df.iloc[START_ROW - 1:][mask]
                
                if not filtered_df.empty:
                    new_wb = Workbook()
                    new_ws = new_wb.active
                    new_ws.title = "DuLieuLoc"
                    
                    # Copy header from template
                    for r in range(1, 4):
                        for c in range(1, template_ws.max_column + 1):
                            src = template_ws.cell(row=r, column=c)
                            dst = new_ws.cell(row=r, column=c)
                            dst.value = src.value
                            helper_copy_cell_format(src, dst)
                    
                    for r_idx, row_data in enumerate(dataframe_to_rows(filtered_df, index=False, header=False), start=4):
                        for c_idx, cell_val in enumerate(row_data, start=1):
                            new_ws.cell(row=r_idx, column=c_idx, value=cell_val)

                    helper_calculate_column_width(new_ws)
                    
                    safe_name = "BLANK" if value == "BLANK" else re.sub(r'[\\/*?:<>|"\t\n\r]+', "_", str(value).strip())[:50]
                    output_path = os.path.join(tmpdir, f"{safe_name}.xlsx")
                    new_wb.save(output_path)
                    zip_f.write(output_path, arcname=os.path.basename(output_path))
                
                local_percent = 20 + (i / total) * 75
                update_progress(local_percent, f"Đang tách file {i}/{total}")

        update_progress(100, "Hoàn tất nén file ZIP")
        zip_buffer.seek(0)
        return zip_buffer

    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 4): {e}")
        logging.error(f"Lỗi Bước 4: {e}")
        return None

# --- GIAO DIỆN STREAMLIT CHÍNH ---

st.set_page_config(page_title="Công cụ Dữ liệu Đất đai", layout="wide", page_icon="📊")

with st.sidebar:
    st.image("https://i.imgur.com/v12A61a.png", width=150)
    st.title("Hướng dẫn sử dụng")
    st.info("**Công cụ 1: Sao chép & Ánh xạ Cột**\n\n- Tải lên file Nguồn và file Đích.\n- Chọn sheet tương ứng.\n- Công cụ sẽ sao chép dữ liệu từ nguồn sang đích theo cấu hình định sẵn.")
    st.info("**Công cụ 2: Làm sạch & Phân loại**\n\n- Tải file Excel gốc, chọn sheet.\n- Công cụ sẽ làm sạch, tô màu và phân loại dữ liệu thành các sheet `Nhóm 1`, `Nhóm 2`...")
    st.info("**Công cụ 3: Tách file theo Thôn**\n\n- Tải file đã xử lý bởi Công cụ 2.\n- Công cụ sẽ tách sheet `Nhóm 2_GDC` thành nhiều file con và nén lại thành tệp ZIP.")
    st.success("Phát triển bởi: **Trường Sinh**\n\nSĐT: **0917.750.555**")

st.title("📊 Tổng hợp Công cụ Hỗ trợ Xử lý Dữ liệu Đất đai")
st.markdown("---")

tab1, tab2, tab3 = st.tabs([
    " CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ CỘT ", 
    " CÔNG CỤ 2: LÀM SẠCH & PHÂN LOẠI ", 
    " CÔNG CỤ 3: TÁCH FILE THEO THÔN "
])

# --- GIAO DIỆN CÔNG CỤ 1 ---
with tab1:
    st.header("Chuyển đổi và sao chép dữ liệu giữa hai file Excel")
    
    col1, col2 = st.columns(2)
    with col1:
        source_file = st.file_uploader("1. Tải lên File Nguồn (lấy dữ liệu)", type=["xlsx", "xls"], key="tool1_source")
        if source_file:
            source_sheets = tool1_get_sheet_names_from_buffer(source_file)
            selected_source_sheet = st.selectbox("2. Chọn Sheet Nguồn:", source_sheets, key="tool1_source_sheet")
    
    with col2:
        dest_file = st.file_uploader("3. Tải lên File Đích (nhận dữ liệu)", type=["xlsx", "xls"], key="tool1_dest")
        if dest_file:
            dest_sheets = tool1_get_sheet_names_from_buffer(dest_file)
            selected_dest_sheet = st.selectbox("4. Chọn Sheet Đích:", dest_sheets, key="tool1_dest_sheet")

    if st.button("BẮT ĐẦU SAO CHÉP DỮ LIỆU", type="primary", key="tool1_start"):
        if source_file and dest_file and selected_source_sheet and selected_dest_sheet:
            progress_bar_1 = st.progress(0)
            status_text_1 = st.empty()
            
            result_buffer = tool1_transform_and_copy(
                source_file, selected_source_sheet,
                dest_file, selected_dest_sheet,
                progress_bar_1, status_text_1
            )
            
            if result_buffer:
                status_text_1.success("✅ HOÀN TẤT! Vui lòng tải file đích đã được cập nhật về.")
                st.download_button(
                    label="📥 Tải về File Đích đã cập nhật",
                    data=result_buffer,
                    file_name=f"[Updated]_{dest_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Vui lòng tải lên cả hai file và chọn sheet tương ứng.")

# --- GIAO DIỆN CÔNG CỤ 2 ---
with tab2:
    st.header("Xử lý file tổng, tạo các nhóm dữ liệu")
    uploaded_file_2 = st.file_uploader("1. Tải lên file Excel cần xử lý", type=["xlsx", "xlsm"], key="tool2_uploader")

    if uploaded_file_2:
        try:
            sheets_2 = tool1_get_sheet_names_from_buffer(uploaded_file_2)
            selected_sheet_2 = st.selectbox("2. Chọn sheet chính chứa dữ liệu:", sheets_2, key="tool2_sheet")

            if st.button("BẮT ĐẦU LÀM SẠCH & PHÂN LOẠI", type="primary", key="tool2_start"):
                progress_bar_2 = st.progress(0)
                status_text_2 = st.empty()
                main_wb_2 = load_workbook(uploaded_file_2)
                
                main_wb_2 = run_step_1_process(main_wb_2, selected_sheet_2, progress_bar_2, status_text_2, 0, 33)
                if main_wb_2:
                    main_wb_2 = run_step_2_clear_fill(main_wb_2, progress_bar_2, status_text_2, 33, 33)
                if main_wb_2:
                    main_wb_2 = run_step_3_split_by_color(main_wb_2, progress_bar_2, status_text_2, 66, 34)

                if main_wb_2:
                    status_text_2.success("✅ HOÀN TẤT! Vui lòng tải file về.")
                    final_buffer_2 = io.BytesIO()
                    main_wb_2.save(final_buffer_2)
                    final_buffer_2.seek(0)
                    st.download_button(label="📥 Tải về File đã xử lý", data=final_buffer_2, file_name=f"[Processed]_{uploaded_file_2.name}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Lỗi: {e}")

# --- GIAO DIỆN CÔNG CỤ 3 ---
with tab3:
    st.header("Tách file từ sheet 'Nhóm 2_GDC' thành nhiều file con")
    uploaded_file_3 = st.file_uploader("1. Tải lên file Excel ĐÃ ĐƯỢC XỬ LÝ bởi Công cụ 2", type=["xlsx", "xlsm"], key="tool3_uploader", help="File này phải chứa sheet 'Nhóm 2_GDC' và 'TongHop'.")

    if uploaded_file_3:
        if st.button("BẮT ĐẦU TÁCH FILE", type="primary", key="tool3_start"):
            progress_bar_3 = st.progress(0)
            status_text_3 = st.empty()
            data_buffer_3 = io.BytesIO(uploaded_file_3.getvalue())
            
            zip_buffer = run_step_4_split_files(data_buffer_3, progress_bar_3, status_text_3, 0, 100)

            if zip_buffer:
                status_text_3.success("✅ HOÀN TẤT! Vui lòng tải gói ZIP về.")
                st.download_button(label="🗂️ Tải về Gói file con (.zip)", data=zip_buffer, file_name="Cac_file_con_theo_thon.zip", mime="application/zip")


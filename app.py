import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
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
# Cấu hình logging để ghi lại các bước xử lý và lỗi có thể xảy ra
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- CẤU HÌNH CHUNG CHO CÁC BƯỚC ---
STEP1_CHECK_COLS = ["D", "E", "F", "I", "J", "L", "M", "R", "S", "T", "U"]
STEP1_START_ROW = 5
STEP1_YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
STEP1_EMPTY_FILL = PatternFill(fill_type=None)

STEP2_TARGET_COL = "G"
STEP2_START_ROW = 5
STEP2_EMPTY_FILL = PatternFill(fill_type=None)

# --- CÁC HÀM HELPER (HỖ TRỢ) ---

def helper_copy_cell_format(src_cell, tgt_cell):
    """(Helper) Sao chép định dạng từ cell nguồn sang cell đích."""
    if src_cell.has_style:
        tgt_cell.font = copy(src_cell.font)
        tgt_cell.border = copy(src_cell.border)
        tgt_cell.fill = copy(src_cell.fill)
        tgt_cell.number_format = copy(src_cell.number_format)
        tgt_cell.protection = copy(src_cell.protection)
        tgt_cell.alignment = copy(src_cell.alignment)

def helper_copy_rows_with_style(src_ws, tgt_ws, max_row=3):
    """(Helper) Copy N hàng đầu tiên (giá trị + định dạng + merge + độ rộng cột)."""
    for row_idx in range(1, max_row + 1):
        for col_idx, src_cell in enumerate(src_ws[row_idx], start=1):
            tgt_cell = tgt_ws.cell(row=row_idx, column=col_idx, value=src_cell.value)
            helper_copy_cell_format(src_cell, tgt_cell)

    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            tgt_ws.column_dimensions[col_letter].width = dim.width

    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.min_row <= max_row:
            tgt_ws.merge_cells(str(merged_range))

def helper_normalize_value(val):
    """(Helper) Chuẩn hóa giá trị: chuyển về str, loại bỏ khoảng trắng, xử lý NaN."""
    if pd.isna(val) or val is None:
        return np.nan
    str_val = str(val).strip()
    str_val = re.sub(r'\s+', ' ', str_val)
    return str_val.lower() if str_val else np.nan

def helper_group_columns_openpyxl(ws):
    """(Helper) Group các cột bằng openpyxl, an toàn cho môi trường online."""
    try:
        # Xóa group cũ (nếu có) để tránh lỗi
        for col in ws.column_dimensions:
            dim = ws.column_dimensions[col]
            if dim.outline_level > 0:
                dim.outline_level = 0
                dim.collapsed = False
        
        ranges_to_group = [("B", "C"), ("G", "H"), ("K", "K"), ("N", "Q"), ("W", "AY")]
        
        for start_col, end_col in ranges_to_group:
            start_idx = column_index_from_string(start_col)
            end_idx = column_index_from_string(end_col)
            
            for c_idx in range(start_idx, end_idx + 1):
                col_letter = get_column_letter(c_idx)
                if col_letter in ws.column_dimensions:
                    ws.column_dimensions[col_letter].outline_level = 1
        
        logging.info("✅ Group cột thành công bằng openpyxl")
    except Exception as e:
        logging.warning(f"⚠️ Không thể group cột bằng openpyxl: {e}")

def helper_calculate_column_width(ws):
    """(Helper) Tính toán độ rộng cột thủ công để thay thế cho auto-fit."""
    for col in range(1, ws.max_column + 1):
        max_length = 0
        column_letter = get_column_letter(col)
        for cell in ws[column_letter]:
            try:
                if cell.value:
                    cell_len = len(str(cell.value))
                    max_length = max(max_length, cell_len)
            except:
                pass
        # Đặt độ rộng hợp lý, tránh quá rộng hoặc quá hẹp
        adjusted_width = min(max(max_length + 2, 8), 60)
        ws.column_dimensions[column_letter].width = adjusted_width

def helper_get_safe_filepath(output_folder, name):
    """(Helper) Tạo tên tệp an toàn, tránh ghi đè khi lưu."""
    counter = 1
    safe_path = os.path.join(output_folder, f"{name}.xlsx")
    while os.path.exists(safe_path):
        safe_path = os.path.join(output_folder, f"{name}_{counter}.xlsx")
        counter += 1
    return safe_path

def helper_cell_has_bg(c):
    """(Helper) Kiểm tra một cell có màu nền hay không."""
    try:
        fg = getattr(c.fill, 'fgColor', None)
        if fg is None or fg.rgb is None:
            return False
        rgb_val = str(fg.rgb).upper()
        # Bỏ qua các màu nền mặc định (đen, trắng, không màu)
        if rgb_val in ('00000000', 'FFFFFFFF', '00FFFFFF', 'FF000000'):
            return False
        return True
    except:
        return False

# --- CÁC HÀM XỬ LÝ CHÍNH THEO TỪNG BƯỚC ---

def run_step_1_process(wb, sheet_name, progress_bar, status_label, base_percent, step_budget):
    """Bước 1: Tìm dòng trống, tô màu, và tách thành 'Nhóm 1', 'Nhóm 2'."""
    try:
        # Cập nhật giao diện
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
            helper_copy_rows_with_style(ws, ws_dst, max_row=4)
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
    """Bước 2: Trong 'Nhóm 2', xóa màu vàng ở hàng nào có dữ liệu ở cột G."""
    try:
        TARGET_SHEET = "Nhóm 2"
        # Cập nhật giao diện
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
    """Bước 3: Tách 'Nhóm 2' thành 'Nhóm 2_TC' (không màu) và 'Nhóm 2_GDC' (có màu)."""
    try:
        TARGET_SHEET = "Nhóm 2"
        # Cập nhật giao diện
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
            helper_copy_rows_with_style(ws_src, ws_dst, max_row=4)
            next_row = 5
            for r in range(5, ws_src.max_row + 1):
                # Kiểm tra màu ở cell cột A
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
    """Bước 4: Tách file từ sheet 'Nhóm 2_GDC' theo cột 'T'."""
    DATA_SHEET = "Nhóm 2_GDC"
    TEMPLATE_SHEET = "TongHop"
    FILTER_COLUMN = "T"
    START_ROW = 5
    
    try:
        # Cập nhật giao diện
        def update_progress(local_percent, step_text=""):
            status_label.info(f"Bước 4: {step_text}...")
            master_percent = base_percent + (local_percent / 100) * step_budget
            progress_bar.progress(int(master_percent))

        update_progress(0, "Đọc dữ liệu từ bộ nhớ")
        wb_main = load_workbook(data_buffer, data_only=True)
        
        if TEMPLATE_SHEET not in wb_main.sheetnames:
            st.error(f"Lỗi (Bước 4): Không tìm thấy sheet mẫu '{TEMPLATE_SHEET}'!")
            return None
        if DATA_SHEET not in wb_main.sheetnames:
            st.error(f"Lỗi (Bước 4): Không tìm thấy sheet dữ liệu '{DATA_SHEET}'!")
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
                    
                    helper_copy_rows_with_style(template_ws, new_ws, max_row=3)
                    
                    for r_idx, row_data in enumerate(dataframe_to_rows(filtered_df, index=False, header=False), start=4):
                        for c_idx, cell_val in enumerate(row_data, start=1):
                            new_ws.cell(row=r_idx, column=c_idx, value=cell_val)

                    helper_group_columns_openpyxl(new_ws)
                    helper_calculate_column_width(new_ws)
                    
                    safe_name = "BLANK" if value == "BLANK" else re.sub(r'[\\/*?:<>|"\t\n\r]+', "_", str(value).strip())[:50]
                    output_path = helper_get_safe_filepath(tmpdir, safe_name)
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

st.set_page_config(page_title="Công cụ Xử lý Dữ liệu Đất đai", layout="wide", page_icon="🚀")

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://i.imgur.com/v12A61a.png", width=150) # Placeholder image
    st.title("Hướng dẫn")
    st.info(
        "**Công cụ 1:** Tải file Excel gốc lên, chọn sheet và nhấn 'Bắt đầu' "
        "để làm sạch, tô màu và phân loại dữ liệu thành các sheet Nhóm 1, Nhóm 2, v.v."
    )
    st.info(
        "**Công cụ 2:** Tải file đã được xử lý bởi Công cụ 1, "
        "ứng dụng sẽ tự động tách sheet `Nhóm 2_GDC` thành nhiều file con và nén lại."
    )
    st.success("Phát triển bởi: **Trường Sinh**\n\nSĐT: **0917.750.555**")

# --- TRANG CHÍNH ---
st.title("🚀 Công cụ Hỗ trợ Chiến dịch Làm sạch Dữ liệu Đất đai")
st.markdown("---")

# --- TẠO HAI TAB CHO HAI CÔNG CỤ ---
tab1, tab2 = st.tabs([" CÔNG CỤ 1: LÀM SẠCH & PHÂN LOẠI ", " CÔNG CỤ 2: TÁCH FILE THEO THÔN "])

# --- GIAO DIỆN CÔNG CỤ 1 ---
with tab1:
    st.header("Xử lý file tổng, tạo các nhóm dữ liệu")
    uploaded_file_1 = st.file_uploader(
        "1. Tải lên file Excel cần xử lý", 
        type=["xlsx", "xlsm"], 
        key="uploader1"
    )

    if uploaded_file_1:
        try:
            wb_sheets = load_workbook(uploaded_file_1, read_only=True)
            sheet_names = wb_sheets.sheetnames
            wb_sheets.close()
            
            selected_sheet = st.selectbox(
                "2. Chọn sheet chính chứa dữ liệu:", 
                sheet_names,
                help="Đây là sheet gốc chứa dữ liệu bạn muốn lọc."
            )

            if st.button("BẮT ĐẦU XỬ LÝ (LÀM SẠCH)", type="primary"):
                progress_bar_1 = st.progress(0)
                status_text_1 = st.empty()
                
                main_wb = load_workbook(uploaded_file_1)
                
                # Chạy các bước 1, 2, 3
                main_wb = run_step_1_process(main_wb, selected_sheet, progress_bar_1, status_text_1, 0, 33)
                if main_wb:
                    main_wb = run_step_2_clear_fill(main_wb, progress_bar_1, status_text_1, 33, 33)
                if main_wb:
                    main_wb = run_step_3_split_by_color(main_wb, progress_bar_1, status_text_1, 66, 34)

                if main_wb:
                    status_text_1.success("✅ HOÀN TẤT XỬ LÝ! Vui lòng tải file về.")
                    
                    # Tạo buffer để tải về
                    final_buffer = io.BytesIO()
                    main_wb.save(final_buffer)
                    final_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Tải về File đã xử lý",
                        data=final_buffer,
                        file_name=f"[Processed]_{uploaded_file_1.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    status_text_1.error("❌ Xử lý thất bại. Vui lòng kiểm tra lại file đầu vào.")

        except Exception as e:
            st.error(f"Lỗi: Không thể đọc file. File có thể bị hỏng hoặc sai định dạng. Chi tiết: {e}")

# --- GIAO DIỆN CÔNG CỤ 2 ---
with tab2:
    st.header("Tách file từ sheet 'Nhóm 2_GDC' thành nhiều file con")
    uploaded_file_2 = st.file_uploader(
        "1. Tải lên file Excel ĐÃ ĐƯỢC XỬ LÝ bởi Công cụ 1",
        type=["xlsx", "xlsm"],
        key="uploader2",
        help="File này phải chứa sheet 'Nhóm 2_GDC' và 'TongHop'."
    )

    if uploaded_file_2:
        if st.button("BẮT ĐẦU XỬ LÝ (TÁCH FILE)", type="primary"):
            progress_bar_2 = st.progress(0)
            status_text_2 = st.empty()
            
            # Đọc dữ liệu từ file đã tải lên vào bộ nhớ
            data_buffer = io.BytesIO(uploaded_file_2.getvalue())
            
            # Chạy bước 4
            zip_file_buffer = run_step_4_split_files(data_buffer, progress_bar_2, status_text_2, 0, 100)

            if zip_file_buffer:
                status_text_2.success("✅ HOÀN TẤT TÁCH FILE! Vui lòng tải gói ZIP về.")
                st.download_button(
                    label="🗂️ Tải về Gói các file con (.zip)",
                    data=zip_file_buffer,
                    file_name="Cac_file_con_theo_thon.zip",
                    mime="application/zip"
                )
            else:
                status_text_2.error("❌ Tách file thất bại. Hãy chắc chắn file đầu vào có sheet 'Nhóm 2_GDC' và 'TongHop'.")

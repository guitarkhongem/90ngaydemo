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

# --- CẤU HÌNH CÔNG CỤ 2: LÀM SẠCH & TÁCH FILE ---
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
        
def helper_copy_rows_with_style(src_ws, tgt_ws, max_row=3):
    """(Helper) Copy N hàng đầu tiên (giá trị + định dạng + merge + độ rộng cột)"""
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

def helper_group_columns_openpyxl(ws):
    """(Helper) Group các cột bằng openpyxl (An toàn cho online)"""
    try:
        # Xóa group cũ (nếu có)
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

def helper_get_safe_filepath(output_folder, name):
    """(Helper) Tạo tên tệp an toàn, tránh ghi đè"""
    counter = 1
    safe_path = os.path.join(output_folder, f"{name}.xlsx")
    while os.path.exists(safe_path):
        safe_path = os.path.join(output_folder, f"{name}_{counter}.xlsx")
        counter += 1
    return safe_path

# --- CÁC HÀM CHO CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ (MỚI BỔ SUNG) ---

def tool1_excel_col_to_index(col_letter):
    index = 0
    for char in col_letter.upper():
        index = index * 26 + (ord(char) - ord('A')) + 1
    return index - 1

def get_sheet_names_from_buffer(file_buffer):
    try:
        # Đảm bảo buffer có thể được đọc lại
        file_buffer.seek(0)
        wb = load_workbook(file_buffer, read_only=True)
        return wb.sheetnames
    except Exception as e:
        st.error(f"Không thể đọc sheet từ file: {e}")
        return []

def tool1_transform_and_copy(source_buffer, source_sheet, dest_buffer, dest_sheet, progress_bar, status_label):
    from openpyxl.styles import Border, Side
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

# --- CÁC HÀM CHO CÔNG CỤ 2: LÀM SẠCH, PHÂN LOẠI & TÁCH FILE (GIỮ NGUYÊN) ---

def run_step_1_process(wb, sheet_name, master_progress_bar, master_status_label, base_percent, step_budget):
    def update_progress_step1(local_percent, step_text=None):
        if step_text:
            master_status_label.info(f"Bước 1: {step_text} ({local_percent:.0f}%)")
        master_percent = base_percent + (local_percent / 100) * step_budget
        master_progress_bar.progress(int(master_percent))
    
    try:
        if sheet_name not in wb.sheetnames:
            st.error(f"Lỗi Bước 1: Không tìm thấy sheet '{sheet_name}'.")
            return None
        ws = wb[sheet_name]

        last_row = ws.max_row
        while last_row > 1 and ws[f"A{last_row}"].value in (None, ""):
            last_row -= 1
        
        update_progress_step1(0, "Đang tìm hàng trống...")
        rows_to_color = set()
        total_check_rows = last_row - STEP1_START_ROW + 1
        
        for i, row_idx in enumerate(range(STEP1_START_ROW, last_row + 1)):
            for col in STEP1_CHECK_COLS:
                cell_value = ws[f"{col}{row_idx}"].value
                if cell_value is None or str(cell_value).strip() == "":
                    rows_to_color.add(row_idx)
                    break
            if i % 100 == 0:
                update_progress_step1((i / max(total_check_rows, 1)) * 10, "Đang tìm hàng trống...")

        update_progress_step1(10, "Đang xoá màu cũ...")
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=last_row), start=1):
            for cell in row:
                cell.fill = STEP1_EMPTY_FILL
            if i % 50 == 0:
                percent = 10 + (i / last_row) * 20
                update_progress_step1(min(percent, 30), "Đang xoá màu cũ...")
        
        update_progress_step1(30, "Đang tô vàng...")
        for idx, row_idx in enumerate(rows_to_color, start=1):
            for cell in ws[row_idx]:
                cell.fill = STEP1_YELLOW_FILL
            if idx % 50 == 0:
                percent = 30 + (idx / max(len(rows_to_color), 1)) * 10
                update_progress_step1(min(percent, 40), "Đang tô vàng hàng trống...")

        update_progress_step1(40, "Đang xuất Nhóm 1...")
        ws_src = wb[sheet_name]
        last_col = ws_src.max_column

        def copy_rows_step1(title, condition_fn, start_percent, end_percent):
            if title in wb.sheetnames:
                wb.remove(wb[title])
            ws_dst = wb.create_sheet(title)
            for r in range(1, 5):
                for c in range(1, last_col + 1):
                    src = ws_src.cell(row=r, column=c)
                    dst = ws_dst.cell(row=r, column=c)
                    dst.value = src.value
                    if src.has_style:
                        helper_copy_cell_format(src, dst)
            next_row = 5
            total_data_rows = last_row - 4
            for i, r in enumerate(range(5, last_row + 1), start=1):
                if condition_fn(r):
                    for c in range(1, last_col + 1):
                        src = ws_src.cell(row=r, column=c)
                        dst = ws_dst.cell(row=next_row, column=c)
                        dst.value = src.value
                        if src.has_style:
                            helper_copy_cell_format(src, dst)
                    next_row += 1
                if i % 20 == 0:
                    progress = start_percent + (i / max(total_data_rows, 1)) * (end_percent - start_percent)
                    update_progress_step1(min(progress, end_percent), f"Đang xử lý {title}...")
            
            helper_calculate_column_width(ws_dst)

        copy_rows_step1("Nhóm 1", lambda r_idx: r_idx not in rows_to_color, 40, 70)
        copy_rows_step1("Nhóm 2", lambda r_idx: r_idx in rows_to_color, 70, 99)
        
        update_progress_step1(100, "Hoàn tất Bước 1!")
        return wb

    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 1): {e}")
        logging.error(f"Lỗi Bước 1: {e}")
        return None

def run_step_2_clear_fill(wb, master_progress_bar, master_status_label, base_percent, step_budget):
    TARGET_SHEET = "Nhóm 2"
    
    try:
        logging.info(f"Bước 2: Bắt đầu xử lý sheet {TARGET_SHEET}")
        if TARGET_SHEET not in wb.sheetnames:
            st.error(f"Lỗi (Bước 2): Không tìm thấy sheet '{TARGET_SHEET}' để xử lý.")
            return None
        ws = wb[TARGET_SHEET]
        last_row = ws.max_row
        while last_row > 1 and ws[f"A{last_row}"].value in (None, ""):
            last_row -= 1
        rows_changed = 0

        for row_idx in range(STEP2_START_ROW, last_row + 1):
            cell_g = ws[f"{STEP2_TARGET_COL}{row_idx}"]
            is_blank = (cell_g.value is None or str(cell_g.value).strip() == "")
            if not is_blank:
                for cell_in_row in ws[row_idx]:
                    cell_in_row.fill = STEP2_EMPTY_FILL
                rows_changed += 1
            if row_idx % 50 == 0:
                local_percent = (row_idx / max(last_row, 1)) * 100
                master_status_label.info(f"Bước 2: Đang xoá màu cột G... ({local_percent:.0f}%)")
                master_percent = base_percent + (local_percent / 100) * step_budget
                master_progress_bar.progress(int(master_percent))

        master_progress_bar.progress(int(base_percent + step_budget))
        logging.info(f"Bước 2: Hoàn tất, đã xoá màu {rows_changed} hàng.")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 2): {e}")
        logging.error(f"Lỗi Bước 2: {e}")
        return None

def run_step_3_split_by_color(wb, master_progress_bar, master_status_label, base_percent, step_budget):
    TARGET_SHEET = "Nhóm 2"
    
    try:
        logging.info(f"Bước 3: Bắt đầu xử lý sheet {TARGET_SHEET}")
        if TARGET_SHEET not in wb.sheetnames:
            st.error(f"Lỗi (Bước 3): Không tìm thấy sheet '{TARGET_SHEET}' để xử lý.")
            return None
        ws_src = wb[TARGET_SHEET]
        last_row = ws_src.max_row
        last_col = ws_src.max_column

        def copy_rows_step3(condition_fn, title):
            if title in wb.sheetnames:
                wb.remove(wb[title])
            ws_dst = wb.create_sheet(title)
            for row in range(1, 5):
                for col in range(1, last_col + 1):
                    cell_src = ws_src.cell(row=row, column=col)
                    cell_dst = ws_dst.cell(row=row, column=col)
                    cell_dst.value = cell_src.value
                    if cell_src.has_style:
                        helper_copy_cell_format(cell_src, cell_dst)
            next_row = 5
            for row in range(5, last_row + 1):
                cell = ws_src.cell(row=row, column=1)
                if condition_fn(cell):
                    for col in range(1, last_col + 1):
                        cell_src = ws_src.cell(row=row, column=col)
                        cell_dst = ws_dst.cell(row=next_row, column=col)
                        cell_dst.value = cell_src.value
                        if cell_src.has_style:
                            helper_copy_cell_format(cell_src, cell_dst)
                    next_row += 1
            
            helper_calculate_column_width(ws_dst)

        total_steps = 2 * (last_row - 4)
        processed = 0

        def update_progress_step3(add, message):
            nonlocal processed
            processed += add
            local_percent = (processed / max(total_steps, 1)) * 100
            master_status_label.info(f"Bước 3: {message} ({local_percent:.0f}%)")
            master_percent = base_percent + (local_percent / 100) * step_budget
            master_progress_bar.progress(int(master_percent))

        copy_rows_step3(lambda c: not helper_cell_has_bg(c), "Nhóm 2_TC")
        update_progress_step3(last_row - 4, "Đang xuất Nhóm 2_TC (không màu)...")

        copy_rows_step3(lambda c: helper_cell_has_bg(c), "Nhóm 2_GDC")
        update_progress_step3(last_row - 4, "Đang xuất Nhóm 2_GDC (có màu)...")

        master_progress_bar.progress(int(base_percent + step_budget))
        logging.info("Bước 3: Hoàn tất, đã tạo 'Nhóm 2_TC' và 'Nhóm 2_GDC'.")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 3): {e}")
        logging.error(f"Lỗi Bước 3: {e}")
        return None

def run_step_4_split_files(
    step4_data_buffer, 
    main_processed_buffer, 
    main_processed_filename, 
    master_progress_bar, 
    master_status_label, 
    base_percent, 
    step_budget
):
    wb_openpyxl = None
    
    DATA_SHEET = "Nhóm 2_GDC"
    TEMPLATE_SHEET = "TongHop"
    FILTER_COLUMN = "T"
    START_ROW = 5
    
    try:
        logging.info("Bước 4 (Online): Bắt đầu xử lý tách file")
        
        try:
            wb_openpyxl = load_workbook(step4_data_buffer, data_only=True)
            if TEMPLATE_SHEET not in wb_openpyxl.sheetnames:
                st.error("Lỗi (Bước 4): Không tìm thấy sheet mẫu 'TongHop'!")
                return None
            if DATA_SHEET not in wb_openpyxl.sheetnames:
                st.error("Lỗi (Bước 4): Không tìm thấy sheet dữ liệu 'Nhóm 2_GDC'!")
                return None
            tonghop_ws = wb_openpyxl["TongHop"]
            
            step4_data_buffer.seek(0)
            df = pd.read_excel(step4_data_buffer, sheet_name=DATA_SHEET, header=None)
            logging.info("Đã tải thành công template và data từ buffer")
        except Exception as e:
            st.error(f"Lỗi (Bước 4): Không thể đọc buffer: {e}")
            logging.error(f"Bước 4: Lỗi đọc buffer: {e}")
            return None

        col_index = column_index_from_string(FILTER_COLUMN) - 1
        start_row_index = START_ROW - 1
        if col_index >= len(df.columns):
            st.error(f"Lỗi (Bước 4): Cột lọc '{FILTER_COLUMN}' không tồn tại!")
            return None
        
        data_col_raw = df.iloc[start_row_index:, col_index]
        data_col = data_col_raw.apply(helper_normalize_value)
        unique_normalized = data_col.dropna().unique().tolist()
        if data_col.isnull().any():
            unique_normalized.append("BLANK")

        total = len(unique_normalized)
        master_status_label.info(f"Bước 4: Chuẩn bị tách {total} file con...")

        with tempfile.TemporaryDirectory() as tmpdir:
            logging.info(f"Đã tạo thư mục tạm: {tmpdir}")
            
            try:
                main_file_path = os.path.join(tmpdir, main_processed_filename)
                with open(main_file_path, 'wb') as f:
                    f.write(main_processed_buffer.getbuffer())
                logging.info(f"Đã lưu file chính vào: {main_file_path}")
            except Exception as e_save_main:
                logging.warning(f"Không thể lưu file chính vào zip: {e_save_main}")

            for i, norm_value in enumerate(unique_normalized, start=1):
                if norm_value == "BLANK":
                    mask = data_col.isnull()
                else:
                    mask = data_col == norm_value
                filtered = df.iloc[start_row_index:][mask]
                if filtered.empty:
                    continue

                new_wb = Workbook()
                new_ws = new_wb.active
                new_ws.title = "DuLieuLoc"
                helper_copy_rows_with_style(tonghop_ws, new_ws, max_row=3)
                for r_idx, row in enumerate(dataframe_to_rows(filtered, index=False, header=False), start=4):
                    for c_idx, value_cell in enumerate(row, start=1):
                        new_ws.cell(row=r_idx, column=c_idx, value=value_cell)
                
                safe_name = "BLANK" if norm_value == "BLANK" else re.sub(r'[\\/*?:<>|"\t\n\r]+', "_", str(norm_value).strip())[:50]
                output_path = helper_get_safe_filepath(tmpdir, safe_name)
                
                try:
                    helper_group_columns_openpyxl(new_ws)
                    helper_calculate_column_width(new_ws)
                    new_wb.save(output_path)
                except Exception as e_openpyxl:
                    logging.error(f"Lỗi openpyxl khi xử lý {output_path}: {e_openpyxl}")
                finally:
                    if new_wb: new_wb.close() 

                local_percent = (i / total) * 100
                master_status_label.info(f"Bước 4: Đang tách file {i}/{total} ({local_percent:.0f}%)")
                master_percent = base_percent + (local_percent / 100) * step_budget
                master_progress_bar.progress(int(master_percent))
            
            master_status_label.info("Bước 4: Đang nén file ZIP...")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                for root, _, files in os.walk(tmpdir):
                    for file in files:
                        zip_f.write(os.path.join(root, file), arcname=file)
            
            zip_buffer.seek(0)
            master_progress_bar.progress(int(base_percent + step_budget))
            logging.info("Đã tạo ZIP buffer thành công.")
            
            if wb_openpyxl: wb_openpyxl.close()
            
            return zip_buffer

    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 4 - Online): {str(e)}")
        logging.error(f"Lỗi Bước 4 (Online): {e}")
        return None
    finally:
        if wb_openpyxl:
            try: wb_openpyxl.close()
            except: pass

# --- GIAO DIỆN STREAMLIT CHÍNH ---



# --- SIDEBAR ---
with st.sidebar:
    
    st.title("Hướng dẫn sử dụng")
    st.info("**Công cụ 1: Sao chép & Ánh xạ Cột**\n\n- Tải lên file Nguồn và file Đích.\n- Chọn sheet tương ứng.\n- Công cụ sẽ sao chép dữ liệu từ nguồn sang đích theo cấu hình định sẵn.")
    st.info("**Công cụ 2: Làm sạch & Tách file**\n\n- Tải file Excel gốc, chọn sheet.\n- Công cụ sẽ tự động chạy toàn bộ quy trình làm sạch, phân loại và tách file.\n- Kết quả trả về gồm file tổng đã xử lý và gói ZIP các file con.")
    st.success("Phát triển bởi: **Trường Sinh**\n\nSĐT: **0917.750.555**")

# --- MAIN PAGE ---
col1, col2 = st.columns([1, 10])
with col1:
    # Left column intentionally left mostly empty; keep a placeholder to satisfy Python's syntax
    st.empty()

with col2:
    # Right column intentionally left empty as layout spacer; placeholder avoids IndentationError
    st.empty()

st.markdown("---")

tab1, tab2 = st.tabs([
    " CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ CỘT ", 
    " CÔNG CỤ 2: LÀM SẠCH & TÁCH FILE THEO THÔN "
])

# --- GIAO DIỆN CÔNG CỤ 1 (MỚI BỔ SUNG) ---
with tab1:
    st.header("Chuyển đổi và sao chép dữ liệu giữa hai file Excel")
    
    col1, col2 = st.columns(2)
    with col1:
        source_file = st.file_uploader("1. Tải lên File Nguồn (lấy dữ liệu)", type=["xlsx", "xls"], key="tool1_source")
        if source_file:
            source_sheets = get_sheet_names_from_buffer(source_file)
            selected_source_sheet = st.selectbox("2. Chọn Sheet Nguồn:", source_sheets, key="tool1_source_sheet")
        st.caption("Ví dụ về file nguồn (dữ liệu thô):")
        
    
    with col2:
        dest_file = st.file_uploader("3. Tải lên File Đích (nhận dữ liệu)", type=["xlsx", "xls"], key="tool1_dest")
        if dest_file:
            dest_sheets = get_sheet_names_from_buffer(dest_file)
            selected_dest_sheet = st.selectbox("4. Chọn Sheet Đích:", dest_sheets, key="tool1_dest_sheet")
        st.caption("Ví dụ về file đích (biểu mẫu có sẵn định dạng):")
        

    st.markdown("---")
    
    if st.button("BẮT ĐẦU SAO CHÉP DỮ LIỆU", type="primary", key="tool1_start"):
        if source_file and dest_file and 'selected_source_sheet' in locals() and 'selected_dest_sheet' in locals():
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

# --- GIAO DIỆN CÔNG CỤ 2 (GỘP) ---
with tab2:
    st.header("Quy trình làm sạch, phân loại và tách file")
    uploaded_file_2 = st.file_uploader("1. Tải lên file Excel gốc cần xử lý", type=["xlsx", "xlsm"], key="tool2_uploader")

    if uploaded_file_2:
        try:
            sheets_2 = get_sheet_names_from_buffer(uploaded_file_2)
            selected_sheet_2 = st.selectbox("2. Chọn sheet chính chứa dữ liệu:", sheets_2, key="tool2_sheet")

            if st.button("BẮT ĐẦU QUY TRÌNH XỬ LÝ & TÁCH FILE", type="primary", key="tool2_start"):
                progress_bar_2 = st.progress(0, text="Bắt đầu...")
                status_text_2 = st.empty()
                
                # BƯỚC 1, 2, 3: LÀM SẠCH VÀ PHÂN LOẠI
                status_text_2.info("Đang tải file vào bộ nhớ...")
                main_wb_2 = load_workbook(uploaded_file_2)
                
                main_wb_2 = run_step_1_process(main_wb_2, selected_sheet_2, progress_bar_2, status_text_2, 0, 25)
                if main_wb_2:
                    main_wb_2 = run_step_2_clear_fill(main_wb_2, progress_bar_2, status_text_2, 25, 10)
                if main_wb_2:
                    main_wb_2 = run_step_3_split_by_color(main_wb_2, progress_bar_2, status_text_2, 35, 15)

                # BƯỚC 4: TÁCH FILE
                zip_buffer = None
                if main_wb_2:
                    processed_buffer = io.BytesIO()
                    main_wb_2.save(processed_buffer)
                    
                    # Cần 2 buffer cho bước 4
                    step4_read_buffer = io.BytesIO(processed_buffer.getvalue())
                    main_processed_buffer = io.BytesIO(processed_buffer.getvalue())
                    
                    main_processed_filename = f"[Processed]_{uploaded_file_2.name}"

                    zip_buffer = run_step_4_split_files(
                        step4_read_buffer,
                        main_processed_buffer,
                        main_processed_filename,
                        progress_bar_2, 
                        status_text_2, 
                        50, 
                        50
                    )

                # HIỂN THỊ KẾT QUẢ
                if main_wb_2 and zip_buffer:
                    status_text_2.success("✅ HOÀN TẤT TOÀN BỘ QUY TRÌNH!")
                    
                    processed_buffer.seek(0)
                    
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        st.download_button(
                            label="📥 Tải về File Tổng đã xử lý", 
                            data=processed_buffer, 
                            file_name=f"[Processed]_{uploaded_file_2.name}", 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with dl_col2:
                         st.download_button(
                             label="🗂️ Tải về Gói file con (.zip)", 
                             data=zip_buffer, 
                             file_name="Cac_file_con_theo_thon.zip", 
                             mime="application/zip"
                        )
                else:
                    status_text_2.error("❌ Quy trình thất bại. Vui lòng kiểm tra lại file đầu vào hoặc định dạng file.")

        except Exception as e:
            st.error(f"Lỗi không xác định: {e}")


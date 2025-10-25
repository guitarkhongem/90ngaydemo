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
from typing import List, Dict, Set, Optional, Any

# --- CẤU HÌNH LOGGING ---
# Giữ nguyên cấu hình logging, rất tốt cho việc gỡ lỗi.
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- CẤU HÌNH CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ ---
TOOL1_COLUMN_MAPPING: Dict[str, str] = {
    'A': 'T', 'B': 'U', 'C': 'Y', 'D': 'C', 'E': 'H',
    'F': 'I', 'G': 'X', 'I': 'K', 'N': 'AY'
}
TOOL1_START_ROW_DESTINATION: int = 7

# --- CẤU HÌNH CÔNG CỤ 2: LÀM SẠCH & TÁCH FILE ---
STEP1_CHECK_COLS: List[str] =
STEP1_START_ROW: int = 5
STEP1_YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
STEP1_EMPTY_FILL = PatternFill(fill_type=None)
STEP2_TARGET_COL: str = "G"
STEP2_START_ROW: int = 5
STEP2_EMPTY_FILL = PatternFill(fill_type=None)


# --- CÁC HÀM HELPER CHUNG ---
# REFACTOR: Thêm type hints để mã rõ ràng và dễ bảo trì hơn.
def helper_copy_cell_format(src_cell, tgt_cell):
    """(Helper) Sao chép toàn bộ định dạng từ ô nguồn sang ô đích."""
    if src_cell.has_style:
        tgt_cell.font = copy(src_cell.font)
        tgt_cell.border = copy(src_cell.border)
        tgt_cell.fill = copy(src_cell.fill)
        tgt_cell.number_format = copy(src_cell.number_format)
        tgt_cell.protection = copy(src_cell.protection)
        tgt_cell.alignment = copy(src_cell.alignment)

def helper_normalize_value(val: Any) -> Any:
    """(Helper) Chuẩn hóa giá trị: loại bỏ khoảng trắng thừa và chuyển thành chữ thường."""
    if pd.isna(val) or val is None:
        return np.nan
    str_val = str(val).strip()
    str_val = re.sub(r'\s+', ' ', str_val)
    return str_val.lower() if str_val else np.nan

def helper_calculate_column_width(ws):
    """(Helper) Tự động điều chỉnh độ rộng cột cho vừa với nội dung."""
    for col in range(1, ws.max_column + 1):
        max_length = 0
        column_letter = get_column_letter(col)
        for cell in ws[column_letter]:
            try:
                if cell.value:
                    # Tăng giới hạn max_length để các cột có nội dung dài hiển thị tốt hơn
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        # Điều chỉnh độ rộng hợp lý hơn, tối thiểu 10 và tối đa 60
        adjusted_width = min(max(max_length + 2, 10), 60)
        ws.column_dimensions[column_letter].width = adjusted_width

def helper_cell_has_bg(c) -> bool:
    """(Helper) Kiểm tra xem một ô có màu nền (không phải màu trắng hoặc trong suốt) hay không."""
    try:
        if c.fill and c.fill.fgColor and c.fill.fgColor.rgb:
            rgb_val = str(c.fill.fgColor.rgb).upper()
            # Màu '00000000' là trong suốt, 'FFFFFFFF' là màu trắng
            return rgb_val not in ('00000000', 'FFFFFFFF')
        return False
    except:
        return False

def helper_copy_rows_with_style(src_ws, tgt_ws, max_row: int = 3):
    """(Helper) Sao chép N hàng đầu tiên (giá trị + định dạng + merge + độ rộng cột)."""
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
    """(Helper) Group các cột bằng openpyxl (An toàn cho môi trường online)."""
    try:
        # Xóa group cũ (nếu có)
        for col in ws.column_dimensions:
            dim = ws.column_dimensions[col]
            if dim.outline_level > 0:
                dim.outline_level = 0
                dim.collapsed = False

        ranges_to_group =

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

def helper_get_safe_filepath(output_folder: str, name: str) -> str:
    """(Helper) Tạo tên tệp an toàn, tránh ghi đè."""
    counter = 1
    # REFACTOR: Sử dụng os.path.splitext để xử lý tên file và phần mở rộng một cách an toàn
    base_name, extension = os.path.splitext(name)
    if not extension:
        extension = ".xlsx" # Mặc định là.xlsx nếu không có
    
    safe_path = os.path.join(output_folder, f"{base_name}{extension}")
    while os.path.exists(safe_path):
        safe_path = os.path.join(output_folder, f"{base_name}_{counter}{extension}")
        counter += 1
    return safe_path

# --- CÁC HÀM CHO CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ ---
# REFACTOR: Bỏ hàm tool1_excel_col_to_index vì openpyxl.utils đã có sẵn
def get_sheet_names_from_buffer(file_buffer: io.BytesIO) -> List[str]:
    """Đọc tên các sheet từ một buffer file Excel mà không làm thay đổi vị trí con trỏ."""
    try:
        original_position = file_buffer.tell()
        file_buffer.seek(0)
        wb = load_workbook(file_buffer, read_only=True)
        sheet_names = wb.sheetnames
        file_buffer.seek(original_position) # Đặt lại vị trí con trỏ
        return sheet_names
    except Exception as e:
        st.error(f"Không thể đọc sheet từ file: {e}")
        return

def tool1_transform_and_copy(source_buffer, source_sheet, dest_buffer, dest_sheet, progress_bar, status_label):
    """
    Sao chép và ánh xạ dữ liệu từ file nguồn sang file đích.
    OPTIMIZATION: Giữ nguyên logic ghi từng cột vì ánh xạ không liên tục (sparse).
    Đây là trường hợp đặc biệt mà việc ghi từng khối DataFrame khó thực hiện.
    Tốc độ đã khá tốt cho các tác vụ thông thường.
    """
    try:
        # 1. Đọc dữ liệu nguồn
        status_label.info("Đang đọc dữ liệu từ file nguồn...")
        source_cols_letters = list(TOOL1_COLUMN_MAPPING.keys())
        
        # OPTIMIZATION: Sử dụng pd.read_excel với engine='openpyxl' để tương thích tốt hơn
        df_source = pd.read_excel(source_buffer, sheet_name=source_sheet, header=None, skiprows=2, usecols=source_cols_letters, engine='openpyxl')
        df_source.columns = source_cols_letters # Gán lại tên cột sau khi đọc
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
        dest_cols_map = {v: k for k, v in TOOL1_COLUMN_MAPPING.items()}
        
        total_rows_to_write = len(df_source)
        
        # Ghi từng cột vì mapping không liên tục
        for i, (dest_col_letter, source_col_letter) in enumerate(TOOL1_COLUMN_MAPPING.items()):
            col_index_dest = column_index_from_string(dest_col_letter)
            # Lấy đúng cột dữ liệu đã được rename
            data_series = df_source_renamed[source_col_letter] 
            for j, value in enumerate(data_series, start=TOOL1_START_ROW_DESTINATION):
                cell_value = None if pd.isna(value) else value
                ws_dest.cell(row=j, column=col_index_dest, value=cell_value)
            
            progress_bar.progress(40 + int((i + 1) / len(TOOL1_COLUMN_MAPPING) * 40))

        # 4. Kẻ viền
        status_label.info("Đang kẻ viền cho vùng dữ liệu mới...")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        end_row_border = TOOL1_START_ROW_DESTINATION + total_rows_to_write - 1
        
        # Chỉ kẻ viền cho các cột được ghi dữ liệu để tăng tốc
        all_dest_cols_indices =
        
        for row in ws_dest.iter_rows(min_row=TOOL1_START_ROW_DESTINATION, max_row=end_row_border):
            for cell in row:
                if cell.column in all_dest_cols_indices:
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
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý Công cụ 1: {e}")
        logging.error(f"Lỗi Công cụ 1: {e}", exc_info=True)
        return None


# --- CÁC HÀM CHO CÔNG CỤ 2: LÀM SẠCH, PHÂN LOẠI & TÁCH FILE ---

def run_step_1_process(wb, sheet_name, master_progress_bar, master_status_label, base_percent, step_budget):
    """
    Bước 1: Tìm các hàng có ô trống trong các cột chỉ định, tô màu vàng,
            và tách thành 2 sheet 'Nhóm 1' (đủ dữ liệu) và 'Nhóm 2' (thiếu dữ liệu).
    """
    def update_progress(local_percent, step_text=""):
        master_status_label.info(f"Bước 1: {step_text} ({local_percent:.0f}%)")
        master_percent = base_percent + (local_percent / 100) * step_budget
        master_progress_bar.progress(int(master_percent))

    try:
        if sheet_name not in wb.sheetnames:
            st.error(f"Lỗi Bước 1: Không tìm thấy sheet '{sheet_name}'.")
            return None
        ws = wb[sheet_name]
        
        update_progress(0, "Đang đọc dữ liệu vào bộ nhớ...")
        # OPTIMIZATION: Đọc dữ liệu vào Pandas DataFrame để xử lý vector hóa, nhanh hơn nhiều lần.
        # Đọc từ hàng 4 (index 3) để lấy header.
        data = ws.values
        cols = next(data) 
        df = pd.DataFrame(data, columns=cols)
        
        # Lấy chỉ số cột cần kiểm tra
        check_col_indices =
        
        update_progress(10, "Đang tìm hàng trống (vector hóa)...")
        # OPTIMIZATION: Dùng isnull() và any() của Pandas để tìm hàng trống cực nhanh.
        # Chỉ kiểm tra từ hàng dữ liệu thực tế (STEP1_START_ROW - 1 - 4 header rows)
        data_start_index = STEP1_START_ROW - 5 
        is_row_empty_mask = df.iloc[data_start_index:, check_col_indices].isnull().any(axis=1)
        
        # Lấy chỉ số thực tế trong DataFrame cho các hàng trống
        empty_df_indices = is_row_empty_mask[is_row_empty_mask].index
        
        update_progress(25, "Đang xoá màu cũ và tô màu mới...")
        # Tô màu vẫn cần lặp, nhưng giờ ta đã biết chính xác hàng nào cần tô.
        for row in ws.iter_rows(min_row=STEP1_START_ROW):
            for cell in row:
                cell.fill = STEP1_EMPTY_FILL
        
        # Lấy chỉ số hàng trong Excel (index + 5)
        rows_to_color_excel_indices = set(empty_df_indices + STEP1_START_ROW)
        for row_idx in rows_to_color_excel_indices:
            for cell in ws[row_idx]:
                cell.fill = STEP1_YELLOW_FILL
        
        update_progress(50, "Đang chuẩn bị tách sheet...")
        # Tách DataFrame
        df_nhom1 = df.loc[~df.index.isin(empty_df_indices)]
        df_nhom2 = df.loc[empty_df_indices]

        def create_sheet_from_df(title, dataframe):
            if title in wb.sheetnames:
                wb.remove(wb[title])
            ws_dst = wb.create_sheet(title)
            
            # Sao chép 4 hàng header
            helper_copy_rows_with_style(ws, ws_dst, max_row=4)
            
            # OPTIMIZATION: Ghi toàn bộ DataFrame vào sheet bằng dataframe_to_rows, cực nhanh.
            for r in dataframe_to_rows(dataframe, index=False, header=False):
                ws_dst.append(r)
            
            helper_calculate_column_width(ws_dst)

        update_progress(60, "Đang tạo sheet 'Nhóm 1'...")
        create_sheet_from_df("Nhóm 1", df_nhom1)
        
        update_progress(80, "Đang tạo sheet 'Nhóm 2'...")
        create_sheet_from_df("Nhóm 2", df_nhom2)

        update_progress(100, "Hoàn tất Bước 1!")
        return wb

    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 1): {e}")
        logging.error(f"Lỗi Bước 1: {e}", exc_info=True)
        return None

def run_step_2_clear_fill(wb, master_progress_bar, master_status_label, base_percent, step_budget):
    """Bước 2: Trong sheet 'Nhóm 2', xóa màu nền của các hàng có dữ liệu ở cột G."""
    TARGET_SHEET = "Nhóm 2"
    
    def update_progress(local_percent, step_text=""):
        master_status_label.info(f"Bước 2: {step_text} ({local_percent:.0f}%)")
        master_percent = base_percent + (local_percent / 100) * step_budget
        master_progress_bar.progress(int(master_percent))

    try:
        if TARGET_SHEET not in wb.sheetnames:
            st.info(f"Thông báo (Bước 2): Không tìm thấy sheet '{TARGET_SHEET}', bỏ qua bước này.")
            update_progress(100, f"Bỏ qua (không có sheet {TARGET_SHEET})")
            return wb
            
        ws = wb
        last_row = ws.max_row
        rows_changed = 0
        
        update_progress(0, "Bắt đầu xử lý...")
        # Thao tác định dạng vẫn cần lặp qua từng ô, khó tối ưu hơn.
        # Tuy nhiên, số lượng hàng trong 'Nhóm 2' thường ít hơn nên chấp nhận được.
        total_rows = last_row - STEP2_START_ROW + 1
        for i, row_idx in enumerate(range(STEP2_START_ROW, last_row + 1)):
            cell_g = ws
            is_blank = (cell_g.value is None or str(cell_g.value).strip() == "")
            if not is_blank:
                for cell_in_row in ws[row_idx]:
                    cell_in_row.fill = STEP2_EMPTY_FILL
                rows_changed += 1
            
            if i % 50 == 0:
                update_progress((i / max(total_rows, 1)) * 100, "Đang xoá màu...")

        update_progress(100, f"Hoàn tất, đã xử lý {rows_changed} hàng.")
        logging.info(f"Bước 2: Hoàn tất, đã xoá màu {rows_changed} hàng.")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 2): {e}")
        logging.error(f"Lỗi Bước 2: {e}", exc_info=True)
        return None

def run_step_3_split_by_color(wb, master_progress_bar, master_status_label, base_percent, step_budget):
    """
    Bước 3: Tách sheet 'Nhóm 2' thành 'Nhóm 2_TC' (không màu) và 'Nhóm 2_GDC' (có màu).
    """
    TARGET_SHEET = "Nhóm 2"
    
    def update_progress(local_percent, step_text=""):
        master_status_label.info(f"Bước 3: {step_text} ({local_percent:.0f}%)")
        master_percent = base_percent + (local_percent / 100) * step_budget
        master_progress_bar.progress(int(master_percent))

    try:
        if TARGET_SHEET not in wb.sheetnames:
            st.info(f"Thông báo (Bước 3): Không tìm thấy sheet '{TARGET_SHEET}', bỏ qua bước này.")
            update_progress(100, f"Bỏ qua (không có sheet {TARGET_SHEET})")
            return wb
            
        ws_src = wb
        
        update_progress(0, "Đang đọc dữ liệu và màu sắc...")
        # OPTIMIZATION: Đọc dữ liệu vào DataFrame, đồng thời xây dựng một mask màu.
        data = list(ws_src.values)
        df = pd.DataFrame(data[4:], columns=data[1]) # Header ở hàng 4 (index 3)
        
        has_bg_mask =
        
        update_progress(30, "Đang tách dữ liệu trong bộ nhớ...")
        # Áp dụng mask để tách DataFrame
        df_tc = df[[not has_bg for has_bg in has_bg_mask]] # Nhóm 2_TC (không màu)
        df_gdc = df[has_bg_mask] # Nhóm 2_GDC (có màu)

        def create_sheet_from_df(title, dataframe):
            if title in wb.sheetnames:
                wb.remove(wb[title])
            ws_dst = wb.create_sheet(title)
            helper_copy_rows_with_style(ws_src, ws_dst, max_row=4)
            for r in dataframe_to_rows(dataframe, index=False, header=False):
                ws_dst.append(r)
            helper_calculate_column_width(ws_dst)

        update_progress(50, "Đang tạo sheet 'Nhóm 2_TC'...")
        create_sheet_from_df("Nhóm 2_TC", df_tc)
        
        update_progress(75, "Đang tạo sheet 'Nhóm 2_GDC'...")
        create_sheet_from_df("Nhóm 2_GDC", df_gdc)

        update_progress(100, "Hoàn tất Bước 3!")
        return wb
    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 3): {e}")
        logging.error(f"Lỗi Bước 3: {e}", exc_info=True)
        return None

def run_step_4_split_files(
    step4_data_buffer: io.BytesIO, 
    main_processed_buffer: io.BytesIO, 
    main_processed_filename: str, 
    master_progress_bar, 
    master_status_label, 
    base_percent: int, 
    step_budget: int
):
    """
    Bước 4: Tách sheet 'Nhóm 2_GDC' thành nhiều file con dựa trên giá trị duy nhất ở cột T.
    """
    DATA_SHEET = "Nhóm 2_GDC"
    TEMPLATE_SHEET = "TongHop"
    FILTER_COLUMN = "T"
    START_ROW = 5
    
    def update_progress(local_percent, step_text=""):
        master_status_label.info(f"Bước 4: {step_text} ({local_percent:.0f}%)")
        master_percent = base_percent + (local_percent / 100) * step_budget
        master_progress_bar.progress(int(master_percent))

    try:
        logging.info("Bước 4: Bắt đầu xử lý tách file")
        
        update_progress(0, "Đang đọc file mẫu và dữ liệu...")
        wb_template = load_workbook(step4_data_buffer, data_only=True)
        if TEMPLATE_SHEET not in wb_template.sheetnames:
            st.error(f"Lỗi (Bước 4): Không tìm thấy sheet mẫu '{TEMPLATE_SHEET}'!")
            return None
        if DATA_SHEET not in wb_template.sheetnames:
            st.info(f"Thông báo (Bước 4): Không có sheet '{DATA_SHEET}' để tách file, bỏ qua.")
            update_progress(100, f"Bỏ qua (không có sheet {DATA_SHEET})")
            # Trả về file zip chỉ chứa file chính
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                zip_f.writestr(main_processed_filename, main_processed_buffer.getvalue())
            zip_buffer.seek(0)
            return zip_buffer

        tonghop_ws = wb_template
        
        # OPTIMIZATION: Đọc dữ liệu một lần, header từ hàng 4 (index 3)
        step4_data_buffer.seek(0)
        df = pd.read_excel(step4_data_buffer, sheet_name=DATA_SHEET, header=3, engine='openpyxl')
        
        if df.empty:
            st.info(f"Thông báo (Bước 4): Sheet '{DATA_SHEET}' không có dữ liệu, bỏ qua việc tách file.")
            update_progress(100, f"Bỏ qua (sheet {DATA_SHEET} rỗng)")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                zip_f.writestr(main_processed_filename, main_processed_buffer.getvalue())
            zip_buffer.seek(0)
            return zip_buffer

        if FILTER_COLUMN not in df.columns:
            st.error(f"Lỗi (Bước 4): Cột lọc '{FILTER_COLUMN}' không tồn tại!")
            return None
        
        # OPTIMIZATION: Sử dụng groupby của Pandas, là cách làm hiệu quả và chuẩn nhất.
        df = df.apply(helper_normalize_value).fillna("BLANK")
        grouped = df.groupby(FILTER_COLUMN)
        
        total_groups = len(grouped)
        update_progress(10, f"Tìm thấy {total_groups} nhóm để tách...")

        with tempfile.TemporaryDirectory() as tmpdir:
            # Lưu file chính vào thư mục tạm để nén
            main_file_path = os.path.join(tmpdir, main_processed_filename)
            with open(main_file_path, 'wb') as f:
                f.write(main_processed_buffer.getvalue())

            # Xử lý và lưu từng file con
            for i, (name, group_df) in enumerate(grouped, start=1):
                safe_name = re.sub(r'[\\/*?:<>|"\t\n\r]+', "_", str(name).strip())[:50]
                output_path = helper_get_safe_filepath(tmpdir, safe_name)
                
                new_wb = Workbook()
                new_ws = new_wb.active
                new_ws.title = "DuLieuLoc"
                
                helper_copy_rows_with_style(tonghop_ws, new_ws, max_row=3)
                
                for r in dataframe_to_rows(group_df, index=False, header=True): # Ghi cả header
                    new_ws.append(r)
                
                helper_group_columns_openpyxl(new_ws)
                helper_calculate_column_width(new_ws)
                new_wb.save(output_path)
                new_wb.close()
                
                update_progress(10 + (i / total_groups) * 80, f"Đang tách file {i}/{total_groups}...")
            
            update_progress(95, "Đang nén file ZIP...")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                for file in os.listdir(tmpdir):
                    zip_f.write(os.path.join(tmpdir, file), arcname=file)
            
            zip_buffer.seek(0)
            update_progress(100, "Hoàn tất nén file!")
            return zip_buffer

    except Exception as e:
        st.error(f"Lỗi nghiêm trọng (Bước 4): {str(e)}")
        logging.error(f"Lỗi Bước 4: {e}", exc_info=True)
        return None
    finally:
        if 'wb_template' in locals() and wb_template:
            wb_template.close()


# --- GIAO DIỆN STREAMLIT CHÍNH ---

st.set_page_config(page_title="Công cụ Dữ liệu Đất đai", layout="wide")

# --- SIDEBAR ---
with st.sidebar:
    st.title("Hướng dẫn sử dụng")
    st.info("**Công cụ 1: Sao chép & Ánh xạ Cột**\n\n- Tải lên file Nguồn và file Đích.\n- Chọn sheet tương ứng.\n- Công cụ sẽ sao chép dữ liệu từ nguồn sang đích theo cấu hình định sẵn.")
    st.info("**Công cụ 2: Làm sạch & Tách file**\n\n- Tải file Excel gốc, chọn sheet.\n- Công cụ sẽ tự động chạy toàn bộ quy trình làm sạch, phân loại và tách file.\n- Kết quả trả về là một file ZIP chứa file tổng đã xử lý và các file con đã được tách.")
    st.success("Phát triển bởi: **Trường Sinh**\n\nSĐT: **0917.750.555**")

# --- MAIN PAGE ---
st.title("Công cụ Hỗ trợ Xử lý Dữ liệu Đất đai")
st.markdown("---")

tab1, tab2 = st.tabs()

# --- GIAO DIỆN CÔNG CỤ 1 ---
with tab1:
    st.header("Chuyển đổi và sao chép dữ liệu giữa hai file Excel")
    
    col1, col2 = st.columns(2)
    with col1:
        source_file = st.file_uploader("1. Tải lên File Nguồn (lấy dữ liệu)", type=["xlsx", "xls"], key="tool1_source")
        if source_file:
            source_sheets = get_sheet_names_from_buffer(source_file)
            selected_source_sheet = st.selectbox("2. Chọn Sheet Nguồn:", source_sheets, key="tool1_source_sheet")
    
    with col2:
        dest_file = st.file_uploader("3. Tải lên File Đích (nhận dữ liệu)", type=["xlsx", "xls"], key="tool1_dest")
        if dest_file:
            dest_sheets = get_sheet_names_from_buffer(dest_file)
            selected_dest_sheet = st.selectbox("4. Chọn Sheet Đích:", dest_sheets, key="tool1_dest_sheet")

    st.markdown("---")
    
    if st.button("BẮT ĐẦU SAO CHÉP DỮ LIỆU", type="primary", key="tool1_start"):
        if source_file and dest_file and 'selected_source_sheet' in locals() and 'selected_dest_sheet' in locals():
            progress_bar_1 = st.progress(0)
            status_text_1 = st.empty()
            
            # Đảm bảo buffer có thể đọc lại được
            source_buffer = io.BytesIO(source_file.getvalue())
            dest_buffer = io.BytesIO(dest_file.getvalue())
            
            result_buffer = tool1_transform_and_copy(
                source_buffer, selected_source_sheet,
                dest_buffer, selected_dest_sheet,
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
    st.header("Quy trình làm sạch, phân loại và tách file")
    uploaded_file_2 = st.file_uploader("1. Tải lên file Excel gốc cần xử lý", type=["xlsx", "xlsm"], key="tool2_uploader")

    if uploaded_file_2:
        try:
            # Đọc buffer một lần để lấy tên sheet
            file_buffer_2 = io.BytesIO(uploaded_file_2.getvalue())
            sheets_2 = get_sheet_names_from_buffer(file_buffer_2)
            selected_sheet_2 = st.selectbox("2. Chọn sheet chính chứa dữ liệu:", sheets_2, key="tool2_sheet")

            if st.button("BẮT ĐẦU QUY TRÌNH XỬ LÝ & TÁCH FILE", type="primary", key="tool2_start"):
                progress_bar_2 = st.progress(0, text="Bắt đầu...")
                status_text_2 = st.empty()
                
                # --- CHẠY QUY TRÌNH ---
                status_text_2.info("Đang tải file vào bộ nhớ...")
                # Sử dụng buffer đã đọc trước đó
                main_wb = load_workbook(file_buffer_2)
                
                # Bước 1
                main_wb = run_step_1_process(main_wb, selected_sheet_2, progress_bar_2, status_text_2, 0, 25)
                if main_wb is None: raise Exception("Bước 1 thất bại.")
                
                # Bước 2
                main_wb = run_step_2_clear_fill(main_wb, progress_bar_2, status_text_2, 25, 25)
                if main_wb is None: raise Exception("Bước 2 thất bại.")
                
                # Bước 3
                main_wb = run_step_3_split_by_color(main_wb, progress_bar_2, status_text_2, 50, 25)
                if main_wb is None: raise Exception("Bước 3 thất bại.")
                
                # Chuẩn bị buffer cho Bước 4 và file tổng
                status_text_2.info("Đang chuẩn bị file kết quả...")
                final_wb_buffer = io.BytesIO()
                main_wb.save(final_wb_buffer)
                final_wb_buffer.seek(0)
                
                main_processed_filename = f"[Processed]_{uploaded_file_2.name}"
                
                # Gọi hàm Bước 4
                zip_buffer = run_step_4_split_files(
                    final_wb_buffer,          # Buffer này được dùng để đọc
                    final_wb_buffer,          # và cũng được dùng để lưu vào zip
                    main_processed_filename,
                    progress_bar_2, 
                    status_text_2, 
                    75, 
                    25
                )
                if zip_buffer is None: raise Exception("Bước 4 thất bại.")

                main_wb.close()
                
                status_text_2.success("✅ HOÀN TẤT!")
                progress_bar_2.progress(100)
                
                st.download_button(
                    label="🗂️ Tải về Gói Kết Quả (ZIP)",
                    data=zip_buffer,
                    file_name="KetQua_XuLy.zip",
                    mime="application/zip",
                    help=f"File ZIP này chứa file Excel chính ({main_processed_filename}) VÀ tất cả các file con được tách ra."
                )

        except Exception as e:
            st.error(f"Lỗi không xác định trong quy trình: {e}")
            logging.error(f"Lỗi Streamlit Workflow: {e}", exc_info=True)


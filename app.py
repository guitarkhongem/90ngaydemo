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
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# --- CẤU HÌNH CÔNG CỤ 1: SAO CHÉP & ÁNH XẠ ---
TOOL1_COLUMN_MAPPING: Dict[str, str] = {
    'A': 'T', 'B': 'U', 'C': 'Y', 'D': 'C', 'E': 'H',
    'F': 'I', 'G': 'X', 'I': 'K', 'N': 'AY'
}
TOOL1_START_ROW_DESTINATION: int = 7
# Đường dẫn file mẫu cố định
TOOL1_TEMPLATE_FILE_PATH: str = "templates/PL3-01-CV2071-QLĐĐ (Cap nhat).xlsx"
# Tên file đích khi tải về
TOOL1_DESTINATION_FILE_NAME: str = "PL3-01-CV2071-QLĐĐ (Cap nhat).xlsx"

# --- CẤU HÌNH CÔNG CỤ 2: LÀM SẠCH & TÁCH FILE ---
STEP1_CHECK_COLS: List[str] = ["D", "E", "F", "I", "J", "L", "M", "R", "S", "T", "U"]
STEP1_START_ROW: int = 5
STEP1_YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
STEP1_EMPTY_FILL = PatternFill(fill_type=None)
STEP2_TARGET_COL: str = "G"
STEP2_START_ROW: int = 5
STEP2_EMPTY_FILL = PatternFill(fill_type=None)

# --- CÁC HÀM HELPER CHUNG ---
# (Giữ nguyên các hàm helper như code gốc: helper_copy_cell_format, helper_normalize_value, 
# helper_calculate_column_width, helper_cell_has_bg, helper_copy_rows_with_style, 
# helper_group_columns_openpyxl, helper_get_safe_filepath)

def get_sheet_names_from_buffer(file_buffer: io.BytesIO) -> List[str]:
    """Đọc tên các sheet từ một buffer file Excel mà không làm thay đổi vị trí con trỏ."""
    try:
        original_position = file_buffer.tell()
        file_buffer.seek(0)
        wb = load_workbook(file_buffer, read_only=True)
        sheet_names = wb.sheetnames
        file_buffer.seek(original_position)
        wb.close()
        return sheet_names
    except Exception as e:
        st.error(f"Không thể đọc sheet từ file: {e}")
        return []

def get_sheet_names_from_path(file_path: str) -> List[str]:
    """Đọc tên các sheet từ file Excel theo đường dẫn."""
    try:
        wb = load_workbook(file_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        return sheet_names
    except Exception as e:
        st.error(f"Không thể đọc sheet từ file mẫu: {e}")
        return []

def tool1_transform_and_copy(source_buffer, source_sheet, dest_sheet, progress_bar, status_label):
    """
    Sao chép và ánh xạ dữ liệu từ file nguồn sang file đích dựa trên file mẫu cố định.
    """
    try:
        # 1. Đọc dữ liệu nguồn
        status_label.info("Đang đọc dữ liệu từ file nguồn...")
        source_cols_letters_list = list(TOOL1_COLUMN_MAPPING.keys())
        source_cols_str = ",".join(source_cols_letters_list)
        
        df_source = pd.read_excel(source_buffer,
                                 sheet_name=source_sheet,
                                 header=None,
                                 skiprows=2,
                                 usecols=source_cols_str,
                                 engine='openpyxl')
        
        sorted_source_cols = sorted(source_cols_letters_list, key=column_index_from_string)
        if len(df_source.columns) != len(sorted_source_cols):
            st.error(f"Lỗi đọc cột: Đọc được {len(df_source.columns)} cột, nhưng mong đợi {len(sorted_source_cols)} cột.")
            logging.error(f"Lỗi mapping cột: Đã đọc {df_source.columns} nhưng key là {sorted_source_cols}")
            return None

        df_source.columns = sorted_source_cols 
        df_source_renamed = df_source.rename(columns=TOOL1_COLUMN_MAPPING)
        progress_bar.progress(20)

        # 2. Mở file mẫu
        status_label.info("Đang mở file mẫu...")
        if not os.path.exists(TOOL1_TEMPLATE_FILE_PATH):
            st.error(f"Lỗi: Không tìm thấy file mẫu tại '{TOOL1_TEMPLATE_FILE_PATH}'.")
            logging.error(f"Không tìm thấy file mẫu tại {TOOL1_TEMPLATE_FILE_PATH}")
            return None
        wb_dest = load_workbook(TOOL1_TEMPLATE_FILE_PATH)
        if dest_sheet not in wb_dest.sheetnames:
            st.error(f"Lỗi: Không tìm thấy sheet '{dest_sheet}' trong file mẫu.")
            logging.error(f"Sheet '{dest_sheet}' không tồn tại trong file mẫu")
            wb_dest.close()
            return None
        ws_dest = wb_dest[dest_sheet]
        progress_bar.progress(40)

        # 3. Ghi dữ liệu
        status_label.info("Đang sao chép dữ liệu...")
        total_rows_to_write = len(df_source)
        
        for i, (source_col_letter_in_map, dest_col_letter) in enumerate(TOOL1_COLUMN_MAPPING.items()):
            col_index_dest = column_index_from_string(dest_col_letter)
            data_series = df_source_renamed[dest_col_letter]
            
            for j, value in enumerate(data_series, start=TOOL1_START_ROW_DESTINATION):
                cell_value = None if pd.isna(value) else value
                ws_dest.cell(row=j, column=col_index_dest, value=cell_value)
            
            progress_bar.progress(40 + int((i + 1) / len(TOOL1_COLUMN_MAPPING) * 40))

        # 4. Kẻ viền
        status_label.info("Đang kẻ viền cho vùng dữ liệu mới...")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        end_row_border = TOOL1_START_ROW_DESTINATION + total_rows_to_write - 1
        all_dest_cols_indices = [column_index_from_string(col) for col in TOOL1_COLUMN_MAPPING.values()]

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
        wb_dest.close()
        return output_buffer

    except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý Công cụ 1: {e}")
        logging.error(f"Lỗi Công cụ 1: {e}", exc_info=True)
        return None

# --- CÁC HÀM CHO CÔNG CỤ 2: LÀM SẠCH, PHÂN LOẠI & TÁCH FILE ---
# (Giữ nguyên tất cả các hàm của Công cụ 2 như code gốc: run_step_1_process, run_step_2_clear_fill, 
# run_step_3_split_by_color, run_step_4_split_files)

# --- GIAO DIỆN STREAMLIT CHÍNH ---
st.set_page_config(page_title="TSCopyRight", layout="wide", page_icon="🚀")

# --- SIDEBAR ---
st.sidebar.title("Hướng dẫn sử dụng")
st.sidebar.markdown("""
- **Kế hoạch số 515/KH-BCA-BNN&MT ngày 31/8/2025 của Bộ Công an và Bộ Nông nghiệp và Môi trường về việc triển khai thực hiện chiến dịch làm giàu, làm sạch cơ sở dữ liệu quốc gia về đất đai.
- **Công văn số 780/UBND-NNMT ngày 04/9/2025 của UBND tỉnh Quảng Trị về việc triển khai Kế hoạch số 515/KH-BCA-BNN&MT.
- **Công văn số 2071/QLĐĐ-TKKKTTĐĐ ngày 05/9/2025 của Cục Quản lý đất đai về việc hướng dẫn tổ chức thực hiện chiến dịch làm giàu, làm sạch cơ sở dữ liệu quốc gia về đất đai.
- **Công văn số 1730/SNNMT-ĐĐBĐVT ngày 08/9/2025 của Sở Nông nghiệp và Môi trường tỉnh Quảng Trị về việc triển khai Kế hoạch số 515/KH-BCA-BNN&MT.
- **Quyết định 1392/QĐ-UBND ngày 10/9/2025, của UBND tỉnh về việc thành lập Tổ công tác.
- **Kế hoạch số 847/KH-UBND ngày 10/9/2025, của UBND tỉnh Quảng Trị về triển khai thực hiện chiến dịch làm giàu, làm sạch cơ sở dữ liệu đất đai.
- **Công văn số 2240/QLĐĐ-TKKKTTĐĐ ngày 19/9/2025, về việc phối hợp với các đơn vị phần mềm trong thực hiện Kế hoạch số 515/KH-BCA-BNN&MT.
- **Công văn số /QLĐĐ-TKKKTTĐĐ Tháng 10 năm 2025, về tài liệu hướng dẫn bổ sung theo Công văn số 2071/QLĐĐ-TKKKTTĐĐ.
""")
st.sidebar.info("Phát triển dựa trên quy trình nghiệp vụ của Trường Sinh - SĐT 0917.750.555.")

# --- MAIN PAGE ---
st.title("Chiến Dịch Xây Dựng Cơ Sở Dữ Liệu Đất Đai")
st.header("Bộ Công cụ Hỗ trợ Dữ liệu")
st.markdown("---")

# --- TẠO 2 TAB CHO 2 CÔNG CỤ ---
tab1, tab2 = st.tabs([
    "Công cụ 1: Sao chép & Ánh xạ Dữ liệu",
    "Công cụ 2: Làm sạch & Tách file (Quy trình chính)"
])

# --- GIAO DIỆN CHO CÔNG CỤ 1 ---
with tab1:
    st.subheader("Sao chép dữ liệu từ File Nguồn sang File Mẫu")
    
    st.markdown("### Bước 1: Tải lên File Nguồn (File chứa dữ liệu)")
    source_file = st.file_uploader("Chọn File Nguồn (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="tool1_source")
    
    source_sheet = None
    dest_sheet = None

    col1, col2 = st.columns(2)
    with col1:
        if source_file:
            source_sheets = get_sheet_names_from_buffer(source_file)
            source_sheet = st.selectbox("Chọn Sheet Nguồn (để đọc):", source_sheets, key="tool1_source_sheet")
    
    with col2:
        try:
            dest_sheets = get_sheet_names_from_path(TOOL1_TEMPLATE_FILE_PATH)
            dest_sheet = st.selectbox("Chọn Sheet Đích (để ghi):", dest_sheets, key="tool1_dest_sheet")
        except Exception as e:
            st.error(f"Không thể đọc file mẫu tại '{TOOL1_TEMPLATE_FILE_PATH}'. Vui lòng kiểm tra!")
            logging.error(f"Lỗi đọc file mẫu: {e}")
            dest_sheet = None

    st.markdown("### Bước 2: Xác nhận")
    start_tool1 = st.button("Bắt đầu Sao chép & Ánh xạ", key="tool1_start")

    if start_tool1:
        if not source_file or not source_sheet or not dest_sheet:
            st.error("Vui lòng tải lên file nguồn và chọn cả hai sheet.")
        else:
            progress_bar_tool1 = st.progress(0)
            status_label_tool1 = st.empty()
            
            try:
                source_file.seek(0)
                result_buffer = tool1_transform_and_copy(
                    source_file, source_sheet, 
                    dest_sheet, 
                    progress_bar_tool1, status_label_tool1
                )
                
                if result_buffer:
                    status_label_tool1.success("✅ HOÀN TẤT!")
                    st.download_button(
                        label="Tải về File Đích đã cập nhật",
                        data=result_buffer,
                        file_name=TOOL1_DESTINATION_FILE_NAME,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    status_label_tool1.error("Xử lý thất bại. Vui lòng kiểm tra log.")
            
            except Exception as e:
                st.error(f"Lỗi nghiêm trọng Công cụ 1: {e}")
                logging.error(f"Lỗi Streamlit Tool 1: {e}", exc_info=True)

# --- GIAO DIỆN CHO CÔNG CỤ 2 ---
# (Giữ nguyên toàn bộ giao diện và logic của Công cụ 2 như code gốc)

with tab2:
    st.subheader("Làm sạch, Phân loại và Tách file tự động")
    
    st.markdown("### Bước 1: Tải lên File Excel")
    uploaded_file_tool2 = st.file_uploader("Chọn file Excel của bạn (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="tool2_uploader")

    if uploaded_file_tool2:
        st.markdown("---")
        st.markdown("### Bước 2: Chọn Sheet")
        try:
            uploaded_file_tool2.seek(0)
            wb_sheets = load_workbook(uploaded_file_tool2, read_only=True)
            sheet_names = wb_sheets.sheetnames
            wb_sheets.close()
            
            selected_sheet_tool2 = st.selectbox("Chọn sheet chính để xử lý:", sheet_names, 
                                               help="Đây là sheet gốc chứa dữ liệu bạn muốn lọc.", 
                                               key="tool2_sheet_select")

            st.markdown("### Bước 3: Xác nhận")
            start_button_tool2 = st.button("Bắt đầu Làm sạch & Tách file", key="tool2_start")
            st.markdown("---")

            if start_button_tool2:
                st.markdown("### Bước 4: Hoàn thành và Tải về")
                progress_bar = st.progress(0)
                status_text_area = st.empty()
                
                try:
                    status_text_area.info("Đang tải file vào bộ nhớ...")
                    uploaded_file_tool2.seek(0)
                    main_wb = load_workbook(uploaded_file_tool2)
                    
                    main_wb = run_step_1_process(main_wb, selected_sheet_tool2, progress_bar, status_text_area, 0, 25)
                    if main_wb is None: raise Exception("Bước 1 thất bại.")
                    
                    main_wb = run_step_2_clear_fill(main_wb, progress_bar, status_text_area, 25, 25)
                    if main_wb is None: raise Exception("Bước 2 thất bại.")
                    
                    main_wb = run_step_3_split_by_color(main_wb, progress_bar, status_text_area, 50, 25)
                    if main_wb is None: raise Exception("Bước 3 thất bại.")
                    
                    status_text_area.info("Đang chuẩn bị file kết quả...")
                    final_wb_buffer = io.BytesIO()
                    main_wb.save(final_wb_buffer)
                    final_wb_buffer.seek(0)
                    
                    step4_read_buffer = io.BytesIO(final_wb_buffer.read())
                    final_wb_buffer.seek(0)
                    
                    main_processed_filename = f"[Processed]_{uploaded_file_tool2.name}"
                    
                    zip_buffer = run_step_4_split_files(
                        step4_read_buffer,
                        final_wb_buffer,
                        main_processed_filename,
                        progress_bar, 
                        status_text_area, 
                        75, 
                        25
                    )
                    if zip_buffer is None: raise Exception("Bước 4 thất bại.")

                    main_wb.close()
                    
                    status_text_area.success("✅ HOÀN TẤT!")
                    progress_bar.progress(100)
                    
                    st.download_button(
                        label="🗂️ Tải về Gói Kết Quả (ZIP)",
                        data=zip_buffer,
                        file_name="KetQua_Thon.zip",
                        mime="application/zip",
                        help=f"File ZIP này chứa file Excel chính ({main_processed_filename}) VÀ tất cả các file con được tách ra từ 'Nhóm 2_GDC'."
                    )
                    
                except Exception as e:
                    st.error(f"Quy trình đã dừng do lỗi: {e}")
                    logging.error(f"Lỗi Streamlit Workflow Tool 2: {e}")

        except Exception as e:
            st.error(f"Không thể đọc file Excel. File có thể bị hỏng hoặc không đúng định dạng: {e}")
            logging.error(f"Lỗi Streamlit Tải file Tool 2: {e}")
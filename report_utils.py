# report_utils.py
# Updated to use Traditional Chinese keys for data lookup

import streamlit as st # Still needed for st.write/info/error/warning inside functions
import openpyxl
from fpdf import FPDF
from pypdf import PdfWriter # Use PdfWriter from pypdf
import io
import subprocess
import tempfile
import os
import platform
import datetime # Not currently used here, but might be if adding date stamps directly

# --- Helper Function ---

def get_value_for_year(results_list, target_year):
    """Helper to find the CSV for a specific year in a list of results."""
    if not results_list: return None
    for item in results_list:
        if item.get('year') == target_year:
            # Return total_csv, or adapt if you need guaranteed/non-guaranteed parts
            return item.get('total_csv', None)
    return None # Or 0 or '--'

# --- Excel Generation Function ---

def create_plan_excel(calculated_data, template_path):
    """
    Generates an Excel file by populating a template with calculated data.
    - Populates B1-B6 parameters (using fixed 5 years for total premium).
    - Leaves A11-A25 (Year Labels) unchanged.
    - Leaves C9/C10, E9/E10 (Withdrawal Descriptions) unchanged.
    - Populates results grid (e.g., B12-B25, D12-D25, F12-F25) for the fixed
      set of years (10-90) defined in the mapping, ignoring UI year selection.

    Args:
        calculated_data (dict): Dictionary containing parameters and scenario results.
                                EXPECTS keys like "無提取", "提取方案 A", "提取方案 B".
        template_path (str): Path to the Excel template file (.xlsx).

    Returns:
        bytes: Content of the generated Excel file, or None if an error occurs.
    """
    st.info("正在生成 Excel 報告...") # Translated message
    if not calculated_data:
        st.error("沒有可用於生成 Excel 的計算數據。") # Translated message
        return None
    if not os.path.exists(template_path):
        st.error(f"找不到 Excel 範本檔案於：{template_path}") # Translated message
        st.info(f"請確保在 '{os.path.dirname(template_path)}' 目錄中存在名為 '{os.path.basename(template_path)}' 的檔案。") # Translated message
        return None

    try:
        workbook = openpyxl.load_workbook(template_path)
        try:
            sheet = workbook['Sheet1']
        except KeyError:
            st.warning("找不到範本工作表 'Sheet1'，將使用活動工作表。") # Translated message
            sheet = workbook.active

        params = calculated_data.get('parameters', {})
        template_years = params.get('report_years', []) # Should be [10, ..., 90]

        # --- Populate Parameter Cells ---
        # B1: Client Name
        try: sheet['B1'] = params.get('client_name', 'N/A')
        except Exception as cell_err: st.warning(f"無法將客戶名稱寫入儲存格 B1：{cell_err}") # Translated message
        # B2 - B4: Fixed (Do nothing)
        # B5: Annual Premium
        try:
            annual_prem_val = params.get('premium');
            if annual_prem_val is not None: sheet['B5'] = float(annual_prem_val)
            else: sheet['B5'] = "N/A"
        except Exception as cell_err: st.warning(f"無法將年繳保費寫入儲存格 B5：{cell_err}") # Translated message
        # B6: Total Premium (Use FIXED_PAYMENT_YEARS from params)
        try:
            fixed_payment_years_from_params = int(params.get('years', 5))
            total_prem = float(params.get('premium', 0)) * fixed_payment_years_from_params
            sheet['B6'] = total_prem
        except Exception as cell_err: st.warning(f"無法將總保費寫入儲存格 B6：{cell_err}") # Translated message

        # --- Populate Withdrawal Scenario Descriptions ---
        # (Assuming these are fixed in the template for now)

        # --- Populate Results Cells for FIXED Years (10-90) ---
        # Assumes: Col B = No Withdrawal, Col D = Withdrawal A, Col F = Withdrawal B
        # *** KEY CHANGE: Use Chinese keys to match app.py ***
        scenario_map = {
            "無提取": 'B',     # Changed from 'no_withdrawal'
            "提取方案 A": 'D', # Changed from 'withdrawal_A'
            "提取方案 B": 'F'  # Changed from 'withdrawal_B'
            # Add more scenarios/columns if needed, ensuring keys match app.py
        }

        cell_map_results_fixed = {}
        start_row = 12
        for i, year in enumerate(template_years):
            current_row = start_row + i
            cell_map_results_fixed[year] = {}
            for scen_key_chinese, col_letter in scenario_map.items():
                 cell_map_results_fixed[year][scen_key_chinese] = f'{col_letter}{current_row}'


        # Iterate through the FIXED template years and populate data
        for year in template_years:
            if year in cell_map_results_fixed:
                year_map = cell_map_results_fixed[year]
                # year_map keys are now the Chinese scenario names
                for scenario_key_chinese, cell_ref in year_map.items():
                    # Check if the Chinese key exists in the calculated data
                    if scenario_key_chinese in calculated_data:
                        results_list = calculated_data.get(scenario_key_chinese)
                        value = get_value_for_year(results_list, year)
                        if value is not None:
                            try: sheet[cell_ref] = float(value)
                            except (ValueError, TypeError): sheet[cell_ref] = str(value) # Fallback
                        else: sheet[cell_ref] = None # Clear cell if no value for that year/scenario
                    else:
                         # Optional: Clear cell if the whole scenario is missing from data
                         sheet[cell_ref] = None


        # --- Save to memory buffer ---
        output_buffer = io.BytesIO()
        workbook.save(output_buffer)
        workbook.close()
        st.info("Excel 範本已成功填入。") # Translated message
        return output_buffer.getvalue()

    except Exception as e:
        st.error(f"創建格式化 Excel 輸出時出錯：{e}") # Translated message
        import traceback
        st.error(traceback.format_exc())
        return None


# --- PDF Generation Function (Using Excel -> PDF Conversion Approach) ---

def create_plan_pdf(calculated_data, template_path_excel, static_pdf_path):
    """
    Generates a PDF by:
    1. Generating the Excel report in memory using create_plan_excel.
    2. Converting the Excel data to PDF (Page 1) using LibreOffice.
    3. Merging the converted PDF (Page 1) with the static PDF (Page 2).

    Args:
        calculated_data (dict): Dictionary containing parameters and scenario results.
        template_path_excel (str): Path to the Excel template file (.xlsx).
        static_pdf_path (str): Path to the static PDF file to append.

    Returns:
        bytes: Content of the final merged PDF file, or None if an error occurs.
    """
    st.info("正在生成 PDF (Excel -> PDF 方法)...") # Translated message
    if not calculated_data:
        st.error("沒有可用於生成 PDF 的計算數據。") # Translated message
        return None

    # 1. Generate Excel data first
    st.info("步驟 1：正在生成中間 Excel 數據...") # Translated message
    excel_bytes = create_plan_excel(calculated_data, template_path_excel)

    if not excel_bytes:
        st.error("生成中間 Excel 數據失敗。無法繼續生成 PDF。") # Translated message
        return None
    st.info("步驟 1：中間 Excel 數據已生成。") # Translated message

    # --- 2. Convert Excel Bytes to PDF Bytes using LibreOffice ---
    st.info("步驟 2：正在透過 LibreOffice 將 Excel 轉換為 PDF (第 1 頁)...") # Translated message
    page1_pdf_bytes = None
    temp_excel_file = None
    output_dir = None
    output_pdf_path = None
    temp_excel_path = None # Define temp_excel_path outside try block

    try:
        # Create a temporary directory
        output_dir = tempfile.mkdtemp()

        # Create a temporary file *within* the directory to write Excel bytes
        # Use delete=False so we can pass the path to subprocess, then clean up manually
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=output_dir) as temp_excel_file:
             temp_excel_file.write(excel_bytes)
             temp_excel_path = temp_excel_file.name # Get the path

        # Determine LibreOffice Path (Adjust as needed)
        soffice_path = None
        if platform.system() == "Windows":
            # Check common Program Files locations
            prog_files = os.environ.get("ProgramFiles", "C:\\Program Files")
            prog_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
            paths_to_check = [
                os.path.join(prog_files, "LibreOffice", "program", "soffice.exe"),
                os.path.join(prog_files_x86, "LibreOffice", "program", "soffice.exe")
            ]
            for path in paths_to_check:
                 if os.path.exists(path):
                     soffice_path = path
                     break
            if not soffice_path: raise FileNotFoundError("LibreOffice not found at default paths.")

        elif platform.system() == "Darwin": # macOS
            soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            if not os.path.exists(soffice_path): raise FileNotFoundError("LibreOffice not found at default macOS path.")
        else: # Linux/Other (Assume 'soffice' is in PATH)
            # Check if soffice is actually callable
            try:
                 result = subprocess.run(["which", "soffice"], capture_output=True, text=True, check=True)
                 soffice_path = result.stdout.strip()
                 if not soffice_path: raise FileNotFoundError # Handle empty output from 'which'
            except (FileNotFoundError, subprocess.CalledProcessError):
                 raise FileNotFoundError("LibreOffice 'soffice' command not found in PATH.")

        # Construct and Run Command
        soffice_command = [
            soffice_path, "--headless", "--invisible", "--nologo",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            temp_excel_path # Use the path obtained earlier
        ]
        st.write(f"執行中：{' '.join(soffice_command)}") # Translated message (Executing)
        timeout_seconds = 60
        result = subprocess.run(soffice_command, capture_output=True, text=True, timeout=timeout_seconds, check=False) # check=False to handle errors manually

        # Construct expected output path
        output_filename = os.path.splitext(os.path.basename(temp_excel_path))[0] + ".pdf"
        output_pdf_path = os.path.join(output_dir, output_filename)

        # Check results
        if result.returncode == 0 and os.path.exists(output_pdf_path):
            st.info("LibreOffice 轉換成功。") # Translated message
            with open(output_pdf_path, "rb") as f:
                page1_pdf_bytes = f.read()
        else:
            st.error(f"LibreOffice 轉換失敗！返回碼：{result.returncode}") # Translated message
            st.error(f"標準錯誤：{result.stderr or '無'}") # Translated message
            st.error(f"標準輸出：{result.stdout or '無'}") # Translated message
            st.error(f"預期輸出路徑：{output_pdf_path} (存在：{os.path.exists(output_pdf_path)})") # Translated message
            page1_pdf_bytes = None

    except FileNotFoundError as fnf_err:
        st.error(f"轉換失敗：找不到 LibreOffice。請確保已安裝 LibreOffice 且路徑正確。錯誤：{fnf_err}") # Translated message
        page1_pdf_bytes = None
    except subprocess.TimeoutExpired:
        st.error(f"LibreOffice 轉換在 {timeout_seconds} 秒後超時。") # Translated message
        page1_pdf_bytes = None
    except Exception as conversion_err:
        st.error(f"Excel 到 PDF 轉換期間發生錯誤：{conversion_err}") # Translated message
        import traceback
        st.error(traceback.format_exc())
        page1_pdf_bytes = None
    finally: # Cleanup
        if temp_excel_path and os.path.exists(temp_excel_path):
            try: os.remove(temp_excel_path)
            except Exception as e: st.warning(f"無法移除暫存 Excel 檔案：{e}") # Translated message
        if output_pdf_path and os.path.exists(output_pdf_path):
             try: os.remove(output_pdf_path)
             except Exception as e: st.warning(f"無法移除暫存 PDF 檔案：{e}") # Translated message
        if output_dir and os.path.exists(output_dir):
             # Try removing directory - might fail if files are still locked briefly
             try: os.rmdir(output_dir)
             except OSError as e: st.warning(f"無法移除暫存目錄（可能稍後自動清理）：{e}") # Translated message


    if not page1_pdf_bytes:
        st.error("Excel 到 PDF 轉換步驟失敗。無法生成最終的合併 PDF。") # Translated message
        return None
    st.info("步驟 2：Excel 到 PDF 轉換完成。") # Translated message

    # --- 3. Merge Converted PDF (Page 1) with Static PDF (Page 2) ---
    st.info("步驟 3：正在合併動態第 1 頁與靜態第 2 頁...") # Translated message
    try:
        merger = PdfWriter()
        merger.append(fileobj=io.BytesIO(page1_pdf_bytes)) # Page 1
        static_content_merged = False
        if static_pdf_path and os.path.exists(static_pdf_path):
             try:
                 merger.append(static_pdf_path)
                 static_content_merged = True # Page 2
             except Exception as merge_err: st.warning(f"無法合併靜態 PDF：{merge_err}。") # Translated message
        else: st.warning(f"找不到靜態 PDF：{static_pdf_path}。最終 PDF 將只有第 1 頁。") # Translated message

        final_pdf_buffer = io.BytesIO()
        merger.write(final_pdf_buffer)
        merger.close()
        st.info(f"PDF 合併成功 ({len(merger.pages)} 頁)。已合併靜態內容：{static_content_merged}") # Translated message
        return final_pdf_buffer.getvalue()

    except Exception as e:
        st.error(f"PDF 合併期間出錯：{e}") # Translated message
        import traceback
        st.error(traceback.format_exc())
        return None
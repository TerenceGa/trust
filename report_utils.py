# report_utils.py
# Updated to use Traditional Chinese keys for data lookup
# ADDED Debug prints for byte buffer sizes
# FIXED NameError by adding find_soffice_path locally
# UPDATED: Removed verbose st.info/st.write messages for cleaner UI. Kept st.error and critical st.warning.
# UPDATED: Dynamically populates cell C10 with Withdrawal Scenario A details.
# UPDATED: Dynamically populates cell E10 with Withdrawal Scenario B details.

import time
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
import traceback # Import traceback for detailed error printing
import shutil # Import shutil for rmtree


# --- Helper Function to Find LibreOffice (Copied Here) ---
def find_soffice_path():
    """Attempts to find the path to the LibreOffice executable."""
    print("--- DEBUG (report_utils): Entering find_soffice_path (local) ---") # DEBUG
    path = None
    try:
        if platform.system() == "Windows":
            print("--- DEBUG (report_utils): Checking Windows paths... ---") # DEBUG
            prog_files = os.environ.get("ProgramFiles", "C:\\Program Files")
            prog_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
            paths_to_check = [
                os.path.join(prog_files, "LibreOffice", "program", "soffice.exe"),
                os.path.join(prog_files_x86, "LibreOffice", "program", "soffice.exe"),
            ]
            for p in paths_to_check:
                if os.path.exists(p):
                    print(f"--- DEBUG (report_utils): Found at: {p} ---") # DEBUG
                    path = p
                    break
            if not path:
                print("--- DEBUG (report_utils): Checking 'where' command... ---") # DEBUG
                try:
                    result = subprocess.run(["where", "soffice.exe"], capture_output=True, text=True, check=True, shell=True)
                    paths = result.stdout.strip().splitlines()
                    if paths and os.path.exists(paths[0]):
                         print(f"--- DEBUG (report_utils): Found via 'where': {paths[0]} ---") # DEBUG
                         path = paths[0]
                except (FileNotFoundError, subprocess.CalledProcessError) as where_err:
                    print(f"--- DEBUG (report_utils): 'where' command failed or not found: {where_err} ---") # DEBUG
        elif platform.system() == "Darwin": # macOS
             print("--- DEBUG (report_utils): Checking macOS paths... ---") # DEBUG
             p = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
             if os.path.exists(p):
                  print(f"--- DEBUG (report_utils): Found at: {p} ---") # DEBUG
                  path = p
             if not path: # Try 'which' command as a fallback
                try:
                    result = subprocess.run(["which", "soffice"], capture_output=True, text=True, check=True)
                    p_which = result.stdout.strip()
                    if p_which and os.path.exists(p_which):
                        print(f"--- DEBUG (report_utils): Found via 'which': {p_which} ---")
                        path = p_which
                except (FileNotFoundError, subprocess.CalledProcessError) as e:
                    print(f"--- DEBUG (report_utils): 'which soffice' failed: {e} ---")
        else: # Linux/Other
             print("--- DEBUG (report_utils): Checking Linux/Other paths using 'which'... ---") # DEBUG
             try:
                result = subprocess.run(["which", "soffice"], capture_output=True, text=True, check=True)
                p_which = result.stdout.strip()
                if p_which and os.path.exists(p_which):
                    print(f"--- DEBUG (report_utils): Found via 'which': {p_which} ---")
                    path = p_which
                else: # Check common install locations if 'which' fails or path is invalid
                    common_paths = ["/usr/bin/soffice", "/usr/local/bin/soffice", "/opt/libreoffice/program/soffice", "/snap/bin/libreoffice.soffice"]
                    for p_common in common_paths:
                        if os.path.exists(p_common):
                            print(f"--- DEBUG (report_utils): Found in common path: {p_common} ---")
                            path = p_common
                            break
             except (FileNotFoundError, subprocess.CalledProcessError) as e:
                print(f"--- DEBUG (report_utils): 'which soffice' failed: {e}. Will check common paths. ---")
                common_paths = ["/usr/bin/soffice", "/usr/local/bin/soffice", "/opt/libreoffice/program/soffice", "/snap/bin/libreoffice.soffice"]
                for p_common in common_paths:
                    if os.path.exists(p_common):
                        print(f"--- DEBUG (report_utils): Found in common path: {p_common} ---")
                        path = p_common
                        break
    except Exception as e:
        print(f"--- ERROR (report_utils): Exception in find_soffice_path: {e} ---") # DEBUG
        print(traceback.format_exc())
        path = None

    print(f"--- DEBUG (report_utils): Exiting find_soffice_path. Found path: {path} ---") # DEBUG
    return path

# --- Helper Function ---
def get_value_for_year(results_list, target_year):
    """Helper to find the CSV for a specific year in a list of results."""
    if not results_list: return None
    for item in results_list:
        if item.get('year') == target_year:
            return item.get('total_csv', None)
    return None

# --- Helper function to generate withdrawal scenario text ---
def get_withdrawal_scenario_text(start_val, amount_val, scenario_letter):
    """Generates descriptive text for a withdrawal scenario."""
    if start_val > 0 and amount_val > 0:
        formatted_amount = int(amount_val) if amount_val == int(amount_val) else amount_val
        return f"从第{start_val}年开始提取\n每年提取{formatted_amount}美金"
    elif start_val > 0 and amount_val <= 0:
        return f"从第{start_val}年开始提取\n提取金额未设定或为零"
    elif start_val <= 0 and amount_val > 0:
        formatted_amount_partial = int(amount_val) if amount_val == int(amount_val) else amount_val
        return f"提取开始年份未设定\n计划每年提取{formatted_amount_partial}美金"
    else:
        return f"未設定提取方案 {scenario_letter}"

# --- Excel Generation Function ---
def create_plan_excel(calculated_data, template_path):
    """
    Generates an Excel file by populating a template with calculated data.
    """
    print("--- DEBUG (report_utils): Entering create_plan_excel ---") # DEBUG
    if not calculated_data:
        print("--- DEBUG (report_utils): No calculated data for Excel. ---") # DEBUG
        st.error("沒有可用於生成 Excel 的計算數據。") # Keep UI error
        return None
    if not os.path.exists(template_path):
        print(f"--- DEBUG (report_utils): Excel template not found: {template_path} ---") # DEBUG
        st.error(f"找不到 Excel 範本檔案於：{template_path}") # Keep UI error
        st.info(f"請確保在 '{os.path.dirname(template_path)}' 目錄中存在名為 '{os.path.basename(template_path)}' 的檔案。")
        return None

    workbook = None
    try:
        print(f"--- DEBUG (report_utils): Loading Excel template: {template_path} ---") # DEBUG
        workbook = openpyxl.load_workbook(template_path)
        try:
            sheet = workbook['Sheet1']
            print("--- DEBUG (report_utils): Found sheet 'Sheet1'. ---") # DEBUG
        except KeyError:
            st.warning("找不到範本工作表 'Sheet1'，將使用活動工作表。") # Keep UI warning
            sheet = workbook.active
            print("--- DEBUG (report_utils): Using active sheet. ---") # DEBUG

        params = calculated_data.get('parameters', {})
        template_years = params.get('report_years', [])
        print(f"--- DEBUG (report_utils): Parameters for Excel: {params}") # DEBUG
        print(f"--- DEBUG (report_utils): Template years for Excel: {template_years}") # DEBUG

        try: sheet['B1'] = params.get('client_name', 'N/A')
        except Exception as cell_err: st.warning(f"無法將客戶名稱寫入儲存格 B1：{cell_err}")
        try:
            annual_prem_val = params.get('premium');
            if annual_prem_val is not None: sheet['B5'] = float(annual_prem_val)
            else: sheet['B5'] = "N/A"
        except Exception as cell_err: st.warning(f"無法將年繳保費寫入儲存格 B5：{cell_err}")
        try:
            fixed_payment_years_from_params = int(params.get('years', 5))
            total_prem = float(params.get('premium', 0)) * fixed_payment_years_from_params
            sheet['B6'] = total_prem
        except Exception as cell_err: st.warning(f"無法將總保費寫入儲存格 B6：{cell_err}")

        # --- Populate custom text for Withdrawal Scenario A in C10:D10 ---
        w_a_start_val = params.get('withdrawal_a_start', 0)
        w_a_amount_val = params.get('withdrawal_a_amount', 0)
        text_for_c10 = get_withdrawal_scenario_text(w_a_start_val, w_a_amount_val, "A")
        try:
            sheet['C10'] = text_for_c10
            # Ensure text wrapping is enabled for C10 in your Excel template
            # from openpyxl.styles import Alignment
            # sheet['C10'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
            print(f"--- DEBUG (report_utils): Wrote to C10: '{text_for_c10}' ---")
        except Exception as cell_err:
            st.warning(f"無法將提取方案A的詳細信息寫入儲存格 C10：{cell_err}")
            print(f"--- WARNING (report_utils): Could not write to C10: {cell_err} ---")

        # --- Populate custom text for Withdrawal Scenario B in E10:F10 ---
        w_b_start_val = params.get('withdrawal_b_start', 0)
        w_b_amount_val = params.get('withdrawal_b_amount', 0)
        text_for_e10 = get_withdrawal_scenario_text(w_b_start_val, w_b_amount_val, "B")
        try:
            sheet['E10'] = text_for_e10
            # Ensure text wrapping is enabled for E10 in your Excel template
            # from openpyxl.styles import Alignment
            # sheet['E10'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
            print(f"--- DEBUG (report_utils): Wrote to E10: '{text_for_e10}' ---")
        except Exception as cell_err:
            st.warning(f"無法將提取方案B的詳細信息寫入儲存格 E10：{cell_err}")
            print(f"--- WARNING (report_utils): Could not write to E10: {cell_err} ---")

        # --- Populate Results Cells ---
        scenario_map = {"無提取": 'B', "提取方案 A": 'D', "提取方案 B": 'F'}
        cell_map_results_fixed = {}
        start_row = 12
        for i, year in enumerate(template_years):
            current_row = start_row + i
            cell_map_results_fixed[year] = {scen_key: f'{col}{current_row}' for scen_key, col in scenario_map.items()}

        print("--- DEBUG (report_utils): Populating Excel results grid... ---") # DEBUG
        for year in template_years:
            if year in cell_map_results_fixed:
                year_map = cell_map_results_fixed[year]
                for scenario_key_chinese, cell_ref in year_map.items():
                    if scenario_key_chinese in calculated_data:
                        results_list = calculated_data.get(scenario_key_chinese)
                        value = get_value_for_year(results_list, year)
                        if value is not None:
                            try: sheet[cell_ref] = float(value)
                            except (ValueError, TypeError): sheet[cell_ref] = str(value)
                        else: sheet[cell_ref] = None
                    else:
                         sheet[cell_ref] = None

        output_buffer = io.BytesIO()
        workbook.save(output_buffer)
        workbook.close()
        workbook = None
        excel_bytes = output_buffer.getvalue()
        print(f"--- DEBUG (report_utils): Excel report generated. Byte size: {len(excel_bytes)} ---")
        return excel_bytes

    except Exception as e:
        print(f"--- ERROR (report_utils): Error creating Excel: {e} ---") # DEBUG
        print(traceback.format_exc())
        st.error(f"創建格式化 Excel 輸出時出錯：{e}") # Keep UI error
        if workbook: workbook.close()
        return None


# --- PDF Generation Function (Using Excel -> PDF Conversion Approach) ---
def create_plan_pdf(calculated_data, template_path_excel, static_pdf_path):
    """
    Generates a PDF by:
    1. Generating the Excel report in memory using create_plan_excel.
    2. Converting the Excel data to PDF (Page 1) using LibreOffice.
    3. Merging the converted PDF (Page 1) with the static PDF (Page 2).
    """
    print("--- DEBUG (report_utils): Entering create_plan_pdf ---") # DEBUG
    if not calculated_data:
        print("--- DEBUG (report_utils): No calculated data for PDF. ---") # DEBUG
        st.error("沒有可用於生成 PDF 的計算數據。") # Keep UI error
        return None

    excel_bytes = create_plan_excel(calculated_data, template_path_excel)

    if not excel_bytes:
        print("--- DEBUG (report_utils): Intermediate Excel generation failed. ---") # DEBUG
        st.error("生成中間 Excel 數據失敗。無法繼續生成 PDF。") # Keep UI error
        return None
    print(f"--- DEBUG (report_utils): Intermediate Excel byte size for PDF: {len(excel_bytes)} ---") # DEBUG

    page1_pdf_bytes = None
    temp_excel_file = None
    output_dir = None
    output_pdf_path = None
    temp_excel_path = None
    soffice_path = find_soffice_path()

    if not soffice_path:
        print("--- DEBUG (report_utils): soffice not found for PDF generation. ---") # DEBUG
        st.error("找不到 LibreOffice 'soffice'。無法將 Excel 轉換為 PDF。") # Keep UI error
        return None

    try:
        output_dir = tempfile.mkdtemp()
        print(f"--- DEBUG (report_utils): Created temp dir for PDF conversion: {output_dir} ---") # DEBUG
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=output_dir) as temp_excel_file:
             temp_excel_file.write(excel_bytes)
             temp_excel_path = temp_excel_file.name
             print(f"--- DEBUG (report_utils): Wrote intermediate Excel to: {temp_excel_path} ---") # DEBUG

        soffice_command = [
            soffice_path, "--headless", "--invisible", "--nologo",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            temp_excel_path
        ]
        print(f"--- DEBUG (report_utils): Running LibreOffice for PDF: {' '.join(soffice_command)} ---") # DEBUG
        timeout_seconds = 60
        result = subprocess.run(soffice_command, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=timeout_seconds, check=False)
        print(f"--- DEBUG (report_utils): LibreOffice PDF conversion RC: {result.returncode} ---") # DEBUG
        print(f"--- DEBUG (report_utils): LO PDF stdout:\n{result.stdout or 'None'}\n---") # DEBUG
        print(f"--- DEBUG (report_utils): LO PDF stderr:\n{result.stderr or 'None'}\n---") # DEBUG

        output_filename = os.path.splitext(os.path.basename(temp_excel_path))[0] + ".pdf"
        output_pdf_path = os.path.join(output_dir, output_filename)
        print(f"--- DEBUG (report_utils): Expected PDF output path: {output_pdf_path} ---") # DEBUG
        time.sleep(1)

        if result.returncode == 0 and os.path.exists(output_pdf_path):
            with open(output_pdf_path, "rb") as f:
                page1_pdf_bytes = f.read()
            print(f"--- DEBUG (report_utils): Read Page 1 PDF bytes. Size: {len(page1_pdf_bytes)} ---") # DEBUG
        else:
            st.error(f"LibreOffice 轉換失敗！返回碼：{result.returncode}") # Keep UI error
            st.error(f"標準錯誤：{result.stderr or '無'}") # Keep UI error
            page1_pdf_bytes = None

    except Exception as conversion_err:
        print(f"--- ERROR (report_utils): Error during Excel->PDF conversion: {conversion_err} ---") # DEBUG
        print(traceback.format_exc())
        st.error(f"Excel 到 PDF 轉換期間發生錯誤：{conversion_err}") # Keep UI error
        page1_pdf_bytes = None
    finally:
        if temp_excel_path and os.path.exists(temp_excel_path):
            try: os.remove(temp_excel_path)
            except Exception as e: st.warning(f"無法移除暫存 Excel 檔案：{e}")
        if output_pdf_path and os.path.exists(output_pdf_path):
             try: os.remove(output_pdf_path)
             except Exception as e: st.warning(f"無法移除暫存 PDF 檔案：{e}")
        if output_dir and os.path.exists(output_dir):
             try: shutil.rmtree(output_dir)
             except OSError as e: st.warning(f"無法移除暫存目錄（可能稍後自動清理）：{e}")


    if not page1_pdf_bytes:
        print("--- DEBUG (report_utils): Page 1 PDF bytes are None. Cannot merge. ---") # DEBUG
        st.error("Excel 到 PDF 轉換步驟失敗。無法生成最終的合併 PDF。") # Keep UI error
        return None
    try:
        merger = PdfWriter()
        merger.append(fileobj=io.BytesIO(page1_pdf_bytes))
        static_content_merged = False
        if static_pdf_path and os.path.exists(static_pdf_path):
             print(f"--- DEBUG (report_utils): Appending static PDF: {static_pdf_path} ---") # DEBUG
             try:
                 merger.append(static_pdf_path)
                 static_content_merged = True
                 print("--- DEBUG (report_utils): Static PDF appended successfully. ---") # DEBUG
             except Exception as merge_err:
                 print(f"--- WARNING (report_utils): Failed to merge static PDF: {merge_err} ---") # DEBUG
                 st.warning(f"無法合併靜態 PDF：{merge_err}。最終 PDF 將只有第 1 頁。") # Keep UI warning
        else:
            print(f"--- WARNING (report_utils): Static PDF not found: {static_pdf_path} ---") # DEBUG
            st.warning(f"找不到靜態 PDF：{static_pdf_path}。最終 PDF 將只有第 1 頁。") # Keep UI warning

        final_pdf_buffer = io.BytesIO()
        merger.write(final_pdf_buffer)
        merger.close()
        final_pdf_bytes = final_pdf_buffer.getvalue()
        print(f"--- DEBUG (report_utils): Final PDF generated. Merged static: {static_content_merged}. Byte size: {len(final_pdf_bytes)} ---")
        return final_pdf_bytes

    except Exception as e:
        print(f"--- ERROR (report_utils): Error during PDF merge: {e} ---") # DEBUG
        print(traceback.format_exc())
        st.error(f"PDF 合併期間出錯：{e}") # Keep UI error
        return None

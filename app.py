# app.py
# Main Streamlit application file - Uses calculation_logic.py
# Simplified UI & Logic: Only writes Premium and Withdrawal details to Excel.
# CORRECTED: Uses single set of cells (F7/F8 assumed) for all withdrawal scenarios.
# ADDED Debug prints for report byte sizes and PDF display

import streamlit as st
# import pandas as pd
import os
import datetime
import report_utils # For PDF/Excel report generation
import calculation_logic # Import the new calculation module
import base64
import shutil # Needed for cleaning up temp dir
import traceback # Import traceback for detailed error printing


# --- Configuration ---
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "data")
# --- Path to the main calculation TOOLBOX ---
# Make sure filename is correct, including extension
CALCULATOR_XLSX_PATH = os.path.join(DATA_DIR, "TRST_Toolbox.xlsx")
# --- Paths for the final REPORT templates ---
EXCEL_TEMPLATE_PATH = os.path.join(DATA_DIR, "TRST_Template.xlsx") # For report_utils
STATIC_PDF_PATH = os.path.join(DATA_DIR, "Static_Info.pdf") # For report_utils

# --- Fixed Parameters ---
FIXED_PAYMENT_YEARS = 5 # This is assumed by the Excel template, not written
FIXED_REPORT_YEARS = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90]

# --- Cell Mapping for Inputs in TRST-toolbox-2025-04-20.xlsx ---
# !! ONLY include parameters actively set by the app !!
# !! Assumes F7 is withdrawal start year, F8 is withdrawal amount !!
INPUT_CELL_MAP = {
    'premium': 'C7',        # Premium Amount
    'withdrawal_start': 'F7', # Single cell for withdrawal start year input
    'withdrawal_amount': 'F8',# Single cell for withdrawal amount input
}

# --- REMOVE OLD MOCK FUNCTIONS ---

# --- New Function to Orchestrate Scenarios using calculation_logic ---
# (Keep this function as is from the previous version)
def generate_all_scenarios(inputs):
    """
    Orchestrates the calculation for No Withdrawal, Withdrawal A, and Withdrawal B
    by calling the run_calculation_scenario function from calculation_logic.py.
    Writes Premium to C7 (example).
    Writes Withdrawal Start/Amount to F7/F8 (example) for each scenario run.

    Args:
        inputs (dict): Dictionary containing user inputs from the sidebar
                       (client_name, premium, w_a_start, w_a_amount,
                        w_b_start, w_b_amount).

    Returns:
        dict: A dictionary containing parameters and results for each scenario,
              suitable for report_utils, or None if any scenario fails.
    """
    print("--- DEBUG: Entered generate_all_scenarios ---") # DEBUG
    # Extract inputs from UI
    client_name = inputs.get('client_name', 'N/A')
    annual_premium = inputs.get('premium', 0)
    w_a_start = inputs.get('w_a_start', 0)
    w_a_amount = inputs.get('w_a_amount', 0)
    w_b_start = inputs.get('w_b_start', 0)
    w_b_amount = inputs.get('w_b_amount', 0)

    # Prepare parameters dictionary for reporting
    report_parameters = {
        'client_name': client_name,
        'premium': annual_premium,
        'years': FIXED_PAYMENT_YEARS, # Still report the assumed payment term
        'report_years': FIXED_REPORT_YEARS,
        'withdrawal_a_start': w_a_start, # Report user's intended scenario A
        'withdrawal_a_amount': w_a_amount,
        'withdrawal_b_start': w_b_start, # Report user's intended scenario B
        'withdrawal_b_amount': w_b_amount,
        'calculation_date': datetime.date.today().isoformat()
    }

    # Create a unique temporary directory for this calculation run
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    temp_dir = os.path.join(DATA_DIR, f"temp_calcs_{timestamp}")
    try:
        os.makedirs(temp_dir, exist_ok=True)
        print(f"--- DEBUG: Created temporary directory: {temp_dir} ---") # DEBUG
        st.info(f"Created temporary directory: {temp_dir}")
    except OSError as e:
        print(f"--- ERROR: Failed to create temporary directory {temp_dir}: {e} ---") # DEBUG
        st.error(f"Failed to create temporary directory {temp_dir}: {e}")
        return None


    results_no_withdrawal = None
    results_withdrawal_A = None
    results_withdrawal_B = None
    calculation_successful = True # Flag to track overall success

    # --- Run Scenario 1: No Withdrawal ---
    print("--- DEBUG: Preparing Scenario: No Withdrawal ---") # DEBUG
    params_no_withdrawal = {
        'premium': annual_premium,
        'withdrawal_start': 0,
        'withdrawal_amount': 0
    }
    results_no_withdrawal = calculation_logic.run_calculation_scenario(
        "No Withdrawal", CALCULATOR_XLSX_PATH, temp_dir,
        params_no_withdrawal, INPUT_CELL_MAP, FIXED_REPORT_YEARS
    )
    if results_no_withdrawal is None:
        print("--- ERROR: Calculation failed for 'No Withdrawal' scenario. ---") # DEBUG
        st.error("Calculation failed for 'No Withdrawal' scenario.")
        calculation_successful = False # Mark as failed

    # --- Run Scenario 2: Withdrawal A ---
    if w_a_start > 0 and w_a_amount > 0 and calculation_successful:
        print("--- DEBUG: Preparing Scenario: Withdrawal A ---") # DEBUG
        params_withdrawal_a = {
            'premium': annual_premium,
            'withdrawal_start': w_a_start,
            'withdrawal_amount': w_a_amount
        }
        results_withdrawal_A = calculation_logic.run_calculation_scenario(
            "Withdrawal A", CALCULATOR_XLSX_PATH, temp_dir,
            params_withdrawal_a, INPUT_CELL_MAP, FIXED_REPORT_YEARS
        )
        if results_withdrawal_A is None:
            print("--- ERROR: Calculation failed for 'Withdrawal A' scenario. ---") # DEBUG
            st.error("Calculation failed for 'Withdrawal A' scenario.")

    # --- Run Scenario 3: Withdrawal B ---
    if w_b_start > 0 and w_b_amount > 0 and calculation_successful:
        print("--- DEBUG: Preparing Scenario: Withdrawal B ---") # DEBUG
        params_withdrawal_b = {
            'premium': annual_premium,
            'withdrawal_start': w_b_start,
            'withdrawal_amount': w_b_amount
        }
        results_withdrawal_B = calculation_logic.run_calculation_scenario(
            "Withdrawal B", CALCULATOR_XLSX_PATH, temp_dir,
            params_withdrawal_b, INPUT_CELL_MAP, FIXED_REPORT_YEARS
        )
        if results_withdrawal_B is None:
            print("--- ERROR: Calculation failed for 'Withdrawal B' scenario. ---") # DEBUG
            st.error("Calculation failed for 'Withdrawal B' scenario.")

    # --- Structure final results ---
    output_to_return = None
    if results_no_withdrawal is not None:
        final_results = {"parameters": report_parameters, "無提取": results_no_withdrawal}
        if results_withdrawal_A is not None:
            final_results["提取方案 A"] = results_withdrawal_A
        if results_withdrawal_B is not None:
            final_results["提取方案 B"] = results_withdrawal_B
        output_to_return = final_results
        print("--- DEBUG: Successfully structured final results. ---") # DEBUG
    else:
        print("--- DEBUG: Base scenario failed, returning None. ---") # DEBUG

    # Clean up temporary directory
    try:
        print(f"--- DEBUG: Attempting to clean up temp directory: {temp_dir} ---") # DEBUG
        shutil.rmtree(temp_dir)
        print(f"--- DEBUG: Cleaned up temporary directory: {temp_dir} ---") # DEBUG
        st.info(f"Cleaned up temporary directory: {temp_dir}")
    except Exception as e:
        print(f"--- WARNING: Could not clean up temp directory {temp_dir}: {e} ---") # DEBUG
        st.warning(f"Could not clean up temp directory {temp_dir}: {e}")

    print("--- DEBUG: Exiting generate_all_scenarios ---") # DEBUG
    return output_to_return


# --- Streamlit App UI ---
# (Keep UI section as is)
st.set_page_config(layout="wide", page_title="保險計劃生成器")
st.title("保誠保險計劃生成器")
top_col1, top_col2, top_col3 = st.columns([0.7, 0.15, 0.15])
st.markdown(f"""
請在側邊欄輸入參數並點擊「計算計劃預測」。
- **供款年期固定為 {FIXED_PAYMENT_YEARS} 年 (由Excel模板預設)。**
- **計算引擎:** 使用 `TRST-toolbox-2025-04-20.xlsx` 配合 LibreOffice。
- **注意:** 計算可能需要一些時間 (每次點擊約 1-3 分鐘)。
""")
def initialize_session_state():
    defaults = {
        'client_name': '尊貴客戶', 'premium': 10000,
        'w_a_amount': 0, 'w_a_start': 0, 'w_b_amount': 0, 'w_b_start': 0,
        'calculated_data': None, 'pdf_bytes': None, 'excel_bytes': None,
        'calculation_running': False
    }
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value
initialize_session_state()
with st.sidebar:
    st.header("計劃參數")
    st.text_input("客戶名稱", value=st.session_state.client_name, key="client_name")
    st.number_input("年繳保費 (美元)", min_value=1000, value=st.session_state.premium, step=1000, key="premium")
    st.metric(label="供款年期 (年)", value=f"{FIXED_PAYMENT_YEARS} (預設)")
    st.subheader("提取方案 (可選)")
    with st.expander("設定提取方案 A"):
         st.number_input("每年提取金額 A (美元)", min_value=0, value=st.session_state.w_a_amount, step=100, key="w_a_amount")
         st.number_input("由保單年度開始提取 A:", min_value=0, value=st.session_state.w_a_start, step=1, key="w_a_start")
    with st.expander("設定提取方案 B"):
         st.number_input("每年提取金額 B (美元)", min_value=0, value=st.session_state.w_b_amount, step=100, key="w_b_amount")
         st.number_input("由保單年度開始提取 B:", min_value=0, value=st.session_state.w_b_start, step=1, key="w_b_start")
    st.caption(f"報告將顯示第 {min(FIXED_REPORT_YEARS)} 年至第 {max(FIXED_REPORT_YEARS)} 年。")
    st.divider()
    calculate_button = st.button("計算計劃預測", type="primary", use_container_width=True, disabled=st.session_state.calculation_running, key="calc_button")
    status_placeholder = st.empty()

# --- Main Area: Calculation Trigger ---
# (Keep this block as is)
if calculate_button and not st.session_state.calculation_running:
    print("--- DEBUG: Calculate button clicked and not already running. ---") # DEBUG
    valid_input = True
    if st.session_state.w_a_start > 0 and st.session_state.w_a_amount <= 0:
        status_placeholder.error("⚠️ 如設定了提取方案 A 的開始年份，請輸入提取金額 A。")
        valid_input = False
    if st.session_state.w_b_start > 0 and st.session_state.w_b_amount <= 0:
        status_placeholder.error("⚠️ 如設定了提取方案 B 的開始年份，請輸入提取金額 B。")
        valid_input = False

    if valid_input:
        print("--- DEBUG: Input valid, setting calculation_running = True ---") # DEBUG
        st.session_state.calculation_running = True
        st.session_state['calculated_data'] = None
        st.session_state['pdf_bytes'] = None
        st.session_state['excel_bytes'] = None
        print("--- DEBUG: Rerunning app to disable button and show status... ---") # DEBUG
        st.rerun()
    else:
         print("--- DEBUG: Input invalid, not starting calculation. ---") # DEBUG


# --- Calculation Execution Block ---
if st.session_state.calculation_running:
    print("--- DEBUG: calculation_running is True, entering execution block. ---") # DEBUG
    with status_placeholder.container():
        st.info("⏳ 計算中... 請稍候 (可能需要1-3分鐘)。")
        st.info("請勿重複點擊按鈕。")

    current_inputs = {
        'client_name': st.session_state.client_name, 'premium': st.session_state.premium,
        'w_a_start': st.session_state.w_a_start, 'w_a_amount': st.session_state.w_a_amount,
        'w_b_start': st.session_state.w_b_start, 'w_b_amount': st.session_state.w_b_amount,
    }
    print(f"--- DEBUG: Calling generate_all_scenarios with inputs: {current_inputs} ---") # DEBUG

    try:
        calculated_data_result = generate_all_scenarios(current_inputs)
        print(f"--- DEBUG: generate_all_scenarios returned: {'Success' if calculated_data_result else 'Failure/None'} ---") # DEBUG

        if calculated_data_result:
            st.session_state['calculated_data'] = calculated_data_result
            params = calculated_data_result.get('parameters', {})
            print("--- DEBUG: Calculation successful, proceeding to report generation. ---") # DEBUG

            with status_placeholder.container():
                st.info("✓ 計算完成。")
                st.info("⏳ 正在生成 PDF 報告...")
                print("--- DEBUG: Calling report_utils.create_plan_pdf ---") # DEBUG
                pdf_bytes_result = report_utils.create_plan_pdf(calculated_data_result, EXCEL_TEMPLATE_PATH, STATIC_PDF_PATH)
                # --- ADDED DEBUG ---
                pdf_size = len(pdf_bytes_result) if pdf_bytes_result else 0
                print(f"--- DEBUG: PDF generation result type: {type(pdf_bytes_result)}, size: {pdf_size} bytes ---")
                # --- END DEBUG ---
                st.session_state['pdf_bytes'] = pdf_bytes_result
                if pdf_bytes_result: st.info("✓ PDF 報告已生成。")
                else: st.warning("⚠️ PDF 報告生成失敗。")

                st.info("⏳ 正在生成 Excel 報告...")
                print("--- DEBUG: Calling report_utils.create_plan_excel ---") # DEBUG
                excel_bytes_result = report_utils.create_plan_excel(calculated_data_result, EXCEL_TEMPLATE_PATH)
                # --- ADDED DEBUG ---
                excel_size = len(excel_bytes_result) if excel_bytes_result else 0
                print(f"--- DEBUG: Excel generation result type: {type(excel_bytes_result)}, size: {excel_size} bytes ---")
                # --- END DEBUG ---
                st.session_state['excel_bytes'] = excel_bytes_result
                if excel_bytes_result: st.info("✓ Excel 報告已生成。")
                else: st.warning("⚠️ Excel 報告生成失敗。")

            if st.session_state['pdf_bytes'] or st.session_state['excel_bytes']:
                 status_placeholder.success("✅ 報告生成完成！")
            else:
                 status_placeholder.error("❌ 計算完成，但 PDF 及 Excel 報告生成均失敗。")

        else:
            print("--- DEBUG: Calculation failed (generate_all_scenarios returned None). ---") # DEBUG
            status_placeholder.error("❌ 計算過程中發生錯誤，無法生成報告。")
            st.session_state['calculated_data'] = None
            st.session_state['pdf_bytes'] = None
            st.session_state['excel_bytes'] = None

    except Exception as calc_err:
        print(f"--- ERROR: Exception during calculation/report block: {calc_err} ---") # DEBUG
        print(traceback.format_exc()) # Print detailed traceback to console
        status_placeholder.error(f"❌ 處理計算或報告時發生意外錯誤: {calc_err}")
        st.session_state['calculated_data'] = None
        st.session_state['pdf_bytes'] = None
        st.session_state['excel_bytes'] = None
    finally:
         print("--- DEBUG: Calculation execution block finished, setting calculation_running = False ---") # DEBUG
         st.session_state.calculation_running = False
         print("--- DEBUG: Rerunning app one last time... ---") # DEBUG
         st.rerun()


# --- Display PDF and Download Buttons ---
# (Keep this section as is)
if st.session_state.get('calculated_data'):
    print("--- DEBUG: Displaying results area because calculated_data exists. ---") # DEBUG
    params = st.session_state['calculated_data'].get('parameters', {})
    client_name_for_file = params.get('client_name','Plan')
    premium_for_file = params.get('premium',0)
    years_for_file = params.get('years',0) # Payment years
    date_for_file = params.get('calculation_date','today')
    safe_client_name = "".join(c for c in client_name_for_file if c.isalnum() or c in (' ', '_')).rstrip() or 'Plan'
    base_filename = f"保險計劃_{safe_client_name}_{premium_for_file}_{years_for_file}年_{date_for_file}"

    pdf_bytes = st.session_state.get('pdf_bytes')
    excel_bytes = st.session_state.get('excel_bytes')

    # --- Download Buttons ---
    if pdf_bytes:
         with top_col2:
             pdf_filename = f"{base_filename}.pdf"
             st.download_button(label="下載 PDF", data=pdf_bytes, file_name=pdf_filename, mime="application/pdf", key="pdf_download_button_top", use_container_width=True)
    if excel_bytes:
         with top_col3:
             excel_filename = f"{base_filename}.xlsx"
             st.download_button(label="下載 Excel", data=excel_bytes, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="excel_download_button_top", use_container_width=True)

    # --- Embed PDF ---
    if pdf_bytes:
        try:
            # --- ADDED DEBUG ---
            print(f"--- DEBUG: Attempting to display PDF. Bytes length: {len(pdf_bytes)} ---")
            base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
            print(f"--- DEBUG: Base64 PDF preview (first 100 chars): {base64_pdf[:100]}... ---")
            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" height="800px" width="100%" style="border: none;"></iframe>'
            # print(f"--- DEBUG: PDF iframe HTML: {pdf_display} ---") # Optional: print full HTML
            # --- END DEBUG ---
            st.markdown(pdf_display, unsafe_allow_html=True)
            print("--- DEBUG: PDF iframe markdown executed. ---") # DEBUG
        except Exception as e:
            print(f"--- ERROR: Displaying PDF failed: {e} ---") # DEBUG
            print(traceback.format_exc()) # Print detailed traceback
            st.error(f"顯示 PDF 時出錯：{e}")
    else:
         # Only show warning if calculation completed but PDF failed
         if st.session_state.get('calculated_data'):
             print("--- DEBUG: Calculated data exists, but no PDF bytes to display. ---") # DEBUG
             st.warning("已完成計算，但無法載入 PDF 預覽。請嘗試下載 PDF (如有)。")

# --- Footer ---
st.divider()
st.caption(f"保險計劃生成器 v0.9.2 - {datetime.date.today().year}") # Incremented version

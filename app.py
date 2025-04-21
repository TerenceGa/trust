# app.py
# Main Streamlit application file - Traditional Chinese UI - PDF Embed (Using st.markdown)

import streamlit as st
import pandas as pd
import os
import datetime
import report_utils # Import the new module for report generation
import base64 # Needed for embedding PDF
# Removed: import streamlit.components.v1 as components - Not needed for this version

# --- Configuration ---
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "data")
EXCEL_TEMPLATE_PATH = os.path.join(DATA_DIR, "TRST_Template.xlsx")
STATIC_PDF_PATH = os.path.join(DATA_DIR, "Static_Info.pdf")

# --- Fixed Parameters ---
FIXED_PAYMENT_YEARS = 5
FIXED_REPORT_YEARS = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 70, 80, 90] # Years 10-90

# --- Core Calculation Logic (Placeholders / Mock) ---
# (Calculation functions remain the same as previous version)
def calculate_csv(annual_premium, payment_years, total_years, withdrawal_config=None):
    results = []
    current_value = 0.0
    withdraw_start_year = 0
    withdraw_amount = 0.0
    if withdrawal_config:
        withdraw_start_year = withdrawal_config.get('start_year', 0)
        withdraw_amount = withdrawal_config.get('amount', 0.0)
    for year in range(1, total_years + 1):
        if year <= FIXED_PAYMENT_YEARS:
            current_value += annual_premium
        if withdraw_start_year > 0 and year >= withdraw_start_year and withdraw_amount > 0:
            if current_value >= withdraw_amount:
                current_value -= withdraw_amount
            else: pass
        current_value = max(0, current_value)
        results.append({'year': year, 'total_csv': round(current_value, 2)})
    return results

def generate_scenarios(client_name, annual_premium,
                       withdrawal_a_start, withdrawal_a_amount,
                       withdrawal_b_start, withdrawal_b_amount):
    max_calc_year = max(FIXED_REPORT_YEARS) if FIXED_REPORT_YEARS else 60
    payment_years = FIXED_PAYMENT_YEARS
    results_no_withdrawal = calculate_csv(annual_premium, payment_years, max_calc_year)
    config_A, scenario_key_A = None, "withdrawal_A"
    scenario_name_A = "提取方案 A"
    if withdrawal_a_start > 0 and withdrawal_a_amount > 0:
         config_A = {'start_year': withdrawal_a_start, 'amount': withdrawal_a_amount}
    results_withdrawal_A = calculate_csv(annual_premium, payment_years, max_calc_year, config_A)
    config_B, scenario_key_B = None, "withdrawal_B"
    scenario_name_B = "提取方案 B"
    if withdrawal_b_start > 0 and withdrawal_b_amount > 0:
        config_B = {'start_year': withdrawal_b_start, 'amount': withdrawal_b_amount}
    results_withdrawal_B = calculate_csv(annual_premium, payment_years, max_calc_year, config_B)
    parameters = {
        'client_name': client_name, 'premium': annual_premium,
        'years': payment_years, 'report_years': FIXED_REPORT_YEARS,
        'withdrawal_a_start': withdrawal_a_start, 'withdrawal_a_amount': withdrawal_a_amount,
        'withdrawal_b_start': withdrawal_b_start, 'withdrawal_b_amount': withdrawal_b_amount,
        'calculation_date': datetime.date.today().isoformat()
    }
    final_results = {"parameters": parameters, "無提取": results_no_withdrawal}
    if config_A: final_results[scenario_name_A] = results_withdrawal_A
    if config_B: final_results[scenario_name_B] = results_withdrawal_B
    return final_results


# --- Streamlit App UI ---
st.set_page_config(layout="wide", page_title="保險計劃生成器")
st.title("保誠保險計劃生成器")

# --- Placeholders for Top-Right Buttons ---
top_col1, top_col2, top_col3 = st.columns([0.7, 0.15, 0.15])

st.markdown(f"請在側邊欄輸入參數並點擊「計算計劃預測」。**供款年期固定為 {FIXED_PAYMENT_YEARS} 年。**")

# --- Initialize Session State ---
# (Session state initialization remains the same)
default_inputs = {
    'client_name': '尊貴客戶', 'premium': 10000,
    'w_a_amount': 0, 'w_a_start': 0,
    'w_b_amount': 0, 'w_b_start': 0,
}
for key, default_value in default_inputs.items():
    if key not in st.session_state:
        st.session_state[key] = default_value
if 'calculated_data' not in st.session_state:
    st.session_state['calculated_data'] = None
if 'pdf_bytes' not in st.session_state:
    st.session_state['pdf_bytes'] = None
if 'excel_bytes' not in st.session_state:
    st.session_state['excel_bytes'] = None


# --- Input Form (Sidebar) ---
# (Sidebar code remains the same)
st.sidebar.header("計劃參數")
st.session_state.client_name = st.sidebar.text_input(
    "客戶名稱", value=st.session_state.client_name
)
st.session_state.premium = st.sidebar.number_input(
    "年繳保費 (美元)", min_value=1000, value=st.session_state.premium, step=1000
)
st.sidebar.metric(label="供款年期 (年)", value=f"{FIXED_PAYMENT_YEARS} (固定)")
st.sidebar.subheader("提取方案 (可選)")
with st.sidebar.expander("設定提取方案 A"):
     st.session_state.w_a_amount = st.number_input("每年提取金額 A (美元)", min_value=0, value=st.session_state.w_a_amount, step=100, key="w_a_amount_inp")
     st.session_state.w_a_start = st.number_input("由保單年度開始提取 A:", min_value=0, value=st.session_state.w_a_start, step=1, key="w_a_start_inp")
with st.sidebar.expander("設定提取方案 B"):
     st.session_state.w_b_amount = st.number_input("每年提取金額 B (美元)", min_value=0, value=st.session_state.w_b_amount, step=100, key="w_b_amount_inp")
     st.session_state.w_b_start = st.number_input("由保單年度開始提取 B:", min_value=0, value=st.session_state.w_b_start, step=1, key="w_b_start_inp")
st.sidebar.caption(f"報告將顯示第 {min(FIXED_REPORT_YEARS)} 年至第 {max(FIXED_REPORT_YEARS)} 年。")
st.sidebar.divider()
calculate_button = st.sidebar.button("計算計劃預測", type="primary", use_container_width=True)
status_placeholder = st.sidebar.empty()


# --- Main Area: Calculation Trigger ---
# (Calculation trigger logic remains the same)
if calculate_button:
    valid_input = True
    if st.session_state.w_a_start > 0 and st.session_state.w_a_amount <= 0:
        status_placeholder.error("⚠️ 如設定了提取方案 A 的開始年份，請輸入提取金額 A。")
        valid_input = False
    if st.session_state.w_b_start > 0 and st.session_state.w_b_amount <= 0:
        status_placeholder.error("⚠️ 如設定了提取方案 B 的開始年份，請輸入提取金額 B。")
        valid_input = False

    if valid_input:
        status_placeholder.info("⏳ 計算中... 請稍候。")
        st.warning(f"計算警告：正在使用簡化的模擬邏輯 (0% 利息, {FIXED_PAYMENT_YEARS} 年供款)。稍後請替換為實際的保誠計算邏輯。")
        try:
            st.session_state['calculated_data'] = None
            st.session_state['pdf_bytes'] = None
            st.session_state['excel_bytes'] = None
            calculated_data_result = generate_scenarios(
                st.session_state.client_name, st.session_state.premium,
                st.session_state.w_a_start, st.session_state.w_a_amount,
                st.session_state.w_b_start, st.session_state.w_b_amount
            )
            st.session_state['calculated_data'] = calculated_data_result
            params = calculated_data_result.get('parameters', {})
            client_name_for_file = params.get('client_name','Plan')
            premium_for_file = params.get('premium',0)
            years_for_file = params.get('years',0)
            date_for_file = params.get('calculation_date','today')
            safe_client_name = "".join(c for c in client_name_for_file if c.isalnum() or c in (' ', '_')).rstrip() or 'Plan'
            pdf_filename = f"保險計劃_{safe_client_name}_{premium_for_file}_{years_for_file}年_{date_for_file}.pdf"
            excel_filename = f"保險計劃_{safe_client_name}_{premium_for_file}_{years_for_file}年_{date_for_file}.xlsx"

            with status_placeholder:
                st.info("⏳ 正在生成 PDF 報告...")
                pdf_bytes_result = report_utils.create_plan_pdf(calculated_data_result, EXCEL_TEMPLATE_PATH, STATIC_PDF_PATH)
                st.session_state['pdf_bytes'] = pdf_bytes_result
            with status_placeholder:
                st.info("⏳ 正在生成 Excel 報告...")
                excel_bytes_result = report_utils.create_plan_excel(calculated_data_result, EXCEL_TEMPLATE_PATH)
                st.session_state['excel_bytes'] = excel_bytes_result

            if st.session_state['pdf_bytes'] and st.session_state['excel_bytes']:
                status_placeholder.success("✅ 計算及報告生成完成！")
            elif st.session_state['pdf_bytes']:
                 status_placeholder.warning("⚠️ 計算完成。Excel 生成失敗。")
            elif st.session_state['excel_bytes']:
                status_placeholder.warning("⚠️ 計算完成。PDF 生成失敗。")
            else:
                 status_placeholder.error("❌ 計算完成，但 PDF 及 Excel 生成均失敗。")
        except Exception as calc_err:
            status_placeholder.error(f"❌ 計算/報告錯誤: {calc_err}")
            st.session_state['calculated_data'] = None
            st.session_state['pdf_bytes'] = None
            st.session_state['excel_bytes'] = None
            import traceback
            st.error("詳細錯誤追蹤:"); st.error(traceback.format_exc())


# --- Display PDF and Download Buttons (if data and reports exist) ---
if st.session_state.get('calculated_data'):
    params = st.session_state['calculated_data'].get('parameters', {})

    # --- Download Buttons (Top Right) ---
    # (Download button logic remains the same)
    if st.session_state.get('pdf_bytes'):
         with top_col2:
             client_name_for_file = params.get('client_name','Plan')
             premium_for_file = params.get('premium',0)
             years_for_file = params.get('years',0)
             date_for_file = params.get('calculation_date','today')
             safe_client_name = "".join(c for c in client_name_for_file if c.isalnum() or c in (' ', '_')).rstrip() or 'Plan'
             pdf_filename = f"保險計劃_{safe_client_name}_{premium_for_file}_{years_for_file}年_{date_for_file}.pdf"
             st.download_button(
                     label="下載 PDF", data=st.session_state['pdf_bytes'], file_name=pdf_filename,
                     mime="application/pdf", key="pdf_download_button_top", use_container_width=True
             )
    if st.session_state.get('excel_bytes'):
         with top_col3:
             client_name_for_file = params.get('client_name','Plan')
             premium_for_file = params.get('premium',0)
             years_for_file = params.get('years',0)
             date_for_file = params.get('calculation_date','today')
             safe_client_name = "".join(c for c in client_name_for_file if c.isalnum() or c in (' ', '_')).rstrip() or 'Plan'
             excel_filename = f"保險計劃_{safe_client_name}_{premium_for_file}_{years_for_file}年_{date_for_file}.xlsx"
             st.download_button(
                 label="下載 Excel", data=st.session_state['excel_bytes'], file_name=excel_filename,
                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                 key="excel_download_button_top", use_container_width=True
             )

    # --- Embed PDF in Main Area using st.markdown ---
    if st.session_state.get('pdf_bytes'):
        try:
            base64_pdf = base64.b64encode(st.session_state['pdf_bytes']).decode('utf-8')
            # Embedding using HTML iframe within st.markdown
            # NOTE: Using a fixed height here. Adjust as needed.
            # Consider adding width="100%" or specific pixel width if needed.
            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" height="800px" width="100%" style="border: none;"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True) # Changed from components.html
        except Exception as e:
            st.error(f"顯示 PDF 時出錯：{e}")
    else:
        st.warning("已完成計算，但無法載入 PDF 預覽。請嘗試下載 PDF。")

elif calculate_button and not st.session_state.get('calculated_data'):
    if not status_placeholder. BpSuccess and not status_placeholder. BpError and not status_placeholder. BpWarning:
         status_placeholder.warning("計算可能失敗或未產生結果。")

# --- Footer ---
st.divider()
st.caption(f"保險計劃生成器 v0.7 - {datetime.date.today().year}") # Incremented version
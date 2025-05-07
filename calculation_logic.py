# calculation_logic.py
# Handles interaction with the Excel calculation toolbox using openpyxl and LibreOffice
# Updated read_results_from_xlsx for specific cell locations.
# ADDED MORE DEBUG PRINT STATEMENTS
# ADDED Formula reading and explicit cell checks

import streamlit as st # For displaying messages/errors during calculation
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import os
import shutil
import subprocess
import platform
import time
import traceback # Import traceback for detailed error printing

# --- Configuration ---
CALCULATOR_SHEET_NAME = "TRST"
YEAR_TO_CELL_MAP = {
    10: 'G74', 15: 'G79', 20: 'G84', 25: 'G89', 30: 'G94',
    35: 'G99', 40: 'G104', 45: 'G109', 50: 'G114', 55: 'G119',
    60: 'G124', 70: 'G134', 80: 'G144', 90: 'G154'
}

# --- Helper Function to Find LibreOffice ---
# (Keep the previous version with debug prints)
def find_soffice_path():
    """Attempts to find the path to the LibreOffice executable."""
    print("--- DEBUG (calc_logic): Entering find_soffice_path ---") # DEBUG
    path = None
    try:
        if platform.system() == "Windows":
            print("--- DEBUG (calc_logic): Checking Windows paths... ---") # DEBUG
            prog_files = os.environ.get("ProgramFiles", "C:\\Program Files")
            prog_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
            paths_to_check = [
                os.path.join(prog_files, "LibreOffice", "program", "soffice.exe"),
                os.path.join(prog_files_x86, "LibreOffice", "program", "soffice.exe"),
            ]
            for p in paths_to_check:
                # print(f"--- DEBUG (calc_logic): Checking path: {p} ---") # DEBUG (Optional)
                if os.path.exists(p):
                    print(f"--- DEBUG (calc_logic): Found at: {p} ---") # DEBUG
                    path = p
                    break
            if not path:
                print("--- DEBUG (calc_logic): Checking 'where' command... ---") # DEBUG
                try:
                    # Use shell=True cautiously, ensure command is safe
                    result = subprocess.run(["where", "soffice.exe"], capture_output=True, text=True, check=True, shell=True)
                    paths = result.stdout.strip().splitlines()
                    if paths and os.path.exists(paths[0]):
                         print(f"--- DEBUG (calc_logic): Found via 'where': {paths[0]} ---") # DEBUG
                         path = paths[0]
                except (FileNotFoundError, subprocess.CalledProcessError) as where_err:
                    print(f"--- DEBUG (calc_logic): 'where' command failed or not found: {where_err} ---") # DEBUG
                    pass # Not found via 'where'
        elif platform.system() == "Darwin": # macOS
             # (Keep macOS logic)
             pass
        else: # Linux/Other
             # (Keep Linux logic)
             pass
    except Exception as e:
        print(f"--- ERROR (calc_logic): Exception in find_soffice_path: {e} ---") # DEBUG
        print(traceback.format_exc())
        path = None

    print(f"--- DEBUG (calc_logic): Exiting find_soffice_path. Found path: {path} ---") # DEBUG
    return path


# --- UPDATED Function to Read Results ---
def read_results_from_xlsx(calculated_xlsx_path, report_years):
    """
    Reads the calculated results from the specified XLSX file for the required years.
    Uses the YEAR_TO_CELL_MAP to read from specific cells.
    Also reads the formula from G74 and checks C37/AH74 directly.
    """
    results = []
    print(f"--- DEBUG (calc_logic): Entering read_results_from_xlsx for: {calculated_xlsx_path} ---") # DEBUG
    st.info(f"Attempting to read results from specific cells in: {calculated_xlsx_path}")
    if not os.path.exists(calculated_xlsx_path):
        print(f"--- ERROR (calc_logic): Calculated file not found: {calculated_xlsx_path} ---") # DEBUG
        st.error(f"Calculated file not found: {calculated_xlsx_path}")
        return None

    workbook = None # Initialize workbook to None
    try:
        # --- Load workbook WITH formulas first to check G74 ---
        print(f"--- DEBUG (calc_logic): Loading workbook {calculated_xlsx_path} (data_only=False to read formula)... ---") # DEBUG
        workbook_formulas = openpyxl.load_workbook(calculated_xlsx_path, data_only=False)
        if CALCULATOR_SHEET_NAME not in workbook_formulas.sheetnames:
             print(f"--- ERROR (calc_logic): Sheet '{CALCULATOR_SHEET_NAME}' not found (formula check). ---") # DEBUG
             workbook_formulas.close()
             # Fallback or error? For now, let's try loading with data_only=True
        else:
            sheet_formulas = workbook_formulas[CALCULATOR_SHEET_NAME]
            try:
                g74_formula = sheet_formulas['G74'].value # If it's a formula, .value gives the string
                print(f"--- DEBUG (calc_logic): Formula read from G74: {g74_formula} (Type: {type(g74_formula)}) ---") # DEBUG
                # Check if it's actually a formula type
                if sheet_formulas['G74'].data_type == 'f':
                     print("--- DEBUG (calc_logic): G74 data type is 'f' (formula). ---")
                else:
                     print(f"--- DEBUG (calc_logic): G74 data type is '{sheet_formulas['G74'].data_type}'. ---")

                # --- Explicitly check C37 and AH74 ---
                c37_val_f = sheet_formulas['C37'].value
                c37_dt_f = sheet_formulas['C37'].data_type
                print(f"--- DEBUG (calc_logic): Read from C37 (formulas): Value={c37_val_f}, Type={c37_dt_f} ---")
                ah74_val_f = sheet_formulas['AH74'].value
                ah74_dt_f = sheet_formulas['AH74'].data_type
                print(f"--- DEBUG (calc_logic): Read from AH74 (formulas): Value={ah74_val_f}, Type={ah74_dt_f} ---")

            except Exception as formula_read_err:
                print(f"--- WARNING (calc_logic): Could not read formula/cells G74/C37/AH74: {formula_read_err} ---") # DEBUG
            workbook_formulas.close() # Close the formula workbook

        # --- Now load workbook with data_only=True to get calculated values ---
        print(f"--- DEBUG (calc_logic): Loading workbook {calculated_xlsx_path} (data_only=True)... ---") # DEBUG
        workbook = openpyxl.load_workbook(calculated_xlsx_path, data_only=True)
        if CALCULATOR_SHEET_NAME not in workbook.sheetnames:
            print(f"--- ERROR (calc_logic): Result Sheet '{CALCULATOR_SHEET_NAME}' not found (data read). ---") # DEBUG
            workbook.close()
            return None
        sheet = workbook[CALCULATOR_SHEET_NAME]
        print(f"--- DEBUG (calc_logic): Successfully opened sheet '{CALCULATOR_SHEET_NAME}' for data reading. ---") # DEBUG
        st.info(f"Successfully opened sheet '{CALCULATOR_SHEET_NAME}'.")

        # --- Check C37 and AH74 values again with data_only=True ---
        try:
             c37_val_d = sheet['C37'].value
             print(f"--- DEBUG (calc_logic): Read from C37 (data_only): Value={c37_val_d} ---")
             ah74_val_d = sheet['AH74'].value
             print(f"--- DEBUG (calc_logic): Read from AH74 (data_only): Value={ah74_val_d} ---")
        except Exception as data_read_err:
             print(f"--- WARNING (calc_logic): Could not read cells C37/AH74 (data_only): {data_read_err} ---") # DEBUG


        # --- Read results from mapped cells ---
        for year in report_years:
            if year in YEAR_TO_CELL_MAP:
                cell_ref = YEAR_TO_CELL_MAP[year]
                try:
                    cell_value = sheet[cell_ref].value
                    total_csv = float(cell_value) if cell_value is not None else 0.0
                    print(f"--- DEBUG (calc_logic): Read year {year} from {cell_ref}: Value={cell_value}, Float={total_csv} ---") # DEBUG (Reduced frequency)
                except KeyError:
                     print(f"--- WARNING (calc_logic): Cell reference '{cell_ref}' for year {year} not found. ---") # DEBUG
                     st.warning(f"Cell reference '{cell_ref}' for year {year} not found in the sheet. Using 0.0.")
                     total_csv = 0.0
                except (ValueError, TypeError):
                    print(f"--- WARNING (calc_logic): Could not convert value '{cell_value}' from {cell_ref} (year {year}) to float. ---") # DEBUG
                    st.warning(f"Could not convert value '{cell_value}' from cell {cell_ref} (year {year}) to float. Using 0.0.")
                    total_csv = 0.0
                results.append({'year': year, 'total_csv': round(total_csv, 2)})
            else:
                print(f"--- WARNING (calc_logic): Report year {year} not found in YEAR_TO_CELL_MAP. ---") # DEBUG
                st.warning(f"Report year {year} not found in YEAR_TO_CELL_MAP. Appending with 0.0.")
                results.append({'year': year, 'total_csv': 0.0})

        workbook.close()
        print(f"--- DEBUG (calc_logic): Successfully read results for {len(results)} years. ---") # DEBUG
        st.info(f"Successfully read results for {len(results)} years from specific cells.")
        return results

    except Exception as e:
        print(f"--- ERROR (calc_logic): Error reading results: {e} ---") # DEBUG
        print(traceback.format_exc()) # Print detailed traceback to console
        st.error(f"Error reading results from {calculated_xlsx_path}: {e}")
        st.error(traceback.format_exc())
        if 'workbook_formulas' in locals() and workbook_formulas: workbook_formulas.close()
        if workbook: workbook.close()
        return None


# --- Main Function to Run a Calculation Scenario ---
# (This function remains the same as the previous version - Intermediate ODS)
def run_calculation_scenario(scenario_name, base_xlsx_path, temp_dir, scenario_params, cell_map, report_years):
    print(f"\n--- DEBUG (calc_logic): === ENTERING run_calculation_scenario for: {scenario_name} === ---")
    st.info(f"--- Processing Scenario: {scenario_name} ---")
    safe_scenario_name = "".join(c for c in scenario_name if c.isalnum() or c in (' ', '_')).rstrip().lower().replace(' ', '_')
    input_temp_xlsx_path = os.path.join(temp_dir, f"input_{safe_scenario_name}.xlsx")
    intermediate_ods_path = os.path.join(temp_dir, f"intermediate_{safe_scenario_name}.ods")
    calculated_temp_xlsx_path = os.path.join(temp_dir, f"calculated_{safe_scenario_name}.xlsx")
    libre_output_dir = temp_dir

    # --- 1. Copy and Write Inputs to XLSX ---
    workbook = None # Initialize workbook to None
    try:
        print(f"--- DEBUG (calc_logic): Step 1: Copying and Writing Inputs for {scenario_name} ---") # DEBUG
        if not os.path.exists(base_xlsx_path):
             print(f"--- ERROR (calc_logic): Base calculator file not found: {base_xlsx_path} ---") # DEBUG
             st.error(f"Base calculator file not found: {base_xlsx_path}")
             return None
        print(f"--- DEBUG (calc_logic): Copying template '{os.path.basename(base_xlsx_path)}' to '{os.path.basename(input_temp_xlsx_path)}' ---") # DEBUG
        st.info(f"Copying template '{os.path.basename(base_xlsx_path)}' to '{os.path.basename(input_temp_xlsx_path)}'")
        shutil.copyfile(base_xlsx_path, input_temp_xlsx_path)

        print("--- DEBUG (calc_logic): Writing inputs to spreadsheet... ---") # DEBUG (Changed message)
        st.info("Writing inputs...")
        workbook = openpyxl.load_workbook(input_temp_xlsx_path)
        if CALCULATOR_SHEET_NAME not in workbook.sheetnames:
            print(f"--- ERROR (calc_logic): Sheet '{CALCULATOR_SHEET_NAME}' not found. ---") # DEBUG
            st.error(f"Sheet '{CALCULATOR_SHEET_NAME}' not found in {input_temp_xlsx_path}")
            workbook.close()
            return None
        sheet = workbook[CALCULATOR_SHEET_NAME]

        for param_name, cell_ref in cell_map.items():
            if param_name in scenario_params:
                try:
                    value_to_write = scenario_params[param_name]
                    if isinstance(value_to_write, (int, float)):
                        sheet[cell_ref].value = value_to_write
                    else:
                        sheet[cell_ref] = str(value_to_write)
                except Exception as write_err:
                     print(f"--- WARNING (calc_logic): Could not write '{param_name}' to cell {cell_ref}: {write_err} ---") # DEBUG
                     st.warning(f"Could not write '{param_name}' to cell {cell_ref}: {write_err}")

        workbook.save(input_temp_xlsx_path)
        workbook.close()
        workbook = None # Ensure it's closed before LibreOffice tries to access it
        print("--- DEBUG (calc_logic): Inputs written successfully. ---") # DEBUG
        st.info("Inputs written successfully.")

    except Exception as e:
        print(f"--- ERROR (calc_logic): Error preparing input file: {e} ---") # DEBUG
        print(traceback.format_exc()) # Print detailed traceback to console
        st.error(f"Error preparing input file for {scenario_name}: {e}")
        st.error(traceback.format_exc())
        if workbook: workbook.close()
        return None

    # --- 2. Trigger Calculation with LibreOffice (Two Steps: XLSX -> ODS -> XLSX) ---
    print(f"--- DEBUG (calc_logic): Step 2: Triggering Calculation for {scenario_name} ---") # DEBUG
    soffice_path = find_soffice_path()
    if not soffice_path:
        print("--- ERROR (calc_logic): LibreOffice 'soffice' command not found. Cannot proceed. ---") # DEBUG
        st.error("LibreOffice 'soffice' command not found. Cannot trigger calculation. Please ensure LibreOffice is installed and in the system PATH or standard location.")
        return None

    if os.path.exists(intermediate_ods_path):
        try: os.remove(intermediate_ods_path)
        except OSError as e: print(f"Warning: Could not remove old ODS file: {e}")
    if os.path.exists(calculated_temp_xlsx_path):
        try: os.remove(calculated_temp_xlsx_path)
        except OSError as e: print(f"Warning: Could not remove old XLSX file: {e}")

    # --- Step 2a: Convert XLSX to ODS (Calculates and Saves as ODS) ---
    ods_conversion_successful = False
    soffice_command_ods = [ soffice_path, "--headless", "--invisible", "--nologo", "--convert-to", "ods", "--outdir", libre_output_dir, input_temp_xlsx_path ]
    try:
        print(f"--- DEBUG (calc_logic): Running LibreOffice command (XLSX -> ODS): {' '.join(soffice_command_ods)} ---") # DEBUG
        st.info(f"Running LibreOffice (Step 1/2: XLSX -> ODS): {' '.join(soffice_command_ods)}")
        timeout_seconds = 90
        result_ods = subprocess.run(soffice_command_ods, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=timeout_seconds, check=False)
        print(f"--- DEBUG (calc_logic): LibreOffice (XLSX->ODS) finished. Return Code: {result_ods.returncode} ---") # DEBUG
        print(f"--- DEBUG (calc_logic): LibreOffice (XLSX->ODS) stdout:\n{result_ods.stdout or 'None'}\n---") # DEBUG
        print(f"--- DEBUG (calc_logic): LibreOffice (XLSX->ODS) stderr:\n{result_ods.stderr or 'None'}\n---") # DEBUG

        expected_ods_filename = os.path.splitext(os.path.basename(input_temp_xlsx_path))[0] + ".ods"
        expected_ods_path = os.path.join(libre_output_dir, expected_ods_filename)
        print(f"--- DEBUG (calc_logic): Expecting intermediate ODS file at: {expected_ods_path} ---") # DEBUG
        time.sleep(1) # Filesystem delay

        if result_ods.returncode == 0 and os.path.exists(expected_ods_path):
            print(f"--- DEBUG (calc_logic): Intermediate ODS conversion successful. File at: {expected_ods_path} ---") # DEBUG
            os.rename(expected_ods_path, intermediate_ods_path)
            print(f"--- DEBUG (calc_logic): Renamed intermediate file to: {intermediate_ods_path} ---") # DEBUG
            ods_conversion_successful = True
        else:
            st.error(f"LibreOffice failed during XLSX -> ODS conversion (RC={result_ods.returncode}). Cannot proceed.")
            st.error(f"Stderr: {result_ods.stderr or 'None'}")
            return None

    except Exception as e:
        print(f"--- ERROR (calc_logic): Error during XLSX -> ODS conversion: {e} ---") # DEBUG
        print(traceback.format_exc())
        st.error(f"Error during LibreOffice XLSX->ODS execution for {scenario_name}: {e}")
        st.error(traceback.format_exc())
        return None

    # --- Step 2b: Convert ODS back to XLSX (Should be a cleaner save) ---
    if not ods_conversion_successful: return None

    xlsx_conversion_successful = False
    soffice_command_xlsx = [ soffice_path, "--headless", "--invisible", "--nologo", "--convert-to", "xlsx", "--outdir", libre_output_dir, intermediate_ods_path ]
    try:
        print(f"--- DEBUG (calc_logic): Running LibreOffice command (ODS -> XLSX): {' '.join(soffice_command_xlsx)} ---") # DEBUG
        st.info(f"Running LibreOffice (Step 2/2: ODS -> XLSX): {' '.join(soffice_command_xlsx)}")
        timeout_seconds = 60
        result_xlsx = subprocess.run(soffice_command_xlsx, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=timeout_seconds, check=False)
        print(f"--- DEBUG (calc_logic): LibreOffice (ODS->XLSX) finished. Return Code: {result_xlsx.returncode} ---") # DEBUG
        print(f"--- DEBUG (calc_logic): LibreOffice (ODS->XLSX) stdout:\n{result_xlsx.stdout or 'None'}\n---") # DEBUG
        print(f"--- DEBUG (calc_logic): LibreOffice (ODS->XLSX) stderr:\n{result_xlsx.stderr or 'None'}\n---") # DEBUG

        expected_xlsx_filename = os.path.splitext(os.path.basename(intermediate_ods_path))[0] + ".xlsx"
        expected_xlsx_path = os.path.join(libre_output_dir, expected_xlsx_filename)
        print(f"--- DEBUG (calc_logic): Expecting final XLSX file at: {expected_xlsx_path} ---") # DEBUG
        time.sleep(1) # Filesystem delay

        if result_xlsx.returncode == 0 and os.path.exists(expected_xlsx_path):
            print(f"--- DEBUG (calc_logic): Final XLSX conversion successful. File at: {expected_xlsx_path} ---") # DEBUG
            os.rename(expected_xlsx_path, calculated_temp_xlsx_path)
            print(f"--- DEBUG (calc_logic): Renamed final calculated file to: {calculated_temp_xlsx_path} ---") # DEBUG
            xlsx_conversion_successful = True
        else:
            st.error(f"LibreOffice failed during ODS -> XLSX conversion (RC={result_xlsx.returncode}). Cannot proceed.")
            st.error(f"Stderr: {result_xlsx.stderr or 'None'}")
            return None

    except Exception as e:
        print(f"--- ERROR (calc_logic): Error during ODS -> XLSX conversion: {e} ---") # DEBUG
        print(traceback.format_exc())
        st.error(f"Error during LibreOffice ODS->XLSX execution for {scenario_name}: {e}")
        st.error(traceback.format_exc())
        return None

    try:
        if os.path.exists(intermediate_ods_path):
            os.remove(intermediate_ods_path)
            print(f"--- DEBUG (calc_logic): Removed intermediate ODS file: {intermediate_ods_path} ---") # DEBUG
    except Exception as e:
        print(f"--- WARNING (calc_logic): Could not remove intermediate ODS file: {e} ---") # DEBUG

    # --- 3. Read Results ---
    if not xlsx_conversion_successful:
        print(f"--- ERROR (calc_logic): Final XLSX conversion failed, cannot read results. ---") # DEBUG
        return None

    print(f"--- DEBUG (calc_logic): Step 3: Reading Results for {scenario_name} ---") # DEBUG
    print(f"--- DEBUG (calc_logic): Reading results from final calculated file: {calculated_temp_xlsx_path} ---") # DEBUG
    st.info(f"Reading results from calculated file: {calculated_temp_xlsx_path}")
    final_results_list = read_results_from_xlsx(calculated_temp_xlsx_path, report_years)

    print(f"--- DEBUG (calc_logic): === EXITING run_calculation_scenario for: {scenario_name}. Returning results: {'Success' if final_results_list else 'Failure/None'} === ---") # DEBUG
    return final_results_list

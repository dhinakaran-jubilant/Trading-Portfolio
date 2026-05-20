from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
import io
from datetime import datetime
import numpy as np
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

app = Flask(__name__)

# Hardcoded credentials for simplicity
HARDCODED_USERS = {
    "admin": {
        "password_hash": generate_password_hash("Admin@123"),
        "name": "System Admin",
        "role": "admin"
    }
}

# Enable CORS for React frontend
CORS(app, resources={r"/*": {"origins": "*"}})

def clean_numeric(val):
    if pd.isna(val) or val == '':
        return 0.0
    s = str(val).strip()
    # Handle parentheses for negative numbers: (10.22) -> -10.22
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    # Remove commas, spaces, and plus signs
    s = s.replace(',', '').replace(' ', '').replace('+', '')
    try:
        return float(s)
    except:
        return 0.0

@app.route('/process-excel', methods=['POST'])
def process_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    try:
        # Read input file
        df = None
        try:
            file.seek(0)
            if file.filename.endswith('.xls'):
                # Using the specific method requested by the user for .xls files
                df = pd.read_csv(file, sep="\t", engine='python', on_bad_lines='skip')
            elif file.filename.endswith('.xlsx'):
                df = pd.read_excel(file, engine='openpyxl')
            else:
                # General fallback
                file.seek(0)
                df = pd.read_csv(file, sep=None, engine='python', on_bad_lines='skip')
        except Exception as e:
            # Final attempt if extension-based logic fails
            try:
                file.seek(0)
                df = pd.read_excel(file, engine='openpyxl')
            except Exception as e_xl:
                try:
                    file.seek(0)
                    df = pd.read_csv(file, sep="\t", engine='python', on_bad_lines='skip')
                except Exception as e2:
                    return jsonify({"error": f"Could not read file: {str(e2)}"}), 400

        if df is None or df.empty:
            return jsonify({"error": "File is empty or could not be parsed"}), 400
        
        # Clean column names (strip whitespace)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Auto-promote headers if they are located in the first few rows (density-based matching)
        candidates = [list(df.columns)]
        for idx in range(min(10, len(df))):
            candidates.append(list(df.iloc[idx]))
            
        header_keys = {
            'symbol', 'qty', 'quantity', 'avg price', 'average price', 'ltp', 'cmp', 
            'investment value', 'current value', 'unrealised p&l', 'unrealised p&l %', 
            'unrealized profit', 'company name', 'value at cost', 'value at market'
        }
        
        best_match_idx = -1
        max_matches = 0
        
        for c_idx, candidate in enumerate(candidates):
            matches = 0
            for val in candidate:
                val_str = str(val).strip().lower()
                if any(k in val_str for k in header_keys):
                    matches += 1
            if matches > max_matches:
                max_matches = matches
                best_match_idx = c_idx
                
        if best_match_idx > 0 and max_matches >= 3:
            actual_row_idx = best_match_idx - 1
            df.columns = [str(val).strip() for val in df.iloc[actual_row_idx].values]
            df = df.iloc[actual_row_idx + 1:].reset_index(drop=True)
            df.columns = [str(c).strip() for c in df.columns]
        
        # Qty value-shifting detection (HDFC anomaly)
        for col_name in df.columns:
            if col_name.lower().strip() == 'qty':
                qty_values = df[col_name]
                # Check if it has 0 non-empty values
                non_empty_count = qty_values.dropna().astype(str).str.strip().replace('', np.nan).dropna().count()
                if non_empty_count == 0:
                    # Find index of the column
                    col_idx = list(df.columns).index(col_name)
                    if col_idx + 1 < len(df.columns):
                        # Shift values from next column
                        df[col_name] = df.iloc[:, col_idx + 1]
                break
        
        # Check if the file matches the new column structure
        new_style_keys = {'symbol', 'ltp', 'investment value', 'investment', "today's p&l", "today's p"}
        has_new_style = any(any(k in str(c).lower() for k in new_style_keys) for c in df.columns)
        
        # Mapping as per main.ipynb and user request
        if has_new_style:
            mapping = {
                'SYMBOL': 'SYMBOL',
                'QTY': 'QTY',
                'AVG PRICE': 'AVG PRICE',
                'LTP': 'LTP',
                'INVESTMENT VALUE': 'INVESTED',
                'CURRENT VALUE': 'CURRENT VALUE',
                'UNREALISED P&L': 'UNREALISED P&L',
                'UNREALISED P&L %': 'UNREALISED P&L %',
                "TODAY'S P&L": 'Day P&L',
                "TODAY'S CHG %": 'Day CHG %'
            }
        else:
            mapping = {
                'Company Name': 'Company Name',
                'Qty': 'Qty',
                'Average Cost Price': 'Avg Price',
                'Current Market Price': 'CMP',
                'Value At Cost': 'Invested',
                'Value At Market Price': 'Market Price',
                'Unrealized Profit/Loss': 'Unrealized P&L',
                'Unrealized Profit/Loss %': 'Unrealized P&L %',
                '% Change over prev close': '% Change'
            }
        
        # Filter and rename - ensure we don't crash if columns are missing
        processed_df = pd.DataFrame()
        for src, target in mapping.items():
            if src in df.columns:
                processed_df[target] = df[src]
            else:
                # Try case-insensitive match
                matches = [c for c in df.columns if c.lower() == src.lower()]
                if matches:
                    processed_df[target] = df[matches[0]]
                else:
                    # Robust fallback for partial matches/truncations
                    found = False
                    src_lower = src.lower()
                    for col in df.columns:
                        col_lower = col.lower()
                        if target in ['Company Name', 'SYMBOL'] and 'symbol' in col_lower:
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['Qty', 'QTY'] and col_lower.startswith('qty'):
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['Avg Price', 'AVG PRICE'] and ('avg' in col_lower or 'average' in col_lower):
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['CMP', 'LTP'] and (col_lower == 'ltp' or 'current market' in col_lower):
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['Invested', 'INVESTED'] and ('investment' in col_lower or 'value at cost' in col_lower):
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['Market Price', 'CURRENT VALUE'] and ('current value' in col_lower or 'market price' in col_lower or col_lower == 'current'):
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['Unrealized P&L', 'UNREALISED P&L'] and 'unreal' in col_lower and '%' not in col_lower and 'pct' not in col_lower:
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['Unrealized P&L %', 'UNREALISED P&L %'] and 'unreal' in col_lower and ('%' in col_lower or 'pct' in col_lower or col_lower.endswith('.1')):
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target == 'Day P&L' and 'today' in col_lower and '%' not in col_lower and 'chg' not in col_lower:
                            processed_df[target] = df[col]
                            found = True
                            break
                        elif target in ['% Change', 'Day CHG %'] and ('change' in col_lower or 'chg' in col_lower or 'today\'s c' in col_lower):
                            processed_df[target] = df[col]
                            found = True
                            break
                    
                    if not found:
                        processed_df[target] = np.nan
        
        # Clean all numeric columns
        if has_new_style:
            numeric_cols = ['QTY', 'AVG PRICE', 'LTP', 'INVESTED', 'CURRENT VALUE', 'UNREALISED P&L', 'UNREALISED P&L %', 'Day P&L', 'Day CHG %']
        else:
            numeric_cols = ['Qty', 'Avg Price', 'CMP', 'Invested', 'Market Price', 'Unrealized P&L', 'Unrealized P&L %', 'Day P&L', '% Change']
            
        for col in numeric_cols:
            if col in processed_df.columns:
                processed_df[col] = processed_df[col].apply(clean_numeric)
        
        if not has_new_style:
            # Calculate Day P&L as per main.ipynb: (% Change * Market Price) / 100
            if '% Change' in processed_df.columns and 'Market Price' in processed_df.columns:
                processed_df['Day P&L'] = (processed_df['% Change'] * processed_df['Market Price']) / 100
            else:
                processed_df['Day P&L'] = 0.0

        # Round values as requested
        if has_new_style:
            # Round to integers (0 decimals)
            round_cols = ['INVESTED', 'CURRENT VALUE', 'UNREALISED P&L', 'UNREALISED P&L %', 'Day P&L']
            for col in round_cols:
                if col in processed_df.columns:
                    processed_df[col] = processed_df[col].round(0)
            
            # Round AVG PRICE and Day CHG % to 2 decimals
            decimals_cols = ['AVG PRICE', 'Day CHG %']
            for col in decimals_cols:
                if col in processed_df.columns:
                    processed_df[col] = processed_df[col].round(2)
        else:
            round_cols = ['Invested', 'Market Price', 'Unrealized P&L', 'Day P&L']
            for col in round_cols:
                if col in processed_df.columns:
                    processed_df[col] = processed_df[col].round(0)
        
        # Sort by Unrealized P&L % descending
        sort_col = 'UNREALISED P&L %' if has_new_style else 'Unrealized P&L %'
        if sort_col in processed_df.columns:
            processed_df = processed_df.sort_values(by=sort_col, ascending=False)
        
        # Reorder columns to match main.ipynb output
        if has_new_style:
            cols = ['SYMBOL', 'QTY', 'AVG PRICE', 'LTP', 'INVESTED', 'CURRENT VALUE', 'UNREALISED P&L', 'UNREALISED P&L %', 'Day P&L', 'Day CHG %']
        else:
            cols = ['Company Name', 'Qty', 'Avg Price', 'CMP', 'Invested', 'Market Price', 'Unrealized P&L', 'Unrealized P&L %', 'Day P&L', '% Change']
        # Ensure all columns exist before reordering
        final_cols = [c for c in cols if c in processed_df.columns]
        processed_df = processed_df[final_cols]
        
        # Calculate Totals
        if has_new_style:
            total_invested = processed_df['INVESTED'].sum() if 'INVESTED' in processed_df.columns else 0
            total_market_value = processed_df['CURRENT VALUE'].sum() if 'CURRENT VALUE' in processed_df.columns else 0
            total_unrealized_pl = processed_df['UNREALISED P&L'].sum() if 'UNREALISED P&L' in processed_df.columns else 0
            total_day_pl = processed_df['Day P&L'].sum() if 'Day P&L' in processed_df.columns else 0
            
            total_row = {
                'SYMBOL': np.nan,
                'QTY': np.nan,
                'AVG PRICE': np.nan,
                'LTP': np.nan,
                'INVESTED': total_invested,
                'CURRENT VALUE': total_market_value,
                'UNREALISED P&L': total_unrealized_pl,
                'UNREALISED P&L %': np.nan,
                'Day P&L': total_day_pl,
                'Day CHG %': np.nan
            }
        else:
            total_invested = processed_df['Invested'].sum() if 'Invested' in processed_df.columns else 0
            total_market_value = processed_df['Market Price'].sum() if 'Market Price' in processed_df.columns else 0
            total_unrealized_pl = processed_df['Unrealized P&L'].sum() if 'Unrealized P&L' in processed_df.columns else 0
            total_day_pl = processed_df['Day P&L'].sum() if 'Day P&L' in processed_df.columns else 0
            
            total_row = {
                'Company Name': 'TOTAL',
                'Qty': np.nan,
                'Avg Price': np.nan,
                'CMP': np.nan,
                'Invested': total_invested,
                'Market Price': total_market_value,
                'Unrealized P&L': total_unrealized_pl,
                'Unrealized P&L %': np.nan,
                'Day P&L': total_day_pl,
                '% Change': np.nan
            }
        
        processed_df = pd.concat([processed_df, pd.DataFrame([total_row])], ignore_index=True)
        
        # Get dynamic title info
        name_val = request.form.get('name', 'Gautham')
        bank_val = request.form.get('bank', 'HDFC')
        
        # Determine Morning/Evening update
        now = datetime.now()
        # Using 13:00 (1 PM) as the cutoff for Morning vs Evening update
        update_status = "Morning" if now.hour < 13 else "Evening"
        current_date = now.strftime("%d.%m.%Y")
        
        # Format title as per user request: ICICI (name) sir as on 13.05.2026 (Morning update)
        title = f"{bank_val} {name_val} sir as on {current_date} ({update_status} Update)"
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=0, startcol=0, sheet_name='Holdings')
            processed_df.to_excel(writer, index=False, startrow=1, sheet_name='Holdings')
            
            worksheet = writer.sheets['Holdings']
            # Auto-fit column widths with padding
            for col in worksheet.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                # Add padding of 2
                worksheet.column_dimensions[column_letter].width = max_length + 2

            # Format Title (Row 1) for both HDFC and standard styles
            title_cell = worksheet['A1']
            # Merge across all columns
            worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(processed_df.columns))
            title_cell.fill = PatternFill(start_color='FF262626', end_color='FF262626', fill_type='solid')
            title_cell.font = Font(color='FFFFFF', bold=True, size=12)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')

            # Center align all headers and values
            center_alignment = Alignment(horizontal='center', vertical='center')
            header_total_fill = PatternFill(start_color='FFD9D9D9', end_color='FFD9D9D9', fill_type='solid')
            bold_font = Font(bold=True)
            
            thin_side = Side(border_style="thin", color="000000")
            thin_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
            
            header_row_idx = 2
            total_row_idx = len(processed_df) + 2
            
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, max_col=len(processed_df.columns)), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    if has_new_style:
                        if row_idx == 2 or row_idx == total_row_idx:
                            cell.border = thin_border
                            cell.alignment = center_alignment
                            cell.fill = header_total_fill
                            cell.font = bold_font
                        elif row_idx > 2:
                            cell.border = thin_border
                            cell.alignment = center_alignment
                    else:
                        cell.border = thin_border
                        if row_idx >= 2:
                            cell.alignment = center_alignment
                            if row_idx == 2 or row_idx == total_row_idx:
                                cell.fill = header_total_fill
                                cell.font = bold_font

            # Set Row Heights
            worksheet.row_dimensions[1].height = 26 # Title is Row 1
            worksheet.row_dimensions[2].height = 38 # Header is Row 2
            for r_idx in range(3, worksheet.max_row + 1):
                worksheet.row_dimensions[r_idx].height = 18

            # Apply formatting
            if has_new_style:
                indian_format = '[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;##,##0'
                integer_format = '#,##0'
                decimal_2_format = '0.00'
                
                # INVESTED, CURRENT VALUE, UNREALISED P&L, Day P&L -> indian_format
                for col_name in ['INVESTED', 'CURRENT VALUE', 'UNREALISED P&L', 'Day P&L']:
                    if col_name in processed_df.columns:
                        c_idx = processed_df.columns.get_loc(col_name) + 1
                        c_letter = get_column_letter(c_idx)
                        for r_idx in range(3, worksheet.max_row + 1): # Data starts at Row 3
                            cell = worksheet[f"{c_letter}{r_idx}"]
                            if cell.value is not None and cell.value != "":
                                cell.number_format = indian_format
                
                # UNREALISED P&L % -> integer_format
                if 'UNREALISED P&L %' in processed_df.columns:
                    c_idx = processed_df.columns.get_loc('UNREALISED P&L %') + 1
                    c_letter = get_column_letter(c_idx)
                    for r_idx in range(3, worksheet.max_row + 1):
                        cell = worksheet[f"{c_letter}{r_idx}"]
                        if cell.value is not None and cell.value != "":
                            cell.number_format = integer_format

                # AVG PRICE, Day CHG % -> decimal_2_format
                for col_name in ['AVG PRICE', 'Day CHG %']:
                    if col_name in processed_df.columns:
                        c_idx = processed_df.columns.get_loc(col_name) + 1
                        c_letter = get_column_letter(c_idx)
                        for r_idx in range(3, worksheet.max_row + 1):
                            cell = worksheet[f"{c_letter}{r_idx}"]
                            if cell.value is not None and cell.value != "":
                                cell.number_format = decimal_2_format
            else:
                indian_format = '[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;##,##0'
                format_cols = ['Invested', 'Market Price', 'Unrealized P&L', 'Day P&L']
                for col_name in format_cols:
                    if col_name in processed_df.columns:
                        c_idx = processed_df.columns.get_loc(col_name) + 1
                        c_letter = get_column_letter(c_idx)
                        for r_idx in range(3, worksheet.max_row + 1):
                            worksheet[f"{c_letter}{r_idx}"].number_format = indian_format

            # Apply Conditional Formatting to 'UNREALISED P&L %' and 'Day P&L'
            cf_cols = ['UNREALISED P&L %'] if has_new_style else ['Unrealized P&L %', 'Day P&L']
            for col_name in cf_cols:
                if col_name in processed_df.columns:
                    col_idx = processed_df.columns.get_loc(col_name) + 1
                    col_letter = get_column_letter(col_idx)
                    start_row = 3
                    end_row = 1 + len(processed_df)
                    if end_row >= start_row:
                        cell_range = f"{col_letter}{start_row}:{col_letter}{end_row}"
                        
                        # Green-Yellow-Red Color Scale
                        rule = ColorScaleRule(
                            start_type='min', start_color='F8696B', # Red for low
                            mid_type='percentile', mid_value=50, mid_color='FFEB84', # Yellow for middle
                            end_type='max', end_color='63BE7B' # Green for high
                        )
                        worksheet.conditional_formatting.add(cell_range, rule)

        output.seek(0)
        
        # Format filename as per user request: name_bank_date.xlsx
        download_filename = f"{name_val}_{bank_val}_{current_date}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=download_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# --- Authentication Routes ---

@app.route('/api/login/', methods=['POST'])
def login():
    try:
        data = request.json
        emp_code = data.get('employee_code', '').strip()
        password = data.get('password', '')
        
        user_data = HARDCODED_USERS.get(emp_code)
        if user_data:
            if check_password_hash(user_data["password_hash"], password):
                return jsonify({
                    'success': True,
                    'user': {
                        'employee_code': emp_code,
                        'name': user_data["name"],
                        'role': user_data["role"],
                        'is_initial_password': False
                    }
                })
        return jsonify({'success': False, 'message': 'Invalid employee code or password'}), 401
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=1501)

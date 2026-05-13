import pandas as pd
import io

def generate_branch_summary(lcc_stream, par_stream):
    df1 = pd.read_excel(lcc_stream)
    df2 = pd.read_excel(par_stream)

    # Clean columns (strip whitespace) BEFORE renaming
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Robust column mapping for LCC (df1)
    df1.rename(columns={
        'Loan Number': 'Loan No',
        'Loan No.': 'Loan No',
        'IRR': 'Interest',
        'Interest Rate': 'Interest'
    }, inplace=True)

    # Ensure Frequency exists and is a single column
    if 'Frequency' not in df1.columns:
        # Check if it's under 'Mode of RePayment' or similar
        for alt in ['Mode of RePayment', 'Repayment Mode', 'Mode']:
            if alt in df1.columns:
                df1.rename(columns={alt: 'Frequency'}, inplace=True)
                break
    
    if 'Frequency' not in df1.columns:
        df1['Frequency'] = 'Monthly'
    
    # Robust column mapping for PAR (df2)
    df2.rename(columns={
        'Loan Number': 'Loan No',
        'Loan No.': 'Loan No',
        'Total DPD': 'DPD',
        'DPD': 'DPD'
    }, inplace=True)

    def clean_currency_raw(series):
        if series is None: return 0
        # Remove currency symbols, commas, and other non-numeric chars (except . and -)
        cleaned = series.astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
        # Replace empty or just '-' with 0
        cleaned = cleaned.replace(['', '-'], '0')
        return pd.to_numeric(cleaned, errors='coerce').fillna(0)

    # Ensure critical columns exist or create them
    if 'Principal Outstanding' in df1.columns:
        df1['Principal Outstanding'] = clean_currency_raw(df1['Principal Outstanding'])
    else:
        # Try to find any column containing 'Principal' or 'Outstanding'
        match = [c for c in df1.columns if 'principal' in c.lower() or 'outstanding' in c.lower()]
        if match:
            df1['Principal Outstanding'] = clean_currency_raw(df1[match[0]])
        else:
            df1['Principal Outstanding'] = 0

    # --- Data Cleaning: Remove Total Rows and empty rows ---
    # Convert Loan No to string for easier filtering
    if 'Loan No' in df1.columns:
        df1['Loan No'] = df1['Loan No'].astype(str).str.strip()
        # Filter out rows that are clearly summary rows
        total_mask = df1['Loan No'].str.upper().str.contains('TOTAL', na=False)
        df1 = df1[~total_mask].copy()
    
    # Filter for active loans with Principal Outstanding strictly greater than 0
    df1 = df1[df1['Principal Outstanding'] > 0].copy()

    # Map DPD from PAR file
    if 'Loan No' in df1.columns and 'Loan No' in df2.columns:
        dpd_col = 'DPD' if 'DPD' in df2.columns else ('Total DPD' if 'Total DPD' in df2.columns else None)
        if dpd_col:
            df1['DPD'] = df1['Loan No'].map(df2.set_index('Loan No')[dpd_col]).fillna(0).astype(int)
        else:
            df1['DPD'] = 0
    else:
        df1['DPD'] = 0

    os_df = df1.copy()

    # Clean remaining numeric columns
    os_df['Loan Amount'] = clean_currency_raw(os_df['Loan Amount'])
    os_df['Installments'] = pd.to_numeric(os_df['Installments'], errors='coerce').fillna(0)
    def clean_interest(val):
        if pd.isna(val): return 0
        if isinstance(val, (int, float)):
            if 0 < val < 1: return val * 100
            return val
        
        s = str(val).strip()
        has_percent = '%' in s
        s = s.replace('%', '').strip()
        num = pd.to_numeric(s, errors='coerce')
        if pd.isna(num): return 0
        
        # If it had a % sign, the number is already the percentage value (e.g., 0.60 for 0.60%)
        # If it didn't have a % sign and is < 1, it's likely a decimal (e.g., 0.15 for 15%)
        if not has_percent and 0 < num < 1:
            return num * 100
        return num

    if 'Interest' in os_df.columns:
        os_df['Interest'] = os_df['Interest'].apply(clean_interest).fillna(0)
    elif 'Cont. IRR (%)' in os_df.columns:
        os_df['Interest'] = os_df['Cont. IRR (%)'].apply(clean_interest).fillna(0)
    else:
        os_df['Interest'] = 0
    
    if 'Frequency' not in os_df.columns:
        os_df['Frequency'] = 'Monthly' # Default
    # Map State from Region
    if 'Region' in os_df.columns:
        os_df['State'] = os_df['Region']
    else:
        os_df['State'] = ''

    # --- Create DPD buckets ---
    os_df['Current'] = os_df.apply(lambda x: x['Principal Outstanding'] if x['DPD'] == 0 else 0, axis=1)
    os_df['8 to 30 days'] = os_df.apply(lambda x: x['Principal Outstanding'] if 1 <= x['DPD'] <= 30 else 0, axis=1)
    os_df['31 to 60 days'] = os_df.apply(lambda x: x['Principal Outstanding'] if 31 <= x['DPD'] <= 60 else 0, axis=1)
    os_df['61 to 90 days'] = os_df.apply(lambda x: x['Principal Outstanding'] if 61 <= x['DPD'] <= 90 else 0, axis=1)
    os_df['91 to 180 days'] = os_df.apply(lambda x: x['Principal Outstanding'] if 91 <= x['DPD'] <= 180 else 0, axis=1)
    os_df['Above 180 days'] = os_df.apply(lambda x: x['Principal Outstanding'] if x['DPD'] > 180 else 0, axis=1)

    bucket_cols = ['Current', '8 to 30 days', '31 to 60 days', '61 to 90 days', '91 to 180 days', 'Above 180 days']

    # --- Dimension Bucketing ---
    # --- Dimension Bucketing ---
    def get_amount_bucket(amt):
        if amt <= 200000: return '1. <= 2 lakhs'
        if amt <= 500000: return '2. 2 lakhs - 5 lakhs'
        if amt <= 1000000: return '3. 5 lakhs - 10 lakhs'
        if amt <= 1500000: return '4. 10 lakhs - 15 lakhs'
        if amt <= 2000000: return '5. 15 lakhs - 20 lakhs'
        if amt <= 2500000: return '6. 20 lakhs - 25 lakhs'
        if amt <= 3500000: return '7. 25 lakhs - 35 lakhs'
        if amt <= 5000000: return '8. 35 lakhs - 50 lakhs'
        if amt <= 7500000: return '9. 50 lakhs - 75 lakhs'
        if amt <= 15000000: return '10. 75 lakhs - 150 lakhs'
        return '11. > 150 lakhs'

    def get_tenor_bucket(t):
        if t <= 12: return '1. upto 12 Month'
        if t <= 18: return '2. 13 to 18 Months'
        if t <= 24: return '3. 19 to 24 Months'
        if t <= 36: return '4. 25 to 36 Months'
        return '5. More than 36 Months'

    def get_interest_bucket(r):
        if r <= 16: return '1. <=16%'
        if r <= 20: return '2. >=16% to 20%'
        if r <= 24: return '3. >=20% to 24%'
        return '4. > 24 %'

    def get_repayment_bucket(m):
        m_str = str(m).strip().lower()
        if 'daily' in m_str or m_str == 'd': return 'Daily'
        if 'fort' in m_str or 'bi' in m_str or m_str == 'f': return 'Fortnightly'
        if 'weekly' in m_str or m_str == 'w': return 'Weekly'
        if '15' in m_str: return '15 Days'
        if 'month' in m_str or m_str == 'm': return 'Monthly'
        return str(m).strip()

    def get_yearly_irr(row):
        try:
            rate = row['Interest']
            mode = str(row['Frequency']).strip()
            if mode == 'Daily': return rate * 365
            if mode == 'Weekly': return rate * 52
            if mode == 'Fortnightly': return rate * 26
            if mode == '15 Days': return rate * 24
            if mode == 'Monthly': return rate * 12
            return rate
        except:
            return 0

    os_df['Frequency'] = os_df['Frequency'].apply(get_repayment_bucket)
    os_df['Yearly IRR'] = os_df.apply(get_yearly_irr, axis=1)
    os_df['Amount Cut'] = os_df['Loan Amount'].apply(get_amount_bucket)
    os_df['Tenor Cut'] = os_df['Installments'].apply(get_tenor_bucket)
    os_df['Interest Cut'] = os_df['Yearly IRR'].apply(get_interest_bucket)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- Formats ---
        formats = {
            'header': workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9EAD3'}),
            'number': workbook.add_format({'num_format': '0.00', 'border': 1}),
            'text': workbook.add_format({'border': 1}),
            'total': workbook.add_format({'bold': True, 'num_format': '0.00', 'border': 2, 'bg_color': '#F4CCCC'}),
            'int': workbook.add_format({'num_format': '0', 'border': 1}),
            'total_int': workbook.add_format({'bold': True, 'num_format': '0', 'border': 2, 'bg_color': '#F4CCCC'})
        }

        # Master lists for fixed rows
        TICKET_SIZE_CATS = [
            '1. <= 2 lakhs', '2. 2 lakhs - 5 lakhs', '3. 5 lakhs - 10 lakhs', '4. 10 lakhs - 15 lakhs',
            '5. 15 lakhs - 20 lakhs', '6. 20 lakhs - 25 lakhs', '7. 25 lakhs - 35 lakhs', '8. 35 lakhs - 50 lakhs',
            '9. 50 lakhs - 75 lakhs', '10. 75 lakhs - 150 lakhs', '11. > 150 lakhs'
        ]
        TENOR_CATS = ['1. upto 12 Month', '2. 13 to 18 Months', '3. 19 to 24 Months', '4. 25 to 36 Months', '5. More than 36 Months']
        IRR_CATS = ['1. <=16%', '2. >=16% to 20%', '3. >=20% to 24%', '4. > 24 %']
        REPAYMENT_CATS = ['Daily', 'Weekly', 'Fortnightly', 'Monthly']

        def write_summary_sheet(df, group_col, sheet_name, extra_cols=None, start_row=0, display_name=None, fixed_categories=None):
            # Aggregation
            agg_map = {
                'Loan No': 'count',
                **{col: 'sum' for col in bucket_cols}
            }
            if extra_cols:
                for ec in extra_cols:
                    if ec in df.columns:
                        agg_map[ec] = 'first'
            
            summary = df.groupby(group_col).agg(agg_map).reset_index()
            
            # Reindex to ensure all fixed categories are present
            if fixed_categories:
                summary = summary.set_index(group_col).reindex(fixed_categories).reset_index()
                summary.fillna(0, inplace=True)
                summary[group_col] = fixed_categories

            summary.rename(columns={'Loan No': 'No of Account'}, inplace=True)
            summary['Total'] = summary[bucket_cols].sum(axis=1)
            
            # Formatting values (to Crores)
            cols_to_format = ['Total'] + bucket_cols
            summary[cols_to_format] = (summary[cols_to_format] / 10**7)
            
            actual_extra = [ec for ec in (extra_cols or []) if ec in summary.columns]
            display_cols = [group_col] + actual_extra + ['No of Account', 'Total'] + bucket_cols
            summary = summary[display_cols]
            
            # Add TOTAL row
            total_row_data = {
                group_col: 'TOTAL',
                'No of Account': summary['No of Account'].sum(),
                **{col: summary[col].sum() for col in cols_to_format}
            }
            for ec in actual_extra: total_row_data[ec] = ''
            total_row = pd.DataFrame([total_row_data])
            final_df = pd.concat([summary, total_row], ignore_index=True)
            
            # Get or create worksheet
            if sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
            else:
                worksheet = workbook.add_worksheet(sheet_name)
                writer.sheets[sheet_name] = worksheet
                worksheet.set_column(0, 0, 30)
                worksheet.set_column(1, 1, 12)
                worksheet.set_column(2, 2, 12)
                worksheet.set_column(3, 15, 10)
            
            # Custom Formats
            header_fmt = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1,
                'text_wrap': True
            })
            data_fmt = workbook.add_format({'border': 1, 'num_format': '0.00'})
            int_fmt = workbook.add_format({'border': 1, 'num_format': '0'})
            text_fmt = workbook.add_format({'border': 1})
            total_fmt = workbook.add_format({'bold': True, 'border': 1, 'num_format': '0.0000'})
            total_int_fmt = workbook.add_format({'bold': True, 'border': 1, 'num_format': '0'})
            center_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            
            # Row heights for wrapped headers
            worksheet.set_row(start_row, 30)
            worksheet.set_row(start_row + 1, 30)

            # Create Merged Headers
            header_title = display_name if display_name else group_col
            worksheet.merge_range(start_row, 0, start_row + 1, 0, header_title, header_fmt)
            col_idx = 1
            for ec in actual_extra:
                worksheet.merge_range(start_row, col_idx, start_row + 1, col_idx, ec, header_fmt)
                col_idx += 1
            
            no_of_acc_col = col_idx
            worksheet.merge_range(start_row, no_of_acc_col, start_row + 1, no_of_acc_col, 'No of Account', header_fmt)
            
            total_start_col = no_of_acc_col + 1
            total_end_col = total_start_col + len(cols_to_format) - 1
            worksheet.merge_range(start_row, total_start_col, start_row, total_end_col, 'Own', header_fmt)
            
            for i, col_name in enumerate(cols_to_format):
                worksheet.write(start_row + 1, total_start_col + i, col_name, header_fmt)
            
            # Write data rows
            for r_idx in range(len(final_df)):
                row_num = r_idx + start_row + 2
                row_values = final_df.iloc[r_idx]
                is_total = row_values[group_col] == 'TOTAL'
                
                for c_idx, col_name in enumerate(final_df.columns):
                    val = row_values[col_name]
                    if col_name == 'No of Account':
                        curr_fmt = total_int_fmt if is_total else int_fmt
                    elif is_total:
                        curr_fmt = total_fmt
                    elif col_name in cols_to_format:
                        curr_fmt = data_fmt
                    else:
                        curr_fmt = text_fmt
                    
                    if not is_total and col_name != group_col and col_name not in actual_extra and val == 0:
                        worksheet.write(row_num, c_idx, '-', center_fmt)
                    else:
                        worksheet.write(row_num, c_idx, val, curr_fmt)
            
            return start_row + 2 + len(final_df) + 1

        # Generate Sheets
        next_row = write_summary_sheet(os_df, 'Branch', 'Portfolio Cuts', extra_cols=['State'], start_row=0)
        next_row = write_summary_sheet(os_df, 'Loan Type', 'Portfolio Cuts', start_row=next_row, display_name='Product')
        next_row = write_summary_sheet(os_df, 'Amount Cut', 'Portfolio Cuts', start_row=next_row, display_name='Ticket size', fixed_categories=TICKET_SIZE_CATS)
        next_row = write_summary_sheet(os_df, 'Tenor Cut', 'Portfolio Cuts', start_row=next_row, display_name='Tenor', fixed_categories=TENOR_CATS)
        next_row = write_summary_sheet(os_df, 'Interest Cut', 'Portfolio Cuts', start_row=next_row, display_name='IRR', fixed_categories=IRR_CATS)
        write_summary_sheet(os_df, 'Frequency', 'Portfolio Cuts', start_row=next_row, display_name='Repayment', fixed_categories=REPAYMENT_CATS)

    output.seek(0)
    return output

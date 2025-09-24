import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import os

st.title("Excel Merger: Quantities & Dates with Product Info")

# --- Load prod_line.xlsx automatically ---
prod_line_path = os.path.join(os.path.dirname(__file__), "prod_line.xlsx")
if not os.path.exists(prod_line_path):
    st.error("prod_line.xlsx not found in app folder!")
    st.stop()

prod_line = pd.read_excel(prod_line_path)
prod_line.columns = prod_line.columns.str.strip()  # normalize

required_prod_cols = ['bar_code','export_code','prod_description','category','price']
for col in required_prod_cols:
    if col not in prod_line.columns:
        st.error(f"Column '{col}' missing in prod_line.xlsx")
        st.stop()

prod_line['bar_code'] = prod_line['bar_code'].astype(str)

# --- Upload multiple Excel files ---
uploaded_files = st.file_uploader(
    "Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    df_list = []
    dates_list = []
    missing_dates_dict = {}  # track missing dates per location

    for uploaded_file in uploaded_files:
        location = uploaded_file.name.split(".")[0]
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()

        if 'კოდი' not in df.columns or 'რაოდენობა' not in df.columns:
            st.error(f"Columns missing in {uploaded_file.name}")
            st.stop()

        if 'თვე' not in df.columns:
            df['თვე'] = None
        if 'წელი' not in df.columns:
            df['წელი'] = None

        # --- Quantities ---
        df_qty = df.groupby('კოდი', as_index=False)['რაოდენობა'].sum().rename(columns={'რაოდენობა': location})
        df_list.append(df_qty)

        # --- Dates ---
        df_dates = df.groupby('კოდი', as_index=False).agg({'თვე':'first', 'წელი':'first'})

        # Track missing dates
        missing_kodes = df_dates[df_dates['თვე'].isna() | df_dates['წელი'].isna()]['კოდი'].astype(str).tolist()
        missing_dates_dict[location] = missing_kodes

        def format_date(row):
            if pd.notnull(row['თვე']) and pd.notnull(row['წელი']):
                return f"{int(row['თვე']):02d}/{int(row['წელი'])}"
            else:
                return ""

        df_dates[location] = df_dates.apply(format_date, axis=1)
        df_dates = df_dates[['კოდი', location]]
        dates_list.append(df_dates)

    # --- Merge Quantities ---
    final_qty = df_list[0]
    for df in df_list[1:]:
        final_qty = pd.merge(final_qty, df, on='კოდი', how='outer')
    final_qty = final_qty.fillna(0)

    # --- Merge Dates ---
    final_dates = dates_list[0]
    for df in dates_list[1:]:
        final_dates = pd.merge(final_dates, df, on='კოდი', how='outer')
    final_dates = final_dates.fillna("")

    # --- Merge with product info ---
    final_qty['კოდი'] = final_qty['კოდი'].astype(str)
    final_dates['კოდი'] = final_dates['კოდი'].astype(str)

    final_qty = pd.merge(final_qty, prod_line, left_on='კოდი', right_on='bar_code', how='left')
    final_dates = pd.merge(final_dates, prod_line, left_on='კოდი', right_on='bar_code', how='left')

    final_qty.drop(columns=['bar_code'], inplace=True)
    final_dates.drop(columns=['bar_code'], inplace=True)

    prod_cols = ['export_code','prod_description','category','price']
    for col in prod_cols:
        if col not in final_qty.columns:
            final_qty[col] = None
        if col not in final_dates.columns:
            final_dates[col] = None

    # --- Reorder columns ---
    qty_cols = ['კოდი'] + prod_cols + [c for c in final_qty.columns if c not in ['კოდი'] + prod_cols]
    final_qty = final_qty[qty_cols]

    date_cols = ['კოდი'] + prod_cols + [c for c in final_dates.columns if c not in ['კოდი'] + prod_cols]
    final_dates = final_dates[date_cols]

    # --- Show tables ---
    st.subheader("Quantities with Product Info")
    st.dataframe(final_qty)

    st.subheader("Matched Dates with Product Info")
    st.dataframe(final_dates)

    # --- Save to Excel in memory ---
    temp_output = BytesIO()
    with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
        final_qty.to_excel(writer, sheet_name='Quantities', index=False)
        final_dates.to_excel(writer, sheet_name='Matched Dates', index=False)

    temp_output.seek(0)
    wb = load_workbook(temp_output)

    # --- Formatting ---
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for col in ws.columns:
            column_letter = col[0].column_letter
            max_length = 0
            for cell in col:
                cell.border = thin_border
                cell.alignment = align_center
                if cell.row == 1:
                    cell.font = header_font
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[column_letter].width = max_length + 2

    # --- Highlight missing dates per location ---
    for sheet_name in ['Quantities', 'Matched Dates']:
        ws = wb[sheet_name]
        for loc, missing_kodes in missing_dates_dict.items():
            # find the column index of this location
            col_idx = None
            for idx, cell in enumerate(ws[1], start=1):
                if cell.value == loc:
                    col_idx = idx
                    break
            if col_idx is None:
                continue
            # loop rows to check missing codes
            for row_idx in range(2, ws.max_row + 1):
                kode = ws.cell(row=row_idx, column=1).value  # column A = 'კოდი'
                if str(kode) in missing_kodes:
                    ws.cell(row=row_idx, column=col_idx).fill = red_fill

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    # --- Download button ---
    st.download_button(
        label="Download Combined Excel with Product Info",
        data=final_output,
        file_name="combined_locations_with_prod_info.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

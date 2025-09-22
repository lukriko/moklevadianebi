import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("Excel Merger: Quantities & Dates")

# --- Upload multiple Excel files ---
uploaded_files = st.file_uploader(
    "Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    df_list = []
    dates_list = []

    for uploaded_file in uploaded_files:
        location = uploaded_file.name.split(".")[0]
        df = pd.read_excel(uploaded_file)

        # Normalize columns
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

        def format_date(row):
            if pd.notnull(row['თვე']) and pd.notnull(row['წელი']):
                return f"{int(row['თვე']):02d}/{int(row['წელი'])}"
            else:
                return ""

        df_dates[location] = df_dates.apply(format_date, axis=1)
        df_dates = df_dates[['კოდი', location]]
        dates_list.append(df_dates)

    # --- Merge ---
    final_qty = df_list[0]
    for df in df_list[1:]:
        final_qty = pd.merge(final_qty, df, on='კოდი', how='outer')
    final_qty = final_qty.fillna(0)

    final_dates = dates_list[0]
    for df in dates_list[1:]:
        final_dates = pd.merge(final_dates, df, on='კოდი', how='outer')
    final_dates = final_dates.fillna("")

    final_qty['კოდი'] = final_qty['კოდი'].astype(str)
    final_dates['კოდი'] = final_dates['კოდი'].astype(str)

    # --- Show tables ---
    st.subheader("Quantities")
    st.dataframe(final_qty)

    st.subheader("Matched Dates")
    st.dataframe(final_dates)

    # --- Save to Excel in memory ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_qty.to_excel(writer, sheet_name='Quantities', index=False)
        final_dates.to_excel(writer, sheet_name='Matched Dates', index=False)

    # Highlight blank cells in Matched Dates
    wb = load_workbook(output)
    ws_dates = wb['Matched Dates']
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for row in range(2, ws_dates.max_row + 1):
        for col in range(2, ws_dates.max_column + 1):
            cell = ws_dates.cell(row=row, column=col)
            if cell.value == "":
                cell.fill = red_fill

    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download Combined Excel",
        data=output,
        file_name="combined_locations.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

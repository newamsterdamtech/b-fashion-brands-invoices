import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(layout="wide")
st.title("Invoice Validation App")

# Aliases for matching variations in header names
COLUMN_ALIASES = {
    'Style No': ['style no', 'article no', 'style number', 'styleno'],
    'PO Nr': ['po nr', 'po no', 'order no', 'order number', 'po number', 'po no.', 'pono'],
    'Quantity': ['qty', 'qty/pc', 'quantity', 'pieces', 'pcs', 'pieces (pcs)', 'pieces   (pcs)'],
    'Price': ['unit price', 'price', 'price/usd', 'unitprice', 'price/usd'],
    'Total': ['amount', 'total', 'amount (usd)', 'total amount', 'amount usd', 'ttl', 'sum']
}
ALL_ALIASES = sum(COLUMN_ALIASES.values(), [])

def clean_col(col):
    if not isinstance(col, str):
        col = str(col)
    return re.sub(r'[^a-z0-9]', '', col.lower())

def looks_like_header(row):
    # If mostly ascii and not company name, and not empty/duplicate, likely a header
    txts = [str(x).strip() for x in row]
    joined = ' '.join(txts).lower()
    if not joined or len(joined) < 10:
        return False
    ascii_ratio = sum(ord(c) < 128 for c in joined) / max(1, len(joined))
    unique_cells = set(txts) - {""}
    if ascii_ratio < 0.5 or len(unique_cells) < 4:
        return False
    if sum("company" in cell.lower() or "ltd" in cell.lower() or "inc" in cell.lower() for cell in txts) > 2:
        return False
    return True

def find_header_row_and_map(df, min_matches=2, max_rows=40):
    for idx, row in enumerate(df.head(max_rows).values):
        if not looks_like_header(row):
            continue
        cleaned = [clean_col(x) for x in row]
        col_map = {}
        for canon, aliases in COLUMN_ALIASES.items():
            for alias in aliases:
                alias_clean = clean_col(alias)
                for j, cell in enumerate(cleaned):
                    if alias_clean == cell or alias_clean in cell or cell in alias_clean:
                        col_map[canon] = j
                        break
                if canon in col_map:
                    break
        if (
            "Style No" in col_map and "PO Nr" in col_map and
            (("Quantity" in col_map) or ("Price" in col_map) or ("Total" in col_map))
        ):
            mapped = {canon: df.iloc[idx, j] for canon, j in col_map.items()}
            st.info(f"Detected header row {idx+1}: {list(df.iloc[idx])}")
            st.info(f"Column mapping: {mapped}")
            return idx, mapped
    return None, None

def truncate_table(df):
    stop_idx = len(df)
    for i, row in df.iterrows():
        if not row["Style No"] or pd.isna(row["Style No"]) or not row["PO Nr"] or pd.isna(row["PO Nr"]):
            stop_idx = i
            break
    return df.iloc[:stop_idx].copy()

def is_numeric_or_empty(val):
    if pd.isna(val): return True
    try:
        float(str(val).replace(',', '').replace(' ', ''))
        return True
    except Exception:
        return False

def read_and_standardize(file, required=['Style No', 'PO Nr', 'Quantity', 'Price', 'Total']):
    excel = pd.ExcelFile(file)
    for sheet_name in excel.sheet_names:
        df = excel.parse(sheet_name, header=None, dtype=str)
        hdr_row, col_map = find_header_row_and_map(df)
        if hdr_row is not None and "Style No" in col_map and "PO Nr" in col_map:
            df2 = pd.read_excel(file, sheet_name=sheet_name, header=hdr_row, dtype=str)
            mapped_cols = {}
            df2_cols_clean = [clean_col(c) for c in df2.columns]
            for canon, alias in col_map.items():
                alias_clean = clean_col(alias)
                found = None
                for c, cc in zip(df2.columns, df2_cols_clean):
                    if alias_clean == cc or alias_clean in cc or cc in alias_clean:
                        found = c
                        break
                if found:
                    mapped_cols[canon] = found
            if "Style No" in mapped_cols and "PO Nr" in mapped_cols:
                # 1. FILTER out summary/total rows with non-numeric values in numeric columns
                num_cols = [mapped_cols.get(x) for x in ['Quantity', 'Price', 'Total'] if x in mapped_cols]
                if num_cols:
                    mask = df2[num_cols].applymap(is_numeric_or_empty).all(axis=1)
                    df2 = df2[mask]
                # 2. Convert numeric columns
                for x in ['Quantity', 'Price', 'Total']:
                    if x in mapped_cols:
                        df2[mapped_cols[x]] = (
                            df2[mapped_cols[x]]
                            .astype(str)
                            .str.replace(',', '', regex=False)
                            .str.replace(' ', '', regex=False)
                            .replace('', np.nan)
                            .astype(float)
                        )
                    else:
                        df2[x] = np.nan
                result = df2[
                    [mapped_cols['Style No'], mapped_cols['PO Nr'], 
                     mapped_cols.get('Quantity', 'Quantity'), 
                     mapped_cols.get('Price', 'Price'), 
                     mapped_cols.get('Total', 'Total')]
                ].copy()
                result.columns = ['Style No', 'PO Nr', 'Quantity', 'Price', 'Total']
                result = truncate_table(result)
                return result
            else:
                st.warning(f"Could not find all required columns in `{file.name if hasattr(file,'name') else sheet_name}`. Found: {mapped_cols}")
    return None

# --- UI ---

st.header("1. Upload Invoice Input Files")
invoice_files = st.file_uploader(
    "Upload one or more invoice Excel files",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    key="invoices"
)

st.header("2. Upload Consolidated Invoice File")
consolidated_file = st.file_uploader(
    "Upload the consolidated Excel file",
    type=["xls", "xlsx"],
    key="consolidated"
)

if invoice_files and consolidated_file:
    consolidated_df = read_and_standardize(consolidated_file)
    if consolidated_df is None:
        st.error("❌ Could not process consolidated file. No suitable header row or columns found.")
    else:
        for file in invoice_files:
            st.subheader(f"Results for {file.name}")
            input_df = read_and_standardize(file)
            if input_df is None:
                st.error("❌ Could not detect header or required columns for this file.")
                continue

            merged = input_df.merge(
                consolidated_df,
                on=["Style No", "PO Nr"],
                how="left",
                suffixes=("_input", "_consolidated")
            )
            merged["Quantity_diff"] = merged["Quantity_input"] - merged["Quantity_consolidated"]
            merged["Price_diff"] = merged["Price_input"] - merged["Price_consolidated"]
            merged["Total_diff"] = merged["Total_input"] - merged["Total_consolidated"]
            # Set diffs to NaN where consolidated value is missing
            for col, refcol in [
                ("Quantity_diff", "Quantity_consolidated"),
                ("Price_diff", "Price_consolidated"),
                ("Total_diff", "Total_consolidated")
            ]:
                merged.loc[merged[refcol].isna(), col] = np.nan

            # Deviations: either value differs or no match found
            deviation_mask = (
                merged["Quantity_consolidated"].isna() |
                merged["Price_consolidated"].isna() |
                merged["Total_consolidated"].isna() |
                (merged["Quantity_diff"].abs() > 0.01) |
                (merged["Price_diff"].abs() > 0.01) |
                (merged["Total_diff"].abs() > 0.01)
            )
            deviated_rows = merged[deviation_mask].copy()

            # Only sum actual diff values (ignore NaN)
            total_deviated = deviated_rows["Total_diff"].dropna().sum() if not deviated_rows.empty else 0.0


            if deviated_rows.empty:
                st.success("✅ No deviations found! All lines match.")
            else:
                st.warning(f"⚠️ {len(deviated_rows)} deviating lines found.")
                st.write(f"**Total deviated amount:** {total_deviated:.2f}")
                st.dataframe(deviated_rows[
                    ["Style No", "PO Nr", "Quantity_input", "Quantity_consolidated", "Quantity_diff",
                     "Price_input", "Price_consolidated", "Price_diff",
                     "Total_input", "Total_consolidated", "Total_diff"]
                ], use_container_width=True)
                csv_buffer = io.StringIO()
                deviated_rows.to_csv(csv_buffer, index=False)
                st.download_button(
                    label="Download Deviating Lines as CSV",
                    data=csv_buffer.getvalue(),
                    file_name=f"{file.name}_deviations.csv",
                    mime="text/csv"
                )
else:
    st.info("Please upload both input invoice files and a consolidated file to begin.")

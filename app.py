# app.py – Monthly Costs Processor
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
st.set_page_config(page_title="Monthly Costs Processor", layout="wide")
st.title("Monthly Costs Processor")

# ── Session state ──────────────────────────────────────────────
if 'storage_processed' not in st.session_state:
    st.session_state.storage_processed = False

orders_df = None
valid_skus = None
owner_map = {}

# ── Step 1: Orders Processing ─────────────────────────────────
st.header("Step 1: Orders Processing")
col1, col2 = st.columns(2)

with col1:
    monthly_file = st.file_uploader("Upload Order Line Level Data", type=['xlsx'])
with col2:
    sku_file = st.file_uploader("Upload SKU List File", type=['xlsx'])

if monthly_file and sku_file:
    try:
        # Read SKU list
        sku_df = pd.read_excel(sku_file, usecols=[0])
        valid_skus = set(sku_df.iloc[:, 0].astype(str).str.strip().str.upper())

        # Read Order file (header row = row 2)
        xl = pd.ExcelFile(monthly_file)
        monthly_df = pd.read_excel(monthly_file, sheet_name=xl.sheet_names[0], header=1)

        monthly_df.columns = monthly_df.columns.str.strip()

        if 'Date Ordered' in monthly_df.columns:
            monthly_df['Date Ordered'] = (
                pd.to_datetime(monthly_df['Date Ordered'])
                  .dt.strftime('%Y-%m-%d %H:%M:%S')
            )

        if 'Location Code' in monthly_df.columns:
            monthly_df['Location Code'] = monthly_df['Location Code'].astype(str)

        monthly_df['Stock Code'] = (
            monthly_df['Stock Code'].astype(str).str.strip().str.upper()
        )
        monthly_df['Total Locations'] = (
            pd.to_numeric(monthly_df['Total Locations'], errors='coerce')
              .fillna(0)
        )

        filtered_df = monthly_df[monthly_df['Stock Code'].isin(valid_skus)]

        columns_to_remove = [
            'Sell Price', 'Line Status', 'Back Order Status',
            'Back Order Placed Date', 'Sell Price (Packs)'
        ]
        filtered_df = filtered_df.drop(
            columns=[c for c in columns_to_remove if c in filtered_df.columns]
        )

        # Charges
        filtered_df['Pick Charge'] = filtered_df['Total Locations'] * 1.59

        filtered_df['Order_Count'] = (
            filtered_df.groupby('Order Number').cumcount() + 1
        )
        filtered_df['Packaging'] = (
            (filtered_df['Order_Count'] % 10 == 1).astype(int)
              * filtered_df['Total Locations']
        )
        filtered_df = filtered_df.drop('Order_Count', axis=1)

        filtered_df['Packaging Charge'] = filtered_df['Packaging'] * 0.97
        filtered_df['Label Charge'] = filtered_df['Packaging'] * 0.39

        for c in ['Pick Charge', 'Packaging Charge', 'Label Charge']:
            filtered_df[c] = filtered_df[c].round(2)

        orders_df = filtered_df

        # Display metrics
        st.header("Orders Processing Results")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Orders", len(monthly_df))
        c2.metric("Valid Orders", len(filtered_df))
        c3.metric("Removed Orders", len(monthly_df) - len(filtered_df))

    except Exception as e:
        st.error(f"An error occurred in Orders processing: {e}")
        st.error("Please check that the files have the expected format and try again.")

# ── Step 2: Storage Processing ────────────────────────────────
if valid_skus is not None:
    st.header("Step 2: Storage Processing")
    col1, col2 = st.columns(2)

    with col1:
        stock_file   = st.file_uploader("Upload Stock Report", type=['xlsx'])
    with col2:
        storage_file = st.file_uploader("Upload Staci Storage Tab", type=['xlsx'])

    if stock_file and storage_file:
        try:
            stock_df   = pd.read_excel(stock_file,   header=1)
            storage_df = pd.read_excel(storage_file, header=1)

            stock_df['Stock Code'] = (
                stock_df['Stock Code'].astype(str).str.strip().str.upper()
            )
            filtered_stock = stock_df[stock_df['Stock Code'].isin(valid_skus)].copy()

            filtered_stock['Pallets'] = filtered_stock['Stock Code'].map(
                storage_df.set_index('Part Number')['Period']
                         .to_dict()
            ).fillna(0)

            filtered_stock['Cost'] = (filtered_stock['Pallets'] * 1.92).round(2)

            # Map owners and add to orders
            owner_map = filtered_stock.set_index('Stock Code')['Responsible Owner'].to_dict()
            if orders_df is not None:
                orders_df['Responsible Owner'] = (
                    orders_df['Stock Code'].map(owner_map).fillna('Unknown')
                )
            st.session_state.storage_processed = True

        except Exception as e:
            st.error(f"An error occurred in Storage processing: {e}")
            st.error("Please check that all files have the expected format and try again.")
            st.session_state.storage_processed = False
else:
    st.info("Please complete Step 1 (Orders Processing) before proceeding to Step 2.")

# ── Step 3: Goods-In Processing & Report Download ─────────────
if valid_skus is not None and st.session_state.storage_processed:
    st.header("Step 3: Goods In Processing")
    goods_in_file = st.file_uploader("Upload Goods In File", type=['xlsx'])

    if goods_in_file:
        try:
            goods_in_df = pd.read_excel(goods_in_file, header=0)

            # Identify part-number column
            part_no_col = next(
                (c for c in goods_in_df.columns
                 if str(c).lower().replace(' ', '') == 'partno'),
                None
            )
            if part_no_col is None:
                st.error("Could not find the part-number column in the Goods In file.")
                st.stop()

            goods_in_df[part_no_col] = (
                goods_in_df[part_no_col].astype(str).str.strip().str.upper()
            )
            filtered_goods_in = goods_in_df[
                goods_in_df[part_no_col].isin(valid_skus)
            ].copy()

            # Keep expected columns even if no rows match
            filtered_goods_in = filtered_goods_in[
                [part_no_col, 'Part Description', 'Qty', 'No Of Containers']
            ]
            filtered_goods_in.columns = [
                'Stock Code', 'Part Description', 'Qty', 'No of Containers'
            ]
            filtered_goods_in['Cost'] = (
                filtered_goods_in['No of Containers'] * 5.26
            ).round(2)
            filtered_goods_in['Responsible Owner'] = (
                filtered_goods_in['Stock Code'].map(owner_map).fillna('Unknown')
            )

            # ── Build Excel workbook – always ───────────────────
            owners = sorted(filtered_stock['Responsible Owner'].fillna('Unknown').unique())
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                currency_fmt = '_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-'

                for owner in owners:
                    owner_orders = orders_df[orders_df['Responsible Owner'] == owner]
                    sheet_name = f"Orders - {owner}"
                    owner_orders.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws_orders = writer.sheets[sheet_name]

                    total_row = len(owner_orders) + 2
                    for col in ['Pick Charge', 'Packaging Charge', 'Label Charge']:
                        col_letter = get_column_letter(owner_orders.columns.get_loc(col) + 1)
                        ws_orders[f"{col_letter}{total_row}"] = \
                            f"=SUM({col_letter}2:{col_letter}{total_row-1})"
                    ws_orders[f"A{total_row}"] = "Total"

                    for col in ['Pick Charge', 'Packaging Charge', 'Label Charge']:
                        col_letter = get_column_letter(owner_orders.columns.get_loc(col) + 1)
                        for r in range(2, total_row + 1):
                            ws_orders[f"{col_letter}{r}"].number_format = currency_fmt

                    owner_storage = filtered_stock[filtered_stock['Responsible Owner'] == owner]
                    sheet_name_st = f"Storage - {owner}"
                    owner_storage.to_excel(writer, sheet_name=sheet_name_st, index=False)
                    ws_storage = writer.sheets[sheet_name_st]

                    st_total_row = len(owner_storage) + 2
                    st_cost_col = get_column_letter(owner_storage.columns.get_loc('Cost') + 1)
                    ws_storage[f"{st_cost_col}{st_total_row}"] = \
                        f"=SUM({st_cost_col}2:{st_cost_col}{st_total_row-1})"
                    ws_storage[f"A{st_total_row}"] = "Total"

                    for r in range(2, st_total_row + 1):
                        ws_storage[f"{st_cost_col}{r}"].number_format = currency_fmt

                    owner_goods_in = filtered_goods_in[filtered_goods_in['Responsible Owner'] == owner]
                    if not owner_goods_in.empty:
                        sheet_name_gi = f"Goods In - {owner}"
                        owner_goods_in.to_excel(writer, sheet_name=sheet_name_gi, index=False)
                        ws_gi = writer.sheets[sheet_name_gi]

                        gi_total_row = len(owner_goods_in) + 2
                        gi_cost_col = get_column_letter(
                            owner_goods_in.columns.get_loc('Cost') + 1
                        )
                        ws_gi[f"{gi_cost_col}{gi_total_row}"] = \
                            f"=SUM({gi_cost_col}2:{gi_cost_col}{gi_total_row-1})"
                        ws_gi[f"A{gi_total_row}"] = "Total"

                        for r in range(2, gi_total_row + 1):
                            ws_gi[f"{gi_cost_col}{r}"].number_format = currency_fmt

                # Auto-size columns on all sheets
                for ws in writer.sheets.values():
                    for col_cells in ws.columns:
                        max_len = max(
                            len(str(cell.value)) if cell.value else 0
                            for cell in col_cells
                        )
                        ws.column_dimensions[
                            get_column_letter(col_cells[0].column)
                        ].width = max_len + 2

            # ── Download button ────────────────────────────────
            output.seek(0)
            st.download_button(
                label="Download Complete Report",
                data=output,
                file_name="monthly_costs_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Info if Goods In tab absent
            if filtered_goods_in.empty:
                st.info(
                    "No matching SKUs found in the Goods In file, "
                    "so that tab was left out of the workbook."
                )

        except Exception as e:
            st.error(f"An error occurred in Goods In processing: {e}")
            st.error("Please check that the file has the expected format and try again.")
else:
    if valid_skus is not None:
        st.info("Please complete Step 2 (Storage Processing) before proceeding to Step 3.")

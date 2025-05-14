import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers

st.set_page_config(page_title="Monthly Costs Processor", layout="wide")

st.title("Monthly Costs Processor")

# Initialize variables
if 'storage_processed' not in st.session_state:
    st.session_state.storage_processed = False
orders_df = None
valid_skus = None

# Step 1: Orders Processing
st.header("Step 1: Orders Processing")
col1, col2 = st.columns(2)

with col1:
    monthly_file = st.file_uploader("Upload Order Line Level Data", type=['xlsx'])
with col2:
    sku_file = st.file_uploader("Upload SKU List File", type=['xlsx'])

if monthly_file and sku_file:
    try:
        # Read the SKU list and use first column
        sku_df = pd.read_excel(sku_file, usecols=[0])
        valid_skus = set(sku_df.iloc[:, 0].astype(str).str.strip().str.upper())
        
        # Read the monthly costs file, skipping the first row and using the second row as headers
        xl = pd.ExcelFile(monthly_file)
        monthly_df = pd.read_excel(monthly_file, sheet_name=xl.sheet_names[0], header=1)
        
        # Clean up column names - strip whitespace and handle case
        monthly_df.columns = monthly_df.columns.str.strip()
        
        # Convert data types appropriately
        if 'Date Ordered' in monthly_df.columns:
            monthly_df['Date Ordered'] = pd.to_datetime(monthly_df['Date Ordered']).dt.strftime('%Y-%m-%d %H:%M:%S')
        
        # Ensure Location Code is treated as string
        if 'Location Code' in monthly_df.columns:
            monthly_df['Location Code'] = monthly_df['Location Code'].astype(str)
        
        # Convert stock codes to uppercase and strip whitespace for comparison
        monthly_df['Stock Code'] = monthly_df['Stock Code'].astype(str).str.strip().str.upper()
        
        # Ensure Total Locations is numeric
        monthly_df['Total Locations'] = pd.to_numeric(monthly_df['Total Locations'], errors='coerce').fillna(0)
        
        # Filter out irrelevant stock codes
        filtered_df = monthly_df[monthly_df['Stock Code'].isin(valid_skus)]
        
        # Remove specified columns if they exist
        columns_to_remove = ['Sell Price', 'Line Status', 'Back Order Status', 'Back Order Placed Date', 'Sell Price (Packs)']
        filtered_df = filtered_df.drop(columns=[col for col in columns_to_remove if col in filtered_df.columns])
        
        # Calculate Pick Charge based on Total Locations
        filtered_df['Pick Charge'] = filtered_df['Total Locations'] * 1.59
        
        # Calculate Packaging
        # Group by Order Number and create a cumulative count within each group
        filtered_df['Order_Count'] = filtered_df.groupby('Order Number').cumcount() + 1
        # Set Packaging to 1 only for rows that are at positions 1, 11, 21, etc. within each order
        filtered_df['Packaging'] = ((filtered_df['Order_Count'] % 10 == 1).astype(int)) * filtered_df['Total Locations']
        # Remove the temporary column
        filtered_df = filtered_df.drop('Order_Count', axis=1)
        
        # Calculate charges
        filtered_df['Packaging Charge'] = filtered_df['Packaging'] * 0.97
        filtered_df['Label Charge'] = filtered_df['Packaging'] * 0.39
        
        # Format currency columns for display
        for col in ['Pick Charge', 'Packaging Charge', 'Label Charge']:
            filtered_df[col] = filtered_df[col].round(2)
        
        # Store orders_df for later use
        orders_df = filtered_df
        
        # Display statistics
        st.header("Orders Processing Results")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total Orders", len(monthly_df))
        with col2:
            st.metric("Valid Orders", len(filtered_df))
        with col3:
            st.metric("Removed Orders", len(monthly_df) - len(filtered_df))
        
    except Exception as e:
        st.error(f"An error occurred in Orders processing: {str(e)}")
        st.error("Please check that the files have the expected format and try again.")

# Step 2: Storage Processing
if valid_skus is not None:  # Only show Step 2 if Step 1 is complete
    st.header("Step 2: Storage Processing")
    col1, col2 = st.columns(2)
    
    with col1:
        stock_file = st.file_uploader("Upload Stock Report", type=['xlsx'])
    with col2:
        storage_file = st.file_uploader("Upload Staci Storage Tab", type=['xlsx'])
    
    if stock_file and storage_file:
        try:
            # Read the stock report with header in row 2
            stock_df = pd.read_excel(stock_file, header=1)
            
            # Read the storage report with header in row 2
            storage_df = pd.read_excel(storage_file, header=1)
            
            # Continue with the rest of the processing
            stock_df['Stock Code'] = stock_df['Stock Code'].astype(str).str.strip().str.upper()
            
            # Filter stock report based on SKU list
            filtered_stock = stock_df[stock_df['Stock Code'].isin(valid_skus)].copy()
            
            # Add Pallets column by matching with storage report
            filtered_stock['Pallets'] = filtered_stock['Stock Code'].map(
                storage_df.set_index('Part Number')['Period'].to_dict()
            ).fillna(0)
            
            # Calculate Cost
            filtered_stock['Cost'] = filtered_stock['Pallets'] * 1.92
            
            # Format currency column
            filtered_stock['Cost'] = filtered_stock['Cost'].round(2)
            
            # Store the processed data for later use
            st.session_state.storage_processed = True
            
        except Exception as e:
            st.error(f"An error occurred in Storage processing: {str(e)}")
            st.error("Please check that all files have the expected format and try again.")
            st.session_state.storage_processed = False
else:
    st.info("Please complete Step 1 (Orders Processing) before proceeding to Step 2 (Storage Processing).")

# Step 3: Goods In Processing
if valid_skus is not None:  # Only show Step 3 if Step 1 is complete
    if st.session_state.storage_processed:  # If Step 2 is complete, show Step 3
        st.header("Step 3: Goods In Processing")
        
        goods_in_file = st.file_uploader("Upload Goods In File", type=['xlsx'])
        
        if goods_in_file:
            try:
                # Read the goods in file with header in row 1
                goods_in_df = pd.read_excel(goods_in_file, header=0)
                
                # Debug: Show column names to identify the correct one
                st.write("Available columns in Goods In file:", list(goods_in_df.columns))
                
                # Find the correct column name for part number (case-insensitive)
                part_no_col = next((col for col in goods_in_df.columns if str(col).lower().replace(' ', '') == 'partno'), None)
                
                if part_no_col is None:
                    st.error("Could not find the part number column in the Goods In file. Please check the file format.")
                    pass
                
                # Convert part numbers to uppercase and strip whitespace for comparison
                goods_in_df[part_no_col] = goods_in_df[part_no_col].astype(str).str.strip().str.upper()
                
                # Filter based on SKU list
                filtered_goods_in = goods_in_df[goods_in_df[part_no_col].isin(valid_skus)].copy()
                
                if len(filtered_goods_in) > 0:
                    # Select and rename required columns
                    filtered_goods_in = filtered_goods_in[[part_no_col, 'Part Description', 'Qty', 'No Of Containers']]
                    filtered_goods_in.columns = ['Stock Code', 'Part Description', 'Qty', 'No of Containers']
                    
                    # Calculate Cost
                    filtered_goods_in['Cost'] = filtered_goods_in['No of Containers'] * 5.26
                    
                    # Format currency column
                    filtered_goods_in['Cost'] = filtered_goods_in['Cost'].round(2)
                    
                    # ─── Build the Excel file ────────────────────────────────────────────────────
                    output = io.BytesIO()

                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        # 1) Orders sheet  ─ always present
                        orders_df.to_excel(writer, sheet_name="Orders", index=False)
                        ws_orders = writer.sheets["Orders"]
                        
                        # Add totals row to Orders
                        total_row = len(orders_df) + 2
                        for col in ['Pick Charge', 'Packaging Charge', 'Label Charge']:
                            col_letter = get_column_letter(orders_df.columns.get_loc(col) + 1)
                            worksheet_orders[f"{col_letter}{total_row}"] = f"=SUM({col_letter}2:{col_letter}{total_row-1})"
                        
                        worksheet_orders[f"A{total_row}"] = "Total"
                        
                        # Format currency columns in Orders
                        currency_format = '_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-'
                        for col in ['Pick Charge', 'Packaging Charge', 'Label Charge']:
                            col_letter = get_column_letter(orders_df.columns.get_loc(col) + 1)
                            for row in range(2, total_row + 1):
                                cell = worksheet_orders[f"{col_letter}{row}"]
                                cell.number_format = currency_format
                        
                        # 2) Storage sheet  ─ always present
                        filtered_stock.to_excel(writer, sheet_name="Storage", index=False)
                         ws_storage = writer.sheets["Storage"]
                        
                        # Add totals row to Storage (only for Cost column)
                        storage_total_row = len(filtered_stock) + 2
                        cost_col = get_column_letter(filtered_stock.columns.get_loc('Cost') + 1)
                        worksheet_storage[f"{cost_col}{storage_total_row}"] = f"=SUM({cost_col}2:{cost_col}{storage_total_row-1})"
                        worksheet_storage[f"A{storage_total_row}"] = "Total"
                        
                        # Format Cost column in Storage
                        for row in range(2, storage_total_row + 1):
                            cell = worksheet_storage[f"{cost_col}{row}"]
                            cell.number_format = currency_format
                        
                      # 3) Goods In sheet  ─ only if there are rows
                        if not filtered_goods_in.empty:
                            filtered_goods_in.to_excel(writer, sheet_name="Goods In", index=False)
                            ws_goods = writer.sheets["Goods In"]
                              
                        # Add totals row to Goods In (only for Cost column)
                        goods_in_total_row = len(filtered_goods_in) + 2
                        cost_col = get_column_letter(filtered_goods_in.columns.get_loc('Cost') + 1)
                        worksheet_goods_in[f"{cost_col}{goods_in_total_row}"] = f"=SUM({cost_col}2:{cost_col}{goods_in_total_row-1})"
                        worksheet_goods_in[f"A{goods_in_total_row}"] = "Total"
                        
                        # Format Cost column in Goods In
                        for row in range(2, goods_in_total_row + 1):
                            cell = worksheet_goods_in[f"{cost_col}{row}"]
                            cell.number_format = currency_format
                        
                        # 4) Column-width loop covering whatever sheets were created
                        for ws in writer.sheets.values():
                            for col in ws.columns:
                                max_len = max(
                                    len(str(cell.value)) if cell.value is not None else 0
                                    for cell in col
                                )
                                ws.column_dimensions[col[0].column_letter].width = max_len + 2

                    # ─── Download button ─────────────────────────────────────────────────────────
                    output.seek(0)
                    st.download_button(
                        label="Download Complete Report",
                        data=output,
                        file_name="monthly_costs_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    # Optional FYI message
                    if filtered_goods_in.empty:
                        st.info(
                            "No matching SKUs found in the Goods In file, so that tab was left out of the report."
                        )
            except Exception as e:
                st.error(f"An error occurred in Goods In processing: {str(e)}")
                st.error("Please check that the file has the expected format and try again.")
    else:
        st.info("Please complete Step 2 (Storage Processing) before proceeding to Step 3 (Goods In Processing).") 

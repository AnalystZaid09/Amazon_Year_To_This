import streamlit as st
import pandas as pd
import zipfile
from pathlib import Path
import io
import calendar
import gc
import numpy as np

import base64

st.set_page_config(page_title="Sales Data Analysis", layout="wide", initial_sidebar_state="expanded")

# Helper function to create download link (base64 approach - works reliably on Streamlit Cloud)
def create_download_link(df, filename, link_text):
    """Generate a download link for a DataFrame as Excel file using base64 encoding.
    This approach doesn't trigger Streamlit reruns and works reliably on Streamlit Cloud."""
    try:
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        b64 = base64.b64encode(output.getvalue()).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}" style="display: inline-block; padding: 0.5rem 1rem; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">📥 {link_text}</a>'
        return href
    except Exception as e:
        return f'<p style="color: red;">Error creating download: {e}</p>'

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
    }
    .filter-box {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    .quarter-badge {
        display: inline-block;
        padding: 0.25rem 0.5rem;
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
        font-size: 0.9rem;
        margin-left: 0.5rem;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">📊 Sales Data Analysis Dashboard</div>', unsafe_allow_html=True)

# File uploaders
st.sidebar.header("Upload Files")

# Add a clear cache button to ensure fresh data processing
if st.sidebar.button("🔄 Clear Cache & Refresh"):
    st.cache_data.clear()
    st.rerun()

zip_files = st.sidebar.file_uploader(
    "Upload ZIP files (B2B & B2C Reports)", 
    type=['zip'], 
    accept_multiple_files=True
)
pm_file = st.sidebar.file_uploader("Upload PM Excel File", type=['xlsx', 'xls'])

# NEW: High Volume Toggle for 50+ files
high_volume_mode = st.sidebar.checkbox(
    "🚀 High Volume Mode (50+ files)", 
    help="Disables detailed raw data view to save memory for very large uploads"
)

if zip_files and pm_file:
    # Process ZIP files
    @st.cache_data
    def process_zip_files(zip_file_list, h_volume=False):
        shipment_dfs = []
        unfiltered_dfs = []
        transaction_counts = {}
        
        # Key columns to load - reduce memory by only loading what's needed for analysis
        relevant_cols = ['Invoice Date', 'Asin', 'Quantity', 'Invoice Amount', 'Order Id', 'Shipment Id', 'Transaction Type']
        
        for uploaded_zip in zip_file_list:
            with zipfile.ZipFile(uploaded_zip, 'r') as z:
                for file_name in z.namelist():
                    if file_name.endswith('/') or not file_name.lower().endswith(('.csv', '.xlsx', '.xls')):
                        continue
                    
                    try:
                        with z.open(file_name) as f:
                            if file_name.lower().endswith('.csv'):
                                try:
                                    # Use usecols to only load what's needed for analysis
                                    df = pd.read_csv(f, low_memory=False, usecols=lambda x: x in relevant_cols)
                                except ValueError:
                                    f.seek(0)
                                    df = pd.read_csv(f, low_memory=False)
                            elif file_name.lower().endswith(('.xlsx', '.xls')):
                                df = pd.read_excel(f)
                            else:
                                continue
                            
                            # Standardize column names
                            if 'Asin' not in df.columns and 'ASIN' in df.columns:
                                df.rename(columns={'ASIN': 'Asin'}, inplace=True)
                            
                            # Track transaction counts for the summary
                            if 'Transaction Type' in df.columns:
                                counts = df['Transaction Type'].str.strip().str.lower().value_counts().to_dict()
                                for t_type, count in counts.items():
                                    transaction_counts[t_type] = transaction_counts.get(t_type, 0) + count
                                
                                # Separate Shipment records immediately
                                is_shipment = df['Transaction Type'].str.strip().str.lower() == 'shipment'
                                ship_df = df[is_shipment].copy()
                                if not ship_df.empty:
                                    shipment_dfs.append(ship_df)
                                
                                # ONLY collect unfiltered data if NOT in High Volume mode
                                if not h_volume:
                                    # For unfiltered data, downcast to save space
                                    for col in df.select_dtypes(include=['object']).columns:
                                        if df[col].nunique() < 100:
                                            df[col] = df[col].astype('category')
                                    unfiltered_dfs.append(df)
                                else:
                                    # In high volume mode, clear local variables aggressively
                                    del df
                                    gc.collect()
                    except Exception as e:
                        st.warning(f"Error reading {file_name}: {e}")
                        continue
        
        # Concatenate smaller lists
        filtered_combined = pd.concat(shipment_dfs, ignore_index=True) if shipment_dfs else pd.DataFrame()
        unfiltered_combined = pd.concat(unfiltered_dfs, ignore_index=True) if unfiltered_dfs and not h_volume else pd.DataFrame()
        
        # Explicitly delete temporary lists and collect garbage
        del shipment_dfs, unfiltered_dfs
        gc.collect()
        
        return filtered_combined, unfiltered_combined, transaction_counts

    def process_data(filtered_df, unfiltered_df, pm_df):
        def add_date_columns(df):
            if df.empty: return df
            df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
            df['Date'] = df['Invoice Date'].dt.date
            df['Month'] = pd.to_datetime(df['Date']).dt.month
            df['Year'] = pd.to_datetime(df['Date']).dt.year
            df['Month_Year'] = pd.to_datetime(df['Date']).dt.strftime('%b-%y')
            df['Quarter'] = pd.cut(df['Month'], bins=[0, 3, 6, 9, 12], labels=['Q1', 'Q2', 'Q3', 'Q4']).astype(str)
            df['Quarter_Year'] = df['Quarter'] + '-' + df['Year'].astype(str)
            return df
        
        filtered_df = add_date_columns(filtered_df)
        unfiltered_df = add_date_columns(unfiltered_df)
        
        pm_cols = pm_df[['ASIN', 'Brand', 'Brand Manager', 'Vendor SKU Codes', 'Product Name']].drop_duplicates(subset=['ASIN'], keep='first')
        
        for df in [filtered_df, unfiltered_df]:
            if not df.empty: df['Asin'] = df['Asin'].astype(str)
        pm_cols['ASIN'] = pm_cols['ASIN'].astype(str)
        
        filtered_df = filtered_df.merge(pm_cols, left_on='Asin', right_on='ASIN', how='left')
        unfiltered_df = unfiltered_df.merge(pm_cols, left_on='Asin', right_on='ASIN', how='left')
        
        gc.collect()
        return filtered_df, unfiltered_df, len(filtered_df), len(unfiltered_df)

    with st.spinner("Processing files..."):
        f_combined, u_combined, transaction_counts = process_zip_files(zip_files, high_volume_mode)
        pm_df = pd.read_excel(pm_file)
        processed_df, unfiltered_combined_df, filtered_count, unfiltered_count = process_data(f_combined, u_combined, pm_df)
        del f_combined, u_combined, pm_df
        gc.collect()
    
    # Show detailed record counts (using counts captured BEFORE merge to avoid PM duplicate inflation)
    col1, col2 = st.columns(2)
    with col1:
        st.success(f"✅ Filtered (Shipment only): **{filtered_count:,}** records")
    with col2:
        st.info(f"📊 Total Records: **{unfiltered_count if not high_volume_mode else filtered_count:,}**")
        if high_volume_mode:
            st.warning("🚀 High Volume Mode is ACTIVE. Unfiltered data view is disabled to prioritize memory.")
    
    # Show transaction type breakdown in expander for debugging
    with st.expander("🔍 Transaction Type Breakdown"):
        st.write("Records by Transaction Type:")
        for trans_type, count in sorted(transaction_counts.items(), key=lambda x: -x[1]):
            st.write(f"  - **{trans_type}**: {count:,}")
    
    # Enhanced Sidebar filters
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎯 Time Period Filters")
    
    # Create filter box styling
    with st.sidebar.container():
        time_period = st.radio(
            "Select View Type",
            ["📅 All Data", "📆 Quarter View", "🗓️ Month View"],
            help="Choose how you want to view your data"
        )
    
    filtered_df = processed_df.copy()
    filter_info = ""
    
    if time_period == "📆 Quarter View":
        st.sidebar.markdown("---")
        
        # Get available years
        years = sorted(processed_df['Year'].dropna().unique(), reverse=True)
        selected_year = st.sidebar.selectbox(
            "📅 Select Year",
            years,
            help="Select the year for quarter analysis"
        )
        
        # Quarter selection with descriptions
        quarter_options = {
            'Q1': 'Q1 (Jan - Mar)',
            'Q2': 'Q2 (Apr - Jun)',
            'Q3': 'Q3 (Jul - Sep)',
            'Q4': 'Q4 (Oct - Dec)'
        }
        
        # Filter available quarters for selected year
        available_quarters = processed_df[processed_df['Year'] == selected_year]['Quarter'].unique()
        available_quarter_options = {k: v for k, v in quarter_options.items() if k in available_quarters}
        
        if available_quarter_options:
            selected_quarter_display = st.sidebar.selectbox(
                "📊 Select Quarter",
                list(available_quarter_options.values()),
                help="Q1: Jan-Mar | Q2: Apr-Jun | Q3: Jul-Sep | Q4: Oct-Dec"
            )
            
            # Get the quarter code (Q1, Q2, Q3, Q4)
            selected_quarter = [k for k, v in quarter_options.items() if v == selected_quarter_display][0]
            
            filtered_df = processed_df[
                (processed_df['Quarter'] == selected_quarter) & 
                (processed_df['Year'] == selected_year)
            ]
            
            # Define month ranges
            quarter_months = {
                'Q1': ['January', 'February', 'March'],
                'Q2': ['April', 'May', 'June'],
                'Q3': ['July', 'August', 'September'],
                'Q4': ['October', 'November', 'December']
            }
            
            filter_info = f"**{selected_quarter} {selected_year}** ({', '.join(quarter_months[selected_quarter])})"
            
            # Show summary for quarter
            st.sidebar.markdown("---")
            st.sidebar.markdown("#### Quarter Summary")
            st.sidebar.metric("Total Records", f"{len(filtered_df):,}")
            st.sidebar.metric("Date Range", f"{filtered_df['Date'].min()} to {filtered_df['Date'].max()}")
        else:
            st.sidebar.warning(f"No data available for {selected_year}")
    
    elif time_period == "🗓️ Month View":
        st.sidebar.markdown("---")
        
        # Get available years
        years = sorted(processed_df['Year'].dropna().unique(), reverse=True)
        selected_year = st.sidebar.selectbox(
            "📅 Select Year",
            years,
            help="Select the year for month analysis"
        )
        
        # Get available months for selected year
        year_data = processed_df[processed_df['Year'] == selected_year]
        available_months = sorted(year_data['Month'].dropna().unique())
        month_names = [calendar.month_name[m] for m in available_months]
        
        if month_names:
            selected_month_name = st.sidebar.selectbox(
                "📊 Select Month",
                month_names,
                help="Choose a specific month to analyze"
            )
            
            # Get month number
            selected_month = list(calendar.month_name).index(selected_month_name)
            
            filtered_df = processed_df[
                (processed_df['Month'] == selected_month) & 
                (processed_df['Year'] == selected_year)
            ]
            
            filter_info = f"**{selected_month_name} {selected_year}**"
            
            # Show summary for month
            st.sidebar.markdown("---")
            st.sidebar.markdown("#### Month Summary")
            st.sidebar.metric("Total Records", f"{len(filtered_df):,}")
            st.sidebar.metric("Date Range", f"{filtered_df['Date'].min()} to {filtered_df['Date'].max()}")
        else:
            st.sidebar.warning(f"No data available for {selected_year}")
    else:
        filter_info = "**All Available Data**"
    
    # Additional filters
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Additional Filters")
    
    # Brand filter with count
    brands = sorted([b for b in filtered_df['Brand'].dropna().unique() if b])
    brand_counts = filtered_df['Brand'].value_counts()
    
    brand_options = ['All Brands'] + [f"{brand} ({brand_counts[brand]:,})" for brand in brands]
    selected_brand_display = st.sidebar.selectbox(
        "🏢 Filter by Brand",
        brand_options,
        help="Select a specific brand or view all brands"
    )
    
    if selected_brand_display != 'All Brands':
        selected_brand = selected_brand_display.split(' (')[0]
        filtered_df = filtered_df[filtered_df['Brand'] == selected_brand]
    
    # Brand Manager filter
    managers = sorted([m for m in filtered_df['Brand Manager'].dropna().unique() if m])
    manager_options = ['All Managers'] + managers
    selected_manager = st.sidebar.selectbox(
        "👤 Filter by Brand Manager",
        manager_options,
        help="Select a specific brand manager"
    )
    
    if selected_manager != 'All Managers':
        filtered_df = filtered_df[filtered_df['Brand Manager'] == selected_manager]
    
    # Display active filters
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📌 Active Filters")
    st.sidebar.info(f"""
    **Period:** {filter_info}
    **Brand:** {selected_brand_display.split(' (')[0]}
    **Manager:** {selected_manager}
    **Records:** {len(filtered_df):,}
    """)
    
    # Main content - Show filter summary
    st.markdown(f"### Current View: {filter_info}")
    st.markdown("---")
    
    # Main content tabs - Tab 5 is conditional based on high_volume_mode
    tabs = ["📈 Summary", "🏢 Brand Analysis", "📦 ASIN Analysis", "📋 Raw Data"]
    if not high_volume_mode:
        tabs.append("📊 Combined Data (Unfiltered)")
    tabs.extend(["📊 Brand Comparison (YoY)", "📦 ASIN Comparison (YoY)"])
    
    tab_list = st.tabs(tabs)
    
    # Map tabs correctly based on whether Tab 5 exists
    tab1 = tab_list[0]
    tab2 = tab_list[1]
    tab3 = tab_list[2]
    tab4 = tab_list[3]
    if not high_volume_mode:
        tab5 = tab_list[4]
        tab6 = tab_list[5]
        tab7 = tab_list[6]
    else:
        tab5 = None
        tab6 = tab_list[4]
        tab7 = tab_list[5]
    
    with tab1:
        st.header("Summary Statistics")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Quantity", f"{filtered_df['Quantity'].sum():,.0f}")
        with col2:
            st.metric("Total Invoice Amount", f"₹{filtered_df['Invoice Amount'].sum():,.2f}")
        with col3:
            st.metric("Unique ASINs", f"{filtered_df['Asin'].nunique():,}")
        with col4:
            st.metric("Unique Brands", f"{filtered_df['Brand'].nunique():,}")
        
        st.subheader("Monthly Trend")
        # Ensure we don't crash if Month_Year is missing or empty
        if 'Month_Year' in filtered_df.columns and not filtered_df.empty:
            monthly_trend = filtered_df.groupby('Month_Year').agg({
                'Quantity': 'sum',
                'Invoice Amount': 'sum'
            }).reset_index()
            
            # Sort by date
            monthly_trend['sort_date'] = pd.to_datetime(monthly_trend['Month_Year'], format='%b-%y')
            monthly_trend = monthly_trend.sort_values('sort_date')
            
            col1, col2 = st.columns(2)
            with col1:
                st.line_chart(monthly_trend.set_index('Month_Year')['Quantity'])
                st.caption("Quantity Trend")
            with col2:
                st.line_chart(monthly_trend.set_index('Month_Year')['Invoice Amount'])
                st.caption("Invoice Amount Trend")
        else:
            st.info("No trend data available for current selection")
    
    with tab2:
        st.header("Brand Analysis")
        
        # Brand pivot
        brand_pivot = pd.pivot_table(
            filtered_df,
            index='Brand',
            values=['Quantity', 'Invoice Amount'],
            aggfunc='sum',
            margins=False
        )
        brand_pivot = brand_pivot.sort_values(by='Quantity', ascending=False)
        brand_pivot.loc['Grand Total'] = brand_pivot.sum()
        brand_pivot = brand_pivot.reset_index()
        
        # Format numbers
        brand_pivot['Invoice Amount'] = brand_pivot['Invoice Amount'].apply(lambda x: f"₹{x:,.2f}")
        brand_pivot['Quantity'] = brand_pivot['Quantity'].apply(lambda x: f"{x:,.0f}")
        
        st.dataframe(brand_pivot, width='stretch', height=600)
        
        # Download link - Excel format (base64 approach for Streamlit Cloud)
        st.markdown(create_download_link(brand_pivot, f"brand_analysis_{time_period}.xlsx", "Download Brand Analysis Excel"), unsafe_allow_html=True)
    
    with tab3:
        st.header("ASIN Analysis")
        
        # ASIN pivot
        asin_pivot = pd.pivot_table(
            filtered_df,
            index=['Asin', 'Brand'],
            values=['Quantity', 'Invoice Amount'],
            aggfunc='sum'
        )
        asin_pivot = asin_pivot.sort_values(by='Quantity', ascending=False)
        
        grand_total = pd.DataFrame(asin_pivot.sum()).T
        grand_total.index = pd.MultiIndex.from_tuples(
            [('Grand Total', '')],
            names=asin_pivot.index.names
        )
        asin_pivot = pd.concat([asin_pivot, grand_total])
        asin_pivot = asin_pivot.reset_index()
        
        # Format numbers
        asin_pivot['Invoice Amount'] = asin_pivot['Invoice Amount'].apply(lambda x: f"₹{x:,.2f}")
        asin_pivot['Quantity'] = asin_pivot['Quantity'].apply(lambda x: f"{x:,.0f}")
        
        st.dataframe(asin_pivot, width='stretch', height=600)
        
        # Download link - Excel format (base64 approach for Streamlit Cloud)
        st.markdown(create_download_link(asin_pivot, f"asin_analysis_{time_period}.xlsx", "Download ASIN Analysis Excel"), unsafe_allow_html=True)
    
    with tab4:
        st.header("Raw/Processed Data")
        
        # Select columns to display
        all_columns = filtered_df.columns.tolist()
        default_columns = ['Invoice Date', 'Asin', 'Brand', 'Product Name', 'Quantity', 
                          'Invoice Amount', 'Month_Year', 'Quarter', 'Year', 'Order Id', 'Shipment Id']
        
        selected_columns = st.multiselect(
            "Select columns to display",
            all_columns,
            default=[col for col in default_columns if col in all_columns]
        )
        
        if selected_columns:
            # Display row limit for large datasets to prevent browser crashes
            row_limit = 50000
            if len(filtered_df) > row_limit:
                st.warning(f"⚠️ Showing only first {row_limit:,} rows for performance. Full data is available in the Excel download below.")
                display_df = filtered_df[selected_columns].head(row_limit).copy()
            else:
                display_df = filtered_df[selected_columns].copy()
            
            st.dataframe(display_df, width='stretch', height=600)
            
            # Download link - Excel format (base64 approach for Streamlit Cloud)
            st.markdown(create_download_link(filtered_df[selected_columns], f"filtered_data_{time_period}.xlsx", "Download ALL Filtered Data Excel"), unsafe_allow_html=True)
        else:
            st.warning("Please select at least one column to display")
    
    if not high_volume_mode and tab5:
        with tab5:
            st.header("Combined Data (Unfiltered)")
            st.info(f"📊 This tab shows ALL {unfiltered_count:,} records without the 'Shipment' transaction type filter.")
        
        # Show transaction type breakdown
        st.subheader("Transaction Type Distribution")
        if not unfiltered_combined_df.empty and 'Transaction Type' in unfiltered_combined_df.columns:
            trans_type_counts = unfiltered_combined_df['Transaction Type'].value_counts().reset_index()
            trans_type_counts.columns = ['Transaction Type', 'Count']
            trans_type_counts['Percentage'] = (trans_type_counts['Count'] / trans_type_counts['Count'].sum() * 100).round(2).astype(str) + '%'
            st.dataframe(trans_type_counts, width='stretch')
        
        st.subheader("All Data")
        
        # Select columns to display
        all_columns_unfiltered = unfiltered_combined_df.columns.tolist()
        default_columns_unfiltered = ['Invoice Date', 'Transaction Type', 'Asin', 'Brand', 'Product Name', 'Quantity', 
                          'Invoice Amount', 'Month_Year', 'Quarter', 'Year', 'Order Id', 'Shipment Id']
        
        selected_columns_unfiltered = st.multiselect(
            "Select columns to display",
            all_columns_unfiltered,
            default=[col for col in default_columns_unfiltered if col in all_columns_unfiltered],
            key="unfiltered_columns"
        )
        
        if selected_columns_unfiltered:
            display_unfiltered_df = unfiltered_combined_df[selected_columns_unfiltered].copy()
            st.dataframe(display_unfiltered_df, width='stretch', height=600)
            
            # Download link - Excel format (base64 approach for Streamlit Cloud)
            st.markdown(create_download_link(display_unfiltered_df, f"combined_unfiltered_data_{time_period}.xlsx", "Download Combined (Unfiltered) Data Excel"), unsafe_allow_html=True)
        else:
            st.warning("Please select at least one column to display")

    # Year-over-Year Comparison Tabs
    with tab6:
        st.header("📊 Brand Comparison (Year-over-Year)")
        
        # Get available years
        available_years = sorted(processed_df['Year'].dropna().unique(), reverse=True)
        
        if len(available_years) >= 2:
            st.markdown("### Select Years to Compare")
            col1, col2 = st.columns(2)
            
            with col1:
                current_year = st.selectbox(
                    "📅 Current Year (to be analyzed)",
                    available_years,
                    index=0,
                    key="brand_current_year"
                )
            
            with col2:
                # Filter out the current year from previous year options
                prev_year_options = [y for y in available_years if y != current_year]
                if prev_year_options:
                    previous_year = st.selectbox(
                        "📅 Previous Year (to compare against)",
                        prev_year_options,
                        index=0,
                        key="brand_previous_year"
                    )
                else:
                    previous_year = None
                    st.warning("No other year available for comparison")
            
            if previous_year:
                # Filter data by years
                current_year_data = processed_df[processed_df['Year'] == current_year]
                previous_year_data = processed_df[processed_df['Year'] == previous_year]
                
                # Create brand pivots for each year
                current_brand_pivot = pd.pivot_table(
                    current_year_data,
                    index='Brand',
                    values=['Quantity', 'Invoice Amount'],
                    aggfunc='sum'
                ).reset_index()
                current_brand_pivot.columns = ['Brand', f'Invoice Amount ({current_year})', f'Quantity ({current_year})']
                
                previous_brand_pivot = pd.pivot_table(
                    previous_year_data,
                    index='Brand',
                    values=['Quantity', 'Invoice Amount'],
                    aggfunc='sum'
                ).reset_index()
                previous_brand_pivot.columns = ['Brand', f'Invoice Amount ({previous_year})', f'Quantity ({previous_year})']
                
                # Merge the two pivots
                brand_comparison = pd.merge(
                    previous_brand_pivot,
                    current_brand_pivot,
                    on='Brand',
                    how='outer'
                ).fillna(0)
                
                # Calculate differences and percentage changes
                brand_comparison['Qty Difference'] = brand_comparison[f'Quantity ({current_year})'] - brand_comparison[f'Quantity ({previous_year})']
                brand_comparison['Qty % Change'] = brand_comparison.apply(
                    lambda row: ((row[f'Quantity ({current_year})'] - row[f'Quantity ({previous_year})']) / row[f'Quantity ({previous_year})'] * 100) 
                    if row[f'Quantity ({previous_year})'] != 0 else (100 if row[f'Quantity ({current_year})'] > 0 else 0), axis=1
                )
                
                brand_comparison['Amount Difference'] = brand_comparison[f'Invoice Amount ({current_year})'] - brand_comparison[f'Invoice Amount ({previous_year})']
                brand_comparison['Amount % Change'] = brand_comparison.apply(
                    lambda row: ((row[f'Invoice Amount ({current_year})'] - row[f'Invoice Amount ({previous_year})']) / row[f'Invoice Amount ({previous_year})'] * 100) 
                    if row[f'Invoice Amount ({previous_year})'] != 0 else (100 if row[f'Invoice Amount ({current_year})'] > 0 else 0), axis=1
                )
                
                # Reorder columns
                brand_comparison = brand_comparison[[
                    'Brand',
                    f'Quantity ({previous_year})', f'Quantity ({current_year})', 'Qty Difference', 'Qty % Change',
                    f'Invoice Amount ({previous_year})', f'Invoice Amount ({current_year})', 'Amount Difference', 'Amount % Change'
                ]]
                
                # Sort by current year quantity descending
                brand_comparison = brand_comparison.sort_values(by=f'Quantity ({current_year})', ascending=False)
                
                # Add Grand Total row
                grand_total = pd.DataFrame({
                    'Brand': ['Grand Total'],
                    f'Quantity ({previous_year})': [brand_comparison[f'Quantity ({previous_year})'].sum()],
                    f'Quantity ({current_year})': [brand_comparison[f'Quantity ({current_year})'].sum()],
                    'Qty Difference': [brand_comparison['Qty Difference'].sum()],
                    'Qty % Change': [
                        (brand_comparison[f'Quantity ({current_year})'].sum() - brand_comparison[f'Quantity ({previous_year})'].sum()) / 
                        brand_comparison[f'Quantity ({previous_year})'].sum() * 100 if brand_comparison[f'Quantity ({previous_year})'].sum() != 0 else 0
                    ],
                    f'Invoice Amount ({previous_year})': [brand_comparison[f'Invoice Amount ({previous_year})'].sum()],
                    f'Invoice Amount ({current_year})': [brand_comparison[f'Invoice Amount ({current_year})'].sum()],
                    'Amount Difference': [brand_comparison['Amount Difference'].sum()],
                    'Amount % Change': [
                        (brand_comparison[f'Invoice Amount ({current_year})'].sum() - brand_comparison[f'Invoice Amount ({previous_year})'].sum()) / 
                        brand_comparison[f'Invoice Amount ({previous_year})'].sum() * 100 if brand_comparison[f'Invoice Amount ({previous_year})'].sum() != 0 else 0
                    ]
                })
                brand_comparison = pd.concat([brand_comparison, grand_total], ignore_index=True)
                
                # Format display dataframe
                display_brand_comparison = brand_comparison.copy()
                display_brand_comparison[f'Quantity ({previous_year})'] = display_brand_comparison[f'Quantity ({previous_year})'].apply(lambda x: f"{x:,.0f}")
                display_brand_comparison[f'Quantity ({current_year})'] = display_brand_comparison[f'Quantity ({current_year})'].apply(lambda x: f"{x:,.0f}")
                display_brand_comparison['Qty Difference'] = display_brand_comparison['Qty Difference'].apply(lambda x: f"{x:+,.0f}")
                display_brand_comparison['Qty % Change'] = display_brand_comparison['Qty % Change'].apply(lambda x: f"{x:+.2f}%")
                display_brand_comparison[f'Invoice Amount ({previous_year})'] = display_brand_comparison[f'Invoice Amount ({previous_year})'].apply(lambda x: f"₹{x:,.2f}")
                display_brand_comparison[f'Invoice Amount ({current_year})'] = display_brand_comparison[f'Invoice Amount ({current_year})'].apply(lambda x: f"₹{x:,.2f}")
                display_brand_comparison['Amount Difference'] = display_brand_comparison['Amount Difference'].apply(lambda x: f"₹{x:+,.2f}")
                display_brand_comparison['Amount % Change'] = display_brand_comparison['Amount % Change'].apply(lambda x: f"{x:+.2f}%")
                
                # Show summary metrics
                st.markdown(f"### Comparison: {current_year} vs {previous_year}")
                metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
                
                total_qty_change = brand_comparison.iloc[-1]['Qty Difference']
                total_qty_pct = brand_comparison.iloc[-1]['Qty % Change']
                total_amt_change = brand_comparison.iloc[-1]['Amount Difference']
                total_amt_pct = brand_comparison.iloc[-1]['Amount % Change']
                
                with metric_col1:
                    st.metric("Total Qty Change", f"{total_qty_change:+,.0f}", f"{total_qty_pct:+.2f}%")
                with metric_col2:
                    st.metric("Total Amount Change", f"₹{total_amt_change:+,.0f}", f"{total_amt_pct:+.2f}%")
                with metric_col3:
                    st.metric(f"Brands in {current_year}", f"{len(current_year_data['Brand'].dropna().unique()):,}")
                with metric_col4:
                    st.metric(f"Brands in {previous_year}", f"{len(previous_year_data['Brand'].dropna().unique()):,}")
                
                st.dataframe(display_brand_comparison, width='stretch', height=600)
                
                # Download link (base64 approach for Streamlit Cloud)
                st.markdown(create_download_link(brand_comparison, f"brand_comparison_{current_year}_vs_{previous_year}.xlsx", "Download Brand Comparison Excel"), unsafe_allow_html=True)
        else:
            st.warning("⚠️ Need at least 2 years of data for comparison. Please upload data from multiple years.")
    
    with tab7:
        st.header("📦 ASIN Comparison (Year-over-Year)")
        
        # Get available years
        available_years_asin = sorted(processed_df['Year'].dropna().unique(), reverse=True)
        
        if len(available_years_asin) >= 2:
            st.markdown("### Select Years to Compare")
            col1, col2 = st.columns(2)
            
            with col1:
                current_year_asin = st.selectbox(
                    "📅 Current Year (to be analyzed)",
                    available_years_asin,
                    index=0,
                    key="asin_current_year"
                )
            
            with col2:
                # Filter out the current year from previous year options
                prev_year_options_asin = [y for y in available_years_asin if y != current_year_asin]
                if prev_year_options_asin:
                    previous_year_asin = st.selectbox(
                        "📅 Previous Year (to compare against)",
                        prev_year_options_asin,
                        index=0,
                        key="asin_previous_year"
                    )
                else:
                    previous_year_asin = None
                    st.warning("No other year available for comparison")
            
            if previous_year_asin:
                # Filter data by years
                current_year_data_asin = processed_df[processed_df['Year'] == current_year_asin]
                previous_year_data_asin = processed_df[processed_df['Year'] == previous_year_asin]
                
                # Create ASIN pivots for each year
                current_asin_pivot = pd.pivot_table(
                    current_year_data_asin,
                    index=['Asin', 'Brand'],
                    values=['Quantity', 'Invoice Amount'],
                    aggfunc='sum'
                ).reset_index()
                current_asin_pivot.columns = ['Asin', 'Brand', f'Invoice Amount ({current_year_asin})', f'Quantity ({current_year_asin})']
                
                previous_asin_pivot = pd.pivot_table(
                    previous_year_data_asin,
                    index=['Asin', 'Brand'],
                    values=['Quantity', 'Invoice Amount'],
                    aggfunc='sum'
                ).reset_index()
                previous_asin_pivot.columns = ['Asin', 'Brand', f'Invoice Amount ({previous_year_asin})', f'Quantity ({previous_year_asin})']
                
                # Merge the two pivots
                asin_comparison = pd.merge(
                    previous_asin_pivot,
                    current_asin_pivot,
                    on=['Asin', 'Brand'],
                    how='outer'
                ).fillna(0)
                
                # Calculate differences and percentage changes
                asin_comparison['Qty Difference'] = asin_comparison[f'Quantity ({current_year_asin})'] - asin_comparison[f'Quantity ({previous_year_asin})']
                asin_comparison['Qty % Change'] = asin_comparison.apply(
                    lambda row: ((row[f'Quantity ({current_year_asin})'] - row[f'Quantity ({previous_year_asin})']) / row[f'Quantity ({previous_year_asin})'] * 100) 
                    if row[f'Quantity ({previous_year_asin})'] != 0 else (100 if row[f'Quantity ({current_year_asin})'] > 0 else 0), axis=1
                )
                
                asin_comparison['Amount Difference'] = asin_comparison[f'Invoice Amount ({current_year_asin})'] - asin_comparison[f'Invoice Amount ({previous_year_asin})']
                asin_comparison['Amount % Change'] = asin_comparison.apply(
                    lambda row: ((row[f'Invoice Amount ({current_year_asin})'] - row[f'Invoice Amount ({previous_year_asin})']) / row[f'Invoice Amount ({previous_year_asin})'] * 100) 
                    if row[f'Invoice Amount ({previous_year_asin})'] != 0 else (100 if row[f'Invoice Amount ({current_year_asin})'] > 0 else 0), axis=1
                )
                
                # Reorder columns
                asin_comparison = asin_comparison[[
                    'Asin', 'Brand',
                    f'Quantity ({previous_year_asin})', f'Quantity ({current_year_asin})', 'Qty Difference', 'Qty % Change',
                    f'Invoice Amount ({previous_year_asin})', f'Invoice Amount ({current_year_asin})', 'Amount Difference', 'Amount % Change'
                ]]
                
                # Sort by current year quantity descending
                asin_comparison = asin_comparison.sort_values(by=f'Quantity ({current_year_asin})', ascending=False)
                
                # Add Grand Total row
                grand_total_asin = pd.DataFrame({
                    'Asin': ['Grand Total'],
                    'Brand': [''],
                    f'Quantity ({previous_year_asin})': [asin_comparison[f'Quantity ({previous_year_asin})'].sum()],
                    f'Quantity ({current_year_asin})': [asin_comparison[f'Quantity ({current_year_asin})'].sum()],
                    'Qty Difference': [asin_comparison['Qty Difference'].sum()],
                    'Qty % Change': [
                        (asin_comparison[f'Quantity ({current_year_asin})'].sum() - asin_comparison[f'Quantity ({previous_year_asin})'].sum()) / 
                        asin_comparison[f'Quantity ({previous_year_asin})'].sum() * 100 if asin_comparison[f'Quantity ({previous_year_asin})'].sum() != 0 else 0
                    ],
                    f'Invoice Amount ({previous_year_asin})': [asin_comparison[f'Invoice Amount ({previous_year_asin})'].sum()],
                    f'Invoice Amount ({current_year_asin})': [asin_comparison[f'Invoice Amount ({current_year_asin})'].sum()],
                    'Amount Difference': [asin_comparison['Amount Difference'].sum()],
                    'Amount % Change': [
                        (asin_comparison[f'Invoice Amount ({current_year_asin})'].sum() - asin_comparison[f'Invoice Amount ({previous_year_asin})'].sum()) / 
                        asin_comparison[f'Invoice Amount ({previous_year_asin})'].sum() * 100 if asin_comparison[f'Invoice Amount ({previous_year_asin})'].sum() != 0 else 0
                    ]
                })
                asin_comparison = pd.concat([asin_comparison, grand_total_asin], ignore_index=True)
                
                # Format display dataframe
                display_asin_comparison = asin_comparison.copy()
                display_asin_comparison[f'Quantity ({previous_year_asin})'] = display_asin_comparison[f'Quantity ({previous_year_asin})'].apply(lambda x: f"{x:,.0f}")
                display_asin_comparison[f'Quantity ({current_year_asin})'] = display_asin_comparison[f'Quantity ({current_year_asin})'].apply(lambda x: f"{x:,.0f}")
                display_asin_comparison['Qty Difference'] = display_asin_comparison['Qty Difference'].apply(lambda x: f"{x:+,.0f}")
                display_asin_comparison['Qty % Change'] = display_asin_comparison['Qty % Change'].apply(lambda x: f"{x:+.2f}%")
                display_asin_comparison[f'Invoice Amount ({previous_year_asin})'] = display_asin_comparison[f'Invoice Amount ({previous_year_asin})'].apply(lambda x: f"₹{x:,.2f}")
                display_asin_comparison[f'Invoice Amount ({current_year_asin})'] = display_asin_comparison[f'Invoice Amount ({current_year_asin})'].apply(lambda x: f"₹{x:,.2f}")
                display_asin_comparison['Amount Difference'] = display_asin_comparison['Amount Difference'].apply(lambda x: f"₹{x:+,.2f}")
                display_asin_comparison['Amount % Change'] = display_asin_comparison['Amount % Change'].apply(lambda x: f"{x:+.2f}%")
                
                # Show summary metrics
                st.markdown(f"### Comparison: {current_year_asin} vs {previous_year_asin}")
                metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
                
                total_qty_change_asin = asin_comparison.iloc[-1]['Qty Difference']
                total_qty_pct_asin = asin_comparison.iloc[-1]['Qty % Change']
                total_amt_change_asin = asin_comparison.iloc[-1]['Amount Difference']
                total_amt_pct_asin = asin_comparison.iloc[-1]['Amount % Change']
                
                with metric_col1:
                    st.metric("Total Qty Change", f"{total_qty_change_asin:+,.0f}", f"{total_qty_pct_asin:+.2f}%")
                with metric_col2:
                    st.metric("Total Amount Change", f"₹{total_amt_change_asin:+,.0f}", f"{total_amt_pct_asin:+.2f}%")
                with metric_col3:
                    st.metric(f"Unique ASINs in {current_year_asin}", f"{len(current_year_data_asin['Asin'].dropna().unique()):,}")
                with metric_col4:
                    st.metric(f"Unique ASINs in {previous_year_asin}", f"{len(previous_year_data_asin['Asin'].dropna().unique()):,}")
                
                st.dataframe(display_asin_comparison, width='stretch', height=600)
                
                # Download link (base64 approach for Streamlit Cloud)
                st.markdown(create_download_link(asin_comparison, f"asin_comparison_{current_year_asin}_vs_{previous_year_asin}.xlsx", "Download ASIN Comparison Excel"), unsafe_allow_html=True)
        else:
            st.warning("⚠️ Need at least 2 years of data for comparison. Please upload data from multiple years.")

else:
    # Landing page with instructions
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style='text-align: center; padding: 2rem;'>
            <h2>👋 Welcome to Sales Data Analysis Dashboard</h2>
            <p style='font-size: 1.1rem; color: #666;'>
                Upload your files to get started with comprehensive sales analysis
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown("""
        ### 📋 Getting Started
        
        **Step 1:** Upload Files
        - Navigate to the sidebar (left side)
        - Upload all B2B and B2C report ZIP files
        - Upload the PM (Product Master) Excel file
        
        **Step 2:** Select Filters
        - Choose time period: All Data, Quarter, or Month
        - **Quarter Definitions:**
          - 🌱 Q1: January - March
          - ☀️ Q2: April - June  
          - 🍂 Q3: July - September
          - ❄️ Q4: October - December
        - Apply additional filters for Brand and Brand Manager
        
        **Step 3:** Analyze Data
        - View Summary statistics and trends
        - Explore Brand-wise analysis
        - Check ASIN-level details
        - Export filtered data as CSV
        
        ### ✨ Key Features
        
        | Feature | Description |
        |---------|-------------|
        | 📊 **Summary Dashboard** | Overview metrics, trends, and KPIs |
        | 🏢 **Brand Analysis** | Sales grouped by brand with totals |
        | 📦 **ASIN Analysis** | Detailed product-level breakdown |
        | 📋 **Raw Data Export** | Customizable data export options |
        | 🎯 **Smart Filters** | Quarter, Month, Brand, and Manager filters |
        
        ### 💡 Tips
        
        - Use Quarter View for quarterly business reviews
        - Use Month View for detailed monthly analysis
        - Download reports for offline analysis
        - Apply multiple filters for targeted insights
        
        """)
        
        st.markdown("---")
        
        st.info("👈 **Ready to begin?** Upload your files using the sidebar on the left!")

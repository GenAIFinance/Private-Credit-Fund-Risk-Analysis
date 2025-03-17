# Private-Credit-Fund-Risk-Analysis
# PCO Exposure Analysis Tool

A Python-based tool for analyzing Private Credit Opportunities (PCO) fund exposure across multiple dimensions, generating comprehensive Excel reports with charts and formatted outputs.

![Excel Report Example](https://via.placeholder.com/800x400.png?text=Sample+Excel+Output)

## Features

- **Automated Data Processing**
  - Processes RISK tab data from NAV summary files
  - Creates 4 new dimensions using issuer mapping:
    - Issuer Name-N
    - Moody's Industry-N
    - Lien-N
    - Regional-N
- **Multi-Dimensional Analysis**
  - Issuer exposure breakdown
  - Lien type distribution
  - Regional exposure
  - Industry sector analysis
- **Automated Reporting**
  - Excel report generation with:
  - Formatted tables with currency/percentage formatting
  - Interactive pie charts
  - Consolidated PCO/SMA views
- **Smart Configuration**
  - Flexible fund configuration
  - Custom exclusion lists
  - Fixed total commitment tracking ($1.67B)

## Installation

1. **Prerequisites**
   - Python 3.8+
   - pandas
   - openpyxl
   - numpy

2. **Setup**
```bash
git clone https://github.com/yourusername/pco-analysis.git
cd pco-analysis
pip install -r requirements.txt
# =================================================================
# USER CONFIGURATION for PCO Exposure
# =================================================================
DATE = "20241231"
INPUT_PATH = r"input_path"
OUTPUT_PATH = r"output_path"
ISSUER_MAPPING_FILE = r"C:\Users\Desktop\file\issuer_mapping.xlsx"

FUND_CONFIG = {
    "DLF1_CAYTOP": {"nav": 785_883_385, "type": "PCO", "return": 0.0107},
    "DLF1_DELATOP": {"nav": 323_768_715, "type": "PCO", "return": 0.0097},
    "DLF1_RAIFLEVTOP": {"nav": 55_038_599, "type": "PCO", "return": 0.0103},
    "BENM": {"nav": 30_546_889, "type": "SMA", "return": 0.0100}
}

TOTAL_COMMITMENT = 1_670_000_000  # $1.67B fixed for all funds

EXCLUSIONS = {
    "exposure": ["FX Forwards", "Cash", "Fee and Expense", "ABL", "PCO Loan", "Repo"]
}

# =================================================================
# CORE IMPLEMENTATION for PCO Exposure
# =================================================================
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

def load_issuer_mapping():
    """Load and validate issuer mapping file, handling duplicate 'Name' entries"""
    try:
        df = pd.read_excel(ISSUER_MAPPING_FILE)
        # Debug: Print column names to verify structure
        print(f"Columns in issuer_mapping.xlsx: {df.columns.tolist()}")
        
        required_cols = ['Name', 'Issuer Name', "Moody's Industry", 'LIEN', 'Regional']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns in mapping file: {', '.join(missing)}")
        
        # Check for duplicate 'Name' values
        duplicates = df[df['Name'].duplicated(keep=False)]['Name'].unique()
        if len(duplicates) > 0:
            print(f"Warning: Duplicate 'Name' values found in {ISSUER_MAPPING_FILE}: {duplicates}")
            print(f"Dropping duplicates, keeping the first occurrence. Please review and clean {ISSUER_MAPPING_FILE}.")
            df = df.drop_duplicates(subset=['Name'], keep='first')
        
        return df.set_index("Name")
    except Exception as e:
        print(f"Mapping load error: {str(e)}")
        return None

def process_exposure(file_path, issuer_map, fund_name):
    """Process exposure analysis with 4 new variables derived from issuer mapping, using 'Name' for lookups, with debugging"""
    try:
        # Load and prepare data from RISK tab
        print(f"Loading RISK tab for {fund_name} from {file_path}")
        df = pd.read_excel(file_path, sheet_name="RISK")
        
        # Debug: Print the columns and first 10 rows to verify structure
        print(f"Columns in RISK tab for {fund_name}: {df.columns.tolist()}")
        print(f"First 10 rows of RISK tab for {fund_name}:\n{df.head(10).to_string()}")
        print(f"Unique Names for {fund_name}: {df['Name'].unique()[:10]}")  # Sample of unique names
        
        # Use 'Name' as the lookup column for all funds
        name_column = 'Name'
        if name_column not in df.columns:
            raise ValueError(f"'Name' column not found in RISK tab for {fund_name}")
        
        print(f"Using '{name_column}' as the lookup column for {fund_name}")
        if df[name_column].isna().all():
            raise ValueError(f"No non-null values in '{name_column}' column in RISK tab for {fund_name}")
        
        # Create new variables by looking up in issuer_mapping
        if issuer_map is not None:
            issuer_map_reset = issuer_map.reset_index()
            
            # 1. Issuer name-N
            df['Issuer name-N'] = df[name_column].map(issuer_map_reset.set_index('Name')['Issuer Name']).fillna('New Investment')
            new_investments = df[df['Issuer name-N'] == 'New Investment'][name_column].unique()
            if len(new_investments) > 0:
                print(f"Warning: 'New Investment' found for {fund_name} - unmapped Names: {new_investments}")
                print(f"Please update {ISSUER_MAPPING_FILE} with these Names and their details.")
            
            # 2. Moody's Industry-N
            df["Moody's Industry-N"] = df[name_column].map(issuer_map_reset.set_index('Name')["Moody's Industry"]).fillna('New Investment')
            new_investments = df[df["Moody's Industry-N"] == 'New Investment'][name_column].unique()
            if len(new_investments) > 0:
                print(f"Warning: 'New Investment' found for {fund_name} - unmapped Names: {new_investments}")
                print(f"Please update {ISSUER_MAPPING_FILE} with these Names and their details.")
            
            # 3. Lien-N
            if 'LIEN' in issuer_map_reset.columns:
                df['Lien-N'] = df[name_column].map(issuer_map_reset.set_index('Name')['LIEN']).fillna('New Investment')
                print(f"Debug - {fund_name} Lien-N mapping successful. Sample values: {df['Lien-N'].head().tolist()}")
            else:
                print(f"Error: 'LIEN' column missing in issuer_mapping.xlsx. Setting 'Lien-N' to 'Unknown'")
                df['Lien-N'] = 'Unknown'
            new_investments = df[df['Lien-N'] == 'New Investment'][name_column].unique()
            if len(new_investments) > 0:
                print(f"Warning: 'New Investment' found for {fund_name} - unmapped Names: {new_investments}")
                print(f"Please update {ISSUER_MAPPING_FILE} with these Names and their details.")
            
            # 4. Regional-N
            df['Regional-N'] = df[name_column].map(issuer_map_reset.set_index('Name')['Regional']).fillna('New Investment')
            new_investments = df[df['Regional-N'] == 'New Investment'][name_column].unique()
            if len(new_investments) > 0:
                print(f"Warning: 'New Investment' found for {fund_name} - unmapped Names: {new_investments}")
                print(f"Please update {ISSUER_MAPPING_FILE} with these Names and their details.")
        
        # Handle FX Forwards and Cash exclusions for all dimensions
        exclusion_mask = (
            df['Issuer name-N'].isin(EXCLUSIONS["exposure"]) |
            df["Moody's Industry-N"].isin(EXCLUSIONS["exposure"]) |
            (df['Lien-N'].isin(EXCLUSIONS["exposure"]) if 'Lien-N' in df.columns else False) |
            df['Regional-N'].isin(EXCLUSIONS["exposure"])
        )
        df = df[~exclusion_mask]
        if df.empty:
            print(f"No exposure data for {fund_name} after excluding Cash and FX Forwards")
            return None
        
        # Debug: Check Lien-N values after exclusions
        if 'Lien-N' in df.columns:
            print(f"Debug - {fund_name} Lien-N unique values after exclusions: {df['Lien-N'].unique()}")
            print(f"Debug - {fund_name} Lien-N value counts after exclusions:\n{df['Lien-N'].value_counts().to_string()}")
        else:
            print(f"Warning: 'Lien-N' column missing after exclusions for {fund_name}")

        # Calculate Market Exposure with corrected formula and debugging
        required_cols = ['Quantity', 'Mkt Price', 'Fx Rate', 'Accrued Interest Book']
        if not all(col in df.columns for col in required_cols):
            raise ValueError(f"Missing required columns {required_cols} in RISK tab for {fund_name}")
        
        # Convert columns to numeric, coercing errors to NaN
        for col in required_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Debug: Print intermediate values
        print(f"Debug - Quantity sample: {df['Quantity'].head().tolist()}")
        print(f"Debug - Mkt Price sample: {df['Mkt Price'].head().tolist()}")
        print(f"Debug - Fx Rate sample: {df['Fx Rate'].head().tolist()}")
        print(f"Debug - Accrued Interest Book sample: {df['Accrued Interest Book'].head().tolist()}")
        
        # Corrected Market Exposure formula: (Quantity * Mkt Price * Fx Rate / 100 + Accrued Interest Book)
        df['Market Exposure'] = (df['Quantity'] * df['Mkt Price'] * df['Fx Rate'] / 100 + df['Accrued Interest Book']).fillna(0)
        print(f"Debug - Calculated Market Exposure sample: {df['Market Exposure'].head().tolist()}")
        
        # Debug: Verify final DataFrame before grouping
        print(f"Final DataFrame for {fund_name} before grouping:\n{df.head().to_string()}")
        
        # Group and calculate exposures using new dimensions
        output = {}
        
        # 1. Issuer Exposure
        issuer_df = df.groupby(['Issuer name-N', "Moody's Industry-N"], as_index=False).agg({
            'Market Exposure': 'sum',
            'Lien-N': 'first',
            'Regional-N': 'first'
        })
        if not issuer_df.empty:
            issuer_df = calculate_percentages(issuer_df, fund_name)
            issuer_df['Fund Name'] = fund_name
            output['Issuer'] = issuer_df
        else:
            print(f"Warning: Issuer Exposure is empty for {fund_name} after grouping")
        
        # 2. Lien Exposure (Limit to 1st, 2nd, 3rd, case-insensitive with variants)
        if 'Lien-N' in df.columns:
            # Debug: Before filtering
            print(f"Debug - {fund_name} DataFrame shape before Lien-N filtering: {df.shape}")
            # Use a more flexible regex to catch variations (e.g., "1st Lien", "First", "1")
            lien_df_filtered = df[df['Lien-N'].str.lower().str.contains('1st|2nd|3rd|first|second|third|[1-3]', na=False)]
            print(f"Debug - {fund_name} DataFrame shape after Lien-N filtering: {lien_df_filtered.shape}")
            print(f"Debug - {fund_name} Filtered Lien-N unique values: {lien_df_filtered['Lien-N'].unique()}")
            
            if not lien_df_filtered.empty:
                lien_df = lien_df_filtered.groupby('Lien-N', as_index=False).agg({'Market Exposure': 'sum'})
                print(f"Debug - {fund_name} Lien DataFrame after grouping: {lien_df}")
                if not lien_df.empty:
                    lien_df = calculate_percentages(lien_df, fund_name)
                    lien_df['Fund Name'] = fund_name
                    output['LIEN'] = lien_df
                else:
                    print(f"Warning: Lien Exposure is empty for {fund_name} after grouping")
            else:
                print(f"Warning: Lien Exposure is empty for {fund_name} after filtering")
        else:
            print(f"Warning: 'Lien-N' column missing in DataFrame for {fund_name}. Skipping Lien Exposure.")
        
        # 3. Regional Exposure (Limit to US, Euro, case-insensitive)
        region_df = df[df['Regional-N'].str.lower().isin(['us', 'euro'])].groupby('Regional-N', as_index=False).agg({'Market Exposure': 'sum'})
        if not region_df.empty:
            region_df = calculate_percentages(region_df, fund_name)
            region_df['Fund Name'] = fund_name
            output['Region'] = region_df
        else:
            print(f"Warning: Region Exposure is empty for {fund_name} after filtering and grouping")
            print(f"Debug - Regional-N values: {df['Regional-N'].unique()}")
        
        # 4. Industry Exposure
        industry_df = df.groupby("Moody's Industry-N", as_index=False).agg({'Market Exposure': 'sum'})
        if not industry_df.empty:
            industry_df = calculate_percentages(industry_df, fund_name)
            industry_df['Fund Name'] = fund_name
            output['Industry'] = industry_df
        else:
            print(f"Warning: Industry Exposure is empty for {fund_name} after grouping")
        
        # Debug: Print output shapes for each dimension if they exist
        if 'Issuer' in output:
            print(f"Output for {fund_name} - Issuer Exposure shape: {output['Issuer'].shape}")
        if 'LIEN' in output:
            print(f"Output for {fund_name} - Lien Exposure shape: {output['LIEN'].shape}")
        if 'Region' in output:
            print(f"Output for {fund_name} - Region Exposure shape: {output['Region'].shape}")
        if 'Industry' in output:
            print(f"Output for {fund_name} - Industry Exposure shape: {output['Industry'].shape}")

        return {"exposure": output}

    except Exception as e:
        print(f"Exposure processing failed for {fund_name}: {str(e)}")
        print(f"DataFrame columns at failure for {fund_name}: {df.columns.tolist() if 'df' in locals() else 'DataFrame not created'}")
        print(f"Issuer map columns: {issuer_map.columns.tolist() if issuer_map is not None else 'Issuer map not loaded'}")
        return None

def calculate_percentages(df, fund_name):
    """Calculate exposure percentages for each dimension using fund-specific NAV without totals in grouping"""
    total_exposure = df['Market Exposure'].sum()
    config = FUND_CONFIG[fund_name]
    nav = config["nav"]
    commitment = TOTAL_COMMITMENT  # $1.67B fixed for all funds
    
    # Calculate percentages for individual rows
    if total_exposure > 0:
        df['Exposure % Gross Total'] = df['Market Exposure'].apply(lambda x: f"{round((x / total_exposure * 100), 2)}%" if total_exposure > 0 else "0.00%")
        df['Exposure % Fund NAV'] = df['Market Exposure'].apply(lambda x: f"{round((x / nav * 100), 2)}%" if nav > 0 else "0.00%")
        df['Exposure % Total Commitment'] = df['Market Exposure'].apply(lambda x: f"{round((x / commitment * 100), 2)}%")
    else:
        df['Exposure % Gross Total'] = "0.00%"
        df['Exposure % Fund NAV'] = "0.00%"
        df['Exposure % Total Commitment'] = "0.00%"
    
    return df

def add_total_row(df, fund_name):
    """Add a total row to the DataFrame without aggregating as a category"""
    total_exposure = df['Market Exposure'].sum()
    config = FUND_CONFIG[fund_name]
    nav = config["nav"]
    commitment = TOTAL_COMMITMENT
    
    totals = pd.Series({
        'Issuer name-N' if 'Issuer name-N' in df.columns else df.columns[0]: 'Total',
        'Market Exposure': total_exposure,
        'Exposure % Gross Total': '100.00%' if total_exposure > 0 else "0.00%",
        'Exposure % Fund NAV': f"{round((total_exposure / nav * 100), 2)}%" if nav > 0 and total_exposure > 0 else "0.00%",
        'Exposure % Total Commitment': f"{round((total_exposure / commitment * 100), 2)}%"
    })
    if "Moody's Industry-N" in df.columns:
        totals["Moody's Industry-N"] = ''
    if 'Lien-N' in df.columns:
        totals['Lien-N'] = ''
    if 'Regional-N' in df.columns:
        totals['Regional-N'] = ''
    if 'Fund Name' in df.columns:
        totals['Fund Name'] = df['Fund Name'].iloc[0] if not df.empty else ''
    
    return pd.concat([df, pd.DataFrame([totals])], ignore_index=True)

def write_exposure_sheet(ws, title, df, start_row):
    """Write a single exposure table to a worksheet with total row handling"""
    if not df.empty:
        # Write header
        ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=14)
        start_row += 1
        
        # Write headers
        headers = df.columns.tolist()
        for col, header in enumerate(headers, 1):
            ws.cell(row=start_row, column=col, value=header).fill = PatternFill(start_color="00008B", fill_type="solid")
            ws.cell(row=start_row, column=col, value=header).font = Font(bold=True, color="FFFFFF")
        
        # Write data
        for i, row in df.iterrows():
            for col, value in enumerate(row, 1):
                cell = ws.cell(row=start_row + i + 1, column=col, value=value)
                if isinstance(value, (int, float)):
                    col_idx = get_column_letter(col)
                    if col_idx in ['F', 'G', 'H', 'I']:  # Market Exposure and percentage columns for Issuer, Lien, Region, Industry
                        if col_idx == 'F':  # Market Exposure
                            cell.number_format = '$#,##0'
                        else:  # Percentage columns
                            cell.number_format = '0.00%'  # Two decimal places
                elif isinstance(value, str) and '%' in value:
                    cell.number_format = '0.00%'  # Ensure percentage strings are formatted correctly
                if i == len(df) - 1 and value == 'Total':  # Bold the total row
                    cell.font = Font(bold=True)
        
        start_row += len(df) + 2
        return start_row
    return start_row

def create_pie_charts(ws, data, start_row):
    """Create pie charts for the combined exposure across LIEN, Region, and Industry dimensions for PCO funds only"""
    # Debug: Print data used for charts
    print(f"Data for charts shape: {data.shape}")
    print(f"Unique Fund Names in chart data: {data['Fund Name'].unique()}")
    
    # Filter data for PCO funds only (exclude BENM)
    pco_data = data[data['Fund Name'].isin(["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"])]
    print(f"PCO data shape after filtering: {pco_data.shape}")
    
    # Exclude "New Investment" from all dimensions
    pco_data = pco_data[pco_data['Lien-N'] != 'New Investment'] if 'Lien-N' in pco_data.columns else pco_data
    pco_data = pco_data[pco_data['Regional-N'] != 'New Investment']
    pco_data = pco_data[pco_data["Moody's Industry-N"] != 'New Investment']
    print(f"PCO data shape after excluding New Investment: {pco_data.shape}")
    
    dimensions = {
        'LIEN': pco_data.groupby('Lien-N')['Market Exposure'].sum().reset_index(name='Market Exposure') if 'Lien-N' in pco_data.columns else pd.DataFrame(),
        'Region': pco_data.groupby('Regional-N')['Market Exposure'].sum().reset_index(name='Market Exposure'),
        'Industry': pco_data.groupby("Moody's Industry-N")['Market Exposure'].sum().reset_index(name='Market Exposure')
    }
    
    chart_start = start_row
    for dim, series in dimensions.items():
        if not series.empty:
            print(f"Chart data for {dim}: {series}")
            labels = series.index if dim != 'Industry' else series["Moody's Industry-N"]
            values = series['Market Exposure']
            
            chart = PieChart()
            chart.title = f"{dim} Exposure"
            chart.add_data(Reference(ws, min_col=2, min_row=chart_start, max_row=chart_start+len(values)-1))
            chart.set_categories(Reference(ws, min_col=1, min_row=chart_start+1, max_row=chart_start+len(labels)))
            chart.dataLabels = DataLabelList(showPercent=True)
            chart.height = 10
            chart.width = 15
            
            # Write data for chart with attachment-compatible structure
            ws.cell(row=chart_start, column=1, value=dim)
            if dim == 'LIEN':
                for i, (label, value) in enumerate(zip(series['Lien-N'], values), 1):
                    ws.cell(row=chart_start+i, column=1, value=label)
                    ws.cell(row=chart_start+i, column=2, value=value)
                    ws.cell(row=chart_start+i, column=3, value=f"{round((value / values.sum() * 100), 2)}%")
                    ws.cell(row=chart_start+i, column=4, value='N/A')
                    ws.cell(row=chart_start+i, column=5, value=f"{round((value / TOTAL_COMMITMENT * 100), 2)}%")
            elif dim == 'Region':
                for i, (label, value) in enumerate(zip(series['Regional-N'], values), 1):
                    ws.cell(row=chart_start+i, column=1, value=label)
                    ws.cell(row=chart_start+i, column=2, value=value)
                    ws.cell(row=chart_start+i, column=3, value=f"{round((value / values.sum() * 100), 2)}%")
                    ws.cell(row=chart_start+i, column=4, value='N/A')
                    ws.cell(row=chart_start+i, column=5, value=f"{round((value / TOTAL_COMMITMENT * 100), 2)}%")
            elif dim == 'Industry':
                for i, (label, value) in enumerate(zip(series["Moody's Industry-N"], values), 1):
                    ws.cell(row=chart_start+i, column=1, value=label)
                    ws.cell(row=chart_start+i, column=2, value=value)
                    ws.cell(row=chart_start+i, column=3, value=f"{round((value / values.sum() * 100), 2)}%")
                    ws.cell(row=chart_start+i, column=4, value='N/A')
                    ws.cell(row=chart_start+i, column=5, value=f"{round((value / TOTAL_COMMITMENT * 100), 2)}%")
            
            # Add total row
            ws.cell(row=chart_start+len(values)+1, column=1, value='Total')
            ws.cell(row=chart_start+len(values)+1, column=2, value=values.sum())
            ws.cell(row=chart_start+len(values)+1, column=3, value='100.00%')
            ws.cell(row=chart_start+len(values)+1, column=4, value='N/A')
            ws.cell(row=chart_start+len(values)+1, column=5, value=f"{round((values.sum() / TOTAL_COMMITMENT * 100), 2)}%")
            
            ws.add_chart(chart, f"D{chart_start}")
            chart_start += len(values) + 15
    
    return chart_start

def generate_report(data):
    """Generate exposure report for PCO and SMA funds"""
    output_file = os.path.join(OUTPUT_PATH, f"PCO_Exposure_Analysis_{DATE}.xlsx")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        workbook = writer.book
        
        # Ensure at least one sheet is created
        if not workbook.sheetnames:
            workbook.create_sheet("Default_Sheet")
        
        # 1. Separate sheets for each exposure type
        exposures = {}
        for fund in FUND_CONFIG:
            if fund in data and data[fund]["exposure"]:
                for dim in ['Issuer', 'LIEN', 'Region', 'Industry']:
                    if data[fund]["exposure"].get(dim) is not None:
                        if dim not in exposures:
                            exposures[dim] = []
                        df = data[fund]["exposure"][dim].copy()
                        df['Fund Name'] = fund
                        exposures[dim].append(df)
        
        # Debug: Check exposures dictionary
        print(f"Exposures dictionary contents: {list(exposures.keys())}")
        for dim in exposures:
            print(f"Dimension {dim} has {len(exposures[dim])} DataFrames")
            for df in exposures[dim]:
                print(f"  - Fund: {df['Fund Name'].iloc[0] if not df.empty else 'Empty DataFrame'}, Shape: {df.shape}")

        if exposures:
            # Issuer Exposure Sheet
            ws_issuer = workbook.create_sheet("Issuer_Exposure")
            start_row = 1
            all_funds = ["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP", "BENM"]
            for fund in all_funds:
                if exposures.get('Issuer') and any(df['Fund Name'].eq(fund).any() for df in exposures['Issuer']):
                    combined_df = pd.concat([df for df in exposures['Issuer'] if df['Fund Name'].eq(fund).any()], ignore_index=True)
                    combined_df = add_total_row(combined_df, fund)
                    start_row = write_exposure_sheet(ws_issuer, f"Issuer Exposure - {fund}", combined_df, start_row)
            if exposures.get('Issuer'):
                issuer_pco_dfs = [df for df in exposures['Issuer'] if df['Fund Name'].isin(["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"]).any()]
                print(f"Line 311 - Issuer PCO DataFrames to concatenate: {len(issuer_pco_dfs)}")
                if issuer_pco_dfs:
                    pco_data = pd.concat(issuer_pco_dfs, ignore_index=True)
                    if not pco_data.empty:
                        # Re-group by Issuer name-N and Moody's Industry-N to sum Market Exposure
                        pco_data = pco_data.groupby(['Issuer name-N', "Moody's Industry-N"], as_index=False).agg({
                            'Market Exposure': 'sum',
                            'Lien-N': 'first' if 'Lien-N' in pco_data.columns else lambda x: 'Unknown',
                            'Regional-N': 'first'
                        })
                        pco_data = calculate_percentages(pco_data, "DLF1_CAYTOP")  # Use CAYTOP's NAV for consistency
                        pco_data['Fund Name'] = "Combined PCO"
                        pco_data = add_total_row(pco_data, "DLF1_CAYTOP")
                        start_row = write_exposure_sheet(ws_issuer, "Issuer Exposure - Combined PCO", pco_data, start_row)
                else:
                    print("No Issuer PCO data to concatenate")

            # Lien Exposure Sheet
            ws_lien = workbook.create_sheet("Lien_Exposure")
            start_row = 1
            for fund in all_funds:
                if exposures.get('LIEN') and any(df['Fund Name'].eq(fund).any() for df in exposures['LIEN']):
                    combined_df = pd.concat([df for df in exposures['LIEN'] if df['Fund Name'].eq(fund).any()], ignore_index=True)
                    print(f"Debug - Lien Exposure for {fund}: {combined_df}")
                    combined_df = add_total_row(combined_df, fund)
                    start_row = write_exposure_sheet(ws_lien, f"Lien Exposure - {fund}", combined_df, start_row)
            if exposures.get('LIEN'):
                lien_pco_dfs = [df for df in exposures['LIEN'] if df['Fund Name'].isin(["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"]).any()]
                print(f"Line 347 - Lien PCO DataFrames to concatenate: {len(lien_pco_dfs)}")
                if lien_pco_dfs:
                    pco_data = pd.concat(lien_pco_dfs, ignore_index=True)
                    if not pco_data.empty:
                        # Re-group by Lien-N to sum Market Exposure
                        if 'Lien-N' in pco_data.columns:
                            pco_data = pco_data.groupby('Lien-N', as_index=False).agg({'Market Exposure': 'sum'})
                            print(f"Debug - Combined PCO Lien Exposure after grouping: {pco_data}")
                        else:
                            print("Warning: 'Lien-N' column missing in combined PCO data for Lien Exposure")
                            pco_data = pd.DataFrame(columns=['Lien-N', 'Market Exposure'])
                        pco_data = calculate_percentages(pco_data, "DLF1_CAYTOP")
                        pco_data['Fund Name'] = "Combined PCO"
                        pco_data = add_total_row(pco_data, "DLF1_CAYTOP")
                        start_row = write_exposure_sheet(ws_lien, "Lien Exposure - Combined PCO", pco_data, start_row)
                else:
                    print("No Lien PCO data to concatenate")

            # Region Exposure Sheet
            ws_region = workbook.create_sheet("Region_Exposure")
            start_row = 1
            for fund in all_funds:
                if exposures.get('Region') and any(df['Fund Name'].eq(fund).any() for df in exposures['Region']):
                    combined_df = pd.concat([df for df in exposures['Region'] if df['Fund Name'].eq(fund).any()], ignore_index=True)
                    combined_df = add_total_row(combined_df, fund)
                    start_row = write_exposure_sheet(ws_region, f"Region Exposure - {fund}", combined_df, start_row)
            if exposures.get('Region'):
                region_pco_dfs = [df for df in exposures['Region'] if df['Fund Name'].isin(["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"]).any()]
                print(f"Line 356 - Region PCO DataFrames to concatenate: {len(region_pco_dfs)}")
                if region_pco_dfs:
                    pco_data = pd.concat(region_pco_dfs, ignore_index=True)
                    if not pco_data.empty:
                        # Re-group by Regional-N to sum Market Exposure
                        pco_data = pco_data.groupby('Regional-N', as_index=False).agg({'Market Exposure': 'sum'})
                        pco_data = calculate_percentages(pco_data, "DLF1_CAYTOP")
                        pco_data['Fund Name'] = "Combined PCO"
                        pco_data = add_total_row(pco_data, "DLF1_CAYTOP")
                        start_row = write_exposure_sheet(ws_region, "Region Exposure - Combined PCO", pco_data, start_row)
                else:
                    print("No Region PCO data to concatenate")

            # Industry Exposure Sheet
            ws_industry = workbook.create_sheet("Industry_Exposure")
            start_row = 1
            for fund in all_funds:
                if exposures.get('Industry') and any(df['Fund Name'].eq(fund).any() for df in exposures['Industry']):
                    combined_df = pd.concat([df for df in exposures['Industry'] if df['Fund Name'].eq(fund).any()], ignore_index=True)
                    combined_df = add_total_row(combined_df, fund)
                    start_row = write_exposure_sheet(ws_industry, f"Industry Exposure - {fund}", combined_df, start_row)
            if exposures.get('Industry'):
                industry_pco_dfs = [df for df in exposures['Industry'] if df['Fund Name'].isin(["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"]).any()]
                print(f"Line 404 - Industry PCO DataFrames to concatenate: {len(industry_pco_dfs)}")
                if industry_pco_dfs:
                    pco_data = pd.concat(industry_pco_dfs, ignore_index=True)
                    if not pco_data.empty:
                        # Re-group by Moody's Industry-N to sum Market Exposure
                        pco_data = pco_data.groupby("Moody's Industry-N", as_index=False).agg({'Market Exposure': 'sum'})
                        pco_data = calculate_percentages(pco_data, "DLF1_CAYTOP")
                        pco_data['Fund Name'] = "Combined PCO"
                        pco_data = add_total_row(pco_data, "DLF1_CAYTOP")
                        start_row = write_exposure_sheet(ws_industry, "Industry Exposure - Combined PCO", pco_data, start_row)
                else:
                    print("No Industry PCO data to concatenate")

            # Top 10 Issuers Exposure Sheet
            ws_top10 = workbook.create_sheet("Top_10_Issuers_Exposure")
            start_row = 1
            
            # For each fund (DLF1_CAYTOP, DLF1_DELATOP, DLF1_RAIFLEVTOP, BENM)
            for fund in all_funds:
                if exposures.get('Issuer') and any(df['Fund Name'].eq(fund).any() for df in exposures['Issuer']):
                    fund_df = pd.concat([df for df in exposures['Issuer'] if df['Fund Name'].eq(fund).any()], ignore_index=True)
                    if not fund_df.empty:
                        # Exclude the 'Total' row and sort by Market Exposure
                        fund_df = fund_df[fund_df['Issuer name-N'] != 'Total'].sort_values(by='Market Exposure', ascending=False)
                        print(f"Debug - {fund} Issuer data before top 10: {len(fund_df)} rows")
                        top10_df = fund_df.head(10).copy()
                        if len(top10_df) > 0:
                            top10_df = calculate_percentages(top10_df, fund)
                            top10_df = top10_df[['Issuer name-N', 'Market Exposure', 'Exposure % Gross Total', 'Exposure % Fund NAV', 'Exposure % Total Commitment', 'Fund Name']]
                            top10_df = add_total_row(top10_df, fund)
                            print(f"Debug - {fund} Top 10 Issuers: {top10_df}")
                            start_row = write_exposure_sheet(ws_top10, f"Top 10 Issuers Exposure - {fund}", top10_df, start_row)
                        else:
                            print(f"Warning: No top 10 issuers found for {fund}")
            
            # Combined PCO funds
            if exposures.get('Issuer'):
                pco_dfs = [df for df in exposures['Issuer'] if df['Fund Name'].isin(["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"]).any()]
                if pco_dfs:
                    combined_pco = pd.concat(pco_dfs, ignore_index=True)
                    if not combined_pco.empty:
                        # Re-group by Issuer name-N to sum Market Exposure
                        combined_pco = combined_pco.groupby('Issuer name-N', as_index=False).agg({
                            'Market Exposure': 'sum',
                            'Fund Name': 'first'  # Preserve the first Fund Name (will be overwritten later)
                        })
                        # Recalculate percentages for combined PCO
                        combined_pco = calculate_percentages(combined_pco, "DLF1_CAYTOP")  # Use CAYTOP's NAV
                        combined_pco['Fund Name'] = "Combined PCO"
                        # Exclude the 'Total' row and sort by Market Exposure
                        combined_pco = combined_pco[combined_pco['Issuer name-N'] != 'Total'].sort_values(by='Market Exposure', ascending=False)
                        print(f"Debug - Combined PCO Issuer data before top 10: {len(combined_pco)} rows")
                        top10_combined = combined_pco.head(10).copy()
                        if len(top10_combined) > 0:
                            top10_combined = add_total_row(top10_combined, "DLF1_CAYTOP")
                            print(f"Debug - Combined PCO Top 10 Issuers: {top10_combined}")
                            start_row = write_exposure_sheet(ws_top10, "Top 10 Issuers Exposure - Combined PCO", top10_combined, start_row)
                        else:
                            print("Warning: No top 10 issuers found for Combined PCO")
                else:
                    print("No Issuer PCO data to concatenate for Top 10")

        # 2. Exposure Charts (Only for PCO funds)
        all_exposures = []
        for fund in ["DLF1_CAYTOP", "DLF1_DELATOP", "DLF1_RAIFLEVTOP"]:  # Exclude BENM
            if fund in data and data[fund]["exposure"]:
                all_exposures.extend([df for dim in ['LIEN', 'Region', 'Industry'] 
                                   for df in [data[fund]["exposure"].get(dim)] if df is not None and not df.empty])
        
        print(f"Line 485 - All exposures for charts: {len(all_exposures)} DataFrames")
        if all_exposures:
            combined_exposure = pd.concat(all_exposures, ignore_index=True)
            ws_charts = workbook.create_sheet("Exposure_Charts")
            start_row = 1
            start_row = create_pie_charts(ws_charts, combined_exposure, start_row)
        else:
            print("No exposure data for PCO funds to create charts")

        # Remove default sheet if other sheets were created
        if "Default_Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
            workbook.remove(workbook["Default_Sheet"])

        # Format all sheets
        for sheet in workbook.worksheets:
            format_sheet(sheet)

def format_sheet(ws):
    """Apply consistent formatting to worksheets"""
    header_fill = PatternFill(start_color="00008B", fill_type="solid")
    bold_font = Font(bold=True, color="FFFFFF")
    
    # Format headers
    for col in ws.iter_cols(min_row=1, max_row=1):
        for cell in col:
            if cell.value:  # Only format cells with values
                cell.font = bold_font
                cell.fill = header_fill

    # Set column widths and apply wrap text
    for col in ws.columns:
        lengths = [len(str(cell.value).replace('\n', ' ').strip()) 
                   for cell in col if cell.value is not None and str(cell.value).strip()]
        max_length = max(lengths) if lengths else 10
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2
        for cell in col:
            cell.alignment = Alignment(wrap_text=True)
        
        # Format numbers in data rows
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    col_idx = cell.column_letter
                    if ws.title in ["Issuer_Exposure", "Lien_Exposure", "Region_Exposure", "Industry_Exposure"]:
                        if col_idx in ['F', 'G', 'H', 'I']:  # Market Exposure and percentage columns
                            if col_idx == 'F':  # Market Exposure
                                cell.number_format = '$#,##0'
                            else:  # Percentage columns
                                cell.number_format = '0.00%'  # Two decimal places
                    elif ws.title == "Exposure_Charts":
                        if col_idx in ['B', 'C', 'D', 'E']:  # Market Exposure and percentage columns
                            if col_idx == 'B':  # Market Exposure
                                cell.number_format = '$#,##0'
                            elif col_idx in ['C', 'E']:  # Percentage columns
                                cell.number_format = '0.00%'
                            elif col_idx == 'D':  # N/A column
                                cell.value = 'N/A'
                    elif ws.title == "Top_10_Issuers_Exposure":
                        if col_idx in ['B', 'C', 'D', 'E']:  # Market Exposure and percentage columns
                            if col_idx == 'B':  # Market Exposure
                                cell.number_format = '$#,##0'
                            else:  # Percentage columns
                                cell.number_format = '0.00%'

def main():
    """Main execution flow for PCO and SMA exposure analysis"""
    issuer_map = load_issuer_mapping()
    if issuer_map is None:
        return

    all_data = {}
    for file in os.listdir(INPUT_PATH):
        if not file.endswith(".xlsx") or "NAVSummary" not in file:
            continue

        try:
            fund_name = file.split("-")[0]
            if fund_name not in FUND_CONFIG:
                continue

            file_path = os.path.join(INPUT_PATH, file)
            print(f"Processing {file} for exposure analysis...")

            exposure = process_exposure(file_path, issuer_map, fund_name)
            if exposure is not None:
                all_data[fund_name] = exposure
            else:
                print(f"No exposure data generated for {fund_name}")

        except Exception as e:
            print(f"Critical error processing {file} for exposure: {str(e)}")

    # Debug: Check all_data contents
    print(f"Funds with processed data: {list(all_data.keys())}")

    if all_data:
        generate_report(all_data)
        print(f"PCO and SMA Exposure report successfully generated: {os.path.join(OUTPUT_PATH, f'PCO_Exposure_Analysis_{DATE}.xlsx')}")
    else:
        print("No valid exposure data processed for PCO and SMA funds - report not generated")

if __name__ == "__main__":
    main()

# excel.py (Cross-platform version for macOS/Linux)

import pandas as pd
import openpyxl
import warnings

warnings.filterwarnings('ignore')


# Load full sheet using pandas
def load_excel_data(file_path: str, sheet_name: str) -> pd.DataFrame:
    """Load Excel data using pandas - works on all platforms"""
    try:
        # For .xlsb files, we need to use a different engine
        if file_path.endswith('.xlsb'):
            return pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb')
        else:
            return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        print(f"❌ Error loading {file_path}: {e}")
        raise


# Basic filtering function
def extract_pivot_views(df: pd.DataFrame, column_name: str, values: list) -> dict:
    output = {}
    for val in values:
        filtered = df[df[column_name] == val].copy()
        output[val] = filtered
    return output


# Cross-platform approach: Extract unique values from raw data
def get_unique_slicer_values(file_path: str, sheet_name: str, slicer_name: str) -> list:
    """
    Cross-platform approach: Extract unique values directly from Excel data
    Since we can't use COM on macOS, we'll read the raw data and find unique values
    """
    print(f"Getting unique values for '{slicer_name}' in sheet '{sheet_name}'")

    try:
        # Load the data
        df = load_excel_data(file_path, sheet_name)

        # Look for the column in various forms (case-insensitive, partial match)
        possible_columns = []
        for col in df.columns:
            col_str = str(col).strip()
            if (slicer_name.lower() in col_str.lower() or
                    col_str.lower() in slicer_name.lower() or
                    col_str.lower() == slicer_name.lower()):
                possible_columns.append(col)

        if not possible_columns:
            # Try exact match first
            if slicer_name in df.columns:
                possible_columns = [slicer_name]

        if possible_columns:
            column = possible_columns[0]

            # Get unique values, excluding NaN and empty strings
            unique_values = df[column].dropna().astype(str).str.strip()
            unique_values = unique_values[unique_values != ''].unique().tolist()

            # Remove any values that look like totals or summaries
            filtered_values = []
            skip_terms = ['total', 'grand total', 'sum', 'average', 'avg', '(blank)', 'blank']
            for val in unique_values:
                if not any(term in val.lower() for term in skip_terms):
                    filtered_values.append(val)

            print(f"Found {len(filtered_values)} unique values for '{slicer_name}'")
            return filtered_values
        else:
            print(f"Column '{slicer_name}' not found. Available columns: {list(df.columns)}")
            return []

    except Exception as e:
        print(f"Error extracting unique values: {e}")
        return []


# Cross-platform pivot simulation
def refresh_pivot_and_read(file_path: str, sheet_name: str, slicer_values: dict) -> dict:
    """
    Cross-platform approach: Simulate pivot table filtering by filtering raw data
    """
    print(f"Applying filters: {slicer_values}")

    try:
        # Load the raw data
        df = load_excel_data(file_path, sheet_name)

        # Apply filters based on slicer values
        filtered_df = df.copy()

        for slicer_name, slicer_value in slicer_values.items():
            # Find the matching column
            matching_columns = []
            for col in filtered_df.columns:
                col_str = str(col).strip()
                if (slicer_name.lower() in col_str.lower() or
                        col_str.lower() in slicer_name.lower() or
                        col_str.lower() == slicer_name.lower()):
                    matching_columns.append(col)

            if matching_columns:
                column = matching_columns[0]

                # Apply filter
                mask = filtered_df[column].astype(str).str.strip() == str(slicer_value).strip()
                filtered_df = filtered_df[mask]

        # Create simulated pivot tables
        pivot_dfs = {}

        if len(filtered_df) > 0:
            # Create a summary table (simulating a pivot)
            pivot_name = f"Filtered_Data_{sheet_name}"
            pivot_dfs[pivot_name] = filtered_df

            # If we have numeric columns, create some basic aggregations
            numeric_columns = filtered_df.select_dtypes(include=['number']).columns
            if len(numeric_columns) > 0:
                summary_name = f"Summary_{sheet_name}"

                # Create a simple summary
                summary_data = []
                for col in numeric_columns:
                    summary_data.append({
                        'Metric': col,
                        'Sum': filtered_df[col].sum(),
                        'Average': filtered_df[col].mean(),
                        'Count': filtered_df[col].count(),
                        'Min': filtered_df[col].min(),
                        'Max': filtered_df[col].max()
                    })

                if summary_data:
                    summary_df = pd.DataFrame(summary_data)
                    pivot_dfs[summary_name] = summary_df

        else:
            print("No data remaining after filtering")
            # Return empty dataframe for consistency
            pivot_dfs[f"Empty_{sheet_name}"] = pd.DataFrame()

        return pivot_dfs

    except Exception as e:
        print(f"Error during filtering: {e}")
        return {}


# Debug function for cross-platform
def debug_excel_structure(file_path: str, sheet_name: str):
    """Debug function that works on all platforms"""
    print(f"Analyzing sheet '{sheet_name}':")

    try:
        # Load the sheet
        df = load_excel_data(file_path, sheet_name)

        print(f"  Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"  Columns: {list(df.columns)}")

        # Check for potential slicer columns (columns with reasonable number of unique values)
        print(f"  Potential slicer columns:")
        for col in df.columns:
            unique_count = df[col].nunique()
            if 1 < unique_count <= 50:  # Reasonable range for slicer values
                sample_values = df[col].dropna().unique()[:3]
                print(
                    f"    - '{col}': {unique_count} unique values {list(sample_values)}{'...' if len(sample_values) == 3 else ''}")

    except Exception as e:
        print(f"Error during debugging: {e}")


# Alternative approach: Try to read multiple sheets and find patterns
def analyze_excel_file_structure(file_path: str):
    """Analyze the entire Excel file structure"""
    print(f"Analyzing Excel file: {file_path}")

    try:
        # Get all sheet names
        if file_path.endswith('.xlsb'):
            xl_file = pd.ExcelFile(file_path, engine='pyxlsb')
        else:
            xl_file = pd.ExcelFile(file_path, engine='openpyxl')

        sheet_names = xl_file.sheet_names
        print(f"Found {len(sheet_names)} sheets: {sheet_names}")

        xl_file.close()

    except Exception as e:
        print(f"Error analyzing file: {e}")
import pandas as pd
import numpy as np

def read_august_data():
    """Read and display data from August Export_SD 2 Sept_modified.xlsx"""
    
    try:
        # Read the Excel file
        file_path = "August Export_SD 2 Sept_modified.xlsx"
        print(f"Reading data from: {file_path}")
        print("=" * 50)
        
        # Get all sheet names
        excel_file = pd.ExcelFile(file_path)
        print(f"Sheet names: {excel_file.sheet_names}")
        print()
        
        # Read each sheet and display information
        for sheet_name in excel_file.sheet_names:
            print(f"Sheet: {sheet_name}")
            print("-" * 30)
            
            # Read the sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Display basic information
            print(f"Shape: {df.shape}")
            print(f"Columns: {list(df.columns)}")
            print(f"Data types:")
            for col, dtype in df.dtypes.items():
                print(f"  {col}: {dtype}")
            
            print(f"\nFirst 5 rows:")
            print(df.head())
            
            print(f"\nLast 5 rows:")
            print(df.tail())
            
            print(f"\nSummary statistics:")
            print(df.describe())
            
            print(f"\nMissing values:")
            print(df.isnull().sum())
            
            print("\n" + "=" * 50 + "\n")
            
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
    except Exception as e:
        print(f"Error reading file: {e}")

if __name__ == "__main__":
    read_august_data()

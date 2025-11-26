import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Define the base directory
base_dir = Path('/Users/bisratgizaw/Downloads/summary dashboard')

# Months to analyze
months = ['Jan', 'February', 'March', 'April', 'May', 'June', 'July', 'August']

print("=" * 100)
print("DETAILED CALL CENTER ANALYSIS - JANUARY TO AUGUST 2025")
print("Analyzing Unique Contacts, Call Outcomes, and Reach Metrics")
print("=" * 100)
print()

# First, let's explore the structure of a sample file
print("📂 EXPLORING DATA STRUCTURE...")
print("=" * 100)

sample_files_checked = 0
for month in months:
    month_dir = base_dir / month
    if not month_dir.exists():
        continue
    
    excel_files = list(month_dir.glob('*.xlsx'))[:2]  # Check first 2 files
    
    for file in excel_files:
        if sample_files_checked >= 3:
            break
        try:
            # Load Excel file
            excel_file = pd.ExcelFile(file)
            
            print(f"\n📄 File: {file.name}")
            print(f"   Month: {month}")
            print(f"   Available Sheets: {excel_file.sheet_names}")
            
            # Check each sheet
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name, nrows=5)
                print(f"\n   📋 Sheet: '{sheet_name}'")
                print(f"      Columns: {list(df.columns)}")
                print(f"      Row count: {len(pd.read_excel(file, sheet_name=sheet_name))}")
            
            sample_files_checked += 1
            
        except Exception as e:
            print(f"   ⚠️  Error reading {file.name}: {str(e)}")
    
    if sample_files_checked >= 3:
        break

print("\n" + "=" * 100)
print("STRUCTURE EXPLORATION COMPLETE")
print("=" * 100)

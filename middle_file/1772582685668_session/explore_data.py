"""
Step 2: Explore Excel data file structure
- View all column names
- View data samples
- View category distributions
- Understand data structure
"""
import pandas as pd
import sys

file_path = "/home/orangels/xm_dev/ls_dev/reportSkillMaker/警情列表_lingao_20241231-20260115_result_case.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

print("=" * 80)
print("1. BASIC INFO")
print("=" * 80)
print(f"Shape: {df.shape}")
print(f"Total rows: {len(df)}")
print(f"Total columns: {len(df.columns)}")

print("\n" + "=" * 80)
print("2. ALL COLUMN NAMES AND DATA TYPES")
print("=" * 80)
for i, (col, dtype) in enumerate(zip(df.columns, df.dtypes)):
    non_null = df[col].notna().sum()
    print(f"  [{i:2d}] {col:40s} | {str(dtype):15s} | non-null: {non_null}")

print("\n" + "=" * 80)
print("3. FIRST 5 ROWS (SAMPLE)")
print("=" * 80)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', 50)
pd.set_option('display.width', 200)
print(df.head(5).to_string())

print("\n" + "=" * 80)
print("4. UNIQUE VALUES FOR CATEGORICAL COLUMNS")
print("=" * 80)

# Check columns that might be categorical (object type or few unique values)
for col in df.columns:
    n_unique = df[col].nunique()
    if n_unique <= 50 and df[col].dtype == 'object':
        print(f"\n--- {col} (unique: {n_unique}) ---")
        val_counts = df[col].value_counts()
        for val, cnt in val_counts.items():
            print(f"    {val}: {cnt}")
    elif n_unique <= 30:
        print(f"\n--- {col} (unique: {n_unique}, dtype: {df[col].dtype}) ---")
        val_counts = df[col].value_counts()
        for val, cnt in val_counts.items():
            print(f"    {val}: {cnt}")

print("\n" + "=" * 80)
print("5. DATE/TIME COLUMNS ANALYSIS")
print("=" * 80)
for col in df.columns:
    if 'date' in col.lower() or 'time' in col.lower() or '时间' in col or '日期' in col:
        print(f"\n--- {col} ---")
        print(f"  dtype: {df[col].dtype}")
        print(f"  min: {df[col].min()}")
        print(f"  max: {df[col].max()}")
        print(f"  sample: {df[col].head(3).tolist()}")

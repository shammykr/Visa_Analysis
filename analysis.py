import pandas as pd
import matplotlib.pyplot as plt

excel_file = 'FYs97-24_NIVDetailTable.xlsx'

# Dictionary to store F1 visa data for India
f1_visa_data = {}

# List of fiscal years to analyze
years = range(1997, 2025)

try:
    # Use ExcelFile to read the file once and iterate through sheets
    with pd.ExcelFile(excel_file) as xls:
        for year in years:
            sheet_name = f'FY{str(year)[-2:]}'
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                country_column = df.columns[0]
                
                f1_column = None
                for col in df.columns:
                    if 'F1' in col.upper().replace('-', ''):
                        f1_column = col
                        break
                
                if f1_column:
                    india_row = df[df[country_column].str.strip().str.lower() == 'india']
                    
                    if not india_row.empty:
                        visa_count = india_row[f1_column].iloc[0]
                        f1_visa_data[year] = int(visa_count)
                    else:
                        print(f"Warning: 'India' not found in sheet '{sheet_name}'")
                else:
                    print(f"Warning: 'F1' visa data not found in sheet '{sheet_name}'")
            else:
                print(f"Sheet '{sheet_name}' not found in the Excel file.")

except FileNotFoundError:
    print(f"Error: The file '{excel_file}' was not found. Please make sure it's in the same directory as the script.")
except Exception as e:
    print(f"An error occurred: {e}")

# Plot the data if any was collected
if f1_visa_data:
    fiscal_years = sorted(f1_visa_data.keys())
    visa_counts = [f1_visa_data[year] for year in fiscal_years]

    plt.figure(figsize=(14, 8))
    plt.bar(fiscal_years, visa_counts, color='skyblue')
    plt.title('Number of F1 Visas Issued to India (1997-2024)', fontsize=16)
    plt.xlabel('Fiscal Year', fontsize=12)
    plt.ylabel('Number of F1 Visas Issued', fontsize=12)
    plt.xticks(fiscal_years, rotation=45)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    # Add data labels
    for i, count in enumerate(visa_counts):
        plt.text(fiscal_years[i], count, str(count), ha='center', va='bottom', fontsize=8, rotation=90)
    
    plt.tight_layout()
    plt.show()
else:
    print("No F1 visa data for India was found in the provided Excel file.")
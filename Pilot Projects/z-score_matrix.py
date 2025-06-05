import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import TwoSlopeNorm
import calendar
import os

# --- POPUP WINDOW FOR SECTOR SELECTION AND DATE RANGE ---
def select_sector_and_dates():
    # Use absolute path to data folder relative to this script's parent directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    root = tk.Tk()
    root.title("Select Sector and Date Range")
    root.geometry("350x220")
    root.resizable(False, False)

    xl = pd.ExcelFile(os.path.join(data_dir, 'Company Names.xlsx'))
    sheet_names = xl.sheet_names

    selected = tk.StringVar()
    selected.set(sheet_names[0])

    start_date_var = tk.StringVar()
    end_date_var = tk.StringVar()

    def on_ok():
        start = start_date_var.get()
        end = end_date_var.get()
        # Basic validation for yyyy/mm format
        if not selected.get():
            messagebox.showwarning("Selection Error", "Please select a sector.")
            return
        if not (len(start) == 7 and start[4] == '/' and start[:4].isdigit() and start[5:7].isdigit()):
            messagebox.showwarning("Input Error", "Start date must be in yyyy/mm format.")
            return
        if not (len(end) == 7 and end[4] == '/' and end[:4].isdigit() and end[5:7].isdigit()):
            messagebox.showwarning("Input Error", "End date must be in yyyy/mm format.")
            return
        if end < start:
            messagebox.showwarning("Input Error", "End date must not be before start date.")
            return
        root.selected_sector = selected.get()
        root.start_date = start
        root.end_date = end
        root.destroy()

    label = tk.Label(root, text="Select sector:")
    label.pack(pady=(15, 5))

    dropdown = ttk.Combobox(root, textvariable=selected, values=sheet_names, state="readonly")
    dropdown.pack(pady=5)

    start_label = tk.Label(root, text="Start date (yyyy/mm):")
    start_label.pack(pady=(10, 2))
    start_entry = tk.Entry(root, textvariable=start_date_var)
    start_entry.pack(pady=2)

    end_label = tk.Label(root, text="End date (yyyy/mm):")
    end_label.pack(pady=(10, 2))
    end_entry = tk.Entry(root, textvariable=end_date_var)
    end_entry.pack(pady=2)

    ok_btn = tk.Button(root, text="OK", command=on_ok)
    ok_btn.pack(pady=10)

    root.mainloop()
    return (
        getattr(root, 'selected_sector', None),
        getattr(root, 'start_date', None),
        getattr(root, 'end_date', None)
    )

# Use absolute paths for data files
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_dir = os.path.join(base_dir, 'data')

sector, start_date, end_date = select_sector_and_dates()
if sector is None or start_date is None or end_date is None:
    raise SystemExit("No sector or date range selected. Exiting.")

# Update all file paths to use the 'data' folder
excel_file_path = os.path.join(data_dir, f"{sector}.xlsx")
df = pd.read_excel(os.path.join(data_dir, 'Company Names.xlsx'), sheet_name=sector, engine='openpyxl')
tickers = df['Ticker']
matrix = pd.DataFrame(index=tickers, columns=tickers)

print("Your request is being processed. Please wait...")

# Convert start and end date to comparable format
start_period = pd.Period(start_date, freq='M')
end_period = pd.Period(end_date, freq='M')

# Find the latest date across all tickers after filtering
latest_date = None
for ticker in tickers:
    try:
        df_ticker = pd.read_excel(excel_file_path, sheet_name=ticker, header=4, engine='openpyxl')
        df_ticker = df_ticker.rename(columns={'Close Adj. Ex. Div.': 'Last Price', 'EPS Basic - TTM': 'EPS',
                                              'Dividend Yield-TTM': 'D/Y', 'Dates': 'Date'})
        df_ticker['Date'] = pd.to_datetime(df_ticker['Date'], dayfirst=True, errors='coerce')
        df_ticker = df_ticker.fillna(0.01)
        df_ticker = df_ticker[df_ticker['Date'].notna()]
        df_ticker['Period'] = df_ticker['Date'].dt.to_period('M')
        df_ticker = df_ticker[(df_ticker['Period'] >= start_period) & (df_ticker['Period'] <= end_period)]
        if not df_ticker.empty:
            max_date = df_ticker['Date'].max()
            if latest_date is None or max_date > latest_date:
                latest_date = max_date
    except Exception:
        continue

# Format the date range string for the plot title
start_year, start_month = int(start_date[:4]), int(start_date[5:7])
end_year, end_month = int(end_date[:4]), int(end_date[5:7])
start_str = f"{calendar.month_name[start_month]} {start_year}"
end_str = f"{calendar.month_name[end_month]} {end_year}"
date_range_str = f"{start_str} - {end_str}"

for ticker1 in tickers:
    for ticker2 in tickers:
        try:
            if ticker1 != ticker2:
                df1 = pd.read_excel(excel_file_path, sheet_name=ticker1, header=4, engine='openpyxl')
                df2 = pd.read_excel(excel_file_path, sheet_name=ticker2, header=4, engine='openpyxl')

                df1 = df1.rename(columns={'Close Adj. Ex. Div.': 'Last Price', 'EPS Basic - TTM': 'EPS',
                                           'Dividend Yield-TTM': 'D/Y', 'Dates': 'Date'})
                df1['Date'] = pd.to_datetime(df1['Date'], dayfirst=True, errors='coerce')
                df1 = df1.fillna(0.01)
                df1 = df1[df1['Date'].notna()]
                df1['Period'] = df1['Date'].dt.to_period('M')
                df1 = df1[(df1['Period'] >= start_period) & (df1['Period'] <= end_period)]
                df1.set_index('Date', inplace=True)

                df2 = df2.rename(columns={'Close Adj. Ex. Div.': 'Last Price', 'EPS Basic - TTM': 'EPS',
                                           'Dividend Yield-TTM': 'D/Y', 'Dates': 'Date'})
                df2['Date'] = pd.to_datetime(df2['Date'], dayfirst=True, errors='coerce')
                df2 = df2.fillna(0.01)
                df2 = df2[df2['Date'].notna()]
                df2['Period'] = df2['Date'].dt.to_period('M')
                df2 = df2[(df2['Period'] >= start_period) & (df2['Period'] <= end_period)]
                df2.set_index('Date', inplace=True)

                # Align on index and fill NaN with 0.01
                df_div = df1[['P/E']].div(df2[['P/E']])
                df_div = df_div.fillna(0.01)
                if df_div.empty:
                    matrix.loc[ticker1, ticker2] = np.nan
                    continue
                df_div = df_div.sort_index()

                PE_mean = df_div['P/E'].mean()
                PE_std = df_div['P/E'].std()
                current_PE = df_div['P/E'].iloc[-1]
                Z = round((current_PE - PE_mean) / PE_std, 2) if PE_std != 0 else np.nan
                matrix.loc[ticker1, ticker2] = Z
            else:
                df_self = pd.read_excel(excel_file_path, sheet_name=ticker1, header=4, engine='openpyxl')
                df_self = df_self.rename(columns={'Close Adj. Ex. Div.': 'Last Price', 'EPS Basic - TTM': 'EPS',
                                                  'Dividend Yield-TTM': 'D/Y', 'Dates': 'Date'})
                df_self['Date'] = pd.to_datetime(df_self['Date'], dayfirst=True, errors='coerce')
                df_self = df_self.fillna(0.01)
                df_self = df_self[df_self['Date'].notna()]
                df_self['Period'] = df_self['Date'].dt.to_period('M')
                df_self = df_self[(df_self['Period'] >= start_period) & (df_self['Period'] <= end_period)]
                if df_self.empty:
                    matrix.loc[ticker1, ticker2] = np.nan
                    continue
                df_self = df_self.sort_values('Date')
                PE_mean = df_self['P/E'].mean()
                PE_std = df_self['P/E'].std()
                current_PE = df_self['P/E'].iloc[-1]
                Z = round((current_PE - PE_mean) / PE_std, 2) if PE_std != 0 else np.nan
                matrix.loc[ticker1, ticker2] = Z
        except Exception as e:
            matrix.loc[ticker1, ticker2] = np.nan

# Sort the matrix by the count of positive and negative values in each row and column
def count_positives(series):
    return (series > 0).sum()

def count_negatives(series):
    return (series < 0).sum()

row_positive_counts = matrix.astype(float).apply(count_positives, axis=1)
row_negative_counts = matrix.astype(float).apply(count_negatives, axis=1)
row_sums = matrix.astype(float).sum(axis=1)

# Sort by most positives (descending), then most negatives (descending), then sum (descending)
sorted_rows = sorted(
    matrix.index,
    key=lambda x: (
        -row_positive_counts[x],           # Most positives first
        -row_negative_counts[x],           # If tie, most negatives next
        -row_sums[x]                       # If still tie, largest sum first
    )
)

# For columns: use the same order as rows for consistency
matrix = matrix.loc[sorted_rows, sorted_rows]

plt.figure(figsize=(10, 8))
norm = TwoSlopeNorm(vmin=matrix.astype(float).min().min(), vcenter=0, vmax=matrix.astype(float).max().max())
ax = sns.heatmap(
    matrix.astype(float),
    annot=True,
    fmt=".2f",
    cmap='RdYlGn_r',
    cbar=True,
    linewidths=0.5,
    linecolor='gray',
    norm=norm
)

for i in range(len(matrix)):
    ax.add_patch(
        plt.Rectangle(
            (i, i), 1, 1, fill=False, edgecolor='black', lw=3
        )
    )

plt.title(f"Comparative Z-Score Matrix\n{date_range_str}", fontweight='bold')
plt.xlabel("Denominator", fontweight='bold')
plt.ylabel("Numerator", fontweight='bold')
plt.tight_layout()
plt.show()

print(matrix)
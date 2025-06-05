import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import time
import calendar

# --- POPUP WINDOW FOR SECTOR SELECTION AND DATE RANGE ---
def select_sector_and_dates():
    root = tk.Tk()
    root.title("Select Sector and Date Range")
    root.geometry("350x210")
    root.resizable(False, False)

    # Use absolute path to data folder relative to this script's parent directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
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
    return getattr(root, 'selected_sector', None), getattr(root, 'start_date', None), getattr(root, 'end_date', None)

# Use absolute paths for data files
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_dir = os.path.join(base_dir, 'data')

sector, start_date, end_date = select_sector_and_dates()
if sector is None or start_date is None or end_date is None:
    raise SystemExit("No sector or date range selected. Exiting.")

# Set file paths based on sector
excel_file_path = os.path.join(data_dir, f"{sector}.xlsx")

# Load the company names data
df = pd.read_excel(os.path.join(data_dir, 'Company Names.xlsx'), sheet_name=sector, engine='openpyxl')  # Use selected sector as sheet

# Initialize columns in the DataFrame
df['P/E'] = np.nan
df['Abs P/E'] = np.nan     
df['Avg P/E'] = np.nan   
df['P/E Std Dev'] = np.nan
df['Z-score P/E'] = np.nan
df['D/Y'] = np.nan
df['Abs D/Y'] = np.nan
df['Avg D/Y'] = np.nan 
df['D/Y Std Dev'] = np.nan
df['Z-score D/Y'] = np.nan

# Convert start and end date to Period for filtering
start_period = pd.Period(start_date, freq='M')
end_period = pd.Period(end_date, freq='M')

# Find the latest date across all tickers after filtering
latest_date = None

# Process each ticker
for index, row in df.iterrows():
    ticker = row['Ticker'] 
    try:
        # Specify header=4 to start reading from the 5th row (0-indexed)
        ticker_data = pd.read_excel(excel_file_path, sheet_name=ticker, header=4, engine='openpyxl')
        ticker_data = ticker_data.rename(columns={'Close Adj. Ex. Div.': 'Last Price', 'EPS Basic - TTM': 'EPS', 'Dividend Yield-TTM': 'D/Y', 'Dates': 'Date'})
        ticker_data['Date'] = pd.to_datetime(ticker_data['Date'], errors='coerce')
        ticker_data = ticker_data.dropna(subset=['Date'])
        ticker_data['Period'] = ticker_data['Date'].dt.to_period('M')
        ticker_data = ticker_data[(ticker_data['Period'] >= start_period) & (ticker_data['Period'] <= end_period)]
        ticker_data = ticker_data.sort_values(by='Date', ascending=False)

        if ticker_data.empty:
            continue

        # Track the latest date for the plot title
        max_date = ticker_data['Date'].max()
        if latest_date is None or max_date > latest_date:
            latest_date = max_date

        most_recent_pe = ticker_data.iloc[0]['P/E']
        most_recent_dy = ticker_data.iloc[0]['D/Y']
        df.at[index, 'P/E'] = most_recent_pe
        df.at[index, 'D/Y'] = most_recent_dy

        avg_pe = ticker_data['P/E'].mean(skipna=True)
        avg_dy = ticker_data['D/Y'].mean(skipna=True)
        df.at[index, 'Avg P/E'] = avg_pe
        df.at[index, 'Avg D/Y'] = avg_dy

        std_pe = ticker_data['P/E'].std(skipna=True)
        std_dy = ticker_data['D/Y'].std(skipna=True)
        df.at[index, 'P/E Std Dev'] = std_pe
        df.at[index, 'D/Y Std Dev'] = std_dy

        if std_pe and not np.isnan(std_pe) and std_pe != 0:
            z_score_pe = (most_recent_pe - avg_pe) / std_pe
            df.at[index, 'Z-score P/E'] = z_score_pe

        if std_dy and not np.isnan(std_dy) and std_dy != 0:
            z_score_dy = (most_recent_dy - avg_dy) / std_dy
            df.at[index, 'Z-score D/Y'] = z_score_dy

        df.at[index, 'Abs P/E'] = abs(most_recent_pe)
        df.at[index, 'Abs D/Y'] = abs(most_recent_dy)

    except Exception as e:
        print(f"Error processing ticker {ticker}: {e}")

# Format the date range string for the plot titles
start_year, start_month = int(start_date[:4]), int(start_date[5:7])
end_year, end_month = int(end_date[:4]), int(end_date[5:7])
start_str = f"{calendar.month_name[start_month]} {start_year}"
end_str = f"{calendar.month_name[end_month]} {end_year}"
date_range_str = f"{start_str} - {end_str}"

yaxis = df['Z-score P/E'].max() if (df['Z-score P/E'].max()) > abs(df['Z-score P/E'].min()) else abs(df['Z-score P/E'].min())

# Calculate axis limits for first scatter plot
x_min = 0
x_max = df['D/Y'].max() + 0.1
y_max = yaxis + 0.1
y_min = yaxis * -1 - 0.1
x_center = (x_min + x_max) / 2

plt.figure(figsize=(12, 8))
sns.scatterplot(data=df, x='D/Y', y='Z-score P/E', hue='Ticker', legend=False, s=350)

plt.axvline(x=x_center, color='black', linestyle='--')
plt.axhline(y=0, color='black', linestyle='--')

plt.xlim(x_max, x_min)  # Invert x-axis: max on left, 0 on right
plt.ylim(y_min, y_max)

for i in range(len(df)):
    plt.text(
        df['D/Y'].iloc[i], df['Z-score P/E'].iloc[i], df['Ticker'].iloc[i]
    )

plt.text(x_min, y_max, 'Expensive', fontweight='bold')
plt.text(x_max, y_min, 'Cheap', fontweight='bold')

plt.xlabel('D/Y', fontweight='bold')
plt.ylabel('Z-score P/E', fontweight='bold')
plt.title(f'Z-score P/E vs D/Y\n{date_range_str}', fontweight='bold')
plt.tight_layout()
plt.show()

# Calculate axis limits for second scatter plot
x_min = 0
x_max = df['Abs D/Y'].max() + 0.1
y_min = 0
y_max = df['Abs P/E'].max() + 0.5
x_center = (x_min + x_max) / 2
y_center = (y_min + y_max) / 2

plt.figure(figsize=(12, 8))
sns.scatterplot(data=df, x='Abs D/Y', y='Abs P/E', hue='Ticker', legend=False, s=350)

plt.axvline(x=x_center, color='black', linestyle='--')
plt.axhline(y=y_center, color='black', linestyle='--')

plt.xlim(x_max, x_min)  # Invert x-axis: max on left, 0 on right
plt.ylim(y_min, y_max)

for i in range(len(df)):
    plt.text(
        df['Abs D/Y'].iloc[i], df['Abs P/E'].iloc[i], df['Ticker'].iloc[i]
    )

plt.text(x_min, y_max, 'Expensive', fontweight='bold')
plt.text(x_max, y_min, 'Cheap', fontweight='bold')

plt.xlabel('Abs D/Y', fontweight='bold')
plt.ylabel('Abs P/E', fontweight='bold')
plt.title(f'Abs P/E vs Abs D/Y\n{date_range_str}', fontweight='bold')
plt.tight_layout()
plt.show()
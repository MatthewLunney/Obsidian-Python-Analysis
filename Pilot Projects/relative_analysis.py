import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import ttk, messagebox
import os

def select_sector():
    # Use absolute path to data folder relative to this script's parent directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    root = tk.Tk()
    root.title("Select Sector")
    root.geometry("350x150")
    root.resizable(False, False)

    xl = pd.ExcelFile(os.path.join(data_dir, 'Company Names.xlsx'))
    sheet_names = xl.sheet_names

    selected_sector = tk.StringVar()
    selected_sector.set(sheet_names[0])

    def on_ok():
        root.selected_sector = selected_sector.get()
        root.destroy()

    label = tk.Label(root, text="Select sector:")
    label.pack(pady=(20, 5))

    dropdown = ttk.Combobox(root, textvariable=selected_sector, values=sheet_names, state="readonly")
    dropdown.pack(pady=5)

    ok_btn = tk.Button(root, text="OK", command=on_ok)
    ok_btn.pack(pady=15)

    root.mainloop()
    return getattr(root, 'selected_sector', None)

def select_tickers_and_dates(sector):
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base_dir, 'data')
    root = tk.Tk()
    root.title("Select Two Stocks and Date Range")
    root.geometry("400x300")
    root.resizable(False, False)

    sector_file = os.path.join(data_dir, f"{sector}.xlsx")
    xl = pd.ExcelFile(sector_file)
    tickers = xl.sheet_names

    ticker1 = tk.StringVar()
    ticker2 = tk.StringVar()
    start_date_var = tk.StringVar()
    end_date_var = tk.StringVar()
    ticker1.set(tickers[0])
    ticker2.set(tickers[1] if len(tickers) > 1 else tickers[0])

    def on_ok():
        if ticker1.get() == ticker2.get():
            messagebox.showwarning("Selection Error", "Please select two different stocks.")
            return
        start = start_date_var.get()
        end = end_date_var.get()
        if not (len(start) == 7 and start[4] == '/' and start[:4].isdigit() and start[5:7].isdigit()):
            messagebox.showwarning("Input Error", "Start date must be in yyyy/mm format.")
            return
        if not (len(end) == 7 and end[4] == '/' and end[:4].isdigit() and end[5:7].isdigit()):
            messagebox.showwarning("Input Error", "End date must be in yyyy/mm format.")
            return
        if end < start:
            messagebox.showwarning("Input Error", "End date must not be before start date.")
            return
        root.ticker1 = ticker1.get()
        root.ticker2 = ticker2.get()
        root.start_date = start
        root.end_date = end
        root.destroy()

    label1 = tk.Label(root, text="Select first stock:")
    label1.pack(pady=(20, 2))
    dropdown1 = ttk.Combobox(root, textvariable=ticker1, values=tickers, state="readonly")
    dropdown1.pack(pady=2)

    label2 = tk.Label(root, text="Select second stock:")
    label2.pack(pady=(10, 2))
    dropdown2 = ttk.Combobox(root, textvariable=ticker2, values=tickers, state="readonly")
    dropdown2.pack(pady=2)

    start_label = tk.Label(root, text="Start date (yyyy/mm):")
    start_label.pack(pady=(10, 2))
    start_entry = tk.Entry(root, textvariable=start_date_var)
    start_entry.pack(pady=2)

    end_label = tk.Label(root, text="End date (yyyy/mm):")
    end_label.pack(pady=(10, 2))
    end_entry = tk.Entry(root, textvariable=end_date_var)
    end_entry.pack(pady=2)

    ok_btn = tk.Button(root, text="OK", command=on_ok)
    ok_btn.pack(pady=15)

    root.mainloop()
    return (
        getattr(root, 'ticker1', None),
        getattr(root, 'ticker2', None),
        getattr(root, 'start_date', None),
        getattr(root, 'end_date', None)
    )

def filter_by_period(df, start, end):
    df = df.copy()
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
    start_period = pd.Period(start, freq='M')
    end_period = pd.Period(end, freq='M')
    df = df[(df['Date'].dt.to_period('M') >= start_period) & (df['Date'].dt.to_period('M') <= end_period)]
    df['Date'] = df['Date'].dt.strftime('%Y/%m')
    return df

# --- Main Program ---
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_dir = os.path.join(base_dir, 'data')

sector = select_sector()
if not sector:
    raise SystemExit("No sector selected. Exiting.")

ticker1, ticker2, start_date, end_date = select_tickers_and_dates(sector)
if not ticker1 or not ticker2 or not start_date or not end_date:
    raise SystemExit("Stocks or date range not selected. Exiting.")

sector_file = os.path.join(data_dir, f"{sector}.xlsx")
usecols = ['Date', 'Close Adj. Ex. Div.', 'P/E', 'EPS Basic - TTM', 'Dividend Yield-TTM']
rename_dict = {
    'Close Adj. Ex. Div.': 'Last Price',
    'P/E': 'P/E',
    'EPS Basic - TTM': 'EPS TTM',
    'Dividend Yield-TTM': 'D/Y TTM'
}
HEADER_ROW = 4  # Adjust as needed

df1 = pd.read_excel(sector_file, sheet_name=ticker1, usecols=usecols, header=HEADER_ROW)
df2 = pd.read_excel(sector_file, sheet_name=ticker2, usecols=usecols, header=HEADER_ROW)

df1 = df1.rename(columns=rename_dict)
df2 = df2.rename(columns=rename_dict)

df1 = filter_by_period(df1, start_date, end_date)
df2 = filter_by_period(df2, start_date, end_date)

df1.set_index('Date', inplace=True)
df2.set_index('Date', inplace=True)

df = df1.div(df2)
df = df.dropna()
df.reset_index(inplace=True)

if df.empty:
    print("No overlapping data between the two stocks after division and dropping NaNs.")
    raise SystemExit("No data to compare. Exiting.")

# Convert 'Date' to datetime and extract month-year for x-axis
df['Month'] = pd.to_datetime(df['Date'], format='%Y/%m')
x = df['Month']
y1 = df['Last Price']
y2 = df['P/E']
y3 = df['EPS TTM']

PE_mean = y2.mean()
PE_std = y2.std()

sns.set_style("dark")
plt.style.use("dark_background")

fig, ax1 = plt.subplots(figsize=(12, 6))
fig.subplots_adjust(right=0.85)

# 1. Plot Price Relative on left y-axis
ln1 = ax1.plot(x, y1, color='white', label='Price Relative', linestyle='-')
ax1.set_ylabel('Price Relative', color='white')
ax1.tick_params(axis='y', labelcolor='white')
ax1.spines['left'].set_color('white')
ax1.tick_params(axis='y', colors='white')
ax1.set_yscale('log')

# 2. Plot P/E Relative on first right y-axis
ax2 = ax1.twinx()
ax2.spines['right'].set_position(('outward', 60))
ln2 = ax2.plot(x, y2, color='yellow', label='P/E Relative', linestyle='-')
ax2.set_ylabel('P/E Relative', color='yellow')
ax2.tick_params(axis='y', labelcolor='yellow')
ax2.spines['right'].set_color('yellow')
ax2.tick_params(axis='y', colors='yellow')
ax2.set_yscale('log')

# Plot mean and ±1 std lines for P/E Relative
ax2.axhline(y=PE_mean, color='blue', linestyle='--', label='Mean P/E Relative')
ax2.axhline(y=PE_mean + PE_std, color='red', linestyle='--', label='+1 Std P/E Relative')
ax2.axhline(y=max(PE_mean - PE_std, 1e-6), color='red', linestyle='--', label='-1 Std P/E Relative')

# 3. Plot EPS TTM Relative on second right y-axis
ax3 = ax1.twinx()
ax3.spines['right'].set_position(('outward', 120))
ln3 = ax3.plot(x, y3, color='green', label='EPS TTM Relative', linestyle='-')
ax3.set_ylabel('EPS TTM Relative', color='green')
ax3.tick_params(axis='y', labelcolor='green')
ax3.spines['right'].set_color('green')
ax3.tick_params(axis='y', colors='green')
ax3.set_yscale('log')

# Combine legends from all axes
lns = ln1 + ln2 + ln3
labels = [l.get_label() for l in lns]
fig.legend(lns, labels, loc='upper left', bbox_to_anchor=(0.1, 0.9))

# Set x-axis labels to every 4th month
xticks = x[::4]
xticklabels = [dt.strftime('%b %Y') for dt in xticks]
ax1.set_xticks(xticks)
ax1.set_xticklabels(xticklabels, rotation=45, ha='right')

plt.xlabel('Month')
plt.title(f'Relative Analysis: {ticker1} / {ticker2}')
plt.tight_layout()
plt.show()
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
import tempfile

def produce_earnings_vs_div_plots(sector, start_date, end_date):
    import calendar

    excel_file_path = os.path.join('data', f"{sector}.xlsx")
    df = pd.read_excel(os.path.join('data', 'Company Names.xlsx'), sheet_name=sector)
    tickers = df['Ticker']

    df['P/E'] = np.nan
    df['D/Y'] = np.nan

    start_period = pd.Period(start_date, freq='M')
    end_period = pd.Period(end_date, freq='M')
    for idx, ticker in enumerate(tickers):
        try:
            tdf = pd.read_excel(excel_file_path, sheet_name=ticker, header=4)
            tdf = tdf.rename(columns={
                'Close Adj. Ex. Div.': 'Last Price',
                'EPS Basic - TTM': 'EPS',
                'Dividend Yield-TTM': 'D/Y',
                'Dates': 'Date'
            })
            tdf['Date'] = pd.to_datetime(tdf['Date'], errors='coerce')
            tdf = tdf.dropna(subset=['Date'])
            tdf['Period'] = tdf['Date'].dt.to_period('M')
            tdf = tdf[(tdf['Period'] >= start_period) & (tdf['Period'] <= end_period)]
            tdf = tdf.sort_values('Date', ascending=False)
            if not tdf.empty:
                df.at[idx, 'P/E'] = tdf.iloc[0]['P/E']
                df.at[idx, 'D/Y'] = tdf.iloc[0]['D/Y']
        except Exception as e:
            print(f"Error processing {ticker}: {e}")

    df['Avg P/E'] = df['P/E'].mean(skipna=True)
    df['P/E Std Dev'] = df['P/E'].std(skipna=True)
    df['Avg D/Y'] = df['D/Y'].mean(skipna=True)
    df['D/Y Std Dev'] = df['D/Y'].std(skipna=True)
    df['Z-score P/E'] = (df['P/E'] - df['Avg P/E']) / df['P/E Std Dev']
    df['Z-score D/Y'] = (df['D/Y'] - df['Avg D/Y']) / df['D/Y Std Dev']
    df['Abs P/E'] = df['P/E'].abs()
    df['Abs D/Y'] = df['D/Y'].abs()

    start_year, start_month = int(start_date[:4]), int(start_date[5:7])
    end_year, end_month = int(end_date[:4]), int(end_date[5:7])
    start_str = f"{calendar.month_name[start_month]} {start_year}"
    end_str = f"{calendar.month_name[end_month]} {end_year}"
    date_range_str = f"{start_str} - {end_str}"

    # First plot: Z-score P/E vs D/Y
    yaxis = df['Z-score P/E'].max() if (df['Z-score P/E'].max()) > abs(df['Z-score P/E'].min()) else abs(df['Z-score P/E'].min())
    x_min = 0
    x_max = df['D/Y'].max() + 0.1
    y_max = yaxis + 0.1
    y_min = yaxis * -1 - 0.1
    x_center = (x_min + x_max) / 2

    plt.figure(figsize=(12, 8))
    sns.scatterplot(data=df, x='D/Y', y='Z-score P/E', hue='Ticker', legend=False, s=350)
    plt.axvline(x=x_center, color='black', linestyle='--')
    plt.axhline(y=0, color='black', linestyle='--')
    plt.xlim(x_max, x_min)
    plt.ylim(y_min, y_max)
    for i in range(len(df)):
        plt.text(df['D/Y'].iloc[i], df['Z-score P/E'].iloc[i], df['Ticker'].iloc[i])
    plt.text(x_min, y_max, 'Expensive', fontweight='bold')
    plt.text(x_max, y_min, 'Cheap', fontweight='bold')
    plt.xlabel('D/Y', fontweight='bold')
    plt.ylabel('Z-score P/E', fontweight='bold')
    plt.title(f'Z-score P/E vs D/Y\n{date_range_str}', fontweight='bold')
    plt.tight_layout()

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile1:
        plot1_path = tmpfile1.name
        plt.savefig(plot1_path, bbox_inches='tight')
    plt.close()

    # Second plot: Abs P/E vs Abs D/Y
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
    plt.xlim(x_max, x_min)
    plt.ylim(y_min, y_max)
    for i in range(len(df)):
        plt.text(df['Abs D/Y'].iloc[i], df['Abs P/E'].iloc[i], df['Ticker'].iloc[i])
    plt.text(x_min, y_max, 'Expensive', fontweight='bold')
    plt.text(x_max, y_min, 'Cheap', fontweight='bold')
    plt.xlabel('Abs D/Y', fontweight='bold')
    plt.ylabel('Abs P/E', fontweight='bold')
    plt.title(f'Abs P/E vs Abs D/Y\n{date_range_str}', fontweight='bold')
    plt.tight_layout()

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile2:
        plot2_path = tmpfile2.name
        plt.savefig(plot2_path, bbox_inches='tight')
    plt.close()

    return plot1_path, plot2_path
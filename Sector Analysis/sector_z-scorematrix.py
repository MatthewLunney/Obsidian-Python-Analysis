import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import TwoSlopeNorm
import os
import tempfile

def produce_zscore_matrix(sector, start_date, end_date):
    import calendar

    excel_file_path = os.path.join('data', f"{sector}.xlsx")
    df = pd.read_excel(os.path.join('data', 'Company Names.xlsx'), sheet_name=sector)
    tickers = df['Ticker']
    matrix = pd.DataFrame(index=tickers, columns=tickers)

    start_period = pd.Period(start_date, freq='M')
    end_period = pd.Period(end_date, freq='M')

    start_year, start_month = int(start_date[:4]), int(start_date[5:7])
    end_year, end_month = int(end_date[:4]), int(end_date[5:7])
    start_str = f"{calendar.month_name[start_month]} {start_year}"
    end_str = f"{calendar.month_name[end_month]} {end_year}"
    date_range_str = f"{start_str} - {end_str}"

    for ticker1 in tickers:
        for ticker2 in tickers:
            try:
                if ticker1 != ticker2:
                    df1 = pd.read_excel(excel_file_path, sheet_name=ticker1, header=4)
                    df2 = pd.read_excel(excel_file_path, sheet_name=ticker2, header=4)

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
                    df_self = pd.read_excel(excel_file_path, sheet_name=ticker1, header=4)
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

    def count_positives(series):
        return (series > 0).sum()
    def count_negatives(series):
        return (series < 0).sum()

    row_positive_counts = matrix.astype(float).apply(count_positives, axis=1)
    row_negative_counts = matrix.astype(float).apply(count_negatives, axis=1)
    row_sums = matrix.astype(float).sum(axis=1)

    sorted_rows = sorted(
        matrix.index,
        key=lambda x: (
            -row_positive_counts[x],
            -row_negative_counts[x],
            -row_sums[x]
        )
    )
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

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
        heatmap_path = tmpfile.name
        plt.savefig(heatmap_path, bbox_inches='tight')
    plt.close()
    return heatmap_path
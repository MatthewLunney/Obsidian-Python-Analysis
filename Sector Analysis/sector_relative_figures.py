import pandas as pd
import matplotlib.pyplot as plt
import os
import tempfile

def produce_relative_figures(sector, start_date, end_date):
    excel_file_path = os.path.join('data', f"{sector}.xlsx")
    df_names = pd.read_excel(os.path.join('data', 'Company Names.xlsx'), sheet_name=sector)
    tickers = df_names['Ticker'].tolist()
    plots = []

    start_period = pd.Period(start_date, freq='M')
    end_period = pd.Period(end_date, freq='M')

    for ticker1 in tickers:
        for ticker2 in tickers:
            if ticker1 == ticker2:
                continue  # Skip self/self
            try:
                df1 = pd.read_excel(excel_file_path, sheet_name=ticker1, header=4)
                df2 = pd.read_excel(excel_file_path, sheet_name=ticker2, header=4)

                df1 = df1.rename(columns={
                    'Close Adj. Ex. Div.': 'Last Price',
                    'EPS Basic - TTM': 'EPS',
                    'P/E - TTM': 'P/E',
                    'Dates': 'Date'
                })
                df2 = df2.rename(columns={
                    'Close Adj. Ex. Div.': 'Last Price',
                    'EPS Basic - TTM': 'EPS',
                    'P/E - TTM': 'P/E',
                    'Dates': 'Date'
                })

                df1['Date'] = pd.to_datetime(df1['Date'], errors='coerce')
                df2['Date'] = pd.to_datetime(df2['Date'], errors='coerce')
                df1 = df1.dropna(subset=['Date'])
                df2 = df2.dropna(subset=['Date'])

                df1['Period'] = df1['Date'].dt.to_period('M')
                df2['Period'] = df2['Date'].dt.to_period('M')
                df1 = df1[(df1['Period'] >= start_period) & (df1['Period'] <= end_period)]
                df2 = df2[(df2['Period'] >= start_period) & (df2['Period'] <= end_period)]

                # Merge on Date
                merged = pd.merge(
                    df1[['Date', 'Last Price', 'EPS', 'P/E']],
                    df2[['Date', 'Last Price', 'EPS', 'P/E']],
                    on='Date',
                    suffixes=(f'_{ticker1}', f'_{ticker2}')
                )
                merged = merged.sort_values('Date')
                merged['Relative Price'] = merged[f'Last Price_{ticker1}'] / merged[f'Last Price_{ticker2}']
                merged['Relative EPS'] = merged[f'EPS_{ticker1}'] / merged[f'EPS_{ticker2}']
                merged['Relative P/E'] = merged[f'P/E_{ticker1}'] / merged[f'P/E_{ticker2}']

                if merged.empty:
                    continue

                fig, ax1 = plt.subplots(figsize=(12, 6))

                color1 = 'black'
                color2 = 'tab:green'
                color3 = 'tab:orange'

                ax1.set_xlabel('Date', fontweight='bold')
                ax1.set_ylabel('Relative Price', color=color1, fontweight='bold')
                l1, = ax1.plot(merged['Date'], merged['Relative Price'], color=color1, label=f"Relative Price {ticker1}/{ticker2}")
                ax1.tick_params(axis='y', labelcolor=color1)
                ax1.set_yscale('log')

                # Second y-axis for Relative P/E
                ax2 = ax1.twinx()
                ax2.set_ylabel('Relative P/E', color=color2, fontweight='bold')
                l2, = ax2.plot(merged['Date'], merged['Relative P/E'], color=color2, label=f"Relative P/E {ticker1}/{ticker2}")
                mean_pe = merged['Relative P/E'].mean()
                std_pe = merged['Relative P/E'].std()
                ax2.axhline(mean_pe, color='blue', linestyle='--', linewidth=1, label='Mean Rel P/E')
                ax2.axhline(mean_pe + std_pe, color='red', linestyle='--', linewidth=1, label='+1 Std Rel P/E')
                ax2.axhline(max(mean_pe - std_pe, 1e-6), color='green', linestyle='--', linewidth=1, label='-1 Std Rel P/E')
                ax2.tick_params(axis='y', labelcolor=color2)
                ax2.set_yscale('log')

                # Third y-axis for Relative EPS
                ax3 = ax1.twinx()
                ax3.spines['right'].set_position(('outward', 60))
                ax3.set_ylabel('Relative EPS', color=color3, fontweight='bold')
                l3, = ax3.plot(merged['Date'], merged['Relative EPS'], color=color3, label=f"Relative EPS {ticker1}/{ticker2}")
                ax3.tick_params(axis='y', labelcolor=color3)
                ax3.set_yscale('log')

                plt.title(f'Relative Analysis: {ticker1} / {ticker2}', fontweight='bold')
                lines = [l1, l2, l3]
                labels = [l.get_label() for l in lines]
                ax1.legend(lines, labels, loc='upper left')
                fig.tight_layout()

                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                    plot_path = tmpfile.name
                    plt.savefig(plot_path, bbox_inches='tight')
                plt.close(fig)

                plots.append((f"{ticker1} / {ticker2}", plot_path))
            except Exception as e:
                print(f"Error processing {ticker1} and {ticker2}: {e}")

    return plots
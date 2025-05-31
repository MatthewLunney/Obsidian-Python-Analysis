import pandas as pd
import matplotlib.pyplot as plt
import os
import tempfile

def produce_individual_analysis(sector, start_date, end_date):
    excel_file_path = os.path.join('data', f"{sector}.xlsx")
    df_names = pd.read_excel(os.path.join('data', 'Company Names.xlsx'), sheet_name=sector)
    tickers = df_names['Ticker'].tolist()
    plots = []

    start_period = pd.Period(start_date, freq='M')
    end_period = pd.Period(end_date, freq='M')

    for ticker in tickers:
        try:
            df = pd.read_excel(excel_file_path, sheet_name=ticker, header=4)
            df = df.rename(columns={
                'Close Adj. Ex. Div.': 'Last Price',
                'EPS Basic - TTM': 'EPS',
                'P/E - TTM': 'P/E',
                'Dates': 'Date'
            })
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df.dropna(subset=['Date'])
            df['Period'] = df['Date'].dt.to_period('M')
            df = df[(df['Period'] >= start_period) & (df['Period'] <= end_period)]
            df = df.sort_values('Date')

            if df.empty:
                continue

            fig, ax1 = plt.subplots(figsize=(12, 6))
            color1 = 'black'
            color2 = 'tab:green'
            color3 = 'tab:orange'

            ax1.set_xlabel('Date', fontweight='bold')
            ax1.set_ylabel('Last Price', color=color1, fontweight='bold')
            l1, = ax1.plot(df['Date'], df['Last Price'], color=color1, label='Last Price')
            ax1.tick_params(axis='y', labelcolor=color1)
            ax1.set_yscale('log')

            # Second y-axis for P/E
            ax2 = ax1.twinx()
            ax2.set_ylabel('P/E', color=color2, fontweight='bold')
            l2, = ax2.plot(df['Date'], df['P/E'], color=color2, label='P/E')
            mean_pe = df['P/E'].mean()
            std_pe = df['P/E'].std()
            ax2.axhline(mean_pe, color='blue', linestyle='--', linewidth=1, label='Mean Rel P/E')
            ax2.axhline(mean_pe + std_pe, color='red', linestyle='--', linewidth=1, label='+1 Std Rel P/E')
            ax2.axhline(max(mean_pe - std_pe, 1e-6), color='green', linestyle='--', linewidth=1, label='-1 Std Rel P/E')
            ax2.tick_params(axis='y', labelcolor=color2)
            ax2.set_yscale('log')

            # Third y-axis for EPS
            ax3 = ax1.twinx()
            ax3.spines['right'].set_position(('outward', 60))
            ax3.set_ylabel('EPS', color=color3, fontweight='bold')
            l3, = ax3.plot(df['Date'], df['EPS'], color=color3, label='EPS')
            ax3.tick_params(axis='y', labelcolor=color3)
            ax3.set_yscale('log')

            plt.title(f'Individual Analysis: {ticker}', fontweight='bold')
            lines = [l1, l2, l3]
            labels = [l.get_label() for l in lines]
            ax1.legend(lines, labels, loc='upper left')
            fig.tight_layout()

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                plot_path = tmpfile.name
                plt.savefig(plot_path, bbox_inches='tight')
            plt.close(fig)

            plots.append((ticker, plot_path))
        except Exception as e:
            print(f"Error processing {ticker}: {e}")

    return plots
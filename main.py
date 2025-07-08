import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

OUTPUT_DIR = Path("customer-demand-graphs")
OUTPUT_DIR.mkdir(exist_ok=True)

def load_data(file_path):
    df = pd.read_excel(file_path, dtype={'creationDate': str, 'capturingTime': str})
    df['creationDate'] = df['creationDate'].astype(str).str.strip("'").str.strip()
    df['capturingTime'] = df['capturingTime'].astype(str).str.strip()
    # Ensure capturingTime always has milliseconds
    df['capturingTime'] = df['capturingTime'].str.replace(
        r'^(\d{2}:\d{2}:\d{2})$', r'\1.000', regex=True
    )
    df['capturingDateTime'] = pd.to_datetime(df['creationDate'] + ' ' + df['capturingTime'], errors='coerce')
    df['creationDate'] = pd.to_datetime(df['creationDate'], errors='coerce')
    df['amountInEuros'] = df['amountIncludingTipInCents'] / 100
    return df

def plot_payments_per_hour(df, date):
    # Use Windows-compatible format codes
    filename = OUTPUT_DIR / f"{date.strftime('%#d.%#m.%Y')}_demand.png"
    if filename.exists():
        print(f"Skipping {filename} (already exists)")
        return

    day_df = df[df['creationDate'].dt.date == date]
    if day_df.empty:
        print(f"No data for {date}")
        return

    day_df = day_df.copy()
    day_df['hour'] = day_df['capturingDateTime'].dt.hour
    hourly = day_df.groupby('hour')['amountInEuros'].sum()

    plt.figure(figsize=(10, 6))
    ax = hourly.plot(kind='bar', color='skyblue')
    plt.title(f"Payments Captured per Hour ({date.strftime('%d.%m.%Y')})")
    plt.xlabel("Hour")
    plt.ylabel("Total Payments (€)")
    plt.xticks(rotation=0)
    plt.tight_layout()

    daily_total = hourly.sum()
    daily_avg = day_df['amountInEuros'].mean()

    # Payment method totals
    method_totals = day_df.groupby('methodCode')['amountInEuros'].sum()
    method_text = "\n".join(
        f"{method}: €{amount:,.2f}" for method, amount in method_totals.items()
    )

    plt.text(
        0.01, 0.01,
        method_text,
        ha='left', va='bottom',
        transform=ax.transAxes,
        fontsize=11, fontweight='normal',
        bbox=dict(facecolor='white', alpha=0.7, edgecolor='none')
    )

    plt.text(
        0.99, 0.06,
        f"Average: €{daily_avg:,.2f}",
        ha='right', va='bottom',
        transform=ax.transAxes,
        fontsize=11, fontweight='normal',
        bbox=dict(facecolor='white', alpha=0.7, edgecolor='none')
    )
    plt.text(
        0.99, 0.01,
        f"Total: €{daily_total:,.2f}",
        ha='right', va='bottom',
        transform=ax.transAxes,
        fontsize=12, fontweight='bold',
        bbox=dict(facecolor='white', alpha=0.7, edgecolor='none')
    )

    plt.savefig(filename)
    plt.close()
    print(f"Saved {filename}")

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select payments_captured.xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not file_path:
        print("No file selected.")
        return

    df = load_data(file_path)
    unique_dates = df['creationDate'].dt.date.unique()
    for date in unique_dates:
        plot_payments_per_hour(df, date)

if __name__ == "__main__":
    main()
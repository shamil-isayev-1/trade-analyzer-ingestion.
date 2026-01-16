import pandas as pd
import os
import openpyxl
from openpyxl.styles import Alignment
import sys

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_PATH = r'C:\Users\user\PycharmProjects\PythonProject\Labaratory\draft\Отчет.csv'
OUTPUT_NAME = 'Trade_Report.xlsx'


def load_data(file_path):
    """Step 1: Data ingestion with proper encoding handling."""
    if not os.path.exists(file_path):
        print(f"ERROR: File not found at {file_path}")
        sys.exit()

    print("File detected. Starting ingestion...")
    # 'utf-8-sig' handles the BOM (Byte Order Mark) from broker CSV exports
    return pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig')


def clean_and_calculate(df):
    """Step 2: Data cleaning, type casting, and P&L calculations."""
    # Drop rows without a Ticker (empty/invalid trades)
    df = df.dropna(subset=['Тикер']).copy()

    # Normalizing decimal separators and casting to numeric
    cols_to_fix = ['Цена за штуку', 'Объем транзакции']
    for col in cols_to_fix:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '.')
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Calculated column for validation
    df["Calculated_Sum"] = df['Количество'] * df['Цена за штуку']
    return df


def save_to_excel_pretty(df, output_name):
    """Step 3: Excel export with visual post-processing."""
    try:
        report_cols = [
            'Дата', 'Тикер', 'Операция', 'Количество',
            'Цена за штуку', 'Объем транзакции', 'Calculated_Sum'
        ]

        # Select available columns
        existing_cols = [c for c in report_cols if c in df.columns]
        df[existing_cols].to_excel(output_name, index=False)

        # Styling via openpyxl
        wb = openpyxl.load_workbook(output_name)
        ws = wb.active

        for column in ws.columns:
            max_len = 0
            col_letter = column[0].column_letter

            for cell in column:
                # Center alignment
                cell.alignment = Alignment(horizontal='center', vertical='center')

                # Auto-fit width calculation
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = max_len + 3

        wb.save(output_name)
        print(f"--- Success! '{output_name}' generated with auto-formatting ---")

    except PermissionError:
        print(f"PERMISSION ERROR: Please close '{output_name}' in Excel and retry.")
    except Exception as e:
        print(f"Unexpected error: {e}")


def main():
    """Main execution pipeline."""
    raw_df = load_data(INPUT_PATH)
    clean_df = clean_and_calculate(raw_df)
    save_to_excel_pretty(clean_df, OUTPUT_NAME)


if __name__ == "__main__":
    main()


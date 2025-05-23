import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback
from datetime import datetime

# Emoji mapping for countries (English names)
EMOJI_MAP = {
    'AZERBAIJAN': '🇦🇿',
    'ARMENIA': '🇦🇲',
    'IVORY COAST': '🇨🇮',
    'ZAMBIA': '🇿🇲',
    'UAE': '🇦🇪',
    'UZBEKISTAN': '🇺🇿',
    'PERU': '🇵🇪'
}

# Source-to-standard column names (case-insensitive)
SOURCE_COLUMNS = {
    'report date': 'Date',
    'sla, %': 'SLA',
    'avg csat': 'CSAT',
    'full resolution sla %': 'FR',
    'sessions': 'Sessions',
    'country': 'Country'
}


def find_valid_sheet(path):
    """Find first sheet containing any known source column."""
    xl = pd.ExcelFile(path, engine='openpyxl')
    for sheet in xl.sheet_names:
        df0 = xl.parse(sheet, nrows=0)
        cols = [c.strip().lower() for c in df0.columns]
        if any(src in cols for src in SOURCE_COLUMNS.keys()):
            return sheet
    return xl.sheet_names[0]


def normalize_and_rename(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names and rename known columns to standard names."""
    rename_map = {}
    for col in df.columns:
        norm = col.strip().lower()
        if norm in SOURCE_COLUMNS:
            rename_map[col] = SOURCE_COLUMNS[norm]
    return df.rename(columns=rename_map)


def generate_report(df: pd.DataFrame) -> str:
    """
    Generate formatted report text from DataFrame with Date, Country, Sessions, SLA, CSAT, FR.
    Assumes df already has standardized columns.
    """
    missing = [req for req in ('Date','Country') if req not in df.columns]
    if missing:
        raise KeyError(f"Missing columns: {missing}. Available: {list(df.columns)}")

    lines = []
    for date, group in df.groupby('Date'):
        lines.append(f"📝{date.strftime('%d/%m/%Y')} #б2бинормал\n")
        for _, row in group.iterrows():
            raw_country = str(row['Country']).split('|')[-1].strip()
            emoji = EMOJI_MAP.get(raw_country.upper(), '')
            sessions = int(row.get('Sessions', 0)) if pd.notnull(row.get('Sessions')) else 0
            sla = row.get('SLA')
            csat = row.get('CSAT')
            fr = row.get('FR')
            sla_str = f"{float(sla):.1%}" if pd.notnull(sla) else 'no'
            csat_str = f"{float(csat):.2f}" if pd.notnull(csat) else 'no'
            fr_str = f"{float(fr):.1%}" if pd.notnull(fr) else 'no'
            # Add country block with blank line after
            lines.append(
                f"{raw_country}{emoji}\n"
                f"Sessions – {sessions} | SLA 5 min – {sla_str} | CSAT – {csat_str} | FR – {fr_str}\n\n"
            )
        # Blank line between dates
        lines.append("\n")
    return ''.join(lines)


def write_error_log(filepath, error_tb):
    downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_path = os.path.join(downloads, f"support_stats_error_{timestamp}.log")
    try:
        with open(log_path, 'w', encoding='utf-8') as log:
            log.write(f"Error processing file: {filepath}\n")
            log.write("Traceback:\n" + error_tb + "\n")
            try:
                xl = pd.ExcelFile(filepath, engine='openpyxl')
                log.write("Sheets: " + ", ".join(xl.sheet_names) + "\n")
                for sheet in xl.sheet_names:
                    cols = xl.parse(sheet, nrows=0).columns.tolist()
                    log.write(f"{sheet} columns: {cols}\n")
            except Exception as e:
                log.write(f"Failed listing sheets/columns: {e}\n")
        return log_path
    except Exception:
        return None


def load_file_and_generate():
    filepath = filedialog.askopenfilename(
        title="Open Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    if not filepath:
        return
    try:
        sheet = find_valid_sheet(filepath)
        df_raw = pd.read_excel(filepath, sheet_name=sheet, engine='openpyxl')
        df = normalize_and_rename(df_raw)
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
        if df['Date'].isna().all():
            raise ValueError("All dates could not be parsed. Check 'Report date' format.")

        report_text = generate_report(df)

        dates = df['Date'].dropna()
        min_date = dates.min().strftime('%Y%m%d')
        max_date = dates.max().strftime('%Y%m%d')
        downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
        filename = f"support_stats_{min_date}_{max_date}.txt"
        save_path = os.path.join(downloads, filename)
        # Write with BOM so Notepad displays UTF-8 emojis
        with open(save_path, 'w', encoding='utf-8-sig') as f:
            f.write(report_text)
        messagebox.showinfo("Success", f"Report saved to Downloads:\n{save_path}")
    except Exception:
        tb = traceback.format_exc()
        log_path = write_error_log(filepath, tb)
        msg = "Failed to process file."
        if log_path:
            msg += f"\nSee log at: {log_path}"
        messagebox.showerror("Error", msg)


if __name__ == '__main__':
    root = tk.Tk()
    root.title("Support Stats Report Generator")
    root.geometry('450x200')

    btn = tk.Button(
        root,
        text="Load Excel and Generate Report",
        command=load_file_and_generate,
        padx=10,
        pady=10
    )
    btn.pack(expand=True)

    root.mainloop()

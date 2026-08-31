import csv
import re
from datetime import datetime, time
from pathlib import Path

import pandas as pd

SOURCE = Path(r"C:\Users\ADMINI~1\Downloads\Registro de salidas diarias 2026 (1) (1).xlsx")
OUTPUT = Path(r"C:\Users\ADMINI~1\Documents\Codex\2026-08-31\nec\work\registro-diario-funcionarios\legacy-records.csv")
MONTHS = [3, 4, 5]
TIME_RE = re.compile(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)")


def times(value):
    if isinstance(value, (datetime, time)):
        return [value.strftime("%H:%M")] if 7 <= value.hour < 24 else []
    if pd.isna(value):
        return []
    return [f"{int(h):02d}:{int(m):02d}" for h, m in TIME_RE.findall(str(value)) if 7 <= int(h) < 24]


def main():
    book = pd.ExcelFile(SOURCE)
    rows = []
    for sheet_name in book.sheet_names:
        if sheet_name == "Hoja1":
            continue
        data = pd.read_excel(SOURCE, sheet_name=sheet_name, header=None)
        header_rows = []
        for i in range(len(data)):
            day_headers = 0
            for value in data.iloc[i].tolist():
                try:
                    number = float(value)
                    if number.is_integer() and 1 <= number <= 31:
                        day_headers += 1
                except (TypeError, ValueError):
                    continue
            if day_headers >= 3:
                header_rows.append(i)
        for block, header_index in enumerate(header_rows[:3]):
            month = MONTHS[block]
            dates = data.iloc[header_index].tolist()
            end = header_rows[block + 1] if block + 1 < len(header_rows) else len(data)
            for col, day in enumerate(dates):
                if pd.isna(day):
                    continue
                try:
                    day_number = int(float(day))
                except (TypeError, ValueError):
                    continue
                if not 1 <= day_number <= 31:
                    continue
                tokens = []
                for value in data.iloc[header_index + 1:end, col].tolist():
                    tokens.extend(times(value))
                for index in range(0, len(tokens) - 1, 2):
                    fecha = f"2026-{month:02d}-{day_number:02d}"
                    salida, regreso = tokens[index:index + 2]
                    rows.append([fecha, salida, sheet_name, "Salida", "Registro importado del Excel", "importado"])
                    rows.append([fecha, regreso, sheet_name, "Regreso", "Registro importado del Excel", "importado"])

    with OUTPUT.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Fecha", "Hora", "Funcionario", "Tipo", "Motivo", "Usuario"])
        writer.writerows(rows)
    print(f"records={len(rows)}")


if __name__ == "__main__":
    main()

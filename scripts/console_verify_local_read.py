from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.revenue_loader import _parse_revenue_workbook


base = Path(r"C:\Users\gaura\OneDrive - RCG Global Services, Inc\FInance\Solutions Revenue")
files = sorted(base.glob("Solutions_Revenue_*.xlsx"))
template = base / "Revenue Output Dashboard Sample.xlsx"

print(f"base={base}")
print(f"template_exists={template.exists()}")
print("monthly_files=")
for path in files:
    print(f"- {path.name}")

parsed_total_rows = 0
for path in files:
    frame = _parse_revenue_workbook(path)
    parsed_total_rows += len(frame)
    months = sorted({str(value.date()) for value in frame["month"].dropna().tolist()}) if not frame.empty else []
    print(f"{path.name}: rows={len(frame)} months={months}")

print(f"parsed_total_rows={parsed_total_rows}")

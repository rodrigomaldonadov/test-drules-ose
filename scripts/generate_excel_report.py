import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.rule_catalog import DEFAULT_GENERATED_DIR, DEFAULT_REPORTS_DIR, DEFAULT_SOURCE_DIR, generate_excel_report


if __name__ == "__main__":
    output_path = DEFAULT_REPORTS_DIR / "reporte_reglas_completo.xlsx"
    result = generate_excel_report(DEFAULT_SOURCE_DIR, DEFAULT_GENERATED_DIR, output_path)
    print(f"Reporte generado: {result}")

from scripts.rule_catalog import associate_impl_with_source, write_report
from pathlib import Path

ROOT = Path(__file__).resolve().parent
IMPL_FILE = ROOT / "rules_invoice_impl.drl"
SOURCE_DIR = ROOT / "rules" / "source"
REPORT_FILE = ROOT / "rule_association_report.json"


if __name__ == "__main__":
    report = associate_impl_with_source(IMPL_FILE, SOURCE_DIR)
    write_report(report, ROOT)
    print(report["summary"])

"""Command-line entrypoint for append-only Excel imports."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.service import update_deal_database


def main() -> None:
    parser = argparse.ArgumentParser(description="Append PE/VC deal records from Excel into SQLite.")
    parser.add_argument("--excel", required=True, help="Path to Excel workbook.")
    parser.add_argument("--db", default="data/database/deals.db", help="Path to SQLite database.")
    args = parser.parse_args()

    result = update_deal_database(args.excel, args.db)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()


from pathlib import Path
import sqlite3
import tempfile
import unittest

import pandas as pd

from src.normalizer import normalize_project_short_name
from src.service import check_deal_exists, get_deal_detail, update_deal_database


class DealCheckerTest(unittest.TestCase):
    def test_append_without_dedup_and_check(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "deals.xlsx"
            db_path = tmp_path / "deals.db"
            df = pd.DataFrame(
                [
                    {"项目简称": " 星火AI ", "城市": "上海", "是否通过": "是"},
                    {"项目简称": "", "城市": "北京", "是否通过": "否"},
                    {"项目简称": "星火AI", "城市": "上海", "是否通过": "是"},
                ]
            )
            with pd.ExcelWriter(excel_path) as writer:
                df.to_excel(writer, sheet_name="通过", index=False)

            first = update_deal_database(str(excel_path), str(db_path))
            second = update_deal_database(str(excel_path), str(db_path))

            self.assertEqual(first["mode"], "append_all")
            self.assertEqual(first["deduplication"], "disabled")
            self.assertEqual(first["inserted_rows"], 2)
            self.assertEqual(first["skipped_empty_name_rows"], 1)
            self.assertEqual(first["skipped_duplicate_rows"], 0)
            self.assertEqual(second["inserted_rows"], 2)
            self.assertEqual(second["skipped_duplicate_rows"], 0)

            exact = check_deal_exists("星火 ai", str(db_path))
            self.assertEqual(exact["result"], "已存在")
            self.assertEqual(exact["match_type"], "exact")
            self.assertEqual(len(exact["matches"]), 4)
            self.assertEqual(exact["matches"][0]["source_sheet"], "通过")
            self.assertEqual(exact["matches"][0]["source_row"], 2)

            contains = check_deal_exists("星火", str(db_path))
            self.assertEqual(contains["result"], "疑似存在")
            self.assertEqual(contains["match_type"], "contains")

            missing = check_deal_exists("不存在项目", str(db_path))
            self.assertEqual(missing["result"], "未查到")
            self.assertEqual(missing["match_type"], "none")

            detail = get_deal_detail(exact["matches"][0]["id"], str(db_path))
            self.assertEqual(detail["project_short_name"], "星火AI")

    def test_normalizer(self):
        self.assertEqual(normalize_project_short_name("《ABC（测试） 123》"), "abc测试123")

    def test_legacy_unique_database_accepts_duplicates_after_init(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "deals.xlsx"
            db_path = tmp_path / "legacy.db"
            df = pd.DataFrame([{"项目简称": "星火AI", "城市": "上海"}])
            with pd.ExcelWriter(excel_path) as writer:
                df.to_excel(writer, sheet_name="通过", index=False)

            conn = sqlite3.connect(db_path)
            try:
                conn.execute(
                    """
                    CREATE TABLE deals (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        record_key TEXT NOT NULL UNIQUE,
                        project_short_name TEXT NOT NULL,
                        normalized_project_short_name TEXT NOT NULL,
                        founded_time TEXT,
                        city TEXT,
                        expected_application TEXT,
                        previous_valuation TEXT,
                        invested_institutions TEXT,
                        pre_money_valuation TEXT,
                        current_round_amount TEXT,
                        financing_deadline TEXT,
                        main_business TEXT,
                        value_description TEXT,
                        revenue TEXT,
                        profit TEXT,
                        deal_source TEXT,
                        pass_status TEXT,
                        notes_or_rejection_reason TEXT,
                        source_file TEXT,
                        source_sheet TEXT,
                        source_row INTEGER,
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL
                    )
                    """
                )
                conn.commit()
            finally:
                conn.close()

            first = update_deal_database(str(excel_path), str(db_path))
            second = update_deal_database(str(excel_path), str(db_path))

            self.assertEqual(first["inserted_rows"], 1)
            self.assertEqual(second["inserted_rows"], 1)
            self.assertEqual(len(check_deal_exists("星火AI", str(db_path))["matches"]), 2)


if __name__ == "__main__":
    unittest.main()


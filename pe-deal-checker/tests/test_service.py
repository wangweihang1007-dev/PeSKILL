from pathlib import Path
import tempfile
import unittest

import pandas as pd

from src.normalizer import normalize_project_short_name
from src.service import check_deal_exists, get_deal_detail, update_deal_database


class DealCheckerTest(unittest.TestCase):
    def test_append_and_check(self):
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

            self.assertEqual(first["inserted_rows"], 1)
            self.assertEqual(first["skipped_empty_name_rows"], 1)
            self.assertEqual(first["skipped_duplicate_rows"], 1)
            self.assertEqual(second["inserted_rows"], 0)
            self.assertEqual(second["skipped_duplicate_rows"], 2)

            exact = check_deal_exists("星火 ai", str(db_path))
            self.assertEqual(exact["result"], "已存在")
            self.assertEqual(exact["match_type"], "exact")
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


if __name__ == "__main__":
    unittest.main()


import csv
import tempfile
import unittest
from pathlib import Path

from db import CatalogDB


class ManualPriceImportTests(unittest.TestCase):
    def test_insert_update_unmatched_and_skipped(self):
        with tempfile.TemporaryDirectory() as temp_name:
            temp = Path(temp_name)
            database = CatalogDB(temp / "test.db")
            try:
                product_csv = temp / "products.csv"
                with product_csv.open("w", encoding="utf-8-sig", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=["分类", "品牌", "型号", "启用"])
                    writer.writeheader()
                    writer.writerow({"分类": "CPU", "品牌": "测试品牌", "型号": "测试型号", "启用": "1"})
                inserted, updated, skipped = database.import_products(product_csv)
                self.assertEqual((inserted, updated, skipped), (1, 0, 0))
                product_id = database.conn.execute("SELECT id FROM products").fetchone()[0]

                price_csv = temp / "prices.csv"
                fields = ["产品ID", "品牌", "型号", "平台", "平台商品ID", "价格类型", "价格", "更新时间"]
                with price_csv.open("w", encoding="utf-8-sig", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=fields)
                    writer.writeheader()
                    writer.writerow({"产品ID": product_id, "平台": "手动渠道", "平台商品ID": "SKU-1", "价格类型": "进货价", "价格": "100", "更新时间": "2026-08-03 10:00:00"})
                    writer.writerow({"品牌": "不存在", "型号": "不存在", "平台": "手动渠道", "平台商品ID": "SKU-X", "价格": "200"})
                    writer.writerow({"产品ID": product_id, "平台": "手动渠道", "平台商品ID": "SKU-S", "价格": ""})
                self.assertEqual(database.import_offers(price_csv), (1, 0, 1, 1))

                with price_csv.open("w", encoding="utf-8-sig", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=fields)
                    writer.writeheader()
                    writer.writerow({"产品ID": product_id, "平台": "手动渠道", "平台商品ID": "SKU-1", "价格类型": "进货价", "价格": "88", "更新时间": "2026-08-03 11:00:00"})
                self.assertEqual(database.import_offers(price_csv), (0, 1, 0, 0))

                offer = database.conn.execute("SELECT price FROM offers WHERE platform_sku='SKU-1'").fetchone()
                self.assertEqual(float(offer[0]), 88.0)
            finally:
                database.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)

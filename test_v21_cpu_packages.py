from __future__ import annotations

import csv
import sqlite3
import tempfile
import unittest
from collections import Counter
from pathlib import Path

from db import CatalogDB
import backend


CPU_CANDIDATES = Path(__file__).resolve().parent / "catalog_v2_research" / "cpu_candidates.csv"


class V21CpuPackageTests(unittest.TestCase):
    def test_existing_template_header_is_upgraded_and_rows_are_preserved(self):
        with tempfile.TemporaryDirectory(prefix="limo_v21_template_") as folder:
            folder_path = Path(folder)
            template = folder_path / "型号主库导入模板.csv"
            with template.open("w", encoding="utf-8-sig", newline="") as fh:
                writer = csv.writer(fh)
                writer.writerow(["分类", "品牌", "系列", "型号", "型号别名"])
                writer.writerow(["CPU", "AMD", "Ryzen 9000", "Ryzen 7 TEST（盒装）", "TEST"])
            previous_root = backend.ROOT
            try:
                backend.ROOT = folder_path
                backend.create_templates()
            finally:
                backend.ROOT = previous_root
            with template.open("r", encoding="utf-8-sig", newline="") as fh:
                rows = list(csv.DictReader(fh))
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["型号"], "Ryzen 7 TEST（盒装）")
            self.assertEqual(rows[0]["基础型号"], "")
            self.assertEqual(rows[0]["包装类型"], "")
            self.assertEqual(rows[0]["保修说明"], "")

    def test_old_database_is_migrated_without_guessing_package(self):
        with tempfile.TemporaryDirectory(prefix="limo_v21_migration_") as folder:
            db_path = Path(folder) / "legacy.db"
            conn = sqlite3.connect(db_path)
            conn.execute(
                """
                CREATE TABLE products (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  category TEXT NOT NULL DEFAULT '', brand TEXT NOT NULL DEFAULT '',
                  model TEXT NOT NULL, spec TEXT NOT NULL DEFAULT '', cost_price REAL,
                  sale_price REAL, default_url TEXT NOT NULL DEFAULT '', stock TEXT NOT NULL DEFAULT '',
                  updated_at TEXT NOT NULL DEFAULT '', compat1 TEXT NOT NULL DEFAULT '',
                  compat2 TEXT NOT NULL DEFAULT '', mpn TEXT NOT NULL DEFAULT '', gtin TEXT NOT NULL DEFAULT '',
                  active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL,
                  UNIQUE(category, brand, model)
                )
                """
            )
            conn.execute(
                "INSERT INTO products(category,brand,model,created_at) VALUES('CPU','AMD','Ryzen 7 7800X3D','2026-01-01')"
            )
            conn.commit()
            conn.close()

            db = CatalogDB(db_path)
            row = db.conn.execute(
                "SELECT base_model,package_type,warranty_note FROM products"
            ).fetchone()
            self.assertEqual(row["base_model"], "Ryzen 7 7800X3D")
            self.assertEqual(row["package_type"], "未区分")
            self.assertEqual(row["warranty_note"], "")
            self.assertEqual(db.conn.execute("PRAGMA user_version").fetchone()[0], 3)
            db.close()

    @unittest.skipUnless(CPU_CANDIDATES.exists(), "cpu_candidates.csv 尚未生成")
    def test_import_52_cpu_package_candidates_and_export(self):
        with tempfile.TemporaryDirectory(prefix="limo_v21_cpu_") as folder:
            folder_path = Path(folder)
            db = CatalogDB(folder_path / "cpu.db")
            inserted, updated, skipped = db.import_products(CPU_CANDIDATES)
            self.assertEqual((inserted, updated, skipped), (52, 0, 0))

            rows = db.conn.execute(
                "SELECT model,base_model,package_type,warranty_note FROM products ORDER BY id"
            ).fetchall()
            self.assertEqual(len(rows), 52)
            self.assertEqual(Counter(r["package_type"] for r in rows), {"盒装": 27, "散片/Tray": 25})
            self.assertTrue(all(r["base_model"] for r in rows))
            self.assertTrue(all(r["warranty_note"] for r in rows))
            self.assertEqual(len({r["model"] for r in rows}), 52)
            self.assertTrue(all("（盒装）" in r["model"] or "（散片）" in r["model"] for r in rows))

            self.assertEqual(db.import_products(CPU_CANDIDATES), (0, 52, 0))
            listed = db.list_products("散片/Tray")
            self.assertEqual(len(listed), 25)
            self.assertIn("base_model", listed[0])
            self.assertIn("warranty_note", listed[0])

            export_path = folder_path / "export.csv"
            self.assertEqual(db.export_products(export_path), 52)
            with export_path.open("r", encoding="utf-8-sig", newline="") as fh:
                exported = list(csv.DictReader(fh))
            self.assertEqual(len(exported), 52)
            self.assertEqual(Counter(r["包装类型"] for r in exported), {"盒装": 27, "散片/Tray": 25})
            self.assertTrue(all(r["基础型号"] and r["保修说明"] for r in exported))
            db.close()

    def test_legacy_csv_remains_importable(self):
        with tempfile.TemporaryDirectory(prefix="limo_v21_legacy_csv_") as folder:
            folder_path = Path(folder)
            csv_path = folder_path / "legacy.csv"
            with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
                writer = csv.writer(fh)
                writer.writerow(["分类", "品牌", "型号", "规格参数", "启用"])
                writer.writerow(["CPU", "Intel", "Core i5-12400F", "6核12线程", 1])
                writer.writerow(["主板", "华硕", "TUF GAMING B850M-PLUS WIFI", "AM5", 1])
            db = CatalogDB(folder_path / "legacy_csv.db")
            self.assertEqual(db.import_products(csv_path), (2, 0, 0))
            cpu = db.conn.execute("SELECT * FROM products WHERE category='CPU'").fetchone()
            board = db.conn.execute("SELECT * FROM products WHERE category='主板'").fetchone()
            self.assertEqual((cpu["base_model"], cpu["package_type"]), ("Core i5-12400F", "未区分"))
            self.assertEqual((board["base_model"], board["package_type"]), ("", ""))
            db.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)

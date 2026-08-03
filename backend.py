from __future__ import annotations

# © 2026 离墨。保留所有权利。
import argparse
import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.dont_write_bytecode = True

from db import CatalogDB


ROOT = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = str(ROOT / "报价单" / "离墨DIY电脑装机配置报价单.xlsm")


def config_path():
    return ROOT / "config.json"


def load_config():
    default = {
        "workbook_path": DEFAULT_WORKBOOK,
        "max_sync_rows": 0,
    }
    if config_path().exists():
        try:
            saved = json.loads(config_path().read_text(encoding="utf-8-sig"))
            saved.pop("platforms", None)
            if saved.get("workbook_path"):
                saved["workbook_path"] = os.path.normpath(saved["workbook_path"])
            default.update(saved)
        except Exception:
            pass
    return default


def save_config(config):
    config_path().write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def create_templates():
    templates = {
        "价格单导入模板.csv": ["产品ID","分类","品牌","型号","MPN","GTIN","平台","平台商品ID","店铺","商品标题","价格类型","价格","链接","库存","绑定状态","匹配置信度","首选","更新时间"],
        "型号主库导入模板.csv": ["分类","品牌","系列","型号","基础型号","包装类型","保修说明","型号别名","规格参数","MPN","GTIN","数据来源","官网链接","在售状态","最后确认时间","手动进货价","手动销售价","商品链接","库存","兼容参数1","兼容参数2","启用","同步报价单"],
        "供应商报价导入模板.csv": ["供应商","供应商优先级","联系方式","产品ID","分类","品牌","型号","MPN","GTIN","供应商商品ID","进货价","库存","链接","报价时间","备注","启用"],
        "渠道SKU绑定模板.csv": ["产品ID","分类","品牌","型号","MPN","GTIN","平台","平台商品ID","店铺","商品标题","价格类型","价格","链接","库存","绑定状态","匹配置信度","首选","更新时间"],
    }
    for name, headers in templates.items():
        path = ROOT / name
        existing_rows = []
        existing_headers = []
        if path.exists():
            for encoding in ("utf-8-sig", "gb18030"):
                try:
                    with path.open("r", encoding=encoding, newline="") as fh:
                        reader = csv.DictReader(fh)
                        existing_headers = reader.fieldnames or []
                        existing_rows = list(reader)
                    break
                except UnicodeDecodeError:
                    continue
        if existing_headers == headers:
            continue
        # 模板属于程序文件。升级表头时按同名列保留用户可能填写的示例行，
        # 新字段留空；使用临时文件替换，避免中途写入失败留下半个 CSV。
        temporary = path.with_suffix(path.suffix + ".tmp")
        with temporary.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for row in existing_rows:
                writer.writerow({header: row.get(header, "") for header in headers})
        temporary.replace(path)


def initialize():
    create_templates()
    CatalogDB(ROOT / "电脑配件数据库.db").close()
    save_config(load_config())


def sync_workbook(db, workbook_path, max_rows):
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"找不到报价单：{workbook_path}")
    if workbook_path.suffix.lower() != ".xlsm":
        raise ValueError("请选择 .xlsm 宏工作簿。")
    rows = db.workbook_rows(None if int(max_rows) <= 0 else int(max_rows))
    backup_dir = ROOT / "备份"
    backup_dir.mkdir(exist_ok=True)
    backup = backup_dir / f"{workbook_path.stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsm"
    shutil.copy2(workbook_path, backup)
    backups = sorted(backup_dir.glob("*.xlsm"), key=lambda p: p.stat().st_mtime, reverse=True)
    for old in backups[10:]:
        try:
            old.unlink()
        except OSError:
            pass
    fd, json_name = tempfile.mkstemp(prefix="pc_catalog_", suffix=".json")
    os.close(fd)
    payload = Path(json_name)
    try:
        payload.write_text(json.dumps([{"values": row} for row in rows], ensure_ascii=False), encoding="utf-8-sig")
        powershell = shutil.which("pwsh.exe") or shutil.which("powershell.exe") or "powershell.exe"
        command = [
            powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
            str(ROOT / "sync_workbook.ps1"), "-WorkbookPath", str(workbook_path),
            "-JsonPath", str(payload), "-MaxRows", str(max_rows),
        ]
        done = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
        if done.returncode != 0:
            raise RuntimeError((done.stderr or done.stdout).strip() or "Excel 同步失败。")
    finally:
        payload.unlink(missing_ok=True)
    return len(rows), backup


def open_workbook(workbook_path):
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"找不到报价单：{workbook_path}")
    if workbook_path.suffix.lower() != ".xlsm":
        raise ValueError("请选择 .xlsm 宏工作簿。")
    os.startfile(str(workbook_path))
    return workbook_path


def output(value):
    print(json.dumps(value, ensure_ascii=False))


def product_output_rows(product_rows):
    rows = []
    for r in product_rows:
        rows.append({
            "id":r["id"],"category":r["category"],"brand":r["brand"],"series":r["series"],
            "model":r["model"],"base_model":r["base_model"],
            "package_type":r["package_type"],"warranty_note":r["warranty_note"],
            "cost":r["cost"],"cost_source":r["cost_source"],
            "reference_price":r["reference_price"],"suppliers":r["supplier_count"],
            "bindings":r["binding_count"],"stock":r["stock"],"lifecycle":r["lifecycle"],
            "active":r["active"],"quote_enabled":r["quote_enabled"],"updated":r["updated_at"],
            # v1.4 compatibility keys
            "price":r["cost"],"offers":r["binding_count"],
        })
    return rows


def product_page_rows(product_rows):
    return [{
        "id":r["id"], "category":r["category"], "brand":r["brand"],
        "series":r["series"], "model":r["model"], "package_type":r["package_type"],
        "cost":r["cost"], "cost_source":r["cost_source"],
        "reference_price":r["reference_price"], "active":r["active"],
        "quote_enabled":r["quote_enabled"],
    } for r in product_rows]


def quote_page_rows(product_rows):
    return [{
        "id":r["id"], "category":r["category"], "brand":r["brand"],
        "model":r["model"], "package_type":r["package_type"], "cost":r["cost"],
    } for r in product_rows]


def product_preview_limit(search, category, quote_state, active_state):
    preview_only = (
        not search.strip()
        and category in {"", "全部分类"}
        and quote_state in {"", "全部", "全部报价状态", "已加入"}
        and active_state in {"", "启用"}
    )
    return 1000 if preview_only else 10000


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--init", action="store_true")
    parser.add_argument("--stats", action="store_true")
    parser.add_argument("--dashboard", action="store_true")
    parser.add_argument("--list", nargs="?", const="", metavar="SEARCH")
    parser.add_argument("--list-products", nargs="?", const="", metavar="SEARCH")
    parser.add_argument("--list-products-advanced", nargs=5, metavar=("SEARCH", "CATEGORY", "QUOTE", "ACTIVE", "MATCH"))
    parser.add_argument("--products-page", nargs=5, metavar=("SEARCH", "CATEGORY", "QUOTE", "ACTIVE", "MATCH"))
    parser.add_argument("--list-supplier-offers", nargs="?", const="", metavar="SEARCH")
    parser.add_argument("--list-bindings", nargs="?", const="", metavar="SEARCH")
    parser.add_argument("--list-prices", nargs="?", const="", metavar="SEARCH")
    parser.add_argument("--import-products", metavar="CSV")
    parser.add_argument("--import-offers", metavar="CSV")
    parser.add_argument("--import-supplier-offers", metavar="CSV")
    parser.add_argument("--import-platform-items", metavar="CSV")
    parser.add_argument("--export", metavar="CSV")
    parser.add_argument("--set-quote-enabled", nargs=2, metavar=("PRODUCT_ID", "ENABLED"))
    parser.add_argument("--set-quote-enabled-batch", nargs=2, metavar=("PRODUCT_IDS", "ENABLED"))
    parser.add_argument("--clear-quote-enabled", action="store_true")
    parser.add_argument("--set-products-active", nargs=2, metavar=("PRODUCT_IDS", "ENABLED"))
    parser.add_argument("--delete-products", metavar="PRODUCT_IDS")
    parser.add_argument("--set-binding-status", nargs=2, metavar=("ITEM_ID", "STATUS"))
    parser.add_argument("--sync", action="store_true")
    parser.add_argument("--open-workbook", action="store_true")
    parser.add_argument("--get-config", action="store_true")
    parser.add_argument("--set-config", nargs=2, metavar=("WORKBOOK", "MAX_ROWS"))
    args = parser.parse_args()
    if args.init:
        initialize()
        output({"ok": True})
        return
    config = load_config()
    if args.get_config:
        output(config)
        return
    if args.set_config:
        workbook, max_rows = args.set_config
        max_rows = max(0, int(max_rows))
        config["workbook_path"] = workbook
        config["max_sync_rows"] = max_rows
        save_config(config)
        output({"ok": True})
        return
    if args.open_workbook:
        path = open_workbook(config["workbook_path"])
        output({"ok": True, "path": str(path)})
        return
    db = CatalogDB(ROOT / "电脑配件数据库.db")
    try:
        if args.products_page is not None:
            search, category, quote_state, active_state, match_mode = args.products_page
            dashboard = db.dashboard_counts()
            dashboard["offers"] = dashboard["prices"]
            full_limit = max(1, int(dashboard["products"]))
            products = db.list_products_advanced(
                search, category, quote_state, active_state, match_mode,
                full_limit,
            )
            quote_products = db.list_products_advanced(
                "", "全部分类", "已加入", "启用", "智能搜索", full_limit,
            )
            output({
                "dashboard": dashboard,
                "products": product_page_rows(products),
                "quote_products": quote_page_rows(quote_products),
            })
        elif args.stats or args.dashboard:
            data = db.dashboard_counts()
            data["offers"] = data["prices"]
            output(data)
        elif args.list_products_advanced is not None or args.list is not None or args.list_products is not None:
            if args.list_products_advanced is not None:
                search, category, quote_state, active_state, match_mode = args.list_products_advanced
                # 首页不带条件时只预览前 1,000 条，避免 WinForms 一次渲染数千行；
                # 只要输入关键词、选择分类或筛选报价状态，就会搜索完整主库。
                product_rows = db.list_products_advanced(
                    search, category, quote_state, active_state, match_mode,
                    product_preview_limit(search, category, quote_state, active_state),
                )
            else:
                search = args.list_products if args.list_products is not None else args.list
                product_rows = db.list_products(search, 1000)
            output(product_output_rows(product_rows))
        elif args.list_supplier_offers is not None:
            output([dict(r) for r in db.list_supplier_offers(args.list_supplier_offers, 1000)])
        elif args.list_bindings is not None:
            output([dict(r) for r in db.list_platform_items(args.list_bindings, 1000)])
        elif args.list_prices is not None:
            output([dict(r) for r in db.list_price_snapshots(args.list_prices, 1000)])
        elif args.import_products:
            inserted, updated, skipped = db.import_products(args.import_products)
            output({"inserted":inserted,"updated":updated,"skipped":skipped})
        elif args.import_offers:
            inserted, updated, skipped, unmatched = db.import_offers(args.import_offers)
            output({"inserted":inserted,"updated":updated,"skipped":skipped,"unmatched":unmatched})
        elif args.import_supplier_offers:
            inserted, updated, skipped, unmatched = db.import_supplier_offers(args.import_supplier_offers)
            output({"inserted":inserted,"updated":updated,"skipped":skipped,"unmatched":unmatched})
        elif args.import_platform_items:
            inserted, updated, skipped, unmatched, priced = db.import_platform_items(args.import_platform_items)
            output({"inserted":inserted,"updated":updated,"skipped":skipped,"unmatched":unmatched,"priced":priced})
        elif args.set_quote_enabled:
            product_id, enabled = args.set_quote_enabled
            output({"ok":True,"enabled":db.set_quote_enabled(product_id, enabled)})
        elif args.set_quote_enabled_batch:
            product_ids, enabled = args.set_quote_enabled_batch
            count = db.set_quote_enabled_batch(product_ids.split(","), enabled)
            output({"ok":True,"count":count,"enabled":bool(int(enabled))})
        elif args.clear_quote_enabled:
            output({"ok":True,"count":db.clear_quote_enabled()})
        elif args.set_products_active:
            product_ids, enabled = args.set_products_active
            count = db.set_products_active(product_ids.split(","), enabled)
            output({"ok":True,"count":count,"enabled":bool(int(enabled))})
        elif args.delete_products:
            count, backup = db.delete_products(args.delete_products.split(","))
            output({"ok":True,"count":count,"backup":str(backup)})
        elif args.set_binding_status:
            item_id, status = args.set_binding_status
            output({"ok":True,"status":db.set_binding_status(item_id, status)})
        elif args.export:
            output({"count":db.export_products(args.export)})
        elif args.sync:
            count, backup = sync_workbook(db, config["workbook_path"], int(config.get("max_sync_rows", 1000)))
            output({"count":count,"backup":str(backup)})
        else:
            parser.error("No command")
    finally:
        db.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False))
        sys.exit(1)

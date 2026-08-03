from __future__ import annotations

# © 2026 离墨。保留所有权利。
import csv
import hashlib
import re
import sqlite3
from datetime import datetime
from pathlib import Path


PRODUCT_FIELDS = {
    "category": ("分类", "配件", "category", "type"),
    "brand": ("品牌", "brand", "manufacturer"),
    "series": ("系列", "产品系列", "series", "family"),
    "model": ("型号", "品牌及型号", "model", "name", "产品名称", "商品名称"),
    "base_model": ("基础型号", "基础产品型号", "base_model", "base model"),
    "package_type": ("包装类型", "包装", "封装类型", "package_type", "package", "packaging"),
    "warranty_note": ("保修说明", "保修信息", "warranty_note", "warranty"),
    "aliases": ("别名", "型号别名", "aliases", "alias"),
    "spec": ("规格参数", "规格", "spec", "specification", "参数"),
    "cost_price": ("手动进货价", "进货价", "成本价", "cost_price", "cost"),
    "sale_price": ("手动销售价", "销售价", "售价", "sale_price"),
    "default_url": ("商品链接", "链接", "url", "product_url"),
    "stock": ("库存", "stock"),
    "updated_at": ("更新时间", "updated_at", "update_time"),
    "compat1": ("兼容参数1", "compat1", "socket", "接口"),
    "compat2": ("兼容参数2", "compat2", "memory_type", "内存类型"),
    "mpn": ("MPN", "mpn", "厂商编号", "厂商料号"),
    "gtin": ("GTIN", "gtin", "EAN", "ean", "UPC", "upc"),
    "source": ("数据来源", "来源", "source"),
    "source_url": ("官网链接", "来源链接", "source_url", "official_url"),
    "lifecycle": ("在售状态", "生命周期", "lifecycle", "status"),
    "last_seen_at": ("最后确认时间", "最后在售时间", "last_seen_at", "last_seen"),
    "active": ("启用", "active", "是否启用"),
    "quote_enabled": ("同步报价单", "报价单启用", "quote_enabled", "sync_enabled"),
}

OFFER_FIELDS = {
    "product_id": ("产品ID", "product_id", "id"),
    "brand": ("品牌", "brand"),
    "model": ("型号", "品牌及型号", "model", "name", "商品名称"),
    "category": ("分类", "配件", "category", "type"),
    "mpn": ("MPN", "mpn", "厂商编号", "厂商料号"),
    "gtin": ("GTIN", "gtin", "EAN", "ean", "UPC", "upc"),
    "platform": ("平台", "platform"),
    "platform_sku": ("平台商品ID", "商品ID", "SKU", "sku", "item_id"),
    "seller": ("店铺", "卖家", "seller", "shop"),
    "title": ("商品标题", "标题", "title"),
    "price": ("价格", "售价", "price"),
    "price_type": ("价格类型", "price_type", "用途"),
    "url": ("链接", "商品链接", "url"),
    "stock": ("库存", "stock"),
    "status": ("绑定状态", "状态", "status"),
    "confidence": ("匹配置信度", "置信度", "confidence"),
    "preferred": ("首选", "preferred", "是否首选"),
    "collected_at": ("采集时间", "更新时间", "collected_at", "update_time"),
}

SUPPLIER_OFFER_FIELDS = {
    "product_id": ("产品ID", "product_id", "id"),
    "brand": ("品牌", "brand"),
    "model": ("型号", "品牌及型号", "model", "name", "商品名称"),
    "category": ("分类", "配件", "category", "type"),
    "mpn": ("MPN", "mpn", "厂商编号", "厂商料号"),
    "gtin": ("GTIN", "gtin", "EAN", "ean", "UPC", "upc"),
    "supplier": ("供应商", "供应商名称", "supplier", "supplier_name"),
    "supplier_sku": ("供应商商品ID", "供应商SKU", "货号", "supplier_sku", "sku"),
    "price": ("进货价", "采购价", "价格", "cost_price", "price"),
    "stock": ("库存", "stock"),
    "url": ("链接", "商品链接", "url"),
    "collected_at": ("报价时间", "更新时间", "采集时间", "collected_at", "update_time"),
    "priority": ("供应商优先级", "优先级", "priority"),
    "contact": ("联系方式", "联系人", "contact"),
    "notes": ("备注", "notes"),
    "active": ("启用", "active", "是否启用"),
}


def _clean(value) -> str:
    return "" if value is None else str(value).strip()


def _money(value):
    text = _clean(value).replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", "-", "."):
        return None
    try:
        return round(float(text), 2)
    except ValueError:
        return None


def _bool(value, default=1):
    text = _clean(value).lower()
    if not text:
        return default
    return 0 if text in {"0", "否", "禁用", "false", "no", "n"} else 1


def _int(value, default=0):
    text = _clean(value)
    if not text:
        return default
    try:
        return int(float(text))
    except ValueError:
        return default


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _read_csv(path: str | Path):
    path = Path(path)
    last_error = None
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            with path.open("r", encoding=encoding, newline="") as fh:
                yield from csv.DictReader(fh)
            return
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f"无法识别 CSV 编码：{last_error}")


def _mapped(row: dict, aliases: dict) -> dict:
    normalized = {_clean(k).lower(): v for k, v in row.items() if k is not None}
    result = {}
    for field, names in aliases.items():
        result[field] = ""
        for name in names:
            key = name.lower()
            if key in normalized:
                result[field] = normalized[key]
                break
    return result


class CatalogDB:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA busy_timeout=5000")
        self._init_schema()

    def close(self):
        self.conn.close()

    def _ensure_product_columns(self):
        existing = {row[1] for row in self.conn.execute("PRAGMA table_info(products)")}
        additions = {
            "series": "TEXT NOT NULL DEFAULT ''",
            "aliases": "TEXT NOT NULL DEFAULT ''",
            "base_model": "TEXT NOT NULL DEFAULT ''",
            "package_type": "TEXT NOT NULL DEFAULT ''",
            "warranty_note": "TEXT NOT NULL DEFAULT ''",
            "source": "TEXT NOT NULL DEFAULT ''",
            "source_url": "TEXT NOT NULL DEFAULT ''",
            "lifecycle": "TEXT NOT NULL DEFAULT '在售'",
            "last_seen_at": "TEXT NOT NULL DEFAULT ''",
            "quote_enabled": "INTEGER NOT NULL DEFAULT 1",
        }
        for name, sql_type in additions.items():
            if name not in existing:
                self.conn.execute(f"ALTER TABLE products ADD COLUMN {name} {sql_type}")
        # 旧数据库没有包装字段。CPU 保留原显示型号，并明确标记为“未区分”，
        # 避免把历史数据误判为盒装或散片；非 CPU 配件无需包装标记。
        self.conn.execute(
            "UPDATE products SET base_model=model WHERE category='CPU' AND base_model=''"
        )
        self.conn.execute(
            "UPDATE products SET package_type='未区分' WHERE category='CPU' AND package_type=''"
        )

    def _init_schema(self):
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              category TEXT NOT NULL DEFAULT '',
              brand TEXT NOT NULL DEFAULT '',
              model TEXT NOT NULL,
              base_model TEXT NOT NULL DEFAULT '',
              package_type TEXT NOT NULL DEFAULT '',
              warranty_note TEXT NOT NULL DEFAULT '',
              spec TEXT NOT NULL DEFAULT '',
              cost_price REAL,
              sale_price REAL,
              default_url TEXT NOT NULL DEFAULT '',
              stock TEXT NOT NULL DEFAULT '',
              updated_at TEXT NOT NULL DEFAULT '',
              compat1 TEXT NOT NULL DEFAULT '',
              compat2 TEXT NOT NULL DEFAULT '',
              mpn TEXT NOT NULL DEFAULT '',
              gtin TEXT NOT NULL DEFAULT '',
              active INTEGER NOT NULL DEFAULT 1,
              created_at TEXT NOT NULL,
              UNIQUE(category, brand, model)
            );
            CREATE INDEX IF NOT EXISTS idx_products_model ON products(model);
            CREATE INDEX IF NOT EXISTS idx_products_brand_model ON products(brand,model);
            CREATE INDEX IF NOT EXISTS idx_products_mpn ON products(mpn);

            CREATE TABLE IF NOT EXISTS offers (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
              platform TEXT NOT NULL DEFAULT '',
              platform_sku TEXT NOT NULL,
              seller TEXT NOT NULL DEFAULT '',
              title TEXT NOT NULL DEFAULT '',
              price REAL NOT NULL,
              url TEXT NOT NULL DEFAULT '',
              stock TEXT NOT NULL DEFAULT '',
              collected_at TEXT NOT NULL,
              UNIQUE(platform, platform_sku)
            );
            CREATE INDEX IF NOT EXISTS idx_offers_product_price ON offers(product_id, price);

            CREATE TABLE IF NOT EXISTS suppliers (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              name TEXT NOT NULL UNIQUE,
              code TEXT NOT NULL DEFAULT '',
              contact TEXT NOT NULL DEFAULT '',
              priority INTEGER NOT NULL DEFAULT 100,
              active INTEGER NOT NULL DEFAULT 1,
              created_at TEXT NOT NULL,
              updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_suppliers_priority ON suppliers(active,priority,name);

            CREATE TABLE IF NOT EXISTS supplier_offers (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
              supplier_id INTEGER NOT NULL REFERENCES suppliers(id) ON DELETE CASCADE,
              supplier_sku TEXT NOT NULL,
              price REAL NOT NULL,
              stock TEXT NOT NULL DEFAULT '',
              url TEXT NOT NULL DEFAULT '',
              notes TEXT NOT NULL DEFAULT '',
              active INTEGER NOT NULL DEFAULT 1,
              collected_at TEXT NOT NULL,
              UNIQUE(supplier_id,supplier_sku)
            );
            CREATE INDEX IF NOT EXISTS idx_supplier_offers_product ON supplier_offers(product_id,active,price);

            CREATE TABLE IF NOT EXISTS platform_items (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
              platform TEXT NOT NULL,
              platform_sku TEXT NOT NULL,
              seller TEXT NOT NULL DEFAULT '',
              title TEXT NOT NULL DEFAULT '',
              url TEXT NOT NULL DEFAULT '',
              status TEXT NOT NULL DEFAULT '已绑定',
              match_confidence INTEGER NOT NULL DEFAULT 100,
              price_role TEXT NOT NULL DEFAULT '参考价',
              preferred INTEGER NOT NULL DEFAULT 0,
              last_checked_at TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL,
              updated_at TEXT NOT NULL,
              UNIQUE(platform,platform_sku)
            );
            CREATE INDEX IF NOT EXISTS idx_platform_items_product ON platform_items(product_id,status,platform);

            CREATE TABLE IF NOT EXISTS price_snapshots (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
              source_type TEXT NOT NULL,
              source_name TEXT NOT NULL,
              source_item_id TEXT NOT NULL DEFAULT '',
              price_type TEXT NOT NULL,
              price REAL NOT NULL,
              stock TEXT NOT NULL DEFAULT '',
              url TEXT NOT NULL DEFAULT '',
              collected_at TEXT NOT NULL,
              UNIQUE(source_type,source_name,source_item_id,price_type,price,collected_at)
            );
            CREATE INDEX IF NOT EXISTS idx_price_snapshots_product ON price_snapshots(product_id,price_type,collected_at DESC);

            """
        )
        self._ensure_product_columns()
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_products_quote_enabled ON products(active,quote_enabled,category)")
        self._migrate_legacy_offers()
        self.conn.execute("PRAGMA user_version=3")
        self.conn.commit()

    def _migrate_legacy_offers(self):
        now = _now()
        self.conn.execute(
            """
            INSERT OR IGNORE INTO platform_items(
              product_id,platform,platform_sku,seller,title,url,status,match_confidence,
              price_role,preferred,last_checked_at,created_at,updated_at)
            SELECT product_id,platform,platform_sku,seller,title,url,'待复核',0,
                   '参考价',0,collected_at,?,?
            FROM offers
            """,
            (now, now),
        )
        self.conn.execute(
            """
            INSERT OR IGNORE INTO price_snapshots(
              product_id,source_type,source_name,source_item_id,price_type,price,stock,url,collected_at)
            SELECT product_id,'平台',platform,platform_sku,'参考价',price,stock,url,collected_at
            FROM offers
            """
        )

    def counts(self):
        data = self.dashboard_counts()
        return data["products"], data["prices"], data["platforms"]

    def dashboard_counts(self):
        return {
            "products": self.conn.execute("SELECT COUNT(*) FROM products").fetchone()[0],
            "active_products": self.conn.execute("SELECT COUNT(*) FROM products WHERE active=1").fetchone()[0],
            "supplier_offers": self.conn.execute("SELECT COUNT(*) FROM supplier_offers WHERE active=1").fetchone()[0],
            "bindings": self.conn.execute("SELECT COUNT(*) FROM platform_items WHERE status<>'已失效'").fetchone()[0],
            "prices": self.conn.execute("SELECT COUNT(*) FROM price_snapshots").fetchone()[0],
            "platforms": self.conn.execute("SELECT COUNT(DISTINCT platform) FROM platform_items WHERE platform<>'' AND status<>'已失效'").fetchone()[0],
            "suppliers": self.conn.execute("SELECT COUNT(*) FROM suppliers WHERE active=1").fetchone()[0],
            "quote_enabled": self.conn.execute("SELECT COUNT(*) FROM products WHERE active=1 AND quote_enabled=1").fetchone()[0],
        }

    def import_products(self, csv_path):
        inserted = updated = skipped = 0
        now = _now()
        for source in _read_csv(csv_path):
            row = _mapped(source, PRODUCT_FIELDS)
            model = _clean(row["model"])
            if not model:
                skipped += 1
                continue
            values = {
                "category": _clean(row["category"]),
                "brand": _clean(row["brand"]),
                "series": _clean(row["series"]),
                "model": model,
                "base_model": _clean(row["base_model"]),
                "package_type": _clean(row["package_type"]),
                "warranty_note": _clean(row["warranty_note"]),
                "aliases": _clean(row["aliases"]),
                "spec": _clean(row["spec"]),
                "cost_price": _money(row["cost_price"]),
                "sale_price": _money(row["sale_price"]),
                "default_url": _clean(row["default_url"]),
                "stock": _clean(row["stock"]),
                "updated_at": _clean(row["updated_at"]) or now,
                "compat1": _clean(row["compat1"]),
                "compat2": _clean(row["compat2"]),
                "mpn": _clean(row["mpn"]),
                "gtin": _clean(row["gtin"]),
                "source": _clean(row["source"]),
                "source_url": _clean(row["source_url"]),
                "lifecycle": _clean(row["lifecycle"]) or "在售",
                "last_seen_at": _clean(row["last_seen_at"]) or now,
                "active": _bool(row["active"]),
                "quote_enabled": _bool(row["quote_enabled"]),
                "created_at": now,
            }
            if values["category"] == "CPU":
                if not values["base_model"]:
                    values["base_model"] = re.sub(
                        r"\s*[（(](?:盒装|散片|Tray)[）)]\s*$", "", model, flags=re.IGNORECASE
                    ).strip() or model
                if not values["package_type"]:
                    values["package_type"] = "未区分"
            old = self.conn.execute(
                "SELECT id FROM products WHERE category=? AND brand=? AND model=?",
                (values["category"], values["brand"], values["model"]),
            ).fetchone()
            self.conn.execute(
                """
                INSERT INTO products(category,brand,series,model,base_model,package_type,warranty_note,aliases,spec,cost_price,sale_price,
                    default_url,stock,updated_at,compat1,compat2,mpn,gtin,source,source_url,
                    lifecycle,last_seen_at,active,quote_enabled,created_at)
                VALUES(:category,:brand,:series,:model,:base_model,:package_type,:warranty_note,:aliases,:spec,:cost_price,:sale_price,
                    :default_url,:stock,:updated_at,:compat1,:compat2,:mpn,:gtin,:source,:source_url,
                    :lifecycle,:last_seen_at,:active,:quote_enabled,:created_at)
                ON CONFLICT(category,brand,model) DO UPDATE SET
                    series=CASE WHEN excluded.series<>'' THEN excluded.series ELSE products.series END,
                    base_model=CASE WHEN excluded.base_model<>'' THEN excluded.base_model ELSE products.base_model END,
                    package_type=CASE WHEN excluded.package_type<>'' THEN excluded.package_type ELSE products.package_type END,
                    warranty_note=CASE WHEN excluded.warranty_note<>'' THEN excluded.warranty_note ELSE products.warranty_note END,
                    aliases=CASE WHEN excluded.aliases<>'' THEN excluded.aliases ELSE products.aliases END,
                    spec=CASE WHEN excluded.spec<>'' THEN excluded.spec ELSE products.spec END,
                    cost_price=COALESCE(excluded.cost_price,products.cost_price),
                    sale_price=COALESCE(excluded.sale_price,products.sale_price),
                    default_url=CASE WHEN excluded.default_url<>'' THEN excluded.default_url ELSE products.default_url END,
                    stock=CASE WHEN excluded.stock<>'' THEN excluded.stock ELSE products.stock END,
                    updated_at=excluded.updated_at,
                    compat1=CASE WHEN excluded.compat1<>'' THEN excluded.compat1 ELSE products.compat1 END,
                    compat2=CASE WHEN excluded.compat2<>'' THEN excluded.compat2 ELSE products.compat2 END,
                    mpn=CASE WHEN excluded.mpn<>'' THEN excluded.mpn ELSE products.mpn END,
                    gtin=CASE WHEN excluded.gtin<>'' THEN excluded.gtin ELSE products.gtin END,
                    source=CASE WHEN excluded.source<>'' THEN excluded.source ELSE products.source END,
                    source_url=CASE WHEN excluded.source_url<>'' THEN excluded.source_url ELSE products.source_url END,
                    lifecycle=CASE WHEN excluded.lifecycle<>'' THEN excluded.lifecycle ELSE products.lifecycle END,
                    last_seen_at=excluded.last_seen_at,active=excluded.active,quote_enabled=excluded.quote_enabled
                """,
                values,
            )
            updated += bool(old)
            inserted += not bool(old)
        self.conn.commit()
        return int(inserted), int(updated), skipped

    def _find_product(self, product_id, brand, model, mpn="", gtin="", category=""):
        if _clean(product_id).isdigit():
            found = self.conn.execute("SELECT id FROM products WHERE id=?", (int(product_id),)).fetchone()
            if found:
                return found[0]
        brand, model, mpn, gtin, category = map(_clean, (brand, model, mpn, gtin, category))
        if gtin:
            matches = self.conn.execute(
                "SELECT id FROM products WHERE gtin=? ORDER BY active DESC,id LIMIT 2", (gtin,)
            ).fetchall()
            if len(matches) == 1:
                return matches[0][0]
        if brand and mpn:
            matches = self.conn.execute(
                "SELECT id FROM products WHERE brand=? AND mpn=? ORDER BY active DESC,id LIMIT 2",
                (brand, mpn),
            ).fetchall()
            if len(matches) == 1:
                return matches[0][0]
        if not model:
            return None
        if category and brand:
            matches = self.conn.execute(
                "SELECT id FROM products WHERE category=? AND brand=? AND model=? ORDER BY active DESC,id LIMIT 2",
                (category, brand, model),
            ).fetchall()
            if len(matches) == 1:
                return matches[0][0]
        found = self.conn.execute(
            "SELECT id FROM products WHERE brand=? AND model=? ORDER BY active DESC,id LIMIT 1",
            (brand, model),
        ).fetchone()
        if found:
            return found[0]
        matches = self.conn.execute(
            "SELECT id FROM products WHERE model=? ORDER BY active DESC,id LIMIT 2", (model,)
        ).fetchall()
        return matches[0][0] if len(matches) == 1 else None

    def _snapshot(self, product_id, source_type, source_name, source_item_id, price_type,
                  price, stock="", url="", collected_at=""):
        price = _money(price)
        if price is None:
            return False
        self.conn.execute(
            """
            INSERT OR IGNORE INTO price_snapshots(
              product_id,source_type,source_name,source_item_id,price_type,price,stock,url,collected_at)
            VALUES(?,?,?,?,?,?,?,?,?)
            """,
            (
                int(product_id), _clean(source_type), _clean(source_name), _clean(source_item_id),
                _clean(price_type) or "参考价", price, _clean(stock), _clean(url),
                _clean(collected_at) or _now(),
            ),
        )
        return True

    def _upsert_platform_item(self, product_id, platform, platform_sku, seller="", title="",
                              url="", status="已绑定", confidence=100, price_role="参考价",
                              preferred=0, checked_at=""):
        now = _now()
        platform = _clean(platform) or "其他"
        sku = _clean(platform_sku)
        if not sku:
            raw_key = "|".join((platform, _clean(url), _clean(title), str(product_id)))
            sku = "AUTO-" + hashlib.sha1(raw_key.encode("utf-8")).hexdigest()[:16]
        old = self.conn.execute(
            "SELECT id FROM platform_items WHERE platform=? AND platform_sku=?", (platform, sku)
        ).fetchone()
        self.conn.execute(
            """
            INSERT INTO platform_items(product_id,platform,platform_sku,seller,title,url,status,
              match_confidence,price_role,preferred,last_checked_at,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(platform,platform_sku) DO UPDATE SET
              product_id=excluded.product_id,seller=CASE WHEN excluded.seller<>'' THEN excluded.seller ELSE platform_items.seller END,
              title=CASE WHEN excluded.title<>'' THEN excluded.title ELSE platform_items.title END,
              url=CASE WHEN excluded.url<>'' THEN excluded.url ELSE platform_items.url END,
              status=excluded.status,match_confidence=excluded.match_confidence,price_role=excluded.price_role,
              preferred=excluded.preferred,last_checked_at=excluded.last_checked_at,updated_at=excluded.updated_at
            """,
            (
                int(product_id), platform, sku, _clean(seller), _clean(title), _clean(url),
                _clean(status) or "已绑定", max(0, min(100, _int(confidence, 100))),
                "进货价" if _clean(price_role) in {"进货价", "采购价", "成本价"} else "参考价",
                _bool(preferred, 0), _clean(checked_at), now, now,
            ),
        )
        return sku, ("updated" if old else "inserted")

    def import_offers(self, csv_path):
        inserted = updated = skipped = unmatched = 0
        for source in _read_csv(csv_path):
            row = _mapped(source, OFFER_FIELDS)
            product_id = self._find_product(row["product_id"], row["brand"], row["model"], row["mpn"], row["gtin"], row["category"])
            price = _money(row["price"])
            if not product_id:
                unmatched += 1
                continue
            if price is None:
                skipped += 1
                continue
            collected = _clean(row["collected_at"]) or _now()
            price_type = "进货价" if _clean(row["price_type"]) in {"进货价", "采购价", "成本价"} else "参考价"
            sku, action = self._upsert_platform_item(
                product_id, row["platform"], row["platform_sku"], row["seller"], row["title"],
                row["url"], row["status"] or "已绑定", row["confidence"] or 100,
                price_type, row["preferred"], collected,
            )
            platform = _clean(row["platform"]) or "其他"
            self.conn.execute(
                """
                INSERT INTO offers(product_id,platform,platform_sku,seller,title,price,url,stock,collected_at)
                VALUES(?,?,?,?,?,?,?,?,?)
                ON CONFLICT(platform,platform_sku) DO UPDATE SET
                  product_id=excluded.product_id,seller=excluded.seller,title=excluded.title,
                  price=excluded.price,url=excluded.url,stock=excluded.stock,collected_at=excluded.collected_at
                """,
                (product_id, platform, sku, _clean(row["seller"]), _clean(row["title"]),
                 price, _clean(row["url"]), _clean(row["stock"]), collected),
            )
            self._snapshot(product_id, "平台", platform, sku, price_type, price, row["stock"], row["url"], collected)
            updated += action == "updated"
            inserted += action == "inserted"
        self.conn.commit()
        return int(inserted), int(updated), skipped, unmatched

    def import_platform_items(self, csv_path):
        inserted = updated = skipped = unmatched = priced = 0
        for source in _read_csv(csv_path):
            row = _mapped(source, OFFER_FIELDS)
            product_id = self._find_product(row["product_id"], row["brand"], row["model"], row["mpn"], row["gtin"], row["category"])
            if not product_id:
                unmatched += 1
                continue
            if not _clean(row["platform_sku"]) and not _clean(row["url"]):
                skipped += 1
                continue
            collected = _clean(row["collected_at"]) or _now()
            price_type = "进货价" if _clean(row["price_type"]) in {"进货价", "采购价", "成本价"} else "参考价"
            sku, action = self._upsert_platform_item(
                product_id, row["platform"], row["platform_sku"], row["seller"], row["title"],
                row["url"], row["status"] or "已绑定", row["confidence"] or 100,
                price_type, row["preferred"], collected,
            )
            price = _money(row["price"])
            if price is not None:
                self._snapshot(product_id, "平台", _clean(row["platform"]) or "其他", sku,
                               price_type, price, row["stock"], row["url"], collected)
                priced += 1
            updated += action == "updated"
            inserted += action == "inserted"
        self.conn.commit()
        return int(inserted), int(updated), skipped, unmatched, priced

    def _upsert_supplier(self, name, priority=100, contact=""):
        now = _now()
        name = _clean(name)
        if not name:
            raise ValueError("供应商名称不能为空。")
        self.conn.execute(
            """
            INSERT INTO suppliers(name,contact,priority,active,created_at,updated_at)
            VALUES(?,?,?,?,?,?)
            ON CONFLICT(name) DO UPDATE SET
              contact=CASE WHEN excluded.contact<>'' THEN excluded.contact ELSE suppliers.contact END,
              priority=excluded.priority,active=1,updated_at=excluded.updated_at
            """,
            (name, _clean(contact), max(1, _int(priority, 100)), 1, now, now),
        )
        return self.conn.execute("SELECT id FROM suppliers WHERE name=?", (name,)).fetchone()[0]

    def import_supplier_offers(self, csv_path):
        inserted = updated = skipped = unmatched = 0
        for source in _read_csv(csv_path):
            row = _mapped(source, SUPPLIER_OFFER_FIELDS)
            product_id = self._find_product(row["product_id"], row["brand"], row["model"], row["mpn"], row["gtin"], row["category"])
            if not product_id:
                unmatched += 1
                continue
            price = _money(row["price"])
            if price is None or not _clean(row["supplier"]):
                skipped += 1
                continue
            supplier_id = self._upsert_supplier(row["supplier"], row["priority"] or 100, row["contact"])
            sku = _clean(row["supplier_sku"])
            if not sku:
                raw_key = f"{supplier_id}|{product_id}|{_clean(row['url'])}"
                sku = "AUTO-" + hashlib.sha1(raw_key.encode("utf-8")).hexdigest()[:16]
            old = self.conn.execute(
                "SELECT id FROM supplier_offers WHERE supplier_id=? AND supplier_sku=?", (supplier_id, sku)
            ).fetchone()
            collected = _clean(row["collected_at"]) or _now()
            self.conn.execute(
                """
                INSERT INTO supplier_offers(product_id,supplier_id,supplier_sku,price,stock,url,notes,active,collected_at)
                VALUES(?,?,?,?,?,?,?,?,?)
                ON CONFLICT(supplier_id,supplier_sku) DO UPDATE SET
                  product_id=excluded.product_id,price=excluded.price,stock=excluded.stock,
                  url=excluded.url,notes=excluded.notes,active=excluded.active,collected_at=excluded.collected_at
                """,
                (product_id, supplier_id, sku, price, _clean(row["stock"]), _clean(row["url"]),
                 _clean(row["notes"]), _bool(row["active"]), collected),
            )
            self._snapshot(product_id, "供应商", _clean(row["supplier"]), sku, "进货价",
                           price, row["stock"], row["url"], collected)
            updated += bool(old)
            inserted += not bool(old)
        self.conn.commit()
        return int(inserted), int(updated), skipped, unmatched

    def upsert_offer(self, product_id, platform, platform_sku, seller, title, price, url,
                     stock="", collected_at="", price_type="参考价",
                     binding_status="待复核", confidence=0):
        price = _money(price)
        if price is None:
            raise ValueError("报价价格为空或格式不正确。")
        collected_at = _clean(collected_at) or _now()
        price_type = "进货价" if _clean(price_type) in {"进货价", "采购价", "成本价"} else "参考价"
        sku, action = self._upsert_platform_item(
            product_id, platform, platform_sku, seller, title, url, binding_status, confidence,
            price_type, 0, collected_at,
        )
        self.conn.execute(
            """
            INSERT INTO offers(product_id,platform,platform_sku,seller,title,price,url,stock,collected_at)
            VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(platform,platform_sku) DO UPDATE SET
              product_id=excluded.product_id,seller=excluded.seller,title=excluded.title,
              price=excluded.price,url=excluded.url,stock=excluded.stock,collected_at=excluded.collected_at
            """,
            (int(product_id), _clean(platform) or "其他", sku, _clean(seller), _clean(title),
             price, _clean(url), _clean(stock), collected_at),
        )
        self._snapshot(product_id, "平台", _clean(platform) or "其他", sku, price_type,
                       price, stock, url, collected_at)
        self.conn.commit()
        return action

    def _cost_for_product(self, product):
        if product["cost_price"] is not None:
            return product["cost_price"], "手动进货价", product["stock"], product["default_url"], product["updated_at"]
        supplier = self.conn.execute(
            """
            SELECT so.price,so.stock,so.url,so.collected_at,s.name
            FROM supplier_offers so JOIN suppliers s ON s.id=so.supplier_id
            WHERE so.product_id=? AND so.active=1 AND s.active=1
            ORDER BY s.priority ASC,so.price ASC,so.collected_at DESC LIMIT 1
            """,
            (product["id"],),
        ).fetchone()
        if supplier:
            return supplier["price"], "供应商：" + supplier["name"], supplier["stock"], supplier["url"], supplier["collected_at"]
        platform_cost = self.conn.execute(
            """
            SELECT ps.price,ps.stock,ps.url,ps.collected_at,pi.platform
            FROM platform_items pi JOIN price_snapshots ps
              ON ps.product_id=pi.product_id AND ps.source_type='平台'
             AND ps.source_name=pi.platform AND ps.source_item_id=pi.platform_sku
            WHERE pi.product_id=? AND pi.status='已绑定' AND pi.price_role='进货价' AND ps.price_type='进货价'
            ORDER BY pi.preferred DESC,ps.collected_at DESC,ps.price ASC LIMIT 1
            """,
            (product["id"],),
        ).fetchone()
        if platform_cost:
            return platform_cost["price"], platform_cost["platform"] + "进货价", platform_cost["stock"], platform_cost["url"], platform_cost["collected_at"]
        return None, "待录入", product["stock"], product["default_url"], product["updated_at"]

    def _reference_for_product(self, product_id):
        return self.conn.execute(
            """
            SELECT ps.price,ps.url,ps.stock,ps.collected_at,ps.source_name
            FROM price_snapshots ps JOIN platform_items pi
              ON pi.product_id=ps.product_id AND pi.platform=ps.source_name
             AND pi.platform_sku=ps.source_item_id
            WHERE ps.product_id=? AND ps.price_type='参考价' AND pi.status='已绑定'
            ORDER BY ps.collected_at DESC,ps.price ASC LIMIT 1
            """,
            (product_id,),
        ).fetchone()

    def workbook_rows(self, limit=None):
        sql = """
            SELECT * FROM products WHERE active=1 AND quote_enabled=1
            ORDER BY CASE category
              WHEN 'CPU' THEN 1 WHEN '主板' THEN 2 WHEN '显卡' THEN 3 WHEN '内存' THEN 4
              WHEN '固态硬盘' THEN 5 WHEN '机械硬盘' THEN 6 WHEN 'CPU散热器' THEN 7
              WHEN '电源' THEN 8 WHEN '机箱' THEN 9 WHEN '机箱风扇' THEN 10
              WHEN '显示器' THEN 11 WHEN '键盘鼠标' THEN 12 WHEN '系统与软件' THEN 13
              WHEN '其他配件' THEN 14 WHEN '装机服务' THEN 15 ELSE 99 END,
              brand,model
            """
        params = ()
        if limit is not None and int(limit) > 0:
            sql += " LIMIT ?"
            params = (int(limit),)
        products = self.conn.execute(sql, params).fetchall()
        resolved = {row["id"]: row for row in self._product_result_rows(products)}
        preferred_urls = {}
        ids = [int(product["id"]) for product in products]
        for start in range(0, len(ids), 800):
            chunk = ids[start:start + 800]
            if not chunk:
                continue
            placeholders = ",".join("?" for _ in chunk)
            for row in self.conn.execute(
                f"""
                SELECT product_id,url FROM platform_items
                WHERE product_id IN ({placeholders}) AND status='已绑定' AND url<>''
                ORDER BY product_id,preferred DESC,updated_at DESC
                """, tuple(chunk)
            ):
                preferred_urls.setdefault(row["product_id"], row["url"])
        result = []
        for product in products:
            row = resolved[product["id"]]
            url = preferred_urls.get(product["id"], "") or row["reference_url"] or row["cost_url"] or product["default_url"]
            stock = row["cost_stock"] or row["reference_stock"] or product["stock"]
            updated = row["cost_updated"] or row["reference_updated"] or product["updated_at"]
            display_name = " ".join(part for part in (product["brand"], product["model"]) if _clean(part))
            result.append([
                display_name, product["category"], product["brand"], product["spec"],
                row["cost"], product["sale_price"], url, stock, updated,
                product["compat1"], product["compat2"],
            ])
        return result

    def _product_result_rows(self, products):
        products = list(products)
        ids = [int(product["id"]) for product in products]
        supplier_costs = {}
        platform_costs = {}
        references = {}
        for start in range(0, len(ids), 800):
            chunk = ids[start:start + 800]
            if not chunk:
                continue
            placeholders = ",".join("?" for _ in chunk)
            for row in self.conn.execute(
                f"""
                SELECT so.product_id,so.price,so.stock,so.url,so.collected_at,s.name
                FROM supplier_offers so JOIN suppliers s ON s.id=so.supplier_id
                WHERE so.product_id IN ({placeholders}) AND so.active=1 AND s.active=1
                ORDER BY so.product_id,s.priority,so.price,so.collected_at DESC
                """, tuple(chunk)
            ):
                supplier_costs.setdefault(row["product_id"], row)
            for row in self.conn.execute(
                f"""
                SELECT pi.product_id,ps.price,ps.stock,ps.url,ps.collected_at,pi.platform
                FROM platform_items pi JOIN price_snapshots ps
                  ON ps.product_id=pi.product_id AND ps.source_type='平台'
                 AND ps.source_name=pi.platform AND ps.source_item_id=pi.platform_sku
                WHERE pi.product_id IN ({placeholders}) AND pi.status='已绑定'
                  AND pi.price_role='进货价' AND ps.price_type='进货价'
                ORDER BY pi.product_id,pi.preferred DESC,ps.collected_at DESC,ps.price ASC
                """, tuple(chunk)
            ):
                platform_costs.setdefault(row["product_id"], row)
            for row in self.conn.execute(
                f"""
                SELECT ps.product_id,ps.price,ps.url,ps.stock,ps.collected_at,ps.source_name
                FROM price_snapshots ps JOIN platform_items pi
                  ON pi.product_id=ps.product_id AND pi.platform=ps.source_name
                 AND pi.platform_sku=ps.source_item_id
                WHERE ps.product_id IN ({placeholders}) AND ps.price_type='参考价' AND pi.status='已绑定'
                ORDER BY ps.product_id,ps.collected_at DESC,ps.price ASC
                """, tuple(chunk)
            ):
                references.setdefault(row["product_id"], row)

        result = []
        for product in products:
            product_id = product["id"]
            if product["cost_price"] is not None:
                cost, source, stock, cost_url, updated = product["cost_price"], "手动进货价", product["stock"], product["default_url"], product["updated_at"]
            elif product_id in supplier_costs:
                row = supplier_costs[product_id]
                cost, source, stock, cost_url, updated = row["price"], "供应商：" + row["name"], row["stock"], row["url"], row["collected_at"]
            elif product_id in platform_costs:
                row = platform_costs[product_id]
                cost, source, stock, cost_url, updated = row["price"], row["platform"] + "进货价", row["stock"], row["url"], row["collected_at"]
            else:
                cost, source, stock, cost_url, updated = None, "待录入", product["stock"], product["default_url"], product["updated_at"]
            reference = references.get(product_id)
            result.append({
                "id": product_id, "category": product["category"], "brand": product["brand"],
                "series": product["series"], "model": product["model"],
                "base_model": product["base_model"], "package_type": product["package_type"],
                "warranty_note": product["warranty_note"], "cost": cost,
                "cost_source": source, "reference_price": reference["price"] if reference else None,
                "cost_stock": stock, "cost_url": cost_url, "cost_updated": updated,
                "reference_stock": reference["stock"] if reference else "",
                "reference_url": reference["url"] if reference else "",
                "reference_updated": reference["collected_at"] if reference else "",
                "supplier_count": product["supplier_count"] if "supplier_count" in product.keys() else 0,
                "binding_count": product["binding_count"] if "binding_count" in product.keys() else 0,
                "stock": stock or product["stock"], "lifecycle": product["lifecycle"] or "在售",
                "active": product["active"], "quote_enabled": product["quote_enabled"],
                "updated_at": updated or product["updated_at"],
            })
        return result

    def list_products(self, search="", limit=1000):
        return self.list_products_advanced(search=search, limit=limit)

    def list_products_advanced(self, search="", category="", quote_state="全部",
                               active_state="启用", match_mode="智能搜索", limit=5000):
        """搜索型号主库。

        普通搜索按空格拆词并要求每个词都命中任意字段；分号或换行分隔的多组
        关键词按 OR 处理，适合一次粘贴多个型号。精确模式匹配型号、品牌+型号、
        MPN 或 GTIN，避免相近型号被误选。
        """
        clauses = []
        params = []
        category = _clean(category)
        quote_state = _clean(quote_state) or "全部"
        active_state = _clean(active_state) or "启用"
        match_mode = _clean(match_mode) or "智能搜索"

        if category and category != "全部分类":
            clauses.append("p.category=?")
            params.append(category)
        if quote_state == "已加入":
            clauses.append("p.quote_enabled=1")
        elif quote_state == "未加入":
            clauses.append("p.quote_enabled=0")
        if active_state == "启用":
            clauses.append("p.active=1")
        elif active_state == "已停用":
            clauses.append("p.active=0")

        query_groups = [part.strip() for part in re.split(r"[\r\n;；]+", _clean(search)) if part.strip()]
        searchable = (
            "p.category", "p.brand", "p.series", "p.model", "p.base_model",
            "p.package_type", "p.warranty_note", "p.aliases", "p.spec", "p.mpn", "p.gtin",
        )
        if query_groups:
            group_sql = []
            for group in query_groups:
                if match_mode == "精确型号":
                    normalized = " ".join(group.split())
                    group_sql.append("(p.model=? OR TRIM(p.brand || ' ' || p.model)=? OR p.mpn=? OR p.gtin=?)")
                    params.extend([normalized, normalized, normalized, normalized])
                    continue
                token_sql = []
                for token in [value for value in re.split(r"\s+", group) if value]:
                    token_sql.append("(" + " OR ".join(field + " LIKE ?" for field in searchable) + ")")
                    params.extend([f"%{token}%"] * len(searchable))
                if token_sql:
                    group_sql.append("(" + " AND ".join(token_sql) + ")")
            if group_sql:
                clauses.append("(" + " OR ".join(group_sql) + ")")

        where = " AND ".join(clauses) if clauses else "1=1"
        safe_limit = max(1, int(limit))
        products = self.conn.execute(
            f"""
            SELECT p.*,
              (SELECT COUNT(*) FROM supplier_offers so WHERE so.product_id=p.id AND so.active=1) supplier_count,
              (SELECT COUNT(*) FROM platform_items pi WHERE pi.product_id=p.id AND pi.status<>'已失效') binding_count
            FROM products p
            WHERE {where}
            ORDER BY p.active DESC,
              CASE p.category
                WHEN 'CPU' THEN 1 WHEN '主板' THEN 2 WHEN '显卡' THEN 3 WHEN '内存' THEN 4
                WHEN '固态硬盘' THEN 5 WHEN '机械硬盘' THEN 6 WHEN 'CPU散热器' THEN 7
                WHEN '电源' THEN 8 WHEN '机箱' THEN 9 WHEN '机箱风扇' THEN 10
                WHEN '显示器' THEN 11 WHEN '键盘鼠标' THEN 12 WHEN '系统与软件' THEN 13
                WHEN '其他配件' THEN 14 WHEN '装机服务' THEN 15 ELSE 99 END,
              p.brand,p.model LIMIT ?
            """,
            tuple(params + [safe_limit]),
        ).fetchall()
        return self._product_result_rows(products)

    @staticmethod
    def _normalize_ids(product_ids):
        result = []
        for value in product_ids:
            try:
                product_id = int(value)
            except (TypeError, ValueError):
                continue
            if product_id > 0 and product_id not in result:
                result.append(product_id)
        if not result:
            raise ValueError("没有有效的型号 ID。")
        return result

    def set_quote_enabled_batch(self, product_ids, enabled):
        ids = self._normalize_ids(product_ids)
        placeholders = ",".join("?" for _ in ids)
        changed = self.conn.execute(
            f"UPDATE products SET quote_enabled=?,updated_at=? WHERE id IN ({placeholders})",
            tuple([1 if _bool(enabled) else 0, _now()] + ids),
        ).rowcount
        self.conn.commit()
        return changed

    def clear_quote_enabled(self):
        changed = self.conn.execute(
            "UPDATE products SET quote_enabled=0,updated_at=? WHERE active=1 AND quote_enabled=1",
            (_now(),),
        ).rowcount
        self.conn.commit()
        return changed

    def set_products_active(self, product_ids, enabled):
        ids = self._normalize_ids(product_ids)
        placeholders = ",".join("?" for _ in ids)
        active = 1 if _bool(enabled) else 0
        changed = self.conn.execute(
            f"UPDATE products SET active=?,quote_enabled=CASE WHEN ?=0 THEN 0 ELSE quote_enabled END,updated_at=? WHERE id IN ({placeholders})",
            tuple([active, active, _now()] + ids),
        ).rowcount
        self.conn.commit()
        return changed

    def delete_products(self, product_ids):
        ids = self._normalize_ids(product_ids)
        placeholders = ",".join("?" for _ in ids)
        existing = self.conn.execute(
            f"SELECT COUNT(*) FROM products WHERE id IN ({placeholders})", tuple(ids)
        ).fetchone()[0]
        if not existing:
            raise ValueError("找不到要删除的型号。")

        backup_dir = self.path.parent / "备份"
        backup_dir.mkdir(exist_ok=True)
        backup = backup_dir / f"删除型号前_{datetime.now():%Y%m%d_%H%M%S}.db"
        backup_conn = sqlite3.connect(backup)
        try:
            self.conn.backup(backup_conn)
        finally:
            backup_conn.close()
        changed = self.conn.execute(
            f"DELETE FROM products WHERE id IN ({placeholders})", tuple(ids)
        ).rowcount
        self.conn.commit()
        return changed, backup

    def list_supplier_offers(self, search="", limit=1000):
        like = f"%{_clean(search)}%"
        return self.conn.execute(
            """
            SELECT so.id,p.id product_id,p.brand,p.model,s.name supplier,so.supplier_sku,
                   so.price,so.stock,so.url,so.collected_at,s.priority,so.active
            FROM supplier_offers so JOIN suppliers s ON s.id=so.supplier_id
              JOIN products p ON p.id=so.product_id
            WHERE p.brand LIKE ? OR p.model LIKE ? OR s.name LIKE ? OR so.supplier_sku LIKE ?
            ORDER BY so.collected_at DESC,s.priority,so.price LIMIT ?
            """,
            (like, like, like, like, int(limit)),
        ).fetchall()

    def list_platform_items(self, search="", limit=1000):
        like = f"%{_clean(search)}%"
        return self.conn.execute(
            """
            SELECT i.id,p.id product_id,p.brand,p.model,i.platform,i.platform_sku,i.seller,
                   i.price_role,i.status,i.match_confidence,i.preferred,i.url,i.last_checked_at
            FROM platform_items i JOIN products p ON p.id=i.product_id
            WHERE p.brand LIKE ? OR p.model LIKE ? OR i.platform LIKE ? OR i.platform_sku LIKE ?
               OR i.seller LIKE ? OR i.title LIKE ?
            ORDER BY i.status,i.platform,p.brand,p.model LIMIT ?
            """,
            (like, like, like, like, like, like, int(limit)),
        ).fetchall()

    def list_price_snapshots(self, search="", limit=1000):
        like = f"%{_clean(search)}%"
        return self.conn.execute(
            """
            SELECT ps.id,p.id product_id,p.brand,p.model,ps.source_type,ps.source_name,
                   ps.source_item_id,ps.price_type,ps.price,ps.stock,ps.collected_at,ps.url
            FROM price_snapshots ps JOIN products p ON p.id=ps.product_id
            WHERE p.brand LIKE ? OR p.model LIKE ? OR ps.source_name LIKE ? OR ps.source_item_id LIKE ?
            ORDER BY ps.collected_at DESC,ps.id DESC LIMIT ?
            """,
            (like, like, like, like, int(limit)),
        ).fetchall()

    def set_quote_enabled(self, product_id, enabled):
        changed = self.conn.execute(
            "UPDATE products SET quote_enabled=?,updated_at=? WHERE id=?",
            (1 if _bool(enabled) else 0, _now(), int(product_id)),
        ).rowcount
        self.conn.commit()
        if not changed:
            raise ValueError("找不到指定型号。")
        return bool(_bool(enabled))

    def set_binding_status(self, item_id, status):
        status = _clean(status)
        if status not in {"已绑定", "待复核", "已失效"}:
            raise ValueError("绑定状态只能是：已绑定、待复核或已失效。")
        changed = self.conn.execute(
            "UPDATE platform_items SET status=?,updated_at=? WHERE id=?",
            (status, _now(), int(item_id)),
        ).rowcount
        self.conn.commit()
        if not changed:
            raise ValueError("找不到指定渠道绑定。")
        return status

    def export_products(self, csv_path):
        rows = self.conn.execute("SELECT * FROM products ORDER BY category,brand,model").fetchall()
        headers = ["产品ID","分类","品牌","系列","型号","基础型号","包装类型","保修说明","型号别名","规格参数","手动进货价","手动销售价",
                   "商品链接","库存","更新时间","兼容参数1","兼容参数2","MPN","GTIN","数据来源",
                   "官网链接","在售状态","最后确认时间","启用","同步报价单"]
        with Path(csv_path).open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh)
            writer.writerow(headers)
            for r in rows:
                writer.writerow([r["id"],r["category"],r["brand"],r["series"],r["model"],r["base_model"],r["package_type"],r["warranty_note"],r["aliases"],
                    r["spec"],r["cost_price"],r["sale_price"],r["default_url"],r["stock"],r["updated_at"],
                    r["compat1"],r["compat2"],r["mpn"],r["gtin"],r["source"],r["source_url"],
                    r["lifecycle"],r["last_seen_at"],r["active"],r["quote_enabled"]])
        return len(rows)

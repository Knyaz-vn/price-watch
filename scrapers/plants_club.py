"""Збирач каталогу plants-club.ua (Клуб Рослин).

У магазину немає фіду, але кожна сторінка категорії містить список товарів у JSON:
ціна в копійках, артикул, виробник, залишок на складі (stock) і ознака актуальності.
Тому сторінки окремих товарів не відкриваються, достатньо пройти категорії.

Порядок роботи:
1. Дерево каталогу береться з меню, вбудованого в будь-яку сторінку сайту. Обходяться лише
   кінцеві категорії, тож кожен товар читається один раз. Якщо меню не прочиталось,
   список береться з карти сайту без сторінок фільтрів.
2. На першій категорії перевіряється, який спосіб читання працює з цього сервера:
   JSON API з різними заголовками або звичайна HTML-сторінка. Далі використовується
   той, що спрацював, а інший лишається запасним для окремих збоїв.
3. Якщо жоден спосіб не працює, обхід зупиняється одразу з поясненням причини.

Запити йдуть по одному, наступний не раніше ніж через delay секунд після початку попереднього.
"""

from __future__ import annotations

import json
import os
import re
import time

import requests
from lxml import etree

SITE = "https://plants-club.ua"
API = "https://resources.plants-club.ua/api/seo"
SITEMAP = f"{SITE}/sitemap.xml?content=category&locale=uk"
BROWSER_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/129.0 Safari/537.36")
NEXT_DATA = re.compile(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S)

# способи читання списку товарів; порядок задає, що пробується першим
API_VARIANTS = {
    "api": {"Accept": "application/json, text/plain, */*", "Origin": SITE, "Referer": SITE + "/"},
    "api-plain": {"Accept": "*/*"},
    "api-html": {"Accept": "text/html,application/xhtml+xml,*/*;q=0.8"},
}


class ScrapeError(RuntimeError):
    pass


class Client:
    def __init__(self, delay: float, log):
        self.delay = delay
        self.log = log
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": BROWSER_UA, "Accept-Language": "uk-UA,uk;q=0.9,en;q=0.6"})
        proxy = os.environ.get("FEED_PROXY", "").strip()
        if proxy:
            self.session.proxies = {"http": proxy, "https": proxy}
        self.requests = 0
        self.last_error = ""
        self._last = 0.0

    def get(self, url: str, headers: dict | None = None, tries: int = 2) -> tuple[requests.Response | None, str]:
        """Повертає (відповідь, опис помилки). Повторює лише мережеві збої і 502–504."""
        err = ""
        for attempt in range(1, tries + 1):
            wait = self.delay - (time.monotonic() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            self.requests += 1
            try:
                r = self.session.get(url, timeout=(15, 60), headers=headers or {})
            except requests.RequestException as exc:
                err = type(exc).__name__
            else:
                if r.status_code == 200:
                    return r, ""
                err = f"HTTP {r.status_code}"
                if r.status_code not in (429, 502, 503, 504):
                    self.last_error = err
                    return None, err
                if r.status_code == 429:
                    time.sleep(30)
            if attempt < tries:
                time.sleep(3 * attempt)
        self.last_error = err
        return None, err


# ---------------------------------------------------------------- список категорій

def next_data(text: str) -> dict | None:
    m = NEXT_DATA.search(text or "")
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return None


def tree_leaves(catalogue: list) -> list[tuple[str, list[str]]]:
    """Кінцеві категорії дерева з ланцюжком батьківських slug."""
    out = []

    def walk(node, path):
        if not isinstance(node, dict) or not node.get("slug"):
            return
        kids = node.get("subCategories") or node.get("children") or []
        here = path + [node["slug"]]
        if kids:
            for k in kids:
                walk(k, here)
        else:
            out.append((node["slug"], path))

    for root in catalogue or []:
        walk(root, [])
    seen, uniq = set(), []
    for slug, path in out:
        if slug not in seen:
            seen.add(slug)
            uniq.append((slug, path))
    return uniq


def categories_from_menu(client: Client) -> list[tuple[str, list[str]]]:
    r, _ = client.get(SITE + "/", {"Accept": "text/html,application/xhtml+xml,*/*;q=0.8"})
    data = next_data(r.text) if r is not None else None
    catalogue = (((data or {}).get("props") or {}).get("initialProps") or {}).get("data", {}).get("catalogue")
    return tree_leaves(catalogue) if isinstance(catalogue, list) else []


def categories_from_sitemap(client: Client) -> list[tuple[str, list[str]]]:
    r, err = client.get(SITEMAP, {"Accept": "application/xml,text/xml,*/*"})
    if r is None:
        return []
    try:
        root = etree.fromstring(r.content, etree.XMLParser(recover=True, resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError:
        return []
    out = []
    for loc in root.iter("{*}loc"):
        path = (loc.text or "").strip().replace(SITE, "").strip("/")
        # сторінки фільтрів дублюють товари звичайних категорій
        if path and "/" not in path and "-search-" not in path and path not in [s for s, _ in out]:
            out.append((path, []))
    return out


# ---------------------------------------------------------------- читання сторінки категорії

def unwrap(payload) -> dict:
    """API і HTML дають однаковий об'єкт, але на різній глибині."""
    if not isinstance(payload, dict):
        return {}
    if isinstance(payload.get("products"), dict):
        return payload
    inner = payload.get("data")
    if isinstance(inner, dict) and isinstance(inner.get("products"), dict):
        return inner
    return payload


def read_page(client: Client, slug: str, page: int, method: str) -> tuple[dict | None, str]:
    if method == "html":
        url = f"{SITE}/{slug}" + (f"/{page}" if page > 1 else "")
        r, err = client.get(url, {"Accept": "text/html,application/xhtml+xml,*/*;q=0.8"})
        if r is None:
            return None, err
        data = next_data(r.text)
        if data is None:
            return None, "на сторінці немає даних каталогу"
        return ((data.get("props") or {}).get("pageProps") or {}).get("data") or {}, ""
    r, err = client.get(f"{API}/{slug}?page={page}", API_VARIANTS[method])
    if r is None:
        return None, err
    try:
        return unwrap(r.json()), ""
    except ValueError:
        return None, "відповідь не JSON"


def has_products(payload: dict | None) -> bool:
    block = (payload or {}).get("products")
    return isinstance(block, dict) and isinstance(block.get("data"), list)


def choose_method(client: Client, slugs: list[str], log) -> tuple[str, dict, str]:
    """Пробує способи на перших категоріях; повертає (спосіб, перша сторінка, slug)."""
    errors = {}
    for slug in slugs[:3]:
        for method in list(API_VARIANTS) + ["html"]:
            payload, err = read_page(client, slug, 1, method)
            if has_products(payload):
                log(f"  Клуб Рослин: читаю способом «{method}»")
                return method, payload, slug
            errors[method] = err or "без списку товарів"
    detail = "; ".join(f"{m}: {e}" for m, e in errors.items())
    raise ScrapeError(f"сайт не віддає списки товарів жодним способом ({detail})")


# ---------------------------------------------------------------- товари

def to_row(p: dict) -> dict | None:
    pid = p.get("id")
    name = (p.get("name") or "").strip()
    price = p.get("price")
    if pid is None or not name or not isinstance(price, (int, float)) or price <= 0:
        return None
    stock = p.get("stock") if isinstance(p.get("stock"), (int, float)) else None
    minimum = p.get("minimum") if isinstance(p.get("minimum"), (int, float)) else 1
    actual = p.get("actuality") is not False
    available = actual and (stock is None or stock >= max(minimum, 1))
    maker = p.get("manufacturer") if isinstance(p.get("manufacturer"), dict) else {}
    cat = p.get("category") if isinstance(p.get("category"), dict) else {}
    return {
        "id": str(pid),
        "sku": str(p.get("sku") or "").strip(),
        "name": name,
        "price": round(price / 100, 2),  # ціни на сайті зберігаються в копійках
        "available": available,
        "stock": stock,
        "brand": (maker.get("name") or "").strip(),
        "category": (cat.get("name") or "").strip(),
        "url": f"{SITE}/{p['slug']}" if p.get("slug") else "",
    }


def scrape(options: dict, log=print) -> list[dict]:
    """Повертає список товарів. options з config.json:
    delay            мінімальний проміжок між запитами, секунд (типово 0.7)
    categories       обійти лише ці категорії (slug з адреси сторінки)
    skip_categories  пропустити розділи разом із вкладеними, наприклад "roslyny"
    max_minutes      зупинитися, якщо обхід триває довше (типово 50)
    """
    client = Client(float(options.get("delay", 0.7)), log)
    skip = {str(s).strip("/ ") for s in options.get("skip_categories") or []}
    max_minutes = float(options.get("max_minutes", 50))
    deadline = time.monotonic() + 60 * max_minutes

    if options.get("categories"):
        cats = [(str(s).strip("/ "), []) for s in options["categories"]]
    else:
        cats = categories_from_menu(client)
        source = "меню сайту"
        if not cats:
            cats = categories_from_sitemap(client)
            source = "карти сайту"
        if not cats:
            why = f", сайт відповів {client.last_error}" if client.last_error else ""
            raise ScrapeError(f"не вдалося отримати список категорій ні з меню, ні з карти сайту{why}")
        log(f"  Клуб Рослин: кінцевих категорій з {source}: {len(cats)}")
    before = len(cats)
    cats = [(s, path) for s, path in cats if s not in skip and not (set(path) & skip)]
    if before != len(cats):
        log(f"  Клуб Рослин: пропущено за skip_categories: {before - len(cats)}")
    if not cats:
        raise ScrapeError("після skip_categories не лишилось жодної категорії")

    method, first, first_slug = choose_method(client, [s for s, _ in cats], log)
    backup = "html" if method != "html" else "api"

    products: dict[str, dict] = {}
    done, failed, empty = 0, [], 0
    for n, (slug, _) in enumerate(cats, 1):
        if time.monotonic() > deadline:
            raise ScrapeError(f"обхід не вклався в {max_minutes:g} хв, оброблено {done} з {len(cats)} категорій; "
                              f"збільште max_minutes або пропустіть зайві розділи через skip_categories")
        page, last = 1, 1
        while True:
            if slug == first_slug and page == 1:
                payload = first
            else:
                payload, _ = read_page(client, slug, page, method)
                if payload is None:  # запасний спосіб лише при збої, а не коли товарів просто немає
                    payload, _ = read_page(client, slug, page, backup)
            if not has_products(payload):
                if page == 1 and payload is not None:
                    empty += 1  # сторінка відкрилась, але власного списку товарів у категорії немає
                else:
                    failed.append(slug if page == 1 else f"{slug} стор. {page}")
                break
            block = payload["products"]
            last = int((block.get("meta") or {}).get("last_page") or 1)
            for p in block.get("data") or []:
                row = to_row(p)
                if row and row["id"] not in products:
                    products[row["id"]] = row
            if page >= last:
                break
            page += 1
        done += 1
        if done == 15 and len(failed) >= 15:
            raise ScrapeError("перші 15 категорій не прочитались, сайт, схоже, обмежує доступ")
        if n % 100 == 0:
            log(f"    оброблено {n} з {len(cats)} категорій, товарів {len(products)}")

    log(f"  Клуб Рослин: товарів {len(products)}, категорій {done}, без товарів {empty}, "
        f"не прочитались {len(failed)}, запитів {client.requests}")
    if failed:
        log(f"  Клуб Рослин: не прочитались: {', '.join(failed[:10])}" + (" ..." if len(failed) > 10 else ""))
    if not products:
        raise ScrapeError("не зібрано жодного товару")
    if len(failed) > max(5, 0.2 * len(cats)):
        raise ScrapeError(f"не прочиталась значна частина категорій ({len(failed)} з {len(cats)}), дані неповні")
    return list(products.values())

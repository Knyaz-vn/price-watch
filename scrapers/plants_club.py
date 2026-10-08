"""Збирач каталогу plants-club.ua (Клуб Рослин).

У магазину немає фіду, але кожна сторінка категорії віддає список товарів у JSON:
ціна в копійках, артикул, виробник, залишок на складі (stock) і ознака актуальності.
Тому заходити на сторінку кожного товару не потрібно, достатньо пройти категорії.

Порядок роботи:
1. Список категорій береться з карти сайту (sitemap.xml?content=category&locale=uk).
2. Для кожної категорії читається JSON із resources.plants-club.ua/api/seo/<slug>?page=N.
   Якщо API не відповідає, ті самі дані беруться з HTML сторінки (__NEXT_DATA__).
3. Товар з'являється в кількох категоріях (у батьківській і в дочірній), тому дублікати
   відкидаються за внутрішнім id магазину.

Запити йдуть по одному, з паузою між ними, щоб не навантажувати сайт.
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
UA = "Mozilla/5.0 (compatible; PriceWatch/1.0)"
NEXT_DATA = re.compile(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S)


class ScrapeError(RuntimeError):
    pass


class Client:
    def __init__(self, delay: float, log):
        self.delay = delay
        self.log = log
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": UA, "Accept-Language": "uk,en;q=0.8"})
        proxy = os.environ.get("FEED_PROXY", "").strip()
        if proxy:
            self.session.proxies = {"http": proxy, "https": proxy}
        self.requests = 0
        self._last = 0.0

    def get(self, url: str, accept: str = "*/*") -> requests.Response | None:
        """GET з паузою між запитами і повторами при збоях. 404 повертає None."""
        for attempt in range(1, 4):
            wait = self.delay - (time.monotonic() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            self.requests += 1
            try:
                r = self.session.get(url, timeout=(15, 60), headers={"Accept": accept})
            except requests.RequestException as exc:
                err = type(exc).__name__
            else:
                if r.status_code == 200:
                    return r
                if r.status_code == 404:
                    return None
                err = f"HTTP {r.status_code}"
                if r.status_code in (401, 403) and attempt == 1 and self.requests <= 2:
                    # перший же запит заборонено: сайт не пускає цей сервер, повтори не допоможуть
                    raise ScrapeError(f"сайт відповів {err}, ймовірно блокує запити з цього сервера")
            self.log(f"    {url[:90]}: {err}, спроба {attempt}")
            time.sleep(5 * attempt)
        return None


def category_slugs(client: Client) -> list[str]:
    r = client.get(SITEMAP, "application/xml,text/xml,*/*")
    if r is None:
        raise ScrapeError("не вдалося отримати карту категорій сайту")
    try:
        root = etree.fromstring(r.content, etree.XMLParser(recover=True, resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError as exc:
        raise ScrapeError(f"карта категорій пошкоджена: {exc}") from exc
    slugs = []
    for loc in root.iter("{*}loc"):
        path = (loc.text or "").strip().replace(SITE, "").strip("/")
        if path and "/" not in path and path not in slugs:
            slugs.append(path)
    if not slugs:
        raise ScrapeError("у карті сайту не знайдено жодної категорії")
    return slugs


def unwrap(payload: dict) -> dict:
    """API і HTML дають однаковий об'єкт, але на різній глибині."""
    if not isinstance(payload, dict):
        return {}
    if isinstance(payload.get("products"), dict):
        return payload
    inner = payload.get("data")
    if isinstance(inner, dict) and isinstance(inner.get("products"), dict):
        return inner
    return payload


def from_html(text: str) -> dict | None:
    m = NEXT_DATA.search(text)
    if not m:
        return None
    try:
        data = json.loads(m.group(1))
    except json.JSONDecodeError:
        return None
    return ((data.get("props") or {}).get("pageProps") or {}).get("data")


def fetch_page(client: Client, slug: str, page: int) -> dict | None:
    r = client.get(f"{API}/{slug}?page={page}", "application/json")
    if r is not None:
        try:
            return unwrap(r.json())
        except ValueError:
            pass
    html_url = f"{SITE}/{slug}" + (f"/{page}" if page > 1 else "")
    r = client.get(html_url, "text/html")
    return from_html(r.text) if r is not None else None


def ancestors(category: dict) -> set[str]:
    out, node, depth = set(), category, 0
    while isinstance(node, dict) and depth < 10:
        if node.get("slug"):
            out.add(node["slug"])
        node, depth = node.get("parent"), depth + 1
    return out


def to_row(p: dict) -> dict | None:
    pid = p.get("id")
    name = (p.get("name") or "").strip()
    price = p.get("price")
    if pid is None or not name or not isinstance(price, (int, float)) or price <= 0:
        return None
    stock = p.get("stock")
    stock = stock if isinstance(stock, (int, float)) else None
    minimum = p.get("minimum") if isinstance(p.get("minimum"), (int, float)) else 1
    actual = p.get("actuality") is not False
    available = actual and (stock is None or stock >= max(minimum, 1))
    maker = p.get("manufacturer") or {}
    cat = p.get("category") or {}
    return {
        "id": str(pid),
        "sku": str(p.get("sku") or "").strip(),
        "name": name,
        "price": round(price / 100, 2),  # ціни на сайті зберігаються в копійках
        "available": available,
        "stock": stock,
        "brand": (maker.get("name") or "").strip() if isinstance(maker, dict) else "",
        "category": (cat.get("name") or "").strip() if isinstance(cat, dict) else "",
        "url": f"{SITE}/{p['slug']}" if p.get("slug") else "",
    }


def scrape(options: dict, log=print) -> list[dict]:
    """Повертає список товарів. options з config.json:
    delay            пауза між запитами, секунд (типово 0.7)
    categories       обійти лише ці категорії (slug зі адреси сторінки)
    skip_categories  пропустити категорії, що лежать усередині цих (наприклад, "roslyny")
    max_minutes      зупинитися, якщо обхід триває довше (типово 40)
    """
    client = Client(float(options.get("delay", 0.7)), log)
    skip = {str(s).strip("/ ") for s in options.get("skip_categories") or []}
    deadline = time.monotonic() + 60 * float(options.get("max_minutes", 40))

    slugs = [str(s).strip("/ ") for s in options.get("categories") or []] or category_slugs(client)
    log(f"  Клуб Рослин: категорій у карті сайту {len(slugs)}")

    products: dict[str, dict] = {}
    done_cats, empty_cats, skipped, failed = 0, 0, 0, []
    for n, slug in enumerate(slugs, 1):
        if slug in skip:
            skipped += 1
            continue
        if time.monotonic() > deadline:
            raise ScrapeError(f"обхід не вклався в {options.get('max_minutes', 40)} хв, оброблено {done_cats} "
                              f"з {len(slugs)} категорій; збільште max_minutes або звузьте categories")
        first = fetch_page(client, slug, 1)
        if first is None:
            failed.append(slug)
            continue
        if skip and ancestors((first.get("data") or {})) & skip:
            skipped += 1
            continue
        block = first.get("products")
        if not isinstance(block, dict) or not block.get("data"):
            empty_cats += 1  # розділ верхнього рівня без власного списку товарів
            continue
        last = int((block.get("meta") or {}).get("last_page") or 1)
        page_data = block
        page = 1
        while True:
            for p in page_data.get("data") or []:
                row = to_row(p)
                if row and row["id"] not in products:
                    products[row["id"]] = row
            if page >= last:
                break
            page += 1
            nxt = fetch_page(client, slug, page)
            page_data = (nxt or {}).get("products") or {}
            if not page_data.get("data"):
                failed.append(f"{slug} стор. {page}")
                break
        done_cats += 1
        if n % 50 == 0:
            log(f"    оброблено {n} з {len(slugs)} категорій, товарів {len(products)}")

    log(f"  Клуб Рослин: товарів {len(products)}, категорій з товарами {done_cats}, без товарів {empty_cats}, "
        f"пропущено {skipped}, запитів {client.requests}")
    if failed:
        log(f"  Клуб Рослин: не прочитались {len(failed)}: {', '.join(failed[:10])}")
    if not products:
        raise ScrapeError("не зібрано жодного товару")
    if failed and len(failed) > max(5, 0.2 * len(slugs)):
        raise ScrapeError(f"не прочиталась значна частина категорій ({len(failed)}), дані неповні")
    return list(products.values())

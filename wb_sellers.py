#!/usr/bin/env python3
"""
wb_sellers.py - собирает продавцов Wildberries из поисковой выдачи
и оценивает, у кого карточки "слабые" (мало фото) при большом ассортименте.

Запуск:
    pip install requests
    python wb_sellers.py            # собрать данные с WB и сохранить CSV
    python wb_sellers.py --cache    # пересчитать фильтры по уже скачанным данным (без запросов)

Результат:
    sellers_filtered.csv  - продавцы, прошедшие фильтры, лучшие сверху
    sellers_all.csv       - все найденные продавцы
    wb_raw.json           - сырые данные (для режима --cache)

Скрипт ходит редко и с паузами, чтобы не ловить 429. Нужен российский IP.
"""
import csv
import json
import os
import random
import sys
import time
from collections import defaultdict
from urllib.parse import quote

try:
    # curl_cffi притворяется настоящим Chrome на уровне TLS - WB так режет меньше
    from curl_cffi import requests
    IMPERSONATE = {"impersonate": "chrome"}
except ImportError:
    import requests
    IMPERSONATE = {}

# ============================ НАСТРОЙКИ ============================

# Поисковые запросы: чем конкретнее товар, тем лучше список.
QUERIES = [
    "постер а3", "постеры на стену", "постер интерьерный", "постер музыкальный",
    "картина на холсте", "картина на холсте 40х50", "картина интерьерная",
    "постеры в рамку", "набор постеров", "плакат на стену",
    "наклейки на стену", "наклейки интерьерные", "набор наклеек", "стикерпак",
    "магнит на холодильник", "открытки набор", "календарь настенный",
    "коврик для мыши", "чехол для телефона принт", "картина по номерам",
]

SORTS = ["popular", "newly"]   # newly = свежие товары, чаще новые магазины
MAX_PAGES = 10                 # страниц на запрос (на странице обычно до 100 товаров)
DELAY = (1.5, 3.0)             # пауза между запросами, секунды (случайная)
DEST = "1259570991"            # регион доставки (взят из вашего запроса)

# Если WB сменит версию API и скрипт перестанет находить товары:
# откройте wildberries.ru, введите запрос, F12 -> Network -> найдите запрос "search?..."
# и скопируйте его адрес без параметров сюда.
API_URL = "https://www.wildberries.ru/__internal/u-search/exactmatch/ru/common/v18/search"

# Фильтры для sellers_filtered.csv
WEAK_PICS = 3            # карточка "слабая", если фото не больше этого числа
MIN_CARDS = 8            # минимум карточек продавца, найденных в выдаче
MIN_WEAK_SHARE = 0.4     # минимум доли слабых карточек
MIN_SELLER_RATING = 4.5  # рейтинг продавца (0 = неизвестен, не отсекаем)
MAX_AVG_FEEDBACKS = 500  # среднее число отзывов на карточку: крупные продавцы уже
                         # хорошо оформлены, малые и средние - ваша цель

# Если приходит 403: положите рядом со скриптом файл cookie.txt, а в него вставьте
# строку Cookie из браузера (F12 -> Network -> запрос "search?..." -> Headers ->
# Request Headers -> Cookie). По желанию - файл ua.txt со строкой User-Agent оттуда же.
def _read(name):
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), name),
                  encoding="utf-8") as f:
            return " ".join(f.read().split())   # убирает переносы строк
    except OSError:
        return ""


COOKIE = _read("cookie.txt")
USER_AGENT = _read("ua.txt")

MAX_FAILS_IN_ROW = 3     # сколько неудач подряд терпеть, прежде чем остановиться

RAW_FILE = "wb_raw.json"
OUT_ALL = "sellers_all.csv"
OUT_FILTERED = "sellers_filtered.csv"

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept": "*/*",
    "Origin": "https://www.wildberries.ru",
    "Referer": "https://www.wildberries.ru/",
}

# ==================================================================


def pick(d, *keys, default=0):
    for k in keys:
        v = d.get(k)
        if v is not None:
            return v
    return default


def fetch_page(session, query, sort, page):
    """Возвращает список товаров, [] если страница пустая, None при ошибке."""
    params = {
        "appType": 1, "autoselectFilters": "false", "curr": "rub", "dest": DEST,
        "hide_vflags": "4294967296", "inheritFilters": "true", "lang": "ru",
        "locale": "ru", "query": query, "resultset": "catalog", "scale": 4,
        "sort": sort, "spp": 30, "suppressSpellcheck": "false", "page": page,
    }
    for attempt in range(1, 6):
        try:
            r = session.get(API_URL, params=params, timeout=25)
        except Exception as e:
            print(f"   сеть: {e}; повтор {attempt}/5")
            time.sleep(5 * attempt)
            continue
        if r.status_code == 429 or r.status_code >= 500:
            wait = 15 * attempt
            print(f"   код {r.status_code}, жду {wait} c (попытка {attempt}/5)")
            time.sleep(wait)
            continue
        if r.status_code != 200:
            print(f"   код {r.status_code}: {r.text[:200]!r}")
            if r.status_code in (403, 498):
                print("   WB блокирует запросы: выключите VPN, установите curl_cffi "
                      "(pip install curl_cffi) или положите Cookie из браузера в файл cookie.txt рядом со скриптом")
            return None
        try:
            data = r.json()
        except ValueError:
            print(f"   ответ не JSON: {r.text[:200]!r}")
            return None
        products = data.get("products") or (data.get("data") or {}).get("products")
        if products is None:
            print(f"   неожиданный формат ответа: {str(data)[:200]!r}")
            print("   возможно, WB сменил версию API - см. комментарий у API_URL")
            return None
        probe(products)
        return products
    return None


_probed = False


def probe(products):
    """При первой успешной выдаче показывает, какие поля отдаёт WB (для отладки)."""
    global _probed
    if _probed or not products:
        return
    _probed = True
    first = products[0]
    print("   ПОЛЯ ТОВАРА:", ", ".join(sorted(first.keys())))
    for need in ("supplierId", "supplier", "pics"):
        if need not in first:
            print(f"   ВНИМАНИЕ: поля '{need}' нет в ответе - скрипту нужна правка, "
                  f"пришлите содержимое wb_probe.json")
    with open("wb_probe.json", "w", encoding="utf-8") as f:
        json.dump(products[:2], f, ensure_ascii=False, indent=2)


def to_item(p, query):
    """Превращает товар из ответа WB в плоский словарь."""
    pid = p.get("id")
    if pid is None:
        return None
    price = 0
    try:
        price = p["sizes"][0]["price"]["product"] / 100
    except (KeyError, IndexError, TypeError):
        pass
    return {
        "id": pid,
        "name": p.get("name", ""),
        "supplier": p.get("supplier", ""),
        "supplier_id": p.get("supplierId"),
        "supplier_rating": pick(p, "supplierRating", default=0),
        "pics": pick(p, "pics", default=0),
        "feedbacks": pick(p, "feedbacks", "nmFeedbacks", default=0),
        "price": price,
        "query": query,
    }


def collect():
    session = requests.Session(**IMPERSONATE)
    session.headers.update(HEADERS)
    if COOKIE:
        session.headers["Cookie"] = COOKIE
        print("Cookie из cookie.txt подключён")
    else:
        print("cookie.txt не найден - запросы идут без cookie")
    if USER_AGENT:
        session.headers["User-Agent"] = USER_AGENT
    items = {}  # id товара -> данные (дедупликация между запросами)
    fails = 0
    try:
        for q in QUERIES:
            for sort in SORTS:
                for page in range(1, MAX_PAGES + 1):
                    print(f"[{q} | {sort} | стр. {page}]  найдено товаров: {len(items)}")
                    products = fetch_page(session, q, sort, page)
                    time.sleep(random.uniform(*DELAY))
                    if products is None:
                        fails += 1
                        if fails >= MAX_FAILS_IN_ROW:
                            print("Слишком много ошибок подряд, останавливаюсь.")
                            raise KeyboardInterrupt
                        break
                    fails = 0
                    if not products:
                        break
                    for p in products:
                        it = to_item(p, q)
                        if it and it["id"] not in items:
                            items[it["id"]] = it
    except KeyboardInterrupt:
        print("\nПрервано, сохраняю то, что успели собрать.")
    with open(RAW_FILE, "w", encoding="utf-8") as f:
        json.dump(list(items.values()), f, ensure_ascii=False)
    return list(items.values())


def aggregate(items):
    sellers = defaultdict(lambda: {
        "cards": 0, "weak": 0, "pics_sum": 0, "fb_sum": 0, "price_sum": 0.0,
        "rating": 0, "name": "", "queries": set(), "example": None,
    })
    for it in items:
        sid = it["supplier_id"]
        if not sid:
            continue
        s = sellers[sid]
        s["cards"] += 1
        s["pics_sum"] += it["pics"] or 0
        s["fb_sum"] += it["feedbacks"] or 0
        s["price_sum"] += it["price"] or 0
        s["weak"] += 1 if (it["pics"] or 0) <= WEAK_PICS else 0
        s["rating"] = it["supplier_rating"] or s["rating"]
        s["name"] = it["supplier"] or s["name"]
        s["queries"].add(it["query"])
        if s["example"] is None or (it["pics"] or 0) < (s["example"]["pics"] or 0):
            s["example"] = it   # самая слабая карточка - лучший пример для письма

    rows = []
    for sid, s in sellers.items():
        n = s["cards"]
        ex = s["example"]
        rows.append({
            "score": s["weak"],
            "supplier": s["name"],
            "supplier_id": sid,
            "seller_url": f"https://www.wildberries.ru/seller/{sid}",
            "cards_seen": n,
            "weak_cards": s["weak"],
            "weak_share": round(s["weak"] / n, 2),
            "avg_pics": round(s["pics_sum"] / n, 1),
            "avg_feedbacks": round(s["fb_sum"] / n),
            "supplier_rating": s["rating"],
            "avg_price_rub": round(s["price_sum"] / n) if s["price_sum"] else "",
            "found_by_queries": "; ".join(sorted(s["queries"])),
            "example_weak_card": f"https://www.wildberries.ru/catalog/{ex['id']}/detail.aspx",
            "example_name": ex["name"],
            "contact_search": "https://yandex.ru/search/?text="
                              + quote(f"{s['name']} wildberries телеграм сайт"),
            "status": "",
        })
    rows.sort(key=lambda r: (r["score"], r["cards_seen"]), reverse=True)
    return rows


def passes(r):
    return (
        r["cards_seen"] >= MIN_CARDS
        and r["weak_share"] >= MIN_WEAK_SHARE
        and (not r["supplier_rating"] or r["supplier_rating"] >= MIN_SELLER_RATING)
        and r["avg_feedbacks"] <= MAX_AVG_FEEDBACKS
    )


def write_csv(path, rows):
    if not rows:
        print(f"{path}: нечего записывать")
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel откроет без кракозябр
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";")
        w.writeheader()
        w.writerows(rows)
    print(f"{path}: {len(rows)} строк")


def main():
    if "--cache" in sys.argv:
        with open(RAW_FILE, encoding="utf-8") as f:
            items = json.load(f)
        print(f"Из кэша: {len(items)} товаров")
    else:
        items = collect()
        print(f"Собрано товаров: {len(items)}")
    rows = aggregate(items)
    write_csv(OUT_ALL, rows)
    write_csv(OUT_FILTERED, [r for r in rows if passes(r)])


if __name__ == "__main__":
    main()
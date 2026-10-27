#!/usr/bin/env python3
"""
wb_sellers_browser.py - то же, что wb_sellers.py, но через настоящий браузер.
Открывает Chrome, заходит на поиск WB как обычный покупатель и читает ответы,
которые сам сайт получает от своего API. Защита от ботов на это не срабатывает.

Установка (один раз):
    pip3 install playwright
    python3 -m playwright install chromium     # нужно, только если у вас нет Google Chrome

Запуск (wb_sellers.py должен лежать в той же папке):
    python3 wb_sellers_browser.py

Окно браузера откроется само - не закрывайте его, можно свернуть.
Если WB покажет капчу или проверку, пройдите её вручную и нажмите Enter в терминале.
"""
import json
import random
import sys
from urllib.parse import quote

from playwright.sync_api import sync_playwright

import wb_sellers as w

PROFILE_DIR = "wb_profile"   # тут браузер хранит cookie между запусками
PAGES = 5                    # страниц на запрос и сортировку (до ~100 товаров на странице)
PAUSE_MS = (2500, 5000)      # пауза между страницами
URL = "https://www.wildberries.ru/catalog/0/search.aspx?search={q}&sort={sort}&page={page}"


def launch(pw):
    try:
        return pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=False,
            viewport={"width": 1280, "height": 900}, locale="ru-RU")
    except Exception as e:
        print(f"Google Chrome не найден ({e.__class__.__name__}), использую встроенный Chromium")
        return pw.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False,
            viewport={"width": 1280, "height": 900}, locale="ru-RU")


def products_from(data):
    if not isinstance(data, dict):
        return None
    return data.get("products") or (data.get("data") or {}).get("products")


def main():
    items = {}
    captured = []

    def on_response(resp):
        u = resp.url
        if "/u-search/" in u and "resultset=catalog" in u:
            try:
                captured.append(resp.json())
            except Exception:
                pass

    asked_for_help = False
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.on("response", on_response)
        try:
            for q in w.QUERIES:
                for sort in w.SORTS:
                    for n in range(1, PAGES + 1):
                        print(f"[{q} | {sort} | стр. {n}]  найдено товаров: {len(items)}")
                        got = None
                        for attempt in (1, 2):
                            captured.clear()
                            try:
                                page.goto(URL.format(q=quote(q), sort=sort, page=n),
                                          wait_until="domcontentloaded", timeout=45000)
                            except Exception as e:
                                print(f"   страница не открылась: {e.__class__.__name__}")
                            for _ in range(40):          # ждём ответ API до 20 секунд
                                if captured:
                                    break
                                page.wait_for_timeout(500)
                            if captured:
                                got = products_from(captured[-1])
                                break
                            if attempt == 1 and not asked_for_help:
                                asked_for_help = True
                                input("   Ответа с товарами нет. Если в окне браузера "
                                      "капча или проверка - пройдите её, потом нажмите Enter: ")
                        page.wait_for_timeout(random.randint(*PAUSE_MS))
                        if not got:
                            break   # пустая страница или конец выдачи - к следующей сортировке
                        w.probe(got)
                        for p in got:
                            it = w.to_item(p, q)
                            if it and it["id"] not in items:
                                items[it["id"]] = it
        except KeyboardInterrupt:
            print("\nПрервано, сохраняю то, что успели собрать.")
        finally:
            pass

    with open(w.RAW_FILE, "w", encoding="utf-8") as f:
        json.dump(list(items.values()), f, ensure_ascii=False)
    print(f"Собрано товаров: {len(items)}")
    rows = w.aggregate(list(items.values()))
    w.write_csv(w.OUT_ALL, rows)
    w.write_csv(w.OUT_FILTERED, [r for r in rows if w.passes(r)])


if __name__ == "__main__":
    main()

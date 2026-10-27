# wb_stocks v1.1
# G2G code
#
# Проставляет остатки на уже созданные карточки Wildberries (склад продавца, FBS).
# Артикулы берутся из Постеры.xlsx (столбец «Артикул продавца»), по ним в WB
# находятся карточки и их штрихкоды (sku); каждому штрихкоду выставляется остаток.
#
# Настройки — в config.json (папка scripts/), раздел "wildberries":
#   {"wildberries": {"api_token": "...", "stock_amount": 100, "warehouse_id": 123456}}
#   - stock_amount  — на сколько пополнять остаток (меняется в «Настройки» приложения);
#   - warehouse_id  — необязательно: если у вас один склад, он выберется сам;
#                     если складов несколько, скрипт спросит и запомнит выбор.
#
# ВАЖНО: токену нужно право «Маркетплейс» (кроме «Контент»), иначе WB вернёт 401/403.
# Остаток ЗАМЕНЯЕТСЯ на указанное число (так работает API WB), а не прибавляется.
#
# Режим «все товары магазина»: если в папке проекта (рядом с Постеры.xlsx) лежит файл остатков
# stocks.xlsx (выгрузка из кабинета WB), скрипт работает по нему: ставит нужное число в столбец
# «Количество», по штрихкодам из столбца «Баркод» отправляет остатки на склад WB и сообщает,
# что файл можно удалить. Постеры.xlsx в этом режиме не нужен.
#
# Запуск из терминала: python3 wb_stocks.py

import json
import re
import time
from pathlib import Path

import requests
from openpyxl import load_workbook

CONTENT_URL = "https://content-api.wildberries.ru"
MARKETPLACE_URL = "https://marketplace-api.wildberries.ru"

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent
CONFIG_FILE = SCRIPT_DIR / "config.json"
EXCEL_FILE = ROOT_DIR / "Постеры.xlsx"

SHEET_NAME = "Товары"
ARTICLES_COLUMN = "Артикул продавца"
STOP_WORDS = ("номер", "идентификатор")  # строки-подсказки в шапке шаблона WB

DEFAULT_AMOUNT = 100
STOCK_BATCH = 500  # WB принимает до 1000 штрихкодов за запрос

HEADERS = {}
LOG_CALLBACK = None
PROGRESS_CALLBACK = None
PROMPT_CALLBACK = None


class WBAuthError(RuntimeError):
    """WB отклонил токен (401/403) — продолжать бессмысленно."""


def _log(*args, **kwargs):
    import builtins
    builtins.print(*args, **kwargs)
    if LOG_CALLBACK:
        try:
            LOG_CALLBACK(" ".join(str(a) for a in args))
        except Exception:
            pass


def _progress(value, message=""):
    if PROGRESS_CALLBACK:
        try:
            PROGRESS_CALLBACK(max(0, min(100, int(value))), message)
        except Exception:
            pass


# --- Настройки ---
def read_config(config_path=None):
    path = Path(config_path) if config_path else CONFIG_FILE
    if not path.exists():
        raise RuntimeError(
            f"Не найден config.json: {path}\nВведите токен через кнопку «Настройки» в приложении."
        )
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать config.json: {e}")
    return path, data if isinstance(data, dict) else {}


def save_warehouse(path, data, warehouse_id):
    block = data.get("wildberries")
    if not isinstance(block, dict):
        block = {}
    block["warehouse_id"] = warehouse_id
    data["wildberries"] = block
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def parse_amount(value, default=DEFAULT_AMOUNT):
    try:
        amount = int(str(value).strip())
    except (TypeError, ValueError):
        return default
    return amount if amount >= 0 else default


# --- Excel ---
def load_articles():
    """Артикулы продавца из Постеры.xlsx (без строк-подсказок шаблона)."""
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
    except PermissionError:
        raise SystemExit("x Не удалось открыть Постеры.xlsx — закройте файл в Excel и запустите снова.")
    try:
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        header_row = column = None
        for r in range(1, min(ws.max_row, 30) + 1):
            for c in ws[r]:
                if isinstance(c.value, str) and c.value.strip().lower() == ARTICLES_COLUMN.lower():
                    header_row, column = r, c.column
                    break
            if header_row:
                break
        if not header_row:
            raise SystemExit(f"x В таблице не найден столбец '{ARTICLES_COLUMN}'")

        articles = []
        for r in range(header_row + 1, ws.max_row + 1):
            value = ws.cell(row=r, column=column).value
            if value is None:
                continue
            article = str(value).strip()
            if not article or any(w in article.lower() for w in STOP_WORDS):
                continue
            if article not in articles:
                articles.append(article)
        return articles
    finally:
        wb.close()


# --- WB API ---
def _check_auth(response):
    if response.status_code == 401:
        raise WBAuthError("WB вернул 401: токен неверный или просрочен.")
    if response.status_code == 403:
        raise WBAuthError(
            "WB вернул 403: у токена нет доступа. Для остатков нужно право «Маркетплейс» "
            "(и «Контент» для поиска карточек) — выпустите токен с этими правами."
        )


def api(method, url, json_body=None, retries=5):
    """Запрос к WB с повторами на 429/5xx и сетевые ошибки."""
    for attempt in range(retries):
        try:
            r = requests.request(method, url, headers=HEADERS, json=json_body, timeout=60)
        except requests.RequestException as e:
            _log(f"!  Сетевая ошибка ({type(e).__name__}), повтор...")
            time.sleep(3 * (attempt + 1))
            continue
        _check_auth(r)
        if r.status_code == 429 or r.status_code >= 500:
            try:
                wait = float(r.headers.get("X-Ratelimit-Retry", 0))
            except ValueError:
                wait = 0
            time.sleep(wait or 5 * (attempt + 1))
            continue
        return r
    raise RuntimeError(f"{method} {url}: нет успешного ответа после {retries} попыток")


def fetch_cards(wanted):
    """{vendorCode: [штрихкоды]}; листает каталог от новых карточек и останавливается,
    когда найдены все нужные артикулы."""
    wanted = set(wanted)
    found = {}
    cursor = {"limit": 100}
    while True:
        payload = {"settings": {
            "sort": {"ascending": False},
            "cursor": cursor,
            "filter": {"withPhoto": -1},
        }}
        r = api("POST", CONTENT_URL + "/content/v2/get/cards/list?locale=ru", json_body=payload)
        r.raise_for_status()
        data = r.json()
        cards = data.get("cards") or []
        for card in cards:
            vendor = str(card.get("vendorCode", "")).strip()
            if vendor in wanted:
                skus = []
                for size in card.get("sizes") or []:
                    skus.extend(str(s) for s in (size.get("skus") or []))
                found[vendor] = skus
        cur = data.get("cursor") or {}
        if wanted.issubset(found.keys()):
            break
        if not cards or cur.get("total", 0) < 100:
            break
        cursor = {"limit": 100, "updatedAt": cur.get("updatedAt"), "nmID": cur.get("nmID")}
        time.sleep(0.7)
    return found


def choose_warehouse(path, data):
    """Склад из config.json, а если его нет — единственный склад продавца или выбор из списка."""
    saved = (data.get("wildberries") or {}).get("warehouse_id")
    r = api("GET", MARKETPLACE_URL + "/api/v3/warehouses")
    if r.status_code != 200:
        raise RuntimeError(f"Не удалось получить список складов ({r.status_code}): {r.text[:200]}")
    warehouses = [w for w in (r.json() or []) if not w.get("isDeleting")]
    if not warehouses:
        raise RuntimeError(
            "У продавца нет ни одного склада (FBS). Создайте склад в личном кабинете WB "
            "(Маркетплейс → Мои склады) и запустите снова."
        )

    by_id = {str(w.get("id")): w for w in warehouses}
    if saved is not None and str(saved) in by_id:
        w = by_id[str(saved)]
        _log(f"Склад: {w.get('name')} (ID {w.get('id')})")
        return w["id"]

    if len(warehouses) == 1:
        w = warehouses[0]
    else:
        lines = [f"  {i}. {w.get('name')} (ID {w.get('id')})" for i, w in enumerate(warehouses, 1)]
        question = "Несколько складов, на какой проставить остатки?\n" + "\n".join(lines) + "\nВведите номер:"
        answer = ""
        if PROMPT_CALLBACK:
            answer = PROMPT_CALLBACK(question, "1") or ""
        else:
            answer = input(question + " ")
        try:
            w = warehouses[int(str(answer).strip()) - 1]
        except (ValueError, IndexError):
            raise RuntimeError("Склад не выбран — остановка.")
    _log(f"Склад: {w.get('name')} (ID {w.get('id')})")
    try:
        save_warehouse(path, data, w["id"])
    except OSError:
        pass
    return w["id"]


def put_stocks(warehouse_id, skus, amount):
    """Возвращает (успешно, ошибка) — ошибка None, если всё прошло."""
    url = f"{MARKETPLACE_URL}/api/v3/stocks/{warehouse_id}"
    body = {"stocks": [{"sku": sku, "amount": amount} for sku in skus]}
    r = api("PUT", url, json_body=body)
    if r.status_code in (200, 204):
        return True, None
    return False, f"{r.status_code}: {r.text[:300]}"


STOCKS_BARCODE_COLUMN = "Баркод"
STOCKS_AMOUNT_COLUMN = "Количество"


def find_stocks_file():
    """stocks.xlsx (или «stocks (1).xlsx» и т.п.), брошенный в папку проекта рядом с Постеры.xlsx."""
    if not ROOT_DIR.is_dir():
        return None
    found = [
        f for f in ROOT_DIR.iterdir()
        if f.is_file()
        and not f.name.startswith(("~$", "."))
        and f.name.lower().startswith("stocks")
        and f.suffix.lower() == ".xlsx"
    ]
    if not found:
        return None
    found.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    if len(found) > 1:
        _log(f"! В папке несколько файлов остатков, беру самый свежий: {found[0].name}")
    return found[0]


def fill_stocks_file(path, amount):
    """Пишет amount в столбец «Количество» файла остатков WB и сохраняет его.

    Возвращает (штрихкоды без повторов, число строк с товарами)."""
    try:
        wb = load_workbook(path)
    except PermissionError:
        raise SystemExit(f"x Не удалось открыть {path.name} — закройте файл в Excel и запустите снова.")
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать {path.name}: {e}")

    try:
        ws = wb.active
        header_row = barcode_col = amount_col = None
        for r in range(1, min(ws.max_row, 10) + 1):
            labels = {
                c.value.strip().lower(): c.column
                for c in ws[r] if isinstance(c.value, str)
            }
            if STOCKS_BARCODE_COLUMN.lower() in labels and STOCKS_AMOUNT_COLUMN.lower() in labels:
                header_row = r
                barcode_col = labels[STOCKS_BARCODE_COLUMN.lower()]
                amount_col = labels[STOCKS_AMOUNT_COLUMN.lower()]
                break
        if not header_row:
            raise RuntimeError(
                f"В {path.name} не найдены столбцы «{STOCKS_BARCODE_COLUMN}» и «{STOCKS_AMOUNT_COLUMN}». "
                "Это должна быть выгрузка остатков из личного кабинета WB."
            )

        skus, seen, rows = [], set(), 0
        for r in range(header_row + 1, ws.max_row + 1):
            raw = ws.cell(row=r, column=barcode_col).value
            if raw is None:
                continue
            if isinstance(raw, float) and raw.is_integer():
                raw = int(raw)
            sku = str(raw).strip()
            if not sku:
                continue
            ws.cell(row=r, column=amount_col).value = amount
            rows += 1
            if sku not in seen:
                seen.add(sku)
                skus.append(sku)

        try:
            wb.save(path)
        except PermissionError:
            raise SystemExit(f"x Не удалось сохранить {path.name} — закройте файл в Excel и запустите снова.")
        return skus, rows
    finally:
        wb.close()


def collect_skus_from_stocks_file(path, amount, result):
    """Режим «все товары магазина»: штрихкоды берутся прямо из stocks.xlsx."""
    result["stocks_file"] = path.name
    _log("=" * 60)
    _log("WB: ОСТАТКИ — ВСЕ ТОВАРЫ ИЗ ФАЙЛА")
    _log("=" * 60)
    _log(f"Найден файл остатков: {path.name}")
    _progress(10, f"Чтение {path.name}")
    skus, rows = fill_stocks_file(path, amount)
    result["articles"] = rows
    _log(f"Товаров в файле: {rows}. В столбец «{STOCKS_AMOUNT_COLUMN}» записано: {amount}.")
    if not skus:
        raise RuntimeError(f"В {path.name} нет ни одного штрихкода — проставлять нечего.")
    return skus


def collect_skus_from_table(amount, result):
    """Обычный режим: артикулы из Постеры.xlsx -> карточки WB -> штрихкоды."""
    if not EXCEL_FILE.exists():
        raise RuntimeError(f"Не найден файл '{EXCEL_FILE}'")
    articles = load_articles()
    result["articles"] = len(articles)
    if not articles:
        raise RuntimeError("В таблице нет артикулов — сначала заполните Постеры.xlsx.")

    _log("=" * 60)
    _log("WB: ОСТАТКИ")
    _log("=" * 60)
    _log(f"Артикулов в таблице: {len(articles)}; остаток на каждый: {amount}")

    _progress(10, "Поиск карточек WB")
    cards = fetch_cards(articles)

    skus = []
    for i, article in enumerate(articles, start=1):
        counter = f"[{i}/{len(articles)}]"
        if article not in cards:
            _log(f"{counter} {article} : карточка не найдена")
            result["not_found_cards"].append(article)
        elif not cards[article]:
            _log(f"{counter} {article} : у карточки нет штрихкода")
            result["no_barcodes"].append(article)
        else:
            _log(f"{counter} {article}")
            skus.extend(cards[article])

    if not skus:
        raise RuntimeError("Не нашлось ни одной карточки со штрихкодом — проставлять нечего.")

    return skus


def main(
    artist=None,
    log_callback=None,
    progress_callback=None,
    prompt_callback=None,
    config_path=None,
    amount=None,
):
    """Проставляет остатки. Возвращает словарь {"success": bool, "error": str|None, ...}."""
    global HEADERS, LOG_CALLBACK, PROGRESS_CALLBACK, PROMPT_CALLBACK
    LOG_CALLBACK = log_callback
    PROGRESS_CALLBACK = progress_callback
    PROMPT_CALLBACK = prompt_callback

    result = {
        "success": False,
        "articles": 0,
        "updated": 0,
        "amount": None,
        "not_found_cards": [],
        "no_barcodes": [],
        "error": None,
    }

    try:
        _progress(0, "Запуск")
        path, data = read_config(config_path)
        wb_block = data.get("wildberries") or {}
        token = str(wb_block.get("api_token") or "").strip()
        if not token:
            raise RuntimeError("В config.json не заполнен wildberries.api_token")
        HEADERS = {"Authorization": token}

        amount = parse_amount(amount if amount is not None else wb_block.get("stock_amount"))
        result["amount"] = amount

        stocks_file = find_stocks_file()
        if stocks_file:
            skus = collect_skus_from_stocks_file(stocks_file, amount, result)
        else:
            skus = collect_skus_from_table(amount, result)

        _progress(40, "Выбор склада")
        warehouse_id = choose_warehouse(path, data)

        failed_batches = []
        batches = [skus[i:i + STOCK_BATCH] for i in range(0, len(skus), STOCK_BATCH)]
        for n, batch in enumerate(batches, start=1):
            _progress(50 + 50 * (n - 1) / len(batches), f"Остатки: {n}/{len(batches)}")
            ok, err = put_stocks(warehouse_id, batch, amount)
            if ok:
                result["updated"] += len(batch)
                _log(f"✓ Остатки проставлены: {len(batch)} шт. (штрихкодов)")
            else:
                failed_batches.append(err)
                _log(f"✗ Не удалось проставить остатки для {len(batch)} штрихкодов: {err}")
            time.sleep(0.5)

        _progress(100, "Завершено")
        problems = []
        if result["not_found_cards"]:
            problems.append(f"не найдены карточки WB: {len(result['not_found_cards'])}")
        if result["no_barcodes"]:
            problems.append(f"нет штрихкодов: {len(result['no_barcodes'])}")
        if failed_batches:
            problems.append(f"WB отклонил запросов: {len(failed_batches)}")
        if problems:
            result["error"] = (
                "Остатки проставлены не полностью — " + "; ".join(problems) + ".\n"
                "Подробности — в логе выше. Повторный запуск безопасен: остаток просто перезапишется."
                + (f"\nФайл {result['stocks_file']} пока не удаляйте." if result.get("stocks_file") else "")
            )
        else:
            result["success"] = True
            if result.get("stocks_file"):
                _log(
                    f"\n✓ Всё готово: остатки ({amount}) проставлены у {result['updated']} штрихкодов. "
                    f"Таблицу {result['stocks_file']} можно удалить."
                )
        return result

    except WBAuthError as e:
        result["error"] = str(e)
        _log(f"\nОШИБКА: {result['error']}")
        return result

    except SystemExit as e:
        result["error"] = re.sub(r"^[x✗]\s+", "", str(e)) or "Операция остановлена."
        _log(f"\nОШИБКА: {result['error']}")
        return result

    except requests.RequestException as e:
        result["error"] = (
            "Не удалось связаться с Wildberries. Проверьте интернет (и VPN, если он включён) "
            f"и запустите снова.\n({type(e).__name__})"
        )
        _log(f"\nОШИБКА: {result['error']}")
        return result

    except Exception as e:
        result["error"] = str(e)
        _log(f"\nОШИБКА: {e}")
        return result


if __name__ == "__main__":
    main()
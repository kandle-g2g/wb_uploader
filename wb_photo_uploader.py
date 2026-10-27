# wb_photo_uploader v1.9
# G2G code
#
# v1.9: в логе перед каждым артикулом выводится счётчик [N/всего],
#   как в pinterest_uploader: [1/5] blonde white frank ocean.
#
# v1.8: скрипт встроен в Posters.app (блок «Дополнительно»):
#   - токен WB больше не хранится в коде — берётся из config.json (папка scripts/),
#     тот же, что использует wb_uploader.py (кнопка «Настройки» в приложении);
#   - вместо input() имя исполнителя приходит из приложения (main(artist=...));
#   - вывод идёт и в консоль, и в лог приложения; в конце возвращается словарь
#     {"success": ..., "error": ...}, как у остальных скриптов;
#   - pandas заменён на openpyxl (в приложение pandas не упаковывается);
#   - при 401/403 загрузка останавливается сразу, а не печатает ✗ на каждое фото;
#   - итог честный: если часть фото не загрузилась или не найдены карточки/папки,
#     success = False и причина перечислена в логе.
#   Запуск из терминала по-прежнему работает: python3 wb_photo_uploader.py
#
# v1.7: полный скан всех карточек WB (у вас их 600+ и постоянно прибавляется)
# заменён на постраничный запрос с досрочной остановкой — как только среди уже
# полученных карточек нашлись ВСЕ артикулы из текущей загрузки (обычно 10-30),
# дальнейшая пагинация не выполняется. Карточки запрашиваются в порядке "от
# недавно обновлённых", поэтому только что залитые/отредактированные карточки
# обычно попадают на первые страницы и поиск останавливается очень рано.
# (Фильтр textSearch по имени исполнителя не прижился — WB по нему ничего не
# находил даже на точном совпадении, поэтому от него отказались.)
#
# v1.6: общие фото (mockups/8.jpg, mockups/cover.jpg) больше не грузятся на
# жёстко заданные места №8/№9 — теперь они всегда идут сразу после реально
# найденных "своих" фото исходника, без пропусков в нумерации, если какого-то
# из "своих" фото не хватает (например, для этой расцветки/альбома часть
# кадров не генерировалась).
#
# Ожидаемое расположение (скрипт лежит в <корень>/scripts/):
#   <корень>/Постеры.xlsx
#   <корень>/scripts/config.json                               — токен WB
#   <корень>/mockups/{8.jpg, cover.jpg}                        — общие для ВСЕХ карточек
#   <корень>/output/<Исполнитель>/<исходник>/{1,2,3,4,5-3|5-4|5-5,7}.jpg  — свои для каждого исходника
#
# ВАЖНО (проверьте перед использованием): порядок фото на карточке WB задан
# списком PHOTO_ORDER ниже — 1,2,3,4, один из 5-3/5-4/5-5 (какой реально
# сгенерирован для этого исходника), 6, затем 7. Сразу после них на карточку
# добавляются два общих файла из <корень>/mockups/ (8.jpg, cover.jpg) —
# в этом порядке, одинаковые для всех карточек.
#
# Нумерация мест на карточке всегда СПЛОШНАЯ, без фиксированных номеров:
# если у какого-то исходника не хватает части "своих" фото (например, для
# этой расцветки/альбома часть кадров не генерировалась), то все следующие
# фото — включая 8.jpg и cover.jpg — просто сдвигаются на освободившиеся
# места. Относительный порядок (сначала свои фото, потом общие) при этом не
# меняется, но конкретный номер места на карточке зависит от того, сколько
# "своих" фото реально нашлось.
# Если на Wildberries нужен другой порядок — поменяйте PHOTO_ORDER и/или
# STATIC_PHOTOS.

import json
import time
from pathlib import Path

import requests
from openpyxl import load_workbook

BASE_URL = "https://content-api.wildberries.ru"

# Заполняются в main() из config.json
API_TOKEN = ""
HEADERS = {}

# ------------------- РАСПОЛОЖЕНИЕ ПРОЕКТА -------------------
SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent
CONFIG_FILE = SCRIPT_DIR / "config.json"

EXCEL_FILE = ROOT_DIR / "Постеры.xlsx"
OUTPUT_FOLDER = ROOT_DIR / "output"
MOCKUPS_FOLDER = ROOT_DIR / "mockups"

# ------------------- НАСТРОЙКИ -------------------
SHEET_NAME = "Товары"
ARTICLES_COLUMN = "Артикул продавца"
STOP_WORDS = ("номер", "идентификатор")  # строки-подсказки в шапке шаблона WB

# Порядок "своих" фото исходника (лежат в output/<Исполнитель>/<исходник>/).
# "5-x" — специальная метка: подставляется тот из 5-3/5-4/5-5.jpg, который
# реально существует у конкретного исходника (генерируется только один).
PHOTO_ORDER = ["1", "2", "3", "4", "5-x", "6", "7"]

# Общие фото (лежат в <корень>/mockups/), одинаковые для всех карточек.
# Порядок в списке важен: они всегда идут СРАЗУ ПОСЛЕ "своих" фото исходника,
# без фиксированных номеров места (см. upload_article). Если у какого-то
# исходника не хватает части "своих" фото — общие файлы просто занимают
# освободившиеся места, разрывов в нумерации на карточке не остаётся.
STATIC_PHOTOS = ["8", "cover"]

FORCE_REUPLOAD = False  # если True — перезагружает всегда, если False - пропускает, если фото уже достаточно

# ------------------- СВЯЗЬ С ПРИЛОЖЕНИЕМ -------------------
LOG_CALLBACK = None
PROGRESS_CALLBACK = None
PROMPT_CALLBACK = None


class WBAuthError(RuntimeError):
    """WB отклонил токен (401/403) — продолжать бессмысленно."""


def _log(*args, **kwargs):
    """Сообщение выводится в Terminal и уходит в лог приложения."""
    import builtins
    builtins.print(*args, **kwargs)
    if LOG_CALLBACK:
        try:
            LOG_CALLBACK(" ".join(str(a) for a in args))
        except Exception:
            pass


def _progress(value, message=""):
    """Прогресс 0..100 для приложения."""
    if PROGRESS_CALLBACK:
        try:
            PROGRESS_CALLBACK(max(0, min(100, int(value))), message)
        except Exception:
            pass


# --- ТОКЕН ---
def load_token(config_path=None):
    """Читает токен WB из config.json: {"wildberries": {"api_token": "..."}}"""
    path = Path(config_path) if config_path else CONFIG_FILE
    if not path.exists():
        raise RuntimeError(
            f"Не найден config.json: {path}\n"
            "Введите токен через кнопку «Настройки» в приложении."
        )
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать config.json: {e}")

    token = ((data.get("wildberries") or {}).get("api_token") or "").strip()
    if not token:
        raise RuntimeError("В config.json не заполнен wildberries.api_token")
    return token


# --- ИСПОЛНИТЕЛЬ И ПАПКА С РЕЗУЛЬТАТАМИ ---
def get_artist_name(artist=None):
    name = (artist or "").strip()
    if not name and PROMPT_CALLBACK:
        name = (PROMPT_CALLBACK("Введите имя исполнителя:", None) or "").strip()
    if not name and not PROMPT_CALLBACK:
        name = input("Введите имя исполнителя: ").strip()
    if not name:
        raise SystemExit("Имя исполнителя не введено — остановка.")
    return name


def find_artist_output_folder(artist_name):
    """Ищет output/<Исполнитель> без учёта регистра (на диске могло сохраниться
    ровно так, как ввели в mockuper.jsx, а тут ввод может отличаться регистром)."""
    if not OUTPUT_FOLDER.exists():
        return None
    target = artist_name.strip().lower()
    for sub in OUTPUT_FOLDER.iterdir():
        if sub.is_dir() and sub.name.strip().lower() == target:
            return sub
    return None


def build_article_folder_index(artist_folder, artist_name):
    """
    article -> Path(папка с фото), по тому же правилу, что и excel_filler.py:
    article = "<имя файла исходника в нижнем регистре> <исполнитель в нижнем регистре>"
    """
    artist_lower = artist_name.strip().lower()
    index = {}
    for sub in artist_folder.iterdir():
        if sub.is_dir():
            article = f"{sub.name.strip().lower()} {artist_lower}"
            index[article] = sub
    return index


def collect_photos(source_folder):
    """Возвращает список путей к фото в порядке PHOTO_ORDER (пропуская то, чего нет)."""
    photos = []
    for key in PHOTO_ORDER:
        if key == "5-x":
            for variant in ("5-3", "5-4", "5-5"):
                candidate = source_folder / f"{variant}.jpg"
                if candidate.exists():
                    photos.append(candidate)
                    break
            continue
        candidate = source_folder / f"{key}.jpg"
        if candidate.exists():
            photos.append(candidate)
    return photos


def collect_static_photos():
    """
    Возвращает список Path для общих файлов из <корень>/mockups/ (8.jpg, cover.jpg)
    в порядке STATIC_PHOTOS. Список одинаков для всех карточек, поэтому
    вычисляется один раз в начале запуска. Номер места на карточке им здесь
    не назначается — они встают сразу после "своих" фото исходника (см.
    upload_article), так что реальный номер зависит от того, сколько "своих"
    фото нашлось у конкретного исходника.
    """
    static_photos = []
    for stem in STATIC_PHOTOS:
        candidate = MOCKUPS_FOLDER / f"{stem}.jpg"
        if candidate.exists():
            static_photos.append(candidate)
        else:
            _log(f"!  Не найден общий файл '{candidate}' - будет пропущен для всех карточек.")
    return static_photos


# --- EXCEL ---
def load_articles():
    """Артикулы из столбца 'Артикул продавца' (лист 'Товары', иначе активный)."""
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
            if not article:
                continue
            if any(w in article.lower() for w in STOP_WORDS):
                continue
            articles.append(article)
        return articles
    finally:
        wb.close()


# --- WILDBERRIES API ---
def _check_auth(response):
    if response.status_code == 401:
        raise WBAuthError("x WB вернул 401: токен неверный или просрочен.")
    if response.status_code == 403:
        raise WBAuthError(
            "x WB вернул 403: у токена нет доступа к разделу «Контент». "
            "Выпустите токен с этим правом."
        )


def request_cards(cursor=None):
    url = BASE_URL + "/content/v2/get/cards/list"
    payload = {
        "settings": {
            "sort": {"ascending": False},  # недавно обновлённые (обычно только что залитые) — первыми
            "cursor": cursor or {},
            "filter": {"withPhoto": -1},
        }
    }
    r = requests.post(url, headers=HEADERS, json=payload, timeout=60)
    _check_auth(r)
    r.raise_for_status()
    return r.json()


def load_all_cards(wanted_articles=None):
    """Возвращает словарь {vendorCode: {'nmid': nmid, 'photo_count': количество_фото}}.

    Если передан wanted_articles — пагинация останавливается, как только все они
    найдены, не дожидаясь конца каталога. Карточки запрашиваются в порядке "от
    недавно обновлённых" (см. request_cards), поэтому только что залитые/недавно
    отредактированные карточки обычно оказываются на первых страницах, и в
    типичном случае (10-30 артикулов из 600+) поиск останавливается очень рано.
    Если что-то из wanted_articles не встретится — цикл всё равно корректно
    дойдёт до конца каталога (курсор закончится), просто дольше."""
    result = {}
    cursor = {}
    wanted = set(wanted_articles) if wanted_articles else None
    while True:
        data = request_cards(cursor)
        for card in data.get("cards", []):
            vendor = str(card.get("vendorCode", "")).strip()
            nmid = card.get("nmID")
            if vendor and nmid:
                result[vendor] = {
                    "nmid": nmid,
                    "photo_count": len(card.get("photos", [])),
                }
        if wanted is not None and wanted.issubset(result.keys()):
            break
        cursor = data.get("cursor")
        if not cursor or cursor.get("total", 0) == 0:
            break
        time.sleep(0.3)
    return result


def upload_photo(nmid, photo_number, photo_path):
    url = BASE_URL + "/content/v3/media/file"
    headers = HEADERS.copy()
    headers["X-Nm-Id"] = str(nmid)
    headers["X-Photo-Number"] = str(photo_number)
    for attempt in range(5):
        try:
            with open(photo_path, "rb") as f:
                files = {"uploadfile": (photo_path.name, f, "image/jpeg")}
                r = requests.post(url, headers=headers, files=files, timeout=120)
        except OSError as e:
            _log(f"   не удалось прочитать файл {photo_path.name}: {e}")
            return False
        except requests.RequestException:
            time.sleep(3)
            continue

        _check_auth(r)
        if r.status_code == 200:
            return True
        if r.status_code in (429, 500, 502, 503, 504):
            time.sleep(3)
            continue
        _log(r.text)
        return False
    return False


def upload_article(article, nmid, photos, static_photos, current_photo_count, counter=""):
    """Возвращает список файлов, которые не удалось загрузить (пустой — всё хорошо)."""
    # Свои фото исходника + общие (mockups) идут ОДНИМ сплошным списком, без
    # пропусков в нумерации: если у исходника не хватает какого-то из "своих"
    # фото, общие файлы просто занимают освободившееся место раньше, а не
    # жёстко №8/№9. Относительный порядок (сначала свои, потом общие) при
    # этом сохраняется всегда.
    all_photos = list(photos) + list(static_photos)
    total_count = len(all_photos)
    if not FORCE_REUPLOAD and current_photo_count >= total_count:
        _log(f"{counter} ⏭ {article} : уже есть {current_photo_count} фото, пропускаем загрузку.")
        return []

    _log(f"{counter} {article}")
    failed = []
    for number, photo_path in enumerate(all_photos, start=1):
        ok = upload_photo(nmid, number, photo_path)
        _log(f"   {'✓' if ok else '✗'} {number} ({photo_path.name})")
        if not ok:
            failed.append(f"{article}: {photo_path.name}")
        time.sleep(0.5)
    return failed


def main(
    artist=None,
    log_callback=None,
    progress_callback=None,
    prompt_callback=None,
    config_path=None,
):
    """Загружает только фото на уже существующие карточки WB.

    Возвращает словарь {"success": bool, "error": str|None, ...}."""
    global API_TOKEN, HEADERS
    global LOG_CALLBACK, PROGRESS_CALLBACK, PROMPT_CALLBACK

    LOG_CALLBACK = log_callback
    PROGRESS_CALLBACK = progress_callback
    PROMPT_CALLBACK = prompt_callback

    result = {
        "success": False,
        "artist": None,
        "articles": 0,
        "failed_photos": [],
        "not_found_cards": [],
        "not_found_folders": [],
        "error": None,
    }

    try:
        _progress(0, "Запуск")
        API_TOKEN = load_token(config_path)
        HEADERS = {"Authorization": API_TOKEN}

        if not EXCEL_FILE.exists():
            raise RuntimeError(f"Не найден файл '{EXCEL_FILE}'")

        artist = get_artist_name(artist)
        result["artist"] = artist
        artist_folder = find_artist_output_folder(artist)
        if not artist_folder:
            raise RuntimeError(
                f"Не найдена папка '{OUTPUT_FOLDER / artist}' — "
                "сначала создайте мокапы для этого исполнителя."
            )

        _log("=" * 60)
        _log("WB PHOTO UPLOADER: только фото")
        _log("=" * 60)

        folder_index = build_article_folder_index(artist_folder, artist)
        static_photos = collect_static_photos()

        _progress(5, "Чтение таблицы")
        articles = load_articles()
        result["articles"] = len(articles)
        _log(f"Артикулов в таблице: {len(articles)}")
        if not articles:
            raise RuntimeError("В таблице нет артикулов — сначала заполните Постеры.xlsx.")

        _progress(10, "Поиск карточек WB")
        _log("Поиск карточек WB начиная с новых...")
        wanted_articles = {a for a in articles if "*" not in a}
        cards_data = load_all_cards(wanted_articles=wanted_articles)
        _log(f"Карточек получено: {len(cards_data)}")
        _progress(20, "Загрузка фото")

        total = len(articles)
        for i, article in enumerate(articles, start=1):
            _progress(20 + 80 * (i - 1) / total, f"Фото: {article}")
            counter = f"[{i}/{total}]"

            if "*" in article:
                _log(f"{counter} ⏭ {article} : содержит '*', пропускаем загрузку")
                continue

            if article not in cards_data:
                _log(f"{counter} {article} : карточка не найдена")
                result["not_found_cards"].append(article)
                continue

            source_folder = folder_index.get(article)
            if not source_folder:
                _log(f"{counter} {article} : не найдена папка с фото в '{artist_folder}'")
                result["not_found_folders"].append(article)
                continue

            photos = collect_photos(source_folder)
            if not photos and not static_photos:
                _log(f"{counter} {article} : в '{source_folder}' и '{MOCKUPS_FOLDER}' не найдено ни одного фото")
                result["not_found_folders"].append(article)
                continue

            nmid = cards_data[article]["nmid"]
            photo_count = cards_data[article]["photo_count"]
            result["failed_photos"].extend(
                upload_article(article, nmid, photos, static_photos, photo_count, counter)
            )
            time.sleep(1)

        _log("\n========== ГОТОВО ==========")
        problems = []
        if result["not_found_cards"]:
            _log("\nНе найдены карточки WB:")
            for item in result["not_found_cards"]:
                _log(item)
            problems.append(f"не найдены карточки WB: {len(result['not_found_cards'])}")
        if result["not_found_folders"]:
            _log("\nНе найдены папки с фото:")
            for item in result["not_found_folders"]:
                _log(item)
            problems.append(f"нет папок с фото: {len(result['not_found_folders'])}")
        if result["failed_photos"]:
            _log("\nНе удалось загрузить фото:")
            for item in result["failed_photos"]:
                _log(item)
            problems.append(f"не загрузилось фото: {len(result['failed_photos'])}")

        _progress(100, "Завершено")
        if problems:
            result["error"] = (
                "Загрузка фото завершена с проблемами — " + "; ".join(problems) + ".\n"
                "Списки по артикулам — в логе выше. Запустите этап ещё раз: "
                "карточки, где фото уже достаточно, будут пропущены."
            )
        else:
            result["success"] = True
        return result

    except WBAuthError as e:
        result["error"] = str(e).lstrip("x ").strip()
        _log(f"\nОШИБКА: {result['error']}")
        return result

    except SystemExit as e:
        result["error"] = str(e) or "Операция остановлена."
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
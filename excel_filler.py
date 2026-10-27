# excel_filler v2.2
# G2G Code
#
# v2.2: все значения вынесены из кода. Стандартные лежат в table_defaults.json (рядом со
#   скриптом), свои хранятся в config.json (раздел "table") и задаются в окне «Настройки» →
#   «Заполнение таблицы». Чего нет в config.json, берётся из table_defaults.json.
#
# Ожидаемое расположение:
#   <корень>/Постеры.xlsx
#   <корень>/source/white/*.jpg
#   <корень>/source/dark/*.jpg
#
# Можно запускать:
#   1) напрямую из Terminal: python3 excel_filler.py
#   2) из будущего GUI: main(artist="Имя исполнителя")
#
# ВАЖНО:
# Логика формирования Excel сохранена. Изменён в основном интерфейс
# запуска, чтобы скрипт больше не зависел от input() при запуске из GUI.

import json
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle


# ------------------- РАСПОЛОЖЕНИЕ ПРОЕКТА -------------------

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent

EXCEL_FILE = ROOT_DIR / "Постеры.xlsx"
SOURCE_FOLDER = ROOT_DIR / "source"
WHITE_FOLDER = SOURCE_FOLDER / "white"
DARK_FOLDER = SOURCE_FOLDER / "dark"


# ------------------- НАСТРОЙКИ -------------------

COLUMN_NAME = "Артикул продавца"
IMAGE_EXTENSIONS = (".jpg", ".jpeg")
STOP_WORDS = ("номер", "идентификатор", "название", "артикул")
# Файл-передатчик от mockuper.jsx.
ALBUM_EXPORT_FILE = SCRIPT_DIR / "postery_album_export.json"

CONFIG_FILE = SCRIPT_DIR / "config.json"

# Значения по умолчанию (все постоянные значения столбцов, описание, шаблоны, пороги)
# лежат в table_defaults.json рядом со скриптом — в коде их нет. Свои значения задаются
# в окне «Настройки» → «Заполнение таблицы» и сохраняются в config.json.
DEFAULTS_FILE = SCRIPT_DIR / "table_defaults.json"


# ------------------- СЛУЖЕБНЫЕ ФУНКЦИИ -------------------

def _log(message, log_callback=None):
    """Печатает сообщение и, если скрипт запущен из GUI, передаёт его туда."""
    print(message)
    if log_callback:
        try:
            log_callback(str(message))
        except Exception:
            # Ошибка GUI-логгера не должна ломать формирование Excel.
            pass


def get_default_table_settings():
    """Стандартные значения из table_defaults.json (рядом со скриптом)."""
    if not DEFAULTS_FILE.exists():
        raise FileNotFoundError(
            f"Не найден файл со стандартными значениями: {DEFAULTS_FILE}. "
            "Перезапустите приложение — оно восстановит файлы в папке scripts."
        )
    try:
        data = json.loads(DEFAULTS_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        raise RuntimeError(f"table_defaults.json повреждён: {e}")
    return data


def load_table_settings(config_path=None, log_callback=None):
    """Значения по умолчанию + то, что пользователь изменил в config.json (раздел "table").

    Некорректные значения не ломают заполнение: берётся значение по умолчанию,
    а в лог пишется предупреждение."""
    defaults = get_default_table_settings()
    settings = json.loads(json.dumps(defaults))  # независимая копия
    path = Path(config_path) if config_path else CONFIG_FILE
    if not path.exists():
        return settings

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        block = data.get("table") if isinstance(data, dict) else None
    except Exception as e:
        _log(f"! Не удалось прочитать config.json — параметры таблицы по умолчанию ({e})", log_callback)
        return settings
    if not isinstance(block, dict):
        return settings

    def warn(name):
        _log(f"! В config.json неверное значение «{name}» — использую стандартное.", log_callback)

    custom_static = block.get("static_values")
    if isinstance(custom_static, dict):
        for key, default in defaults["static_values"].items():
            if key not in custom_static:
                continue
            value = custom_static[key]
            is_number = isinstance(default, (int, float)) and not isinstance(default, bool)
            if is_number:
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    settings["static_values"][key] = value
                else:
                    warn(key)
            elif isinstance(value, str):
                settings["static_values"][key] = value
            else:
                warn(key)

    for key in ("description", "poster_title", "size_label"):
        if key in block:
            if isinstance(block[key], str) and (block[key].strip() or key == "size_label"):
                settings[key] = block[key]
            else:
                warn(key)

    if "highlight_color" in block:
        color = str(block["highlight_color"]).lstrip("#").strip()
        if re.fullmatch(r"[0-9A-Fa-f]{6}", color):
            settings["highlight_color"] = color.upper()
        else:
            warn("highlight_color")

    for key in ("max_name_len", "max_article_len"):
        if key in block:
            value = block[key]
            if isinstance(value, int) and not isinstance(value, bool) and value >= 1:
                settings[key] = value
            else:
                warn(key)

    changed = sum(1 for k, v in settings.items() if v != defaults[k])
    if changed:
        _log("Параметры таблицы: есть изменения относительно стандартных (окно «Настройки»).", log_callback)
    return settings


def is_valid_article(value):
    if value is None:
        return False

    s = str(value).strip()
    if not s:
        return False

    lower = s.lower()
    return not any(word in lower for word in STOP_WORDS)


def find_header_row(file_path, column_name):
    wb = load_workbook(file_path, data_only=True)
    try:
        ws = wb.active

        for row_index, row in enumerate(ws.iter_rows(values_only=True)):
            for cell in row:
                if isinstance(cell, str) and column_name.lower() in cell.lower():
                    return row_index
    finally:
        wb.close()

    raise Exception(
        f"Не найдена строка заголовка с колонкой '{column_name}'"
    )


def get_column_index(ws, header_row, column_name):
    for cell in ws[header_row + 1]:
        if cell.value and column_name.lower() in str(cell.value).lower():
            return cell.column

    raise Exception(
        f"Столбец '{column_name}' не найден в строке {header_row + 1}"
    )


def get_image_names():
    """
    Уникальные имена файлов (без расширения)
    из source/white и source/dark вместе.
    """
    names = set()

    for folder in (WHITE_FOLDER, DARK_FOLDER):
        if not folder.exists():
            raise FileNotFoundError(
                f"Папка '{folder}' не найдена"
            )

        for file in folder.iterdir():
            if file.is_file() and file.suffix.lower() in IMAGE_EXTENSIONS:
                names.add(file.stem)

    return sorted(names)


def get_artist_name():
    """
    Оставлено для совместимости с ручным запуском из Terminal.
    GUI будет передавать artist напрямую в main().
    """
    name = input("Введите имя исполнителя: ").strip()
    return name or "Без_имени"


def load_album_export(log_callback=None):
    """
    Читает postery_album_export.json от mockuper.jsx.

    Возвращает:
        (artist, albums_dict)

    Если файла нет или он повреждён:
        (None, {})
    """
    if not ALBUM_EXPORT_FILE.exists():
        return None, {}

    try:
        with open(ALBUM_EXPORT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        artist = data.get("artist")
        albums = data.get("albums", {})

        if not isinstance(albums, dict):
            albums = {}

        return artist, albums

    except Exception as e:
        _log(
            f"! Не удалось прочитать файл названий альбомов "
            f"от Photoshop-скрипта: {e}",
            log_callback,
        )
        return None, {}


def cleanup_album_export(log_callback=None):
    """Удаляет временный файл после успешного использования."""
    try:
        if ALBUM_EXPORT_FILE.exists():
            ALBUM_EXPORT_FILE.unlink()

    except Exception as e:
        _log(
            f"! Не удалось удалить временный файл названий альбомов: {e}",
            log_callback,
        )


def clear_all_rows(ws, header_row, log_callback=None):
    """
    Удаляет все строки данных начиная с header_row + 2.
    Строка непосредственно под заголовком сохраняется.
    """
    for row in range(ws.max_row, header_row + 1, -1):
        if row >= header_row + 2:
            ws.delete_rows(row)

    _log(
        "Таблица очищена (строки с данными удалены, "
        "строка под заголовком сохранена)",
        log_callback,
    )


def rebuild_table(
    excel_path,
    artist_name,
    file_names,
    album_map=None,
    log_callback=None,
    settings=None,
):
    """
    Полностью перестраивает таблицу.

    Основная логика оставлена прежней:
    - очищает старые строки;
    - создаёт строки по уникальным именам файлов;
    - заполняет статические значения;
    - заполняет описание;
    - использует album_map от mockuper.jsx;
    - создаёт формулы;
    - применяет форматирование и условное окрашивание.
    """
    album_map = album_map or {}
    settings = settings or get_default_table_settings()
    static_values = settings["static_values"]

    header_row = find_header_row(excel_path, COLUMN_NAME)

    wb = load_workbook(excel_path)
    ws = wb.active

    clear_all_rows(ws, header_row, log_callback)

    all_columns = list(static_values.keys()) + [
        "Описание",
        "Наименование",
        "Наименование постера",
        COLUMN_NAME,
        "Цвет",
    ]

    col_indices = {
        col: get_column_index(ws, header_row, col)
        for col in all_columns
    }

    # replace, а не format: в тексте описания могут быть любые фигурные скобки
    description = settings["description"].replace("{artist}", artist_name)
    start_row = header_row + 3
    artist_lower = artist_name.lower()

    first_unmatched_index = next(
        (
            i
            for i, f in enumerate(file_names)
            if not album_map.get(f)
        ),
        None,
    )

    for i, file_stem in enumerate(file_names):
        row = start_row + i

        article = f"{file_stem.lower()} {artist_lower}"
        ws.cell(
            row=row,
            column=col_indices[COLUMN_NAME],
            value=article,
        )

        album_name = album_map.get(file_stem)

        def make_title(album):
            return (
                settings["poster_title"]
                .replace("{artist}", artist_name)
                .replace("{album}", album)
            )

        if album_name:
            poster_name = make_title(album_name)
        elif i == first_unmatched_index:
            poster_name = make_title("")
        else:
            poster_name = ""

        ws.cell(
            row=row,
            column=col_indices["Наименование постера"],
            value=poster_name,
        )

        for col_name, val in static_values.items():
            ws.cell(
                row=row,
                column=col_indices[col_name],
                value=val,
            )

        ws.cell(
            row=row,
            column=col_indices["Описание"],
            value=description,
        )

        # Цвет остаётся пустым.

    # Формулы:
    # Наименование = Наименование постера + " А3 " + Цвет
    poster_letter = get_column_letter(
        col_indices["Наименование постера"]
    )
    color_letter = get_column_letter(
        col_indices["Цвет"]
    )
    name_col = col_indices["Наименование"]
    size_label = settings["size_label"].replace('"', '""')  # кавычки внутри формулы Excel

    for row in range(
        start_row,
        start_row + len(file_names),
    ):
        ws.cell(row=row, column=name_col).value = (
            f'={poster_letter}{row}&" {size_label} "&{color_letter}{row}'
        )

    # Серое выделение колонок.
    fill = PatternFill(
        start_color=settings["highlight_color"],
        end_color=settings["highlight_color"],
        fill_type="solid",
    )

    for row in range(
        start_row,
        start_row + len(file_names),
    ):
        for col_name in ("Наименование", "Цвет", "Наименование постера"):
            col_idx = col_indices.get(col_name)

            if col_idx:
                ws.cell(
                    row=row,
                    column=col_idx,
                ).fill = fill

    # Условное форматирование (пороги — из настроек, по умолчанию 60 и 72):
    # Наименование >= max_name_len символов
    # Артикул продавца >= max_article_len символов
    if file_names:
        last_row = start_row + len(file_names) - 1

        for col_idx, max_len in (
            (name_col, settings["max_name_len"]),
            (col_indices[COLUMN_NAME], settings["max_article_len"]),
        ):
            col_letter = get_column_letter(col_idx)

            range_str = (
                f"{col_letter}{start_row}:"
                f"{col_letter}{last_row}"
            )

            red_fill = PatternFill(
                start_color="FF0000",
                end_color="FF0000",
                fill_type="solid",
            )

            dxf = DifferentialStyle(fill=red_fill)

            rule = Rule(
                type="expression",
                dxf=dxf,
                formula=[
                    f"LEN({col_letter}{start_row})>={max_len}"
                ],
            )

            ws.conditional_formatting.add(
                range_str,
                rule,
            )

    wb.save(excel_path)

    _log(
        f"Таблица перестроена: добавлено "
        f"{len(file_names)} уникальных строк, "
        f"начиная со строки {start_row}.",
        log_callback,
    )

    if file_names:
        matched = sum(
            1 for f in file_names
            if album_map.get(f)
        )

        if matched:
            _log(
                f"   'Наименование постера' заполнено автоматически "
                f"для {matched} из {len(file_names)} строк "
                f"(по данным от mockuper.jsx).",
                log_callback,
            )

        if matched < len(file_names):
            missing = len(file_names) - matched

            _log(
                f"   {missing} строк без совпадения по названию "
                f"альбома - заполните вручную "
                f"(шаблон 'Постер {artist_name} - \"\"' "
                f"стоит в первой строке без совпадения).",
                log_callback,
            )


# ------------------- ГЛАВНАЯ ФУНКЦИЯ -------------------

def main(
    artist=None,
    log_callback=None,
    config_path=None,
):
    """
    Основная функция.

    Для GUI:
        result = main(
            artist="Pink Floyd",
            log_callback=print,
        )

    Для Terminal:
        main()

    Возвращает словарь:
        {
            "success": True/False,
            "artist": "...",
            "file": "...",
            "file_count": 10,
            "matched_albums": 10,
            "error": None
        }

    Это позволит будущему GUI понимать результат работы
    без разбора текста print().
    """

    result = {
        "success": False,
        "artist": None,
        "file": str(EXCEL_FILE),
        "file_count": 0,
        "matched_albums": 0,
        "error": None,
    }

    _log("=" * 60, log_callback)
    _log(
        "ЗАПУСК ПОЛНОЙ ПЕРЕЗАПИСИ EXCEL-ФАЙЛА "
        "(только уникальные имена)",
        log_callback,
    )
    _log("=" * 60, log_callback)

    try:
        # Проверяем Excel.
        if not EXCEL_FILE.exists():
            message = f"Не найден файл '{EXCEL_FILE}'"
            _log(f"x {message}", log_callback)

            result["error"] = message
            return result

        # Сначала пытаемся получить artist + album_map
        # от mockuper.jsx.
        export_artist, album_map = load_album_export(
            log_callback=log_callback
        )

        if export_artist:
            artist = export_artist

            _log(
                f"Найден файл с названиями альбомов от mockuper.jsx "
                f"(исполнитель: {artist}, "
                f"альбомов: {len(album_map)}) - "
                f"использую их автоматически.",
                log_callback,
            )

        # Если GUI передал artist — используем его.
        # Если ничего не передали — сохраняем старое поведение
        # через input().
        elif artist is None:
            artist = get_artist_name()

        artist = str(artist).strip() or "Без_имени"
        result["artist"] = artist

        _log(
            f"\nПоиск изображений в: "
            f"{WHITE_FOLDER} и {DARK_FOLDER}",
            log_callback,
        )

        try:
            file_names = get_image_names()

        except FileNotFoundError as e:
            message = str(e)
            _log(f"x {message}", log_callback)

            result["error"] = message
            return result

        result["file_count"] = len(file_names)

        _log(
            f"   Найдено уникальных файлов: "
            f"{len(file_names)}",
            log_callback,
        )

        # Нет исходников — очищаем таблицу.
        if not file_names:
            _log(
                "! Не найдено ни одного файла. "
                "Таблица будет очищена "
                "(сохранена строка под заголовком).",
                log_callback,
            )

            header_row = find_header_row(
                EXCEL_FILE,
                COLUMN_NAME,
            )

            wb = load_workbook(EXCEL_FILE)

            try:
                clear_all_rows(
                    wb.active,
                    header_row,
                    log_callback,
                )
                wb.save(EXCEL_FILE)

            finally:
                wb.close()

            if export_artist:
                cleanup_album_export(
                    log_callback
                )

            _log(
                "Таблица очищена.",
                log_callback,
            )

            result["success"] = True
            return result

        _log(
            f"\nПерестройка таблицы в {EXCEL_FILE}...",
            log_callback,
        )

        settings = load_table_settings(config_path, log_callback)

        rebuild_table(
            EXCEL_FILE,
            artist,
            file_names,
            album_map,
            log_callback=log_callback,
            settings=settings,
        )

        result["matched_albums"] = sum(
            1 for f in file_names
            if album_map.get(f)
        )

        # Удаляем JSON только после успешной перестройки Excel.
        if export_artist:
            cleanup_album_export(
                log_callback
            )

        result["success"] = True

        _log("\n" + "=" * 60, log_callback)
        _log(
            "ГОТОВО! Excel-файл полностью обновлён.",
            log_callback,
        )
        _log("=" * 60, log_callback)

        return result

    except Exception as e:
        # Не удаляем postery_album_export.json при неожиданной
        # ошибке — это позволит повторить запуск и не потерять
        # данные, пришедшие из Photoshop.
        result["error"] = str(e)

        _log(
            f"\nОШИБКА: {e}",
            log_callback,
        )

        return result


# ------------------- ЗАПУСК ИЗ TERMINAL -------------------

if __name__ == "__main__":
    result = main()

    # Для ручного запуска оставляем понятный финальный статус.
    if result["success"]:
        print("\nРезультат: УСПЕШНО")
    else:
        print(
            f"\nРезультат: ОШИБКА — "
            f"{result['error']}"
        )
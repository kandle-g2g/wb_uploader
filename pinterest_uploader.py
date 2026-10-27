# pinterest_uploader v1.5
# G2G Code
#
# v1.5: в строке загрузки убрана подпись папки (white/dark) — только [N/всего] и имя файла.
#
# v1.4: логика входа в Pinterest и создания доски возвращена к рабочей v1.1:
#   - pinterest.login() вызывается как раньше, без проверки сессии и без
#     очистки кэша pinterest_creds/ (в v1.3 это ломало рабочую сессию);
#   - папка pinterest_creds/ заранее не создаётся — её создаёт py3pin;
#   - при 401 на создании доски выводится понятное сообщение, а не повтор.
#   Всё остальное (интеграция с приложением из v1.2) сохранено.
#
# v1.2: скрипт встроен в Posters.app (блок «Дополнительно»):
#   - данные аккаунта (email, username, password) больше не хранятся в коде —
#     берутся из config.json (папка scripts/), раздел "pinterest":
#       {"pinterest": {"email": "...", "username": "...", "password": "..."}}
#     Если чего-то не хватает, приложение спрашивает это перед запуском;
#   - картинки берутся из <корень>/source/white и <корень>/source/dark
#     (jpg/jpeg/png) — теми же файлами, с которых делаются мокапы;
#   - исполнитель приходит из приложения (main(artist=...)), доска всегда
#     называется "<исполнитель> Albums";
#   - временные JPG для загрузки создаются во временной папке системы,
#     а не рядом с исходниками (раньше при сбое там мог остаться мусор);
#   - вывод идёт и в консоль, и в лог приложения; в конце возвращается словарь
#     {"success": ..., "error": ...}, как у остальных скриптов.
#   Запуск из терминала по-прежнему работает: python3 pinterest_uploader.py
#
# Ожидаемое расположение (скрипт лежит в <корень>/scripts/):
#   scripts/config.json        — данные аккаунта Pinterest
#   scripts/pinterest_creds/   — кэш сессии py3pin (создаётся автоматически)
#   <корень>/source/white, <корень>/source/dark — изображения для загрузки

import json
import tempfile
import time
from pathlib import Path

from py3pin.Pinterest import Pinterest
from PIL import Image

# --- Расположение проекта ---
SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent
CONFIG_FILE = SCRIPT_DIR / "config.json"
CRED_ROOT = SCRIPT_DIR / "pinterest_creds"
SOURCE_FOLDER = ROOT_DIR / "source"
SOURCE_SUBFOLDERS = ("white", "dark")  # "Other Albums" не берём — это не постеры исполнителя

IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png")
WAIT_BETWEEN_PINS = 20  # секунд, чтобы не словить рейт-лимит
BOARD_SEARCH_RETRIES = 3
BOARD_SEARCH_WAIT = 10  # секунд

# Заполняется в main() из config.json
CREDS = {}

# --- Связь с приложением ---
LOG_CALLBACK = None
PROGRESS_CALLBACK = None
PROMPT_CALLBACK = None


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


def load_credentials(config_path=None):
    """Читает данные Pinterest из config.json (раздел "pinterest")."""
    path = Path(config_path) if config_path else CONFIG_FILE
    if not path.exists():
        raise RuntimeError(f"Не найден config.json: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать config.json: {e}")

    block = data.get("pinterest") or {}
    creds = {key: str(block.get(key) or "").strip() for key in ("email", "username", "password")}
    missing = [f"pinterest.{key}" for key, value in creds.items() if not value]
    if missing:
        raise RuntimeError("В config.json не заполнено: " + ", ".join(missing))
    return creds


def get_pinterest():
    # Как в v1.1: папку pinterest_creds создаёт сам py3pin
    return Pinterest(
        email=CREDS["email"],
        password=CREDS["password"],
        username=CREDS["username"],
        cred_root=str(CRED_ROOT),
    )


def login(pinterest):
    """Вход как в v1.1: просто pinterest.login(), только с понятными подсказками."""
    try:
        pinterest.login()
    except Exception as e:
        text = str(e)
        hint = (
            "Не удалось войти в Pinterest. Проверьте почту, имя пользователя и пароль "
            "(кнопка «Настройки» в шапке окна) и откройте Pinterest в браузере — возможно, "
            "аккаунт просит подтвердить вход."
        )
        low = text.lower()
        if "chrome" in low or "driver" in low or "selenium" in low:
            hint = (
                "Не удалось запустить браузер для входа в Pinterest. Установите Google Chrome "
                "и проверьте интернет (при первом запуске подтягивается драйвер браузера)."
            )
        raise RuntimeError(f"{hint}\n({type(e).__name__}: {text[:300]})")
    return pinterest


def get_board_id_by_name(pinterest_instance, board_name):
    boards = pinterest_instance.boards(username=CREDS["username"])
    target = board_name.strip().lower()
    for board in boards:
        if board.get("name", "").strip().lower() == target:
            return board.get("id")
    return None


def convert_to_jpg(input_path: Path, output_path: Path) -> Path:
    with Image.open(input_path) as img:
        if img.mode in ("RGBA", "LA", "P"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "RGBA":
                background.paste(img, mask=img.split()[3])
            else:
                background.paste(img)
            img = background
        elif img.mode != "RGB":
            img = img.convert("RGB")
        img.save(output_path, "JPEG", quality=95)
    return output_path


def find_or_create_board(pinterest, artist_name):
    board_name = f"{artist_name} Albums"
    _log(f"Будет использована доска: '{board_name}'")

    board_id = get_board_id_by_name(pinterest, board_name)
    if board_id is not None:
        _log(f"Найдена доска '{board_name}' с ID: {board_id}")
        return pinterest, board_id

    _log(f"Доска '{board_name}' не найдена. Создаём новую...")
    try:
        response = pinterest.create_board(name=board_name, description=f"{artist_name} albums posters")
    except Exception as e:
        if "401" in str(e) or "unauthorized" in str(e).lower():
            raise RuntimeError(
                "Pinterest не разрешил создать доску (401 Unauthorized): сессия есть, но она "
                "не принята для записи. Проверьте, что email/username/password в config.json "
                "те же, с которыми скрипт работал раньше (кэш сессии хранится по email), "
                "и откройте Pinterest в браузере — возможно, нужно подтвердить вход.\n"
                f"({str(e)[:300]})"
            )
        raise

    board_id = None
    try:
        data = response.json()
        board_id = data.get("id") or data.get("board", {}).get("id") or data.get("data", {}).get("id")
        if board_id:
            _log(f"Доска создана, ID из ответа: {board_id}")
    except Exception as e:
        _log(f"Не удалось разобрать ответ: {e}")

    if board_id:
        return pinterest, board_id

    # Ответ не дал ID напрямую — пробуем найти доску заново со свежей сессией
    _log("Не удалось извлечь ID. Пытаемся найти доску с обновлённой сессией...")
    for attempt in range(1, BOARD_SEARCH_RETRIES + 1):
        _log(f"Попытка {attempt}/{BOARD_SEARCH_RETRIES}: ждём {BOARD_SEARCH_WAIT} секунд...")
        time.sleep(BOARD_SEARCH_WAIT)
        fresh_pinterest = get_pinterest()
        _log("Вход в аккаунт...")
        login(fresh_pinterest)
        board_id = get_board_id_by_name(fresh_pinterest, board_name)
        if board_id:
            _log(f"Доска найдена после попытки {attempt}, ID: {board_id}")
            return fresh_pinterest, board_id

    raise RuntimeError(
        f"Не удалось найти доску '{board_name}' после {BOARD_SEARCH_RETRIES} попыток. "
        "Убедитесь, что доска создана, и запустите снова."
    )


def collect_images():
    """Изображения из source/white и source/dark (без скрытых файлов macOS)."""
    if not SOURCE_FOLDER.is_dir():
        raise RuntimeError(f"Не найдена папка '{SOURCE_FOLDER}'.")

    files = []
    for name in SOURCE_SUBFOLDERS:
        folder = SOURCE_FOLDER / name
        if not folder.is_dir():
            _log(f"!  Не найдена папка '{folder}' — пропускаем.")
            continue
        files.extend(sorted(
            f for f in folder.iterdir()
            if f.is_file() and not f.name.startswith(".") and f.suffix.lower() in IMAGE_EXTENSIONS
        ))
    return files


def upload_images(pinterest, board_id, artist_name, image_files, failed):
    total = len(image_files)
    uploaded = 0
    _log(f"Найдено {total} изображений. Начинаем загрузку...")

    with tempfile.TemporaryDirectory(prefix="posters_pinterest_") as tmp:
        tmp_dir = Path(tmp)
        for idx, file_path in enumerate(image_files, start=1):
            _progress(20 + 80 * (idx - 1) / total, f"Pinterest: {file_path.stem}")
            title = file_path.stem
            description = f"{artist_name} - {title} poster"

            try:
                converted_path = convert_to_jpg(file_path, tmp_dir / f"{idx}.jpg")
                _log(f"[{idx}/{total}] {file_path.name}")
                pinterest.upload_pin(
                    board_id=board_id,
                    image_file=str(converted_path),
                    title=title,
                    description=description,
                )
                uploaded += 1
                _log("✓")
            except Exception as e:
                _log(f"  (!) Ошибка при загрузке {file_path.name}: {e}")
                failed.append(file_path.name)

            if idx < total:
                _log(f"Ждём {WAIT_BETWEEN_PINS} секунд...")
                time.sleep(WAIT_BETWEEN_PINS)

    return uploaded


def main(
    artist=None,
    log_callback=None,
    progress_callback=None,
    prompt_callback=None,
    config_path=None,
):
    """Загружает изображения из source/white и source/dark на доску "<исполнитель> Albums".

    Возвращает словарь {"success": bool, "error": str|None, ...}."""
    global CREDS, LOG_CALLBACK, PROGRESS_CALLBACK, PROMPT_CALLBACK

    LOG_CALLBACK = log_callback
    PROGRESS_CALLBACK = progress_callback
    PROMPT_CALLBACK = prompt_callback

    result = {
        "success": False,
        "artist": None,
        "uploaded": 0,
        "failed": [],
        "error": None,
    }

    try:
        _progress(0, "Запуск")
        CREDS = load_credentials(config_path)

        name = (artist or "").strip()
        if not name and PROMPT_CALLBACK:
            name = (PROMPT_CALLBACK("Введите имя исполнителя:", None) or "").strip()
        if not name and not PROMPT_CALLBACK:
            name = input("Введите имя исполнителя: ").strip()
        if not name:
            raise RuntimeError("Имя исполнителя не введено — остановка.")
        result["artist"] = name

        image_files = collect_images()
        if not image_files:
            raise RuntimeError(
                f"В '{SOURCE_FOLDER}' (подпапки white и dark) нет изображений jpg/jpeg/png. "
                "Если там только PSD или TIFF — сначала сохраните их в JPG "
                "(блок «Дополнительно» → TIFF и JPG)."
            )

        _log("=" * 60)
        _log("PINTEREST UPLOADER")
        _log("=" * 60)

        _progress(5, "Вход в Pinterest")
        _log("Вход в аккаунт...")
        pinterest = get_pinterest()
        login(pinterest)

        _progress(15, "Поиск доски")
        pinterest, board_id = find_or_create_board(pinterest, name)

        result["uploaded"] = upload_images(pinterest, board_id, name, image_files, result["failed"])

        _log("\n--- Загрузка на Pinterest завершена ---\n")
        _progress(100, "Завершено")
        if result["failed"]:
            _log("Не загружены:")
            for item in result["failed"]:
                _log(item)
            result["error"] = (
                f"Загружено {result['uploaded']} из {len(image_files)}; "
                f"не удалось загрузить: {len(result['failed'])}.\n"
                "Список файлов — в логе выше. Повторный запуск загрузит ВСЕ изображения "
                "заново и создаст дубли уже загруженных пинов."
            )
        else:
            result["success"] = True
        return result

    except SystemExit as e:
        result["error"] = str(e) or "Операция остановлена."
        _log(f"\nОШИБКА: {result['error']}")
        return result

    except Exception as e:
        result["error"] = str(e)
        _log(f"\nОШИБКА: {e}")
        return result


if __name__ == "__main__":
    main()
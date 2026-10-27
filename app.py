#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import datetime
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Posters"
CONFIG_NAME = "config.json"
EXCEL_NAME = "Постеры.xlsx"


# Настройки самого приложения (не путать с config.json проекта): запоминаем
# вручную выбранные Mockups и Photoshop, чтобы не спрашивать каждый раз.
SETTINGS_FILE = Path.home() / "Library" / "Application Support" / "Posters" / "settings.json"

# Скрипты упакованы внутри .app (см. Posters.spec, datas) и при запуске
# копируются в <Mockups>/scripts — именно там их ожидают mockuper.jsx и др.
BUNDLED_DIR_NAME = "bundled_scripts"
SCRIPT_SUFFIXES = (".py", ".jsx", ".json")

ERROR_LOG = SETTINGS_FILE.parent / "error.log"

# --- Автообновление через GitHub Releases ---
# Две строки ниже при сборке на GitHub подставляются автоматически (см. build.yml):
APP_VERSION = "0.0.0"  # версия = имя релиза на GitHub, например v1.2.0
GITHUB_REPO = ""  # "логин/репозиторий"; пусто — проверка обновлений выключена
UPDATE_TIMEOUT = 10  # секунд на запрос к GitHub

# Запускается отдельным процессом: ждёт закрытия приложения, подменяет .app и открывает заново.
# Если замена не удалась, возвращает старую версию.
_SWAP_SCRIPT = """#!/bin/sh
PID="$1"; NEW="$2"; APP="$3"; TMP="$4"
while kill -0 "$PID" 2>/dev/null; do sleep 0.5; done
BACKUP="$APP.old"
rm -rf "$BACKUP"
if mv "$APP" "$BACKUP"; then
  if ditto "$NEW" "$APP"; then
    xattr -cr "$APP" 2>/dev/null
    rm -rf "$BACKUP"
  else
    rm -rf "$APP"
    mv "$BACKUP" "$APP"
  fi
fi
open "$APP"
case "$TMP" in */posters_update_*) rm -rf "$TMP" ;; esac
"""


def _version_tuple(text):
    nums = [int(n) for n in re.findall(r"\d+", str(text))[:4]]
    return tuple(nums + [0] * (4 - len(nums)))


def _update_asset_name():
    import platform
    label = "AppleSilicon" if platform.machine() == "arm64" else "Intel"
    return f"Posters-{label}.zip"


def _fetch_latest_release():
    import requests
    r = requests.get(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "User-Agent": "Posters-updater"},
        timeout=UPDATE_TIMEOUT,
    )
    r.raise_for_status()
    return r.json()


def _app_bundle_path():
    """Путь к запущенному Posters.app; None, если запуск из исходников."""
    if not getattr(sys, "frozen", False):
        return None
    for parent in Path(sys.executable).resolve().parents:
        if parent.suffix == ".app":
            return parent
    return None


class UserError(Exception):
    """Понятная человеку проблема (не заполнена таблица, нет прав, не найден Photoshop…).
    Показывается обычным текстом, без технического traceback."""

    def __init__(self, message, title="Ошибка"):
        super().__init__(message)
        self.title = title


def _strip_marker(text):
    """Убирает служебные префиксы скриптов вроде 'x ' и '✗ '."""
    return re.sub(r"^[x✗]\s+", "", str(text).strip())


def _explain_error(exc):
    """Превращает исключение в (заголовок, понятный текст, непредвиденная_ли_ошибка)."""
    if isinstance(exc, UserError):
        return exc.title, str(exc), False

    if isinstance(exc, SystemExit):
        text = _strip_marker(exc.code if exc.code is not None else "")
        if not text or text.isdigit():
            text = "Операция остановлена. Подробности — в логе."
        return "Операция остановлена", text, False

    if isinstance(exc, PermissionError):
        target = f"\n{exc.filename}" if getattr(exc, "filename", None) else ""
        return (
            "Нет доступа",
            "Нет доступа к файлу или папке." + target + "\n\n"
            "Если это Постеры.xlsx — закройте его в Excel. Если файл в Документах — "
            "разрешите Posters доступ: Системные настройки → Конфиденциальность и "
            "безопасность → Файлы и папки (или Полный доступ к диску).",
            False,
        )

    if isinstance(exc, FileNotFoundError):
        filename = getattr(exc, "filename", None)
        text = f"Не найден файл или папка:\n{filename}" if filename else str(exc)
        return "Файл не найден", text, False

    if isinstance(exc, ModuleNotFoundError):
        return (
            "Приложение собрано неполностью",
            f"Внутри приложения не хватает модуля «{exc.name}». "
            "Пересоберите приложение (build_mac.sh) или напишите разработчику.",
            True,
        )

    if isinstance(exc, json.JSONDecodeError):
        return "Повреждённый файл", f"Не удалось прочитать JSON-файл: {exc}", False

    try:
        import requests
        if isinstance(exc, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
            return (
                "Нет связи с сервером",
                "Не удалось связаться с сервером. Проверьте интернет (и VPN, если он включён) "
                "и запустите этап снова.",
                False,
            )
    except ImportError:
        pass

    return (
        "Непредвиденная ошибка",
        f"{type(exc).__name__}: {exc}\n\n"
        "Подробности — в логе окна (кнопка «Копировать лог») и в файле:\n"
        f"{ERROR_LOG}",
        True,
    )


def _write_error_log(context, exc, tb=None):
    """Дописывает ошибку в error.log (для разбора, если человек пришлёт файл)."""
    try:
        ERROR_LOG.parent.mkdir(parents=True, exist_ok=True)
        if ERROR_LOG.exists() and ERROR_LOG.stat().st_size > 1_000_000:
            ERROR_LOG.write_text("", encoding="utf-8")
        trace = "".join(
            traceback.format_exception(type(exc), exc, tb or exc.__traceback__)
        )
        with ERROR_LOG.open("a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {context}\n{trace}\n")
    except Exception:
        pass


def _load_settings():
    try:
        with SETTINGS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_settings(**updates):
    data = _load_settings()
    data.update(updates)
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        with SETTINGS_FILE.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write("\n")
    except OSError:
        pass


def _bundled_scripts_dir():
    """Где лежат скрипты, упакованные в приложение (или scripts/ при запуске из исходников)."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        return base / BUNDLED_DIR_NAME
    return Path(__file__).resolve().parent / "scripts"


def _is_project_root(path):
    """Папка Mockups: внутри обязательно mockups/ (PSD-шаблоны) и что-то из
    source/, scripts/, Постеры.xlsx."""
    try:
        p = Path(path)
        return (p / "mockups").is_dir() and (
            (p / "source").is_dir()
            or (p / "scripts").is_dir()
            or (p / EXCEL_NAME).exists()
        )
    except OSError:
        return False


def _candidate_project_dirs():
    """
    Поиск идёт от расположения самого приложения. Штатная раскладка:
        Posters_APP/
        ├── Posters.app
        ├── Albums/
        └── Mockups/

    При отладке исполняемый файл может лежать в Mockups/dist/Posters.app/...,
    поэтому проверяются и родители dist, а также папка скрипта и cwd.
    """
    candidates = []

    if getattr(sys, "frozen", False):
        exe = Path(sys.executable).resolve()
        # .../Posters_APP/Posters.app/Contents/MacOS/Posters
        for i in (3, 4, 5):
            if len(exe.parents) > i:
                candidates.append(exe.parents[i])

    candidates.extend([
        Path(__file__).resolve().parent,
        Path.cwd().resolve(),
    ])

    seen = set()
    for p in candidates:
        if p in seen:
            continue
        seen.add(p)
        yield p


def _find_project_root():
    # 1) ранее выбранная вручную папка (если она ещё существует и подходит)
    saved = _load_settings().get("project_root")
    if saved and _is_project_root(saved):
        return Path(saved)

    # 2) рядом с приложением: <папка приложения>/Mockups, затем сама папка
    errors = []
    for base in _candidate_project_dirs():
        for p in (base / "Mockups", base):
            if _is_project_root(p):
                return p
        errors.append(str(base))

    raise FileNotFoundError(
        "Не найдена папка Mockups рядом с приложением.\n\n"
        "Ожидается:\n"
        "Posters_APP/\n"
        "├── Posters.app\n"
        "└── Mockups/\n"
        "    ├── mockups/   (PSD-шаблоны)\n"
        "    ├── source/\n"
        "    └── Постеры.xlsx\n\n"
        "Проверены:\n" + "\n".join(errors)
    )


def _find_photoshop():
    """Ищет установленный Photoshop. Имя у него с годом (например,
    "Adobe Photoshop 2026"), поэтому "Adobe Photoshop" без года не подходит.
    Сначала проверяется вручную выбранный ранее путь."""
    import glob

    saved = _load_settings().get("photoshop_app")
    if saved and Path(saved).is_dir():
        return saved

    patterns = [
        "/Applications/Adobe Photoshop*/Adobe Photoshop*.app",
        "/Applications/Adobe Photoshop*.app",
        str(Path.home() / "Applications/Adobe Photoshop*/Adobe Photoshop*.app"),
    ]
    found = []
    for pattern in patterns:
        found.extend(glob.glob(pattern))
    # самая свежая версия — последняя по имени
    found = sorted(set(found))
    return found[-1] if found else None


def _fonts_dir():
    """Папка fonts/ с упакованными шрифтами (в .app или рядом с app.py при запуске из исходников)."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        return base / "fonts"
    return Path(__file__).resolve().parent / "fonts"


def _register_bundled_fonts():
    """Регистрирует шрифты из fonts/ в macOS только на время работы приложения
    (в систему они не устанавливаются). Нужно вызвать ДО создания окна Tk.
    Если что-то не вышло — молча продолжаем: заголовок покажется запасным шрифтом."""
    if sys.platform != "darwin":
        return
    folder = _fonts_dir()
    if not folder.is_dir():
        return
    try:
        import ctypes
        import ctypes.util

        cf = ctypes.CDLL(ctypes.util.find_library("CoreFoundation"))
        ct = ctypes.CDLL(ctypes.util.find_library("CoreText"))

        cf.CFURLCreateFromFileSystemRepresentation.restype = ctypes.c_void_p
        cf.CFURLCreateFromFileSystemRepresentation.argtypes = [
            ctypes.c_void_p, ctypes.c_char_p, ctypes.c_long, ctypes.c_bool]
        cf.CFRelease.argtypes = [ctypes.c_void_p]
        ct.CTFontManagerRegisterFontsForURL.restype = ctypes.c_bool
        ct.CTFontManagerRegisterFontsForURL.argtypes = [
            ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p]

        K_SCOPE_PROCESS = 1  # kCTFontManagerScopeProcess
        for font in sorted(folder.iterdir()):
            if font.suffix.lower() not in (".otf", ".ttf"):
                continue
            raw = str(font).encode("utf-8")
            url = cf.CFURLCreateFromFileSystemRepresentation(None, raw, len(raw), False)
            if url:
                ct.CTFontManagerRegisterFontsForURL(url, K_SCOPE_PROCESS, None)
                cf.CFRelease(url)
    except Exception:
        pass


def _title_font(root, size):
    """Шрифт названия: Actay Wide, если он установлен в macOS, иначе Helvetica (жирный).
    Файл шрифта в приложение не вшивается — он коммерческий (см. лицензию)."""
    try:
        import tkinter.font as tkfont
        families = list(tkfont.families(root))
    except Exception:
        families = []
    if "Actay Wide" in families:
        return ("Actay Wide", size)
    similar = sorted(
        (f for f in families if "actay" in f.lower() and "wide" in f.lower()),
        key=lambda f: ("bold" not in f.lower(), len(f)),
    )
    if similar:
        return (similar[0], size)
    return ("Helvetica", size, "bold")


# Сочетания клавиш буфера обмена. В Tk на macOS Cmd+C/V не работают при русской раскладке,
# а Ctrl+C/V не привязаны вовсе, поэтому обрабатываем оба модификатора сами:
# по символу (латиница и кириллица) и по коду клавиши macOS.
_CLIP_KEYSYMS = {
    "<<Copy>>": {"c", "с"},
    "<<Paste>>": {"v", "м"},
    "<<Cut>>": {"x", "ч"},
    "select_all": {"a", "ф"},
}
_CLIP_MAC_KEYCODES = {8: "<<Copy>>", 9: "<<Paste>>", 7: "<<Cut>>", 0: "select_all"}


def enable_clipboard_shortcuts(widget):
    """Включает Cmd/Ctrl + C, V, X, A в поле ввода при любой раскладке."""
    def handler(event):
        action = None
        keysym = (event.keysym or "").lower()
        for name, keys in _CLIP_KEYSYMS.items():
            if keysym in keys:
                action = name
                break
        if action is None:
            action = _CLIP_MAC_KEYCODES.get(event.keycode)
        if action is None:
            return None
        if action == "select_all":
            if isinstance(widget, tk.Text):
                widget.tag_add("sel", "1.0", "end-1c")
                widget.mark_set("insert", "end-1c")
            else:
                widget.select_range(0, "end")
                widget.icursor("end")
        else:
            widget.event_generate(action)
        return "break"

    widget.bind("<Command-KeyPress>", handler)
    widget.bind("<Control-KeyPress>", handler)
    _attach_context_menu(widget, editable=True)
    return widget


def _attach_context_menu(widget, editable):
    """Меню по правой кнопке мыши: вырезать / копировать / вставить / выделить всё.
    Для лога (editable=False) — только копировать и выделить всё."""
    menu = tk.Menu(widget, tearoff=0)

    def select_all():
        if isinstance(widget, tk.Text):
            widget.tag_add("sel", "1.0", "end-1c")
        else:
            widget.select_range(0, "end")
            widget.icursor("end")

    def has_selection():
        try:
            if isinstance(widget, tk.Text):
                return bool(widget.tag_ranges("sel"))
            return widget.selection_present()
        except tk.TclError:
            return False

    def do(action):
        widget.focus_set()
        widget.event_generate(action)

    if editable:
        menu.add_command(label="Вырезать", command=lambda: do("<<Cut>>"))
    menu.add_command(label="Копировать", command=lambda: do("<<Copy>>"))
    if editable:
        menu.add_command(label="Вставить", command=lambda: do("<<Paste>>"))
    menu.add_separator()
    menu.add_command(label="Выделить всё", command=lambda: (widget.focus_set(), select_all()))

    def show(event):
        widget.focus_set()
        selected = has_selection()
        state = "normal" if selected else "disabled"
        menu.entryconfigure("Копировать", state=state)
        if editable:
            menu.entryconfigure("Вырезать", state=state)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    # macOS: правая кнопка = Button-2, а также Ctrl+клик; Windows/Linux: Button-3
    for seq in ("<Button-2>", "<Button-3>", "<Control-Button-1>"):
        widget.bind(seq, show)


class BigButton(tk.Canvas):
    """Высокая кнопка для главных этапов.

    У обычной ttk-кнопки на macOS корпус нельзя сделать выше (отступы лишь
    сдвигают текст), поэтому кнопка рисуется сама. Для остального кода она
    ведёт себя как обычная: command=..., configure(state="normal"/"disabled")."""

    COLORS = {  # состояние: (заливка, цвет текста)
        "normal": ("#636363", "#FFFFFF"),
        "hover": ("#989898", "#FFFFFF"),
        "pressed": ("#484848", "#FFFFFF"),
        "disabled": ("#989898", "#D4D4D4"),
    }

    def __init__(self, parent, text, command, height=56, radius=12,
                 font=("Helvetica", 14, "bold")):
        try:
            bg = ttk.Style().lookup("TFrame", "background") or ttk.Style().lookup("TLabelframe", "background")
            super().__init__(parent, height=height, highlightthickness=0, bd=0, bg=bg)
        except tk.TclError:
            super().__init__(parent, height=height, highlightthickness=0, bd=0)
        self._text = text
        self._command = command
        self._radius = radius
        self._font = font
        self._state = "normal"
        self._hover = False
        self._pressed = False
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._set_flags(hover=True))
        self.bind("<Leave>", lambda e: self._set_flags(hover=False, pressed=False))
        self.bind("<ButtonPress-1>", lambda e: self._set_flags(pressed=True))
        self.bind("<ButtonRelease-1>", self._on_release)

    def configure(self, cnf=None, **kw):
        if "state" in kw:
            self._state = kw.pop("state")
            try:
                super().configure(cursor="" if self._state == "disabled" else "pointinghand")
            except tk.TclError:
                pass
            self._draw()
        if cnf or kw:
            return super().configure(cnf, **kw)

    config = configure

    def _set_flags(self, hover=None, pressed=None):
        if hover is not None:
            self._hover = hover
        if pressed is not None:
            self._pressed = pressed
        self._draw()

    def _on_release(self, event):
        was_pressed = self._pressed
        self._set_flags(pressed=False)
        inside = 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height()
        if was_pressed and inside and self._state != "disabled" and self._command:
            self._command()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        if self._state == "disabled":
            key = "disabled"
        elif self._pressed:
            key = "pressed"
        elif self._hover:
            key = "hover"
        else:
            key = "normal"
        fill, fg = self.COLORS[key]
        r = min(self._radius, h // 2)
        x1, y1, x2, y2 = 1, 1, w - 1, h - 1
        pts = [x1 + r, y1, x1 + r, y1, x2 - r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
               x2, y1 + r, x2, y2 - r, x2, y2 - r, x2, y2, x2 - r, y2, x2 - r, y2,
               x1 + r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y2 - r, x1, y1 + r,
               x1, y1 + r, x1, y1]
        self.create_polygon(pts, smooth=True, fill=fill, outline=fill)
        self.create_text(w / 2, h / 2, text=self._text, fill=fg, font=self._font)


class ScrollableFrame(ttk.Frame):
    """Прокручиваемая область: виджеты кладут в .inner.
    Высота — по содержимому, но не больше max_height (дальше появляется прокрутка)."""

    def __init__(self, parent, max_height=600, width=None):
        super().__init__(parent)
        self._max_height = max_height
        try:
            bg = ttk.Style().lookup("TFrame", "background")
            self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=bg) if bg else tk.Canvas(
                self, highlightthickness=0, bd=0
            )
        except tk.TclError:
            self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        if width:
            self.canvas.configure(width=width)
        self._unit = 4  # пикселей за один «шаг» прокрутки — мелкий шаг даёт плавное движение
        self.canvas.configure(yscrollincrement=self._unit)
        self._last_region = None
        self._last_height = None
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.refresh())
        self.canvas.bind(
            "<Configure>", lambda e: self.canvas.itemconfigure(self._window, width=e.width)
        )
        # колесо мыши ловим на уровне окна, иначе оно не работает над вложенными виджетами
        parent.winfo_toplevel().bind("<MouseWheel>", self._wheel, add="+")

    def refresh(self):
        # меняем область и высоту только если они правда изменились — иначе
        # <Configure> запускает сам себя, и содержимое «дрожит»
        region = self.canvas.bbox("all")
        height = min(self.inner.winfo_reqheight(), self._max_height)
        if region != self._last_region:
            self._last_region = region
            self.canvas.configure(scrollregion=region)
        if height != self._last_height:
            self._last_height = height
            self.canvas.configure(height=height)

    def _wheel(self, event):
        if isinstance(event.widget, tk.Text):
            return  # многострочное поле прокручивается само, вторая прокрутка даёт рывки
        delta = event.delta
        if not delta:
            return
        first, last = self.canvas.yview()
        if first <= 0.001 and last >= 0.999:
            return  # всё помещается — прокручивать нечего
        # на краях ничего не делаем: лишние события дёргают картинку и ползунок
        if (delta > 0 and first <= 0.001) or (delta < 0 and last >= 0.999):
            return
        pixels = delta * 48 / 120 if abs(delta) >= 120 else delta * 6  # колесо Windows / macOS
        pixels = max(-120, min(120, pixels))
        units = int(-pixels / self._unit) or (-1 if delta > 0 else 1)
        self.canvas.yview_scroll(units, "units")


SETTINGS_CONTENT_WIDTH = 1040  # ширина содержимого окна «Настройки», px (было ≈ 520)

STEP_BUTTON_HEIGHT = 84  # высота кнопок этапов в пикселях (обычная кнопка ≈ 28)


class PromptCancelled(SystemExit):
    """Пользователь нажал «Отмена» в диалоге ввода — скрипт нужно остановить."""


class JobCancelled(SystemExit):
    """Пользователь нажал «Отмена» в окне приложения — выполняющийся скрипт обрывается.

    Наследуется от SystemExit (а не от Exception), чтобы его не глотали
    `except Exception:` внутри скриптов."""


def _interrupt_thread(thread):
    """Просит интерпретатор бросить JobCancelled в рабочем потоке.

    Исключение сработает, как только поток вернётся из текущего сетевого вызова
    (обрыв или таймаут requests), — следующих загрузок он уже не сделает."""
    try:
        import ctypes
        ctypes.pythonapi.PyThreadState_SetAsyncExc(
            ctypes.c_ulong(thread.ident), ctypes.py_object(JobCancelled)
        )
    except Exception:
        pass


class PostersApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        screen_h = self.root.winfo_screenheight()
        height = max(680, min(820, screen_h - 110))  # не выше экрана (с запасом на Dock и меню)
        self.root.geometry(f"940x{height}")
        self.root.minsize(840, 640)
        # ошибки в обработчиках кнопок иначе уходят в никуда (у .app нет консоли)
        self.root.report_callback_exception = self._on_tk_exception

        self.project_root = None
        self.startup_error = None
        self.log_queue = queue.Queue()
        self.busy = False
        self._prompt_cancelled = False
        self._worker_thread = None
        self._run_id = 0  # номер текущего запуска: сообщения «старых» потоков игнорируются
        self._osascript_proc = None

        try:
            self.project_root = _find_project_root()
        except Exception as exc:
            self.startup_error = str(exc)

        self.artist_var = tk.StringVar()
        self.project_var = tk.StringVar(
            value=str(self.project_root) if self.project_root else "Проект не найден"
        )
        self.status_var = tk.StringVar(value="Готово")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()
        if self.project_root:
            self._adopt_project(self.project_root, source="найден автоматически")
        else:
            self.log("Папка Mockups рядом с приложением не найдена.")
            self.status_var.set("Проект не найден")
            self.root.after(400, self._startup_manual_prompt)
        self._poll_log_queue()
        self.root.after(3000, self._start_update_check)  # тихая проверка обновлений

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text="POSTERS", font=_title_font(self.root, 26)).pack(side="left")
        self.settings_btn = ttk.Button(
            header, text="Настройки", command=self.open_token_dialog
        )
        self.settings_btn.pack(side="right", anchor="n")
        ttk.Label(
            outer,
            text="Подготовка мокапов → Excel → Wildberries",
        ).pack(anchor="w", pady=(0, 8))

        # --- ряд 1: Исполнитель (слева) и Проект (справа) ---
        top = ttk.Frame(outer)
        top.pack(fill="x")
        top.columnconfigure(0, weight=1, uniform="top")
        top.columnconfigure(1, weight=1, uniform="top")

        artist_frame = ttk.LabelFrame(top, text="Исполнитель", padding=14)
        artist_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        artist_frame.columnconfigure(0, weight=1)
        ttk.Label(artist_frame, text="Имя исполнителя:").grid(row=0, column=0, sticky="w")
        self.artist_entry = ttk.Entry(artist_frame, textvariable=self.artist_var)
        enable_clipboard_shortcuts(self.artist_entry)
        self.artist_entry.grid(row=1, column=0, sticky="ew", pady=(4, 8))
        ttk.Button(
            artist_frame,
            text="Загрузить из config.json",
            command=self.load_config_artist,
        ).grid(row=2, column=0, sticky="w")

        project_frame = ttk.LabelFrame(top, text="Проект", padding=14)
        project_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        project_frame.columnconfigure(0, weight=1)
        ttk.Label(project_frame, text="Папка проекта:").grid(row=0, column=0, sticky="w")
        ttk.Label(
            project_frame, textvariable=self.project_var, wraplength=360, justify="left"
        ).grid(row=1, column=0, sticky="w", pady=(4, 8))
        project_buttons = ttk.Frame(project_frame)
        project_buttons.grid(row=2, column=0, sticky="w")
        self.choose_btn = ttk.Button(
            project_buttons, text="Найти автоматически", command=self.choose_project
        )
        self.choose_btn.pack(side="left")
        self.manual_btn = ttk.Button(
            project_buttons, text="Выбрать вручную…", command=self.choose_project_manually
        )
        self.manual_btn.pack(side="left", padx=(8, 0))

        # --- ряд 2: этапы подряд со стрелками (кнопки рисуются сами, см. BigButton) ---
        steps = ttk.Frame(outer)
        steps.pack(fill="x", pady=(12, 0))
        for col in (0, 2, 4):
            steps.columnconfigure(col, weight=1, uniform="step")
        for col in (1, 3):
            steps.columnconfigure(col, minsize=44)

        step_font = _title_font(self.root, 14)
        self.mockup_btn = BigButton(
            steps, "Создать мокапы", self.run_mockups, height=STEP_BUTTON_HEIGHT, font=step_font
        )
        self.mockup_btn.grid(row=0, column=0, sticky="ew")
        self.excel_btn = BigButton(
            steps, "Заполнить Excel", self.run_excel, height=STEP_BUTTON_HEIGHT, font=step_font
        )
        self.excel_btn.grid(row=0, column=2, sticky="ew")
        self.wb_btn = BigButton(
            steps, "Загрузить на WB", self.run_wb, height=STEP_BUTTON_HEIGHT, font=step_font
        )
        self.wb_btn.grid(row=0, column=4, sticky="ew")
        for col in (1, 3):
            ttk.Label(steps, text="→", font=("Helvetica", 26), anchor="center").grid(
                row=0, column=col, sticky="nsew"
            )

        # --- ряд 3: Дополнительно — четыре кнопки в одну строку ---
        extra = ttk.LabelFrame(outer, text="Дополнительно", padding=14)
        extra.pack(fill="x", pady=(12, 0))
        for col in range(4):
            extra.columnconfigure(col, weight=1, uniform="extra")

        def add_extra(col, attr, text, command, hint):
            btn = ttk.Button(extra, text=text, command=command)
            btn.grid(row=0, column=col, sticky="ew", padx=4, ipady=6)
            setattr(self, attr, btn)
            ttk.Label(
                extra, text=hint, foreground="gray", wraplength=185, justify="left"
            ).grid(row=1, column=col, sticky="nw", padx=4, pady=(6, 0))

        add_extra(0, "tiff_btn", "Сохранить исходники", self.run_tiff,
                  "PSD → TIFF и JPG. Папку выбираете в окне Photoshop.")
        add_extra(1, "wb_photos_btn", "Загрузить фото", self.run_wb_photos,
                  "Фото на уже созданные карточки WB (артикулы из Постеры.xlsx).")
        add_extra(2, "pinterest_btn", "Загрузить на Pinterest", self.run_pinterest,
                  "source/white и dark → доска «<исполнитель> Albums». Аккаунт — в «Настройки».")
        add_extra(3, "stocks_btn", "Проставить остатки", self.run_stocks,
                  "Остаток на склад WB по артикулам из таблицы; если в папке есть stocks.xlsx — всем товарам из него. Число — в «Настройки».")

        ttk.Separator(outer).pack(fill="x", pady=8)
        status_row = ttk.Frame(outer)
        status_row.pack(fill="x")
        ttk.Label(status_row, textvariable=self.status_var).pack(side="left")
        self.cancel_btn = ttk.Button(
            status_row, text="Отмена", command=self.cancel_run, state="disabled"
        )
        self.cancel_btn.pack(side="right")

        self.progress = ttk.Progressbar(
            outer, variable=self.progress_var, maximum=100
        )
        self.progress.pack(fill="x", pady=(6, 8))

        bottom = ttk.Frame(outer)
        bottom.pack(side="bottom", fill="x", pady=(8, 0))

        ttk.Button(bottom, text="Открыть Excel", command=self.open_excel).pack(side="left")
        ttk.Button(bottom, text="Копировать лог", command=self.copy_log).pack(side="left", padx=(8, 0))
        ttk.Button(bottom, text="Очистить лог", command=self.clear_log).pack(side="right")

        log_frame = ttk.LabelFrame(outer, text="Лог", padding=8)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            log_frame, height=6, wrap="word", state="disabled", font=("Menlo", 11)
        )
        self.log_text.pack(side="left", fill="both", expand=True)
        _attach_context_menu(self.log_text, editable=False)

        scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)


    @staticmethod
    def _raise_if_cancelled():
        """Скрипты постоянно пишут в лог/прогресс — это удобная точка остановки."""
        if getattr(threading.current_thread(), "posters_cancelled", False):
            raise JobCancelled("Остановлено пользователем.")

    def log(self, message):
        self._raise_if_cancelled()
        self.log_queue.put(("log", str(message)))

    def set_status(self, message):
        self._raise_if_cancelled()
        self.log_queue.put(("status", str(message)))

    def set_progress(self, value):
        self._raise_if_cancelled()
        self.log_queue.put(("progress", float(value)))

    def _poll_log_queue(self):
        try:
            while True:
                kind, value = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(value)
                elif kind == "status":
                    self.status_var.set(value)
                elif kind == "progress":
                    self.progress_var.set(max(0, min(100, value)))
                elif kind == "done":
                    if value == self._run_id:  # «done» от отменённого запуска не трогаем
                        self._set_busy(False)
                elif kind == "error":
                    title, text = value if isinstance(value, tuple) else ("Ошибка", value)
                    messagebox.showerror(title, text)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def copy_log(self):
        """Copy the complete visible log to the macOS clipboard."""
        self.log_text.configure(state="normal")
        text = self.log_text.get("1.0", "end-1c")
        self.log_text.configure(state="disabled")

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self.status_var.set("Лог скопирован в буфер обмена")

    def _set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self.choose_btn, self.manual_btn, self.mockup_btn, self.excel_btn,
                    self.wb_btn, self.tiff_btn, self.wb_photos_btn, self.pinterest_btn,
                    self.stocks_btn, self.settings_btn):
            btn.configure(state=state)
        self.cancel_btn.configure(state="normal" if busy else "disabled")

    def _require_project(self):
        if self.project_root is None:
            try:
                self._adopt_project(_find_project_root(), source="найден автоматически")
            except Exception:
                if not self._pick_project_manually():
                    return False
        return True

    def _config_path(self):
        """config.json лежит в <Mockups>/scripts/ (в корне Mockups — только папки и таблица)."""
        if self.project_root is None:
            self.project_root = _find_project_root()
        scripts = self.project_root / "scripts"
        scripts.mkdir(exist_ok=True)
        return scripts / CONFIG_NAME

    def _load_config(self):
        path = self._config_path()
        if not path.exists():
            path.write_text(
                '{\n'
                '  "wildberries": {\n'
                '    "api_token": ""\n'
                '  },\n'
                '  "project": {\n'
                '    "artist": ""\n'
                '  }\n'
                '}\n',
                encoding="utf-8",
            )
            self.log(f"Создан config.json: {path}")
        try:
            with path.open("r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            raise UserError(
                f"Файл config.json повреждён (строка {e.lineno}): {e.msg}\n\n{path}\n\n"
                "Исправьте его или удалите — приложение создаст новый.",
                "Ошибка config.json",
            )

    def _artist(self):
        artist = self.artist_var.get().strip()
        if artist:
            return artist

        try:
            config = self._load_config()
            artist = config.get("project", {}).get("artist", "").strip()
        except Exception:
            artist = ""

        if artist:
            self.artist_var.set(artist)
            return artist

        messagebox.showwarning(
            "Исполнитель",
            "Укажите исполнителя в поле выше или в config.json.",
        )
        return None

    def _scripts_dir(self):
        if self.project_root is None:
            self.project_root = _find_project_root()
        path = self.project_root / "scripts"
        if not path.is_dir():
            raise FileNotFoundError(f"Не найдена папка scripts: {path}")
        return path

    def choose_project(self):
        if self.busy:
            return
        try:
            self._adopt_project(_find_project_root(), source="найден автоматически")
            self.status_var.set("Проект найден автоматически")
        except Exception:
            self.log("Автоматически папку Mockups найти не удалось — выберите её вручную.")
            self._pick_project_manually()

    def choose_project_manually(self):
        if self.busy:
            return
        self._pick_project_manually()

    def _startup_manual_prompt(self):
        if self.project_root is not None:
            return
        if messagebox.askyesno(
            "Posters",
            "Не удалось найти папку Mockups рядом с приложением.\n"
            "Выбрать её вручную?",
            parent=self.root,
        ):
            self._pick_project_manually()

    def _pick_project_manually(self):
        """Диалог выбора папки Mockups (только из главного потока)."""
        docs = Path.home() / "Documents"
        initial = str(docs if docs.is_dir() else Path.home())
        while True:
            chosen = filedialog.askdirectory(
                title="Выберите папку Mockups", initialdir=initial, parent=self.root
            )
            if not chosen:
                return False
            if _is_project_root(chosen):
                self._adopt_project(Path(chosen), source="выбран вручную", remember=True)
                self.status_var.set("Проект выбран вручную")
                return True
            if not messagebox.askretrycancel(
                "Папка не подходит",
                "В выбранной папке нет подпапки mockups (PSD-шаблоны) "
                "вместе с source, scripts или Постеры.xlsx.\n\nВыбрать другую папку?",
                parent=self.root,
            ):
                return False
            initial = str(Path(chosen).parent)

    def _adopt_project(self, root, source="", remember=False):
        """Принять папку Mockups: показать в окне, запомнить и разложить скрипты."""
        self.project_root = Path(root)
        self.project_var.set(str(self.project_root))
        self.log(f"Проект ({source}): {self.project_root}")
        if remember:
            _save_settings(project_root=str(self.project_root))
        try:
            self._sync_scripts()
            self._migrate_legacy_config()
        except Exception as e:
            self.log(f"⚠ Не удалось подготовить папку scripts: {e}")

    def _migrate_legacy_config(self):
        """Старые версии держали config.json в корне Mockups — переносим в scripts/."""
        old = self.project_root / CONFIG_NAME
        new = self._config_path()
        if not old.exists():
            return
        if not new.exists():
            shutil.move(str(old), str(new))
            self.log("config.json перенесён из корня Mockups в scripts/")
        else:
            self.log(
                "⚠ В корне Mockups остался старый config.json — он больше не используется "
                "(актуальный лежит в scripts/). Его можно удалить."
            )

    def _sync_scripts(self):
        """Копирует скрипты из приложения в <Mockups>/scripts (обновляет, если изменились)."""
        src = _bundled_scripts_dir()
        dst = self.project_root / "scripts"
        if not src.is_dir():
            self.log(f"⚠ Внутри приложения не найдены скрипты: {src}")
            return
        dst.mkdir(exist_ok=True)
        if src.resolve() == dst.resolve():
            return  # запуск из исходников — копировать нечего
        updated = []
        for f in sorted(src.iterdir()):
            if f.suffix not in SCRIPT_SUFFIXES:
                continue
            target = dst / f.name
            if not target.exists() or target.read_bytes() != f.read_bytes():
                shutil.copy2(f, target)
                updated.append(f.name)
        if updated:
            self.log("Скрипты обновлены в scripts/: " + ", ".join(updated))

    def _call_on_main(self, func):
        """Выполнить func в главном потоке (tkinter) и вернуть результат — для вызова из воркеров."""
        holder = {}
        done = threading.Event()

        def run():
            try:
                holder["value"] = func()
            except Exception as exc:
                holder["error"] = exc
            finally:
                done.set()

        self.root.after(0, run)
        done.wait()
        if "error" in holder:
            raise holder["error"]
        return holder.get("value")

    def _ask_photoshop_manually(self):
        """Photoshop не найден в «Программах» — просим указать .app вручную (до 3 попыток)."""
        def ask():
            return filedialog.askopenfilename(
                title="Выберите Adobe Photoshop",
                initialdir="/Applications",
                filetypes=[("Приложение", "*.app")],
                parent=self.root,
            )

        self._call_on_main(lambda: messagebox.showinfo(
            "Photoshop",
            "Adobe Photoshop не найден в «Программах».\nУкажите приложение Photoshop вручную.",
            parent=self.root,
        ))
        for _ in range(3):
            path = self._call_on_main(ask)
            if not path:
                return None
            p = Path(path)
            if p.suffix == ".app" and p.is_dir() and "photoshop" in p.name.lower():
                _save_settings(photoshop_app=str(p))
                return str(p)
            self._call_on_main(lambda: messagebox.showerror(
                "Photoshop", "Это не похоже на Adobe Photoshop. Попробуйте ещё раз.",
                parent=self.root,
            ))
        return None

    def _ensure_wb_token(self):
        """Токена WB нет в config.json — спрашиваем (до 3 раз) и сохраняем."""
        if self._read_token():
            return
        for _ in range(3):
            raw = self._gui_prompt(
                "Токен Wildberries API не найден в config.json.\n"
                "Вставьте токен:"
            )
            if raw is None:
                break
            token = self._clean_token(raw)
            if token:
                self._save_token(token)
                self.log(f"Токен WB сохранён: …{token[-6:]}")
                return
        raise UserError(
            "Токен WB не задан. Его можно ввести через кнопку «Настройки» в шапке окна.",
            "Нужен токен WB",
        )

    def load_config_artist(self):
        if not self._require_project():
            return

        try:
            config = self._load_config()
            artist = config.get("project", {}).get("artist", "").strip()
            if not artist:
                messagebox.showinfo(
                    "Исполнитель",
                    "В config.json пока не указан project.artist.",
                )
                return
            self.artist_var.set(artist)
            self.log(f"Исполнитель из config.json: {artist}")
        except Exception as e:
            messagebox.showerror("Ошибка config.json", str(e))

    # ---------- Mockups / Photoshop ----------

    def run_mockups(self):
        if not self._require_project():
            return
        artist = self._artist()
        if not artist:
            return

        jsx = self._find_script("mockuper.jsx")
        if not jsx:
            return

        self._run_background(
            "Создание мокапов", self._mockup_worker, jsx, artist
        )

    def _mockup_worker(self, jsx, artist):
        self.set_status("Запуск Photoshop…")
        self.set_progress(0)
        self.log("Запускаю mockuper.jsx…")
        self.log(f"Исполнитель: {artist}")
        self._write_artist_to_config(artist)

        self._run_photoshop_script(jsx)

        self.set_progress(100)
        self.log("✓ Этап мокапов завершён.")
        self.set_status("Мокапы готовы")

    def _run_photoshop_script(self, jsx):
        """Запускает .jsx в Photoshop через AppleScript. Общий для мокапов и «Дополнительно»."""
        ps_app = _find_photoshop() or self._ask_photoshop_manually()
        if not ps_app:
            raise UserError(
                "Adobe Photoshop не найден и не выбран вручную.\n"
                "Установите Photoshop и запустите этап снова.",
                "Photoshop не найден",
            )
        self.log(f"Photoshop: {ps_app}")
        ps_app_esc = ps_app.replace("\\", "\\\\").replace('"', '\\"')

        # Имя приложения подставляем литералом (полный путь), чтобы AppleScript
        # загрузил словарь Photoshop. Файл-объект Photoshop не принимает
        # (ошибка -1728), поэтому отдаём ему JS-код, который сам запускает .jsx
        # через $.evalFile ($.fileName внутри скрипта будет путём к скрипту).
        # Код передаём аргументом (argv): кириллица и кавычки не ломаются.
        js_code = "$.evalFile(new File(" + json.dumps(str(Path(jsx).resolve())) + "));"

        applescript = f"""
on run argv
    set jsCode to item 1 of argv
    tell application "{ps_app_esc}"
        activate
        -- по умолчанию AppleScript ждёт ответ всего 2 минуты (ошибка -1712),
        -- а мокапы с ручными правками идут дольше
        with timeout of 86400 seconds
            do javascript jsCode
        end timeout
    end tell
end run
"""

        self.set_progress(10)

        try:
            proc = subprocess.Popen(
                ["osascript", "-e", applescript, js_code],
                cwd=str(self.project_root),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            self._osascript_proc = proc
            try:
                stdout, stderr = proc.communicate()
            finally:
                self._osascript_proc = None
            result = subprocess.CompletedProcess(proc.args, proc.returncode, stdout, stderr)
        except FileNotFoundError:
            raise UserError(
                "Не найдена системная утилита osascript — приложение работает только на macOS.",
                "Нужен macOS",
            )

        if result.stdout.strip():
            self.log(result.stdout.strip())
        if result.stderr.strip():
            self.log(result.stderr.strip())

        if result.returncode != 0:
            raise UserError(
                self._explain_osascript(result.returncode, result.stderr, result.stdout),
                "Ошибка Photoshop",
            )

    @staticmethod
    def _explain_osascript(code, stderr, stdout):
        err = (stderr or "").strip()
        low = err.lower()
        if "-1743" in err or "not authorized" in low:
            return (
                "macOS не разрешил приложению управлять Photoshop.\n\n"
                "Откройте: Системные настройки → Конфиденциальность и безопасность → "
                "Автоматизация — и включите Photoshop для Posters. Если Posters в списке "
                "нет, перезапустите приложение и запустите этап ещё раз: macOS снова "
                "покажет запрос."
            )
        if "-128" in err or "user canceled" in low or "user cancelled" in low:
            return "Операция отменена."
        if "-1712" in err:
            return "Photoshop слишком долго не отвечал (тайм-аут). Проверьте, не висит ли в нём окно с вопросом."
        if "-600" in err or "isn’t running" in low or "isn't running" in low:
            return "Не удалось запустить Photoshop. Откройте его вручную и повторите."
        if "-1728" in err:
            return (
                "Photoshop не принял команду запуска скрипта (-1728). Убедитесь, что Photoshop "
                "запущен и в нём нет открытых диалоговых окон, и повторите."
            )
        details = err or (stdout or "").strip() or "без сообщения"
        return f"Photoshop не смог выполнить скрипт (код {code}).\n\n{details[:800]}"

    def _find_script(self, name):
        """Путь к скрипту в <Mockups>/scripts или None (с понятным сообщением)."""
        try:
            path = self._scripts_dir() / name
        except FileNotFoundError as e:
            messagebox.showerror("Папка scripts не найдена", str(e))
            return None
        if not path.exists():
            messagebox.showerror(
                "Скрипт не найден",
                f"Не найден {name}:\n{path}\n\n"
                "Перезапустите приложение — оно восстановит скрипты в папке scripts. "
                "Если не помогло, приложение собрано без этого файла.",
            )
            return None
        return path

    @staticmethod
    def _script_error_text(result, prefix, hint):
        error = result.get("error") if isinstance(result, dict) else None
        text = _strip_marker(error) if error is not None else ""
        if not text or text.isdigit():  # у скриптов бывает «ошибка» вида 1 — смысл в логе
            return f"{prefix} {hint}"
        if "401" in text:
            text += "\n\nТокен можно заменить через кнопку «Настройки» в шапке окна."
        return text

    def _on_tk_exception(self, exc_type, exc, tb):
        """Любая необработанная ошибка в обработчике кнопки — в лог и окно, а не в пустоту."""
        err = exc if exc is not None else exc_type()
        title, text, unexpected = _explain_error(err)
        _write_error_log("интерфейс", err, tb)
        try:
            self.log(f"✗ {text}")
            if unexpected:
                self.log("Технические подробности:\n" + "".join(
                    traceback.format_exception(exc_type, exc, tb)))
            messagebox.showerror(title, text)
        except Exception:
            pass

    # ---------- Дополнительно ----------

    def run_wb_photos(self):
        if not self._require_project():
            return
        artist = self._artist()
        if not artist:
            return
        script = self._find_script("wb_photo_uploader.py")
        if not script:
            return
        self._run_background(
            "Загрузка фото на Wildberries", self._wb_photos_worker, script, artist
        )

    def _wb_photos_worker(self, script, artist):
        self.set_status("Загрузка фото на Wildberries…")
        self.set_progress(0)
        self.log("Запускаю wb_photo_uploader.py…")
        self.log(f"Исполнитель: {artist}")
        self._write_artist_to_config(artist)
        self._ensure_wb_token()

        if str(script.parent) not in sys.path:
            sys.path.insert(0, str(script.parent))

        import importlib
        import wb_photo_uploader
        importlib.reload(wb_photo_uploader)

        result = wb_photo_uploader.main(
            artist=artist,
            log_callback=self.log,
            progress_callback=self._wb_progress,
            prompt_callback=self._gui_prompt,
            config_path=str(self._config_path()),
        )

        if not isinstance(result, dict) or not result.get("success"):
            raise UserError(self._script_error_text(
                result,
                "Загрузка фото остановлена.",
                "Причины — в логе выше. Проверьте, что мокапы созданы (папка output/<исполнитель>), "
                "карточки уже есть на WB и Постеры.xlsx не открыт в Excel.",
            ), "Фото не загружены")

        self.set_progress(100)
        self.log("✓ Фото загружены.")
        self.set_status("Фото загружены")

    def run_stocks(self):
        if not self._require_project():
            return
        script = self._find_script("wb_stocks.py")
        if not script:
            return
        self._run_background("Проставление остатков на Wildberries", self._stocks_worker, script)

    def _stocks_worker(self, script):
        amount = self._read_stock_amount()
        self.set_status("Проставляю остатки…")
        self.set_progress(0)
        self.log("Запускаю wb_stocks.py…")
        self.log(f"Остаток на каждый артикул: {amount} (число меняется в «Настройки»).")
        self._ensure_wb_token()

        if str(script.parent) not in sys.path:
            sys.path.insert(0, str(script.parent))

        import importlib
        import wb_stocks
        importlib.reload(wb_stocks)

        result = wb_stocks.main(
            log_callback=self.log,
            progress_callback=self._wb_progress,
            prompt_callback=self._gui_prompt,
            config_path=str(self._config_path()),
            amount=amount,
        )

        if not isinstance(result, dict) or not result.get("success"):
            raise UserError(self._script_error_text(
                result,
                "Проставление остатков остановлено.",
                "Причины — в логе выше. Проверьте, что карточки уже созданы на WB, а у токена есть "
                "право «Маркетплейс» (кнопка «Настройки»).",
            ), "Остатки не проставлены")

        self.set_progress(100)
        if isinstance(result, dict) and result.get("stocks_file"):
            # итоговое сообщение («можно удалять таблицу») скрипт уже вывел сам
            self.set_status("Остатки проставлены — файл остатков можно удалить")
        else:
            self.log(f"✓ Остатки проставлены ({amount} на каждый артикул).")
            self.set_status("Остатки проставлены")

    def run_pinterest(self):
        if not self._require_project():
            return
        artist = self._artist()
        if not artist:
            return
        script = self._find_script("pinterest_uploader.py")
        if not script:
            return
        self._run_background(
            "Загрузка на Pinterest", self._pinterest_worker, script, artist
        )

    def _pinterest_worker(self, script, artist):
        self.set_status("Загрузка на Pinterest…")
        self.set_progress(0)
        self.log("Запускаю pinterest_uploader.py…")
        self.log(f"Исполнитель: {artist}")
        self._write_artist_to_config(artist)
        self._ensure_pinterest_credentials()

        if str(script.parent) not in sys.path:
            sys.path.insert(0, str(script.parent))

        import importlib
        import pinterest_uploader
        importlib.reload(pinterest_uploader)

        result = pinterest_uploader.main(
            artist=artist,
            log_callback=self.log,
            progress_callback=self._wb_progress,
            prompt_callback=self._gui_prompt,
            config_path=str(self._config_path()),
        )

        if not isinstance(result, dict) or not result.get("success"):
            raise UserError(self._script_error_text(
                result,
                "Загрузка на Pinterest остановлена.",
                "Причины — в логе выше. Проверьте интернет, данные аккаунта (кнопка «Настройки») и что "
                "в source/white и source/dark лежат jpg/png.",
            ), "Pinterest: загрузка не завершена")

        self.set_progress(100)
        self.log("✓ Исходники загружены на Pinterest.")
        self.set_status("Pinterest готов")

    def run_tiff(self):
        if not self._require_project():
            return
        jsx = self._find_script("tiff_jpg_saver.jsx")
        if not jsx:
            return
        self._run_background(
            "Сохранение исходников в TIFF и JPG", self._tiff_worker, jsx
        )

    def _tiff_worker(self, jsx):
        self.set_status("Запуск Photoshop…")
        self.set_progress(0)
        self.log("Запускаю tiff_jpg_saver.jsx…")
        self.log("Выберите папку с PSD в окне Photoshop.")

        self._run_photoshop_script(jsx)

        self.set_progress(100)
        self.log("✓ Сохранение TIFF и JPG завершено.")
        self.set_status("TIFF и JPG готовы")

    # ---------- Excel ----------

    def run_excel(self):
        if not self._require_project():
            return
        artist = self._artist()
        if not artist:
            return

        script = self._find_script("excel_filler.py")
        if not script:
            return

        self._run_background(
            "Заполнение Excel", self._excel_worker, script, artist
        )

    def _excel_worker(self, script, artist):
        self.set_status("Заполняю Excel…")
        self.set_progress(0)
        self.log("Запускаю excel_filler.py…")
        self._write_artist_to_config(artist)

        if str(script.parent) not in sys.path:
            sys.path.insert(0, str(script.parent))

        import importlib
        import excel_filler
        importlib.reload(excel_filler)

        result = excel_filler.main(
            artist=artist,
            log_callback=self.log,
            config_path=str(self._config_path()),
        )

        if not isinstance(result, dict) or not result.get("success"):
            raise UserError(self._script_error_text(
                result,
                "Заполнение Excel остановлено.",
                "Подробности — в логе выше. Проверьте, что Постеры.xlsx не открыт в Excel "
                "и что в папке source лежат постеры, затем запустите этап снова.",
            ), "Excel не заполнен")

        self.set_progress(100)
        self.log("✓ Excel заполнен.")
        self.set_status("Excel готов")

    # ---------- Wildberries ----------

    def run_wb(self):
        if not self._require_project():
            return
        artist = self._artist()
        if not artist:
            return

        script = self._find_script("wb_uploader.py")
        if not script:
            return

        self._run_background(
            "Загрузка на Wildberries", self._wb_worker, script, artist
        )

    def _wb_worker(self, script, artist):
        self.set_status("Загрузка на Wildberries…")
        self.set_progress(0)
        self.log("Запускаю wb_uploader.py…")
        self._write_artist_to_config(artist)
        self._ensure_wb_token()

        if str(script.parent) not in sys.path:
            sys.path.insert(0, str(script.parent))

        import importlib
        import wb_uploader
        importlib.reload(wb_uploader)

        result = wb_uploader.main(
            artist=artist,
            log_callback=self.log,
            progress_callback=self._wb_progress,
            prompt_callback=self._gui_prompt,
            config_path=str(self._config_path()),
        )

        if not isinstance(result, dict) or not result.get("success"):
            raise UserError(self._script_error_text(
                result,
                "Загрузка на Wildberries остановлена.",
                "Причины перечислены в логе выше. Чаще всего это незаполненные ячейки "
                "'Цвет' или 'Наименование' в Постеры.xlsx — исправьте их, сохраните таблицу "
                "в Excel и запустите этап снова (уже выполненное скрипт пропустит).",
            ), "Загрузка на WB остановлена")

        self.set_progress(100)
        self.log("✓ Загрузка на Wildberries завершена.")
        self.set_status("Готово")

    def _wb_progress(self, value, message=""):
        """Скрипты WB присылают готовый прогресс 0..100 и короткий статус."""
        try:
            self.set_progress(float(value))
        except (TypeError, ValueError):
            pass
        if message:
            self.set_status(str(message))

    def _gui_prompt(self, message, default=None, secret=False):
        """Диалог ввода для скриптов. «Отмена» останавливает скрипт (PromptCancelled)."""
        result_holder = {"value": default, "cancelled": False}
        done = threading.Event()

        def show():
            try:
                value = self._ask_string(str(message), default, secret)
                if value is None:
                    result_holder["cancelled"] = True
                else:
                    result_holder["value"] = value
            except Exception:
                result_holder["cancelled"] = True
            finally:
                done.set()  # иначе рабочий поток зависнет, если диалог не открылся

        self.root.after(0, show)
        done.wait()
        if result_holder["cancelled"]:
            self._prompt_cancelled = True
            raise PromptCancelled("Остановлено пользователем.")
        return result_holder["value"]

    def _ask_string(self, message, default=None, secret=False):
        """Окно ввода: поле в 3 раза шире стандартного, работают Cmd/Ctrl + C/V/X/A.
        Возвращает строку или None, если нажали «Отмена» / закрыли окно."""
        win = tk.Toplevel(self.root)
        win.title("Posters")
        win.transient(self.root)
        win.resizable(True, False)

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=message, wraplength=640, justify="left").pack(anchor="w", pady=(0, 8))

        var = tk.StringVar(value="" if default is None else str(default))
        entry = ttk.Entry(frame, textvariable=var, width=90, show="•" if secret else "")
        entry.pack(fill="x")
        enable_clipboard_shortcuts(entry)

        result = {"value": None}

        def ok(_event=None):
            result["value"] = var.get()
            win.destroy()

        def cancel(_event=None):
            result["value"] = None
            win.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="OK", command=ok).pack(side="right")
        ttk.Button(buttons, text="Отмена", command=cancel).pack(side="right", padx=(0, 8))

        win.bind("<Return>", ok)
        win.bind("<Escape>", cancel)
        win.protocol("WM_DELETE_WINDOW", cancel)

        win.update_idletasks()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - win.winfo_reqwidth()) // 2)
        y = self.root.winfo_rooty() + 80
        win.geometry(f"+{x}+{y}")
        try:
            win.grab_set()
        except tk.TclError:
            pass
        entry.focus_set()
        entry.select_range(0, "end")
        self.root.wait_window(win)
        return result["value"]

    # ---------- WB token ----------

    @staticmethod
    def _clean_token(raw):
        # при копировании токен часто захватывает пробелы, переносы и кавычки
        return "".join(raw.split()).strip("\"'")

    def _read_token(self):
        try:
            config = self._load_config()
            return ((config.get("wildberries") or {}).get("api_token") or "").strip()
        except Exception:
            return ""

    DEFAULT_STOCK_AMOUNT = 100

    def _read_stock_amount(self):
        try:
            block = self._load_config().get("wildberries") or {}
            value = int(str(block.get("stock_amount")).strip())
            return value if value >= 0 else self.DEFAULT_STOCK_AMOUNT
        except Exception:
            return self.DEFAULT_STOCK_AMOUNT

    def _update_config(self, section, values):
        """Обновляет config[section] переданными значениями, остальное содержимое не трогает."""
        path = self._config_path()
        config = {}
        if path.exists():
            try:
                with path.open("r", encoding="utf-8") as f:
                    config = json.load(f)
            except Exception:
                # не теряем повреждённый файл: сохраняем копию рядом
                try:
                    shutil.copy2(path, path.with_name(path.name + ".bak"))
                    self.log("⚠ config.json был повреждён — копия сохранена как config.json.bak")
                except OSError:
                    pass
                config = {}
        if not isinstance(config, dict):
            config = {}
        block = config.get(section)
        if not isinstance(block, dict):
            block = {}
        block.update(values)
        config[section] = block
        with path.open("w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
            f.write("\n")

    def _save_token(self, token):
        self._update_config("wildberries", {"api_token": token})

    def _read_pinterest(self):
        try:
            block = self._load_config().get("pinterest")
        except Exception:
            return {}
        if not isinstance(block, dict):
            return {}
        return {k: str(block.get(k) or "").strip() for k in ("email", "username", "password")}

    def _ensure_pinterest_credentials(self):
        """Недостающие данные Pinterest спрашиваем (до 3 раз на поле) и сохраняем в config.json."""
        fields = (
            ("email", "почту (email) аккаунта Pinterest", False),
            ("username", "имя пользователя Pinterest (то, что в ссылке на профиль)", False),
            ("password", "пароль от аккаунта Pinterest", True),
        )
        creds = self._read_pinterest()
        for key, label, secret in fields:
            if creds.get(key):
                continue
            for _ in range(3):
                raw = self._gui_prompt(
                    f"В config.json не найдено: {label}.\nВведите значение:", secret=secret
                )
                if raw is None:
                    raise UserError(
                        "Загрузка на Pinterest отменена: данные аккаунта не введены. "
                        "Их можно заполнить заранее через кнопку «Настройки» в шапке окна.",
                        "Нужны данные Pinterest",
                    )
                value = raw.strip("\r\n") if secret else raw.strip()
                if key == "username":
                    value = value.lstrip("@")
                if value.strip():  # пробелы вместо значения не принимаем
                    self._update_config("pinterest", {key: value})
                    self.log(f"Сохранено в config.json: pinterest.{key}")
                    break
            else:
                raise UserError(
                    f"Не задано: {label}. Заполните через кнопку «Настройки» в шапке окна.",
                    "Нужны данные Pinterest",
                )

    # ---------- обновления ----------
    def _start_update_check(self, manual=False):
        if not GITHUB_REPO:
            if manual:
                messagebox.showinfo("Обновления", "В этой сборке проверка обновлений не настроена.")
            return
        threading.Thread(target=self._update_check_worker, args=(manual,), daemon=True).start()

    def _update_check_worker(self, manual):
        try:
            release = _fetch_latest_release()
        except Exception as exc:
            if manual:
                msg = ("Не удалось проверить обновления. Проверьте интернет и попробуйте позже.\n"
                       f"({type(exc).__name__})")
                self.root.after(0, lambda: messagebox.showwarning("Обновления", msg))
            return

        tag = str(release.get("tag_name") or "")
        if _version_tuple(tag) <= _version_tuple(APP_VERSION):
            if manual:
                msg = f"У вас последняя версия ({APP_VERSION})."
                self.root.after(0, lambda: messagebox.showinfo("Обновления", msg))
            return

        wanted = _update_asset_name()
        asset = next((a for a in release.get("assets") or [] if a.get("name") == wanted), None)
        self.root.after(0, lambda: self._offer_update(release, tag, asset, manual))

    def _offer_update(self, release, tag, asset, manual):
        if self.busy:
            return
        if not asset:
            # релиз опубликован, а файл для этого Mac ещё собирается на GitHub
            if manual:
                messagebox.showinfo(
                    "Обновления",
                    f"Вышла версия {tag}, но файл обновления ещё собирается "
                    "(обычно около 10 минут). Попробуйте позже.",
                )
            return

        notes = str(release.get("body") or "").strip()
        if len(notes) > 500:
            notes = notes[:500] + "…"
        text = f"Вышла новая версия {tag} (у вас {APP_VERSION})."
        if notes:
            text += "\n\nЧто нового:\n" + notes
        text += "\n\nОбновить сейчас? Приложение закроется и откроется заново. Ваши настройки сохранятся."
        if messagebox.askyesno("Доступно обновление", text):
            self._run_background(
                "Обновление приложения", self._update_worker,
                asset.get("browser_download_url"), int(asset.get("size") or 0), tag,
            )

    def _update_worker(self, url, size, tag):
        import tempfile
        import requests

        app = _app_bundle_path()
        if app is None:
            raise UserError("Автообновление работает только в собранном приложении Posters.app.", "Обновление")
        if "AppTranslocation" in str(app):
            raise UserError(
                "Приложение запущено прямо из «Загрузок» или архива, поэтому заменить себя не может. "
                "Перенесите Posters.app в папку «Программы», откройте оттуда и повторите обновление.",
                "Обновление",
            )
        if not os.access(app.parent, os.W_OK):
            raise UserError(
                f"Нет прав на запись в папку:\n{app.parent}\n"
                "Перенесите Posters.app в «Программы» (или в свою папку) и повторите обновление.",
                "Обновление",
            )

        tmp = Path(tempfile.mkdtemp(prefix="posters_update_"))
        try:
            zip_path = tmp / "update.zip"
            self.log(f"Скачиваю версию {tag}…")
            with requests.get(url, stream=True, timeout=UPDATE_TIMEOUT * 3) as r:
                r.raise_for_status()
                total = size or int(r.headers.get("Content-Length") or 0)
                got = 0
                with open(zip_path, "wb") as f:
                    for chunk in r.iter_content(256 * 1024):
                        f.write(chunk)
                        got += len(chunk)
                        if total:
                            self.set_progress(min(90, got * 90 / total))

            self.log("Распаковываю…")
            unpack = tmp / "new"
            unpack.mkdir()
            res = subprocess.run(
                ["ditto", "-x", "-k", str(zip_path), str(unpack)], capture_output=True, text=True
            )
            new_app = unpack / "Posters.app"
            if res.returncode != 0 or not new_app.is_dir():
                raise UserError("Не удалось распаковать скачанное обновление.", "Обновление")

            script = tmp / "swap.sh"
            script.write_text(_SWAP_SCRIPT, encoding="utf-8")
            subprocess.Popen(
                ["/bin/sh", str(script), str(os.getpid()), str(new_app), str(app), str(tmp)],
                start_new_session=True,
                stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        except BaseException:
            shutil.rmtree(tmp, ignore_errors=True)
            raise

        self.set_progress(100)
        self.log("Устанавливаю и перезапускаю приложение…")
        self.root.after(800, self.root.destroy)

    def open_token_dialog(self):
        """Окно «Настройки»: токен WB, остатки и данные аккаунта Pinterest. Пустое поле = не менять."""
        if self.busy or not self._require_project():
            return

        current = self._read_token()
        current_amount = self._read_stock_amount()
        pin = self._read_pinterest()

        win = tk.Toplevel(self.root)
        win.title("Настройки")
        win.resizable(False, False)
        win.transient(self.root)

        screen_h = self.root.winfo_screenheight()
        content_w = min(SETTINGS_CONTENT_WIDTH, self.root.winfo_screenwidth() - 160)
        wrap = content_w - 40
        scroll = ScrollableFrame(win, max_height=max(380, screen_h - 280), width=content_w)
        scroll.pack(fill="both", expand=True, padx=20, pady=(20, 0))
        frame = scroll.inner

        if GITHUB_REPO:  # блок версии показываем, только если включено автообновление
            version_row = ttk.Frame(frame)
            version_row.pack(fill="x", pady=(0, 12))
            ttk.Label(version_row, text=f"Версия приложения: {APP_VERSION}", foreground="gray").pack(side="left")
            ttk.Button(
                version_row, text="Проверить обновления",
                command=lambda: (win.destroy(), self._start_update_check(manual=True)),
            ).pack(side="right")

        ttk.Label(
            frame, text="Токен Wildberries API", font=("Helvetica", 14, "bold")
        ).pack(anchor="w")
        status = f"Сейчас сохранён: …{current[-6:]}" if current else "Токен пока не задан"
        ttk.Label(frame, text=status).pack(anchor="w", pady=(2, 8))

        token_var = tk.StringVar()
        token_entry = ttk.Entry(frame, textvariable=token_var, width=58, show="•")
        token_entry.pack(fill="x")
        enable_clipboard_shortcuts(token_entry)

        ttk.Separator(frame).pack(fill="x", pady=16)

        ttk.Label(
            frame, text="Остатки Wildberries", font=("Helvetica", 14, "bold")
        ).pack(anchor="w")
        amount_row = ttk.Frame(frame)
        amount_row.pack(fill="x", pady=(8, 0))
        ttk.Label(amount_row, text="Остаток на каждый артикул, шт.").pack(side="left")
        amount_var = tk.StringVar(value=str(current_amount))
        amount_entry = ttk.Entry(amount_row, textvariable=amount_var, width=8)
        amount_entry.pack(side="left", padx=(10, 0))
        enable_clipboard_shortcuts(amount_entry)
        ttk.Label(
            frame,
            text="Кнопка «Проставить остатки» ставит именно это число (заменяет текущий остаток, "
                 "а не прибавляет). Число можно менять в любой момент.",
            foreground="gray",
            wraplength=wrap,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        ttk.Separator(frame).pack(fill="x", pady=16)

        ttk.Label(
            frame, text="Аккаунт Pinterest", font=("Helvetica", 14, "bold")
        ).pack(anchor="w")
        if pin.get("email") or pin.get("username"):
            pin_status = "Сейчас: " + " / ".join(
                v for v in (pin.get("email"), pin.get("username")) if v
            ) + ("  (пароль сохранён)" if pin.get("password") else "  (пароль не задан)")
        else:
            pin_status = "Данные пока не заданы"
        ttk.Label(frame, text=pin_status).pack(anchor="w", pady=(2, 8))

        grid = ttk.Frame(frame)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        email_var, user_var, pass_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        ttk.Label(grid, text="Почта").grid(row=0, column=0, sticky="w", pady=2)
        enable_clipboard_shortcuts(ttk.Entry(grid, textvariable=email_var)).grid(row=0, column=1, sticky="ew", padx=(10, 0))
        ttk.Label(grid, text="Имя пользователя").grid(row=1, column=0, sticky="w", pady=2)
        enable_clipboard_shortcuts(ttk.Entry(grid, textvariable=user_var)).grid(row=1, column=1, sticky="ew", padx=(10, 0))
        ttk.Label(grid, text="Пароль").grid(row=2, column=0, sticky="w", pady=2)
        pass_entry = ttk.Entry(grid, textvariable=pass_var, show="•")
        enable_clipboard_shortcuts(pass_entry)
        pass_entry.grid(row=2, column=1, sticky="ew", padx=(10, 0))

        show_var = tk.BooleanVar(value=False)

        def toggle_show():
            mark = "" if show_var.get() else "•"
            token_entry.configure(show=mark)
            pass_entry.configure(show=mark)

        ttk.Checkbutton(
            frame, text="Показать токен и пароль", variable=show_var, command=toggle_show
        ).pack(anchor="w", pady=(8, 0))

        ttk.Label(
            frame,
            text="Заполните только то, что нужно изменить: пустые поля не затрагиваются.\n"
                 "Данные записываются в scripts/config.json на этом компьютере.",
            foreground="gray",
        ).pack(anchor="w", pady=(8, 14))

        ttk.Separator(frame).pack(fill="x", pady=(0, 12))
        table_block = self._build_table_settings_block(frame, win, scroll)

        def save():
            token = self._clean_token(token_var.get())
            email = email_var.get().strip()
            username = user_var.get().strip().lstrip("@")
            password = pass_var.get().strip("\r\n")
            if not password.strip():
                password = ""

            amount_text = amount_var.get().strip()
            if not amount_text.isdigit():
                messagebox.showwarning(
                    "Настройки", "Остаток — целое число 0 или больше (например, 100).", parent=win
                )
                return
            new_amount = int(amount_text)
            amount_changed = new_amount != current_amount

            table_new = None
            if table_block is not None:
                try:
                    table_new = table_block["collect"]()
                except ValueError as e:
                    messagebox.showwarning("Заполнение таблицы", str(e), parent=win)
                    return
            table_changed = table_new is not None and table_new != table_block["current"]

            if not (token or email or username or password or amount_changed or table_changed):
                messagebox.showwarning(
                    "Настройки", "Ничего не изменено. Заполните нужное или нажмите «Отмена».", parent=win
                )
                return
            if token and token.count(".") != 2:
                if not messagebox.askyesno(
                    "Токен",
                    "Это не похоже на токен WB (он состоит из трёх частей через точку).\n"
                    "Сохранить всё равно?",
                    parent=win,
                ):
                    return
            if email and "@" not in email:
                if not messagebox.askyesno(
                    "Почта", "Это не похоже на адрес почты. Сохранить всё равно?", parent=win
                ):
                    return

            saved = []
            try:
                if token:
                    self._save_token(token)
                    saved.append(f"токен WB (…{token[-6:]})")
                if amount_changed:
                    self._update_config("wildberries", {"stock_amount": new_amount})
                    saved.append(f"остаток WB: {new_amount}")
                if table_changed:
                    self._update_config("table", table_new)
                    saved.append("параметры заполнения таблицы")
                pin_values = {
                    k: v for k, v in (("email", email), ("username", username), ("password", password)) if v
                }
                if pin_values:
                    self._update_config("pinterest", pin_values)
                    saved.append("Pinterest: " + ", ".join(pin_values))
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{e}", parent=win)
                return
            self.log("Сохранено: " + "; ".join(saved))
            self.status_var.set("Настройки сохранены")
            win.destroy()

        buttons = ttk.Frame(win, padding=(20, 12, 20, 20))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Сохранить", command=save).pack(side="right")
        ttk.Button(buttons, text="Отмена", command=win.destroy).pack(side="right", padx=(0, 8))

        win.bind("<Return>", lambda e: None if isinstance(e.widget, tk.Text) else save())
        win.bind("<Escape>", lambda e: win.destroy())

        win.update_idletasks()
        x = max(10, self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2)
        y = self.root.winfo_rooty() + 60
        win.geometry(f"+{x}+{y}")
        token_entry.focus_set()
        win.grab_set()

    def _build_table_settings_block(self, parent, win, scroll):
        """Сворачиваемый блок «Заполнение таблицы» в окне «Настройки».

        Показывает все параметры, которые excel_filler.py подставляет в Постеры.xlsx.
        Возвращает {"collect": функция, "current": значения} или None, если скрипт
        не удалось загрузить."""
        wrap = int(scroll.canvas.cget("width")) - 40
        try:
            script = self._scripts_dir() / "excel_filler.py"
            if str(script.parent) not in sys.path:
                sys.path.insert(0, str(script.parent))
            import importlib
            import excel_filler
            importlib.reload(excel_filler)
            defaults = excel_filler.get_default_table_settings()
            current = excel_filler.load_table_settings(str(self._config_path()))
        except Exception as exc:
            ttk.Label(
                parent,
                text="Заполнение таблицы: не удалось загрузить параметры "
                     f"({type(exc).__name__}). Перезапустите приложение.",
                foreground="gray",
                wraplength=wrap,
                justify="left",
            ).pack(anchor="w")
            return None

        title = "Заполнение таблицы"
        toggle = ttk.Button(parent, text=f"▸  {title}")
        toggle.pack(fill="x")
        body = ttk.Frame(parent)
        opened = {"value": False}

        def toggle_block():
            opened["value"] = not opened["value"]
            if opened["value"]:
                body.pack(fill="x", pady=(10, 0))
                toggle.configure(text=f"▾  {title}")
            else:
                body.pack_forget()
                toggle.configure(text=f"▸  {title}")
            win.update_idletasks()
            scroll.refresh()
            if opened["value"]:
                # прокручиваем так, чтобы заголовок блока оказался вверху видимой области
                scroll.canvas.update_idletasks()
                total = max(1, scroll.inner.winfo_reqheight())
                scroll.canvas.yview_moveto(toggle.winfo_y() / total)

        toggle.configure(command=toggle_block)

        def fmt(value):
            return str(value)

        ttk.Label(
            body,
            text="Эти значения подставляются в Постеры.xlsx при нажатии «Заполнить Excel». "
                 "Изменения вступают в силу после «Сохранить».",
            foreground="gray",
            wraplength=wrap,
            justify="left",
        ).pack(anchor="w", pady=(0, 8))

        bold = ("Helvetica", 12, "bold")

        # --- постоянные значения столбцов ---
        ttk.Label(body, text="Значения столбцов", font=bold).pack(anchor="w")
        grid = ttk.Frame(body)
        grid.pack(fill="x", pady=(4, 0))
        grid.columnconfigure(1, weight=1)
        static_vars = {}
        for row, (key, default) in enumerate(defaults["static_values"].items()):
            ttk.Label(grid, text=key).grid(row=row, column=0, sticky="w", pady=2)
            var = tk.StringVar(value=fmt(current["static_values"].get(key, default)))
            entry = ttk.Entry(grid, textvariable=var)
            enable_clipboard_shortcuts(entry)
            entry.grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=2)
            static_vars[key] = (var, default)

        # --- описание ---
        ttk.Label(body, text="Описание товара", font=bold).pack(anchor="w", pady=(12, 0))
        ttk.Label(
            body, text="{artist} заменяется на имя исполнителя.", foreground="gray"
        ).pack(anchor="w")
        desc_text = tk.Text(body, height=9, width=1, wrap="word", font=("Helvetica", 12))
        desc_text.insert("1.0", current["description"])
        enable_clipboard_shortcuts(desc_text)
        desc_text.pack(fill="x", pady=(4, 0))

        # --- названия и оформление ---
        ttk.Label(body, text="Названия и оформление", font=bold).pack(anchor="w", pady=(12, 0))
        misc = ttk.Frame(body)
        misc.pack(fill="x", pady=(4, 0))
        misc.columnconfigure(1, weight=1)

        title_var = tk.StringVar(value=current["poster_title"])
        size_var = tk.StringVar(value=current["size_label"])
        color_var = tk.StringVar(value=current["highlight_color"])
        name_len_var = tk.StringVar(value=fmt(current["max_name_len"]))
        article_len_var = tk.StringVar(value=fmt(current["max_article_len"]))

        def add_row(row, label, var, hint=None):
            ttk.Label(misc, text=label).grid(row=row, column=0, sticky="w", pady=2)
            entry = ttk.Entry(misc, textvariable=var)
            enable_clipboard_shortcuts(entry)
            entry.grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=2)
            if hint:
                ttk.Label(misc, text=hint, foreground="gray").grid(
                    row=row + 1, column=1, sticky="w", padx=(10, 0)
                )
            return entry

        add_row(0, "Название постера", title_var,
                "{artist} — исполнитель, {album} — название альбома")
        add_row(2, "Размер в «Наименование»", size_var,
                "Получится: <название постера> <размер> <цвет>")
        add_row(4, "Цвет выделения столбцов (HEX)", color_var)
        swatch = tk.Label(misc, width=3, relief="solid", borderwidth=1)
        swatch.grid(row=4, column=2, padx=(8, 0))

        def update_swatch(*_):
            color = color_var.get().lstrip("#").strip()
            if re.fullmatch(r"[0-9A-Fa-f]{6}", color):
                swatch.configure(bg="#" + color)

        color_var.trace_add("write", update_swatch)
        update_swatch()

        add_row(5, "Красная ячейка: «Наименование» от, симв.", name_len_var)
        add_row(6, "Красная ячейка: «Артикул продавца» от, симв.", article_len_var)

        def reset():
            for key, (var, default) in static_vars.items():
                var.set(fmt(default))
            desc_text.delete("1.0", "end")
            desc_text.insert("1.0", defaults["description"])
            title_var.set(defaults["poster_title"])
            size_var.set(defaults["size_label"])
            color_var.set(defaults["highlight_color"])
            name_len_var.set(fmt(defaults["max_name_len"]))
            article_len_var.set(fmt(defaults["max_article_len"]))

        ttk.Button(body, text="Сбросить к стандартным", command=reset).pack(
            anchor="w", pady=(12, 0)
        )

        def collect():
            """Значения из полей; ValueError с понятным текстом, если что-то введено неверно."""
            import math

            static = {}
            for key, (var, default) in static_vars.items():
                text = var.get().strip()
                if isinstance(default, (int, float)) and not isinstance(default, bool):
                    number_text = text.replace(",", ".")
                    try:
                        number = int(number_text)
                    except ValueError:
                        try:
                            number = float(number_text)
                        except ValueError:
                            raise ValueError(f"«{key}»: введите число (например, {default}).")
                    if isinstance(number, float) and not math.isfinite(number):
                        raise ValueError(f"«{key}»: введите обычное число.")
                    static[key] = number
                else:
                    static[key] = text

            description = desc_text.get("1.0", "end-1c").strip()
            if not description:
                raise ValueError("Описание товара не может быть пустым.")
            poster_title = title_var.get().strip()
            if not poster_title:
                raise ValueError("Название постера не может быть пустым.")
            color = color_var.get().lstrip("#").strip()
            if not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
                raise ValueError("Цвет выделения — шесть символов HEX, например C0C0C0.")

            lengths = []
            for label, var in (("«Наименование»", name_len_var), ("«Артикул продавца»", article_len_var)):
                text = var.get().strip()
                if not text.isdigit() or int(text) < 1:
                    raise ValueError(f"Длина для {label}: целое число 1 или больше.")
                lengths.append(int(text))

            return {
                "static_values": static,
                "description": description,
                "poster_title": poster_title,
                "size_label": size_var.get().strip(),
                "highlight_color": color.upper(),
                "max_name_len": lengths[0],
                "max_article_len": lengths[1],
            }

        return {"collect": collect, "current": current}

    # ---------- Config ----------

    def _write_artist_to_config(self, artist):
        path = self._config_path()

        try:
            if path.exists():
                with path.open("r", encoding="utf-8") as f:
                    config = json.load(f)
            else:
                config = {}

            if not isinstance(config, dict):
                config = {}

            project = config.get("project")
            if not isinstance(project, dict):
                project = {}

            project["artist"] = artist
            config["project"] = project

            with path.open("w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                f.write("\n")

        except Exception as e:
            self.log(f"⚠ Не удалось сохранить project.artist: {e}")

    # ---------- Background ----------

    def _run_background(self, title, worker, *args):
        if self.busy:
            return

        self._set_busy(True)
        self._prompt_cancelled = False
        self._run_id += 1
        run_id = self._run_id
        self.clear_log()
        self.log(f"=== {title} ===")
        self.set_status(title + "…")
        self.set_progress(0)

        title_for_log = title

        def target():
            try:
                worker(*args)
                self.log("=== Готово ===")
            except (Exception, SystemExit) as e:
                if getattr(threading.current_thread(), "posters_cancelled", False):
                    return  # нажали «Отмена» — сообщение уже выведено в cancel_run
                if self._prompt_cancelled:
                    self.log("⏹ Остановлено пользователем (нажата «Отмена»).")
                    self.set_status("Остановлено")
                    return
                title, text, unexpected = _explain_error(e)
                self.log(f"✗ {text}")
                if unexpected:
                    self.log("Технические подробности:\n" + traceback.format_exc())
                _write_error_log(title_for_log, e)
                self.set_status("Ошибка")
                self.log_queue.put(("error", (title, text)))
            finally:
                self.log_queue.put(("done", run_id))

        self._worker_thread = threading.Thread(target=target, daemon=True)
        self._worker_thread.start()

    def cancel_run(self):
        """Кнопка «Отмена»: обрывает запущенный скрипт и сразу освобождает кнопки."""
        if not self.busy:
            return

        thread = self._worker_thread
        self._run_id += 1  # всё, что поток отправит дальше, будет проигнорировано

        if thread is not None:
            thread.posters_cancelled = True
        proc = self._osascript_proc
        if proc is not None:
            try:
                proc.terminate()
            except Exception:
                pass
        if thread is not None and thread.is_alive():
            _interrupt_thread(thread)

        self.log("⏹ Остановлено пользователем.")
        self.set_status("Остановлено")
        self.set_progress(0)
        self._set_busy(False)

    def open_excel(self):
        if not self._require_project():
            return

        excel = self.project_root / EXCEL_NAME
        if not excel.exists():
            messagebox.showwarning("Excel", f"Не найден файл:\n{excel}")
            return

        subprocess.Popen(["open", str(excel)], cwd=str(self.project_root))


def main():
    _register_bundled_fonts()
    root = tk.Tk()
    try:
        style = ttk.Style()
        # POSTERS_THEME=clam|default|aqua — для проверки отрисовки; по умолчанию aqua
        wanted = os.environ.get("POSTERS_THEME", "aqua")
        if wanted in style.theme_names():
            style.theme_use(wanted)
    except Exception:
        pass

    try:
        PostersApp(root)
    except Exception as exc:
        _write_error_log("запуск приложения", exc)
        messagebox.showerror(
            "Posters не удалось запустить",
            f"{type(exc).__name__}: {exc}\n\nПодробности записаны в файл:\n{ERROR_LOG}",
        )
        return
    root.mainloop()


if __name__ == "__main__":
    main()
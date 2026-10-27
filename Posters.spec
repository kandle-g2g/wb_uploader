# -*- mode: python ; coding: utf-8 -*-
# Обычно запускается через build_mac.sh (он выставляет POSTERS_ARCH = arm64 или x86_64).
# Вручную из папки Mockups:  python3 -m PyInstaller --noconfirm --clean Posters.spec
import os

HERE = SPECPATH  # папка со spec-файлом

# py3pin (Pinterest) и selenium, который он может использовать для входа, подгружаются
# скриптом в рантайме — собираем их целиком вместе с данными и бинарниками.
import importlib.util
from PyInstaller.utils.hooks import collect_all

# Если библиотеки нет в Python, которым идёт сборка, PyInstaller молча её пропускает,
# и приложение падает уже у пользователя («не хватает модуля …»). Поэтому проверяем
# заранее и останавливаем сборку с понятной подсказкой.
_REQUIRED = ("openpyxl", "et_xmlfile", "requests", "PIL", "py3pin", "selenium")
_missing = [m for m in _REQUIRED if importlib.util.find_spec(m) is None]
if _missing:
    raise SystemExit(
        "Для сборки не хватает библиотек: " + ", ".join(_missing) + ".\n"
        "Установите их тем же Python, которым запускается сборка:\n"
        "  python3 -m pip install openpyxl requests pillow py3pin selenium\n"
        "и запустите build_mac.sh снова."
    )

_datas, _bins, _hidden = [], [], []
# py3pin/selenium — как раньше; openpyxl собираем целиком (у него много подмодулей,
# которые подгружаются динамически)
for _pkg in ("py3pin", "selenium", "openpyxl"):
    _d, _b, _h = collect_all(_pkg)
    _datas += _d
    _bins += _b
    _hidden += _h

a = Analysis(
    [os.path.join(HERE, "app.py")],
    pathex=[],
    binaries=_bins,
    # Скрипты кладутся внутрь .app; при запуске приложение копирует их в Mockups/scripts
    datas=_datas + [
        (os.path.join(HERE, "scripts", "mockuper.jsx"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "excel_filler.py"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "wb_uploader.py"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "tiff_jpg_saver.jsx"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "wb_photo_uploader.py"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "wb_stocks.py"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "table_defaults.json"), "bundled_scripts"),
        (os.path.join(HERE, "scripts", "pinterest_uploader.py"), "bundled_scripts"),
        # шрифт заголовка (app.py регистрирует его при запуске, в систему не ставит)
        (os.path.join(HERE, "fonts", "ActayWide-Bold.otf"), "fonts"),
    ],
    # scripts/*.py подгружаются приложением в рантайме, PyInstaller их не видит,
    # поэтому всё, что они импортируют, перечисляем вручную.
    hiddenimports=_hidden + [
        "tkinter", "tkinter.ttk", "tkinter.simpledialog", "tkinter.filedialog", "tkinter.messagebox",
        "requests", "certifi", "urllib3", "idna", "charset_normalizer",
        "openpyxl", "et_xmlfile",
        "PIL", "PIL.Image", "py3pin", "py3pin.Pinterest",
        "openpyxl.formatting.rule", "openpyxl.styles.differential",
        "difflib", "re", "time", "datetime", "json", "builtins", "importlib",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Posters",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=os.environ.get("POSTERS_ARCH") or None,  # arm64 / x86_64; пусто = как у вашего Python
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, upx_exclude=[], name="Posters")
app = BUNDLE(
    coll,
    name="Posters.app",
    icon=os.path.join(HERE, "Posters.icns"),
    bundle_identifier="com.local.posters",
    info_plist={
        "CFBundleName": "Posters",
        "NSHighResolutionCapable": True,
        # без этого ключа macOS может молча не показать запрос
        # «Posters хочет управлять Adobe Photoshop»
        "NSAppleEventsUsageDescription": "Posters запускает скрипты мокапов в Adobe Photoshop.",
        # понятные тексты в системных запросах доступа к папкам
        "NSDocumentsFolderUsageDescription": "Posters работает с папкой Mockups: исходники, мокапы и таблица.",
        "NSDesktopFolderUsageDescription": "Posters работает с папкой Mockups: исходники, мокапы и таблица.",
        "NSDownloadsFolderUsageDescription": "Posters работает с папкой Mockups: исходники, мокапы и таблица.",
        "NSRemovableVolumesUsageDescription": "Posters работает с папкой Mockups на внешнем диске.",
    },
)

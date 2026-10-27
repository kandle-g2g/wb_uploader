#!/bin/zsh
# Сборка Posters.app для macOS.
#
# Запуск из папки Mockups:
#   chmod +x build_mac.sh && ./build_mac.sh
#
# На Mac с чипом Apple собираются ДВЕ версии — для Apple Silicon (нативно) и для Intel
# (через Rosetta); на Intel-маке — только Intel. Результат:
#   Posters_APP/Release/AppleSilicon/Posters.app  + Posters-AppleSilicon.zip
#   Posters_APP/Release/Intel/Posters.app         + Posters-Intel.zip
#   Posters_APP/Posters.app                       — копия версии под ваш Mac
#
# Переменные окружения (необязательно):
#   PYTHON=/путь/к/python3     какой Python использовать (по умолчанию python3 из PATH)
#   ARCHS="arm64"              собрать только одну архитектуру (arm64 или x86_64)
set -euo pipefail
cd "$(dirname "$0")"
ROOT="$(pwd)"
DEST="$(cd "$ROOT/.." && pwd)"
RELEASE="$DEST/Release"
ENV_DIR="$ROOT/.build_env"

die() { echo "ОШИБКА: $*" >&2; exit 1; }

# ---------- Python ----------
PYTHON="${PYTHON:-$(command -v python3 || true)}"
[[ -n "$PYTHON" ]] || die "Python 3 не найден. Установите Python 3.12 или 3.13 с python.org (universal2)."
PY_BIN="$("$PYTHON" -c 'import sys; print(sys.executable)')"
echo "Python: $PY_BIN"
"$PYTHON" --version

PY_ARCHS="$(lipo -archs "$PY_BIN" 2>/dev/null || echo "")"
echo "Архитектуры Python: ${PY_ARCHS:-не определены}"

HOST_ARCH="$(uname -m)"
if [[ -n "${ARCHS:-}" ]]; then
  BUILD_ARCHS=(${=ARCHS})
elif [[ "$HOST_ARCH" == "arm64" ]]; then
  BUILD_ARCHS=(arm64 x86_64)
else
  BUILD_ARCHS=(x86_64)
fi
echo "Будут собраны: ${BUILD_ARCHS[*]}"

# ---------- проверки до долгой сборки ----------
for a in "${BUILD_ARCHS[@]}"; do
  [[ "$PY_ARCHS" == *"$a"* ]] || die "Этот Python не содержит архитектуру $a (есть: ${PY_ARCHS:-?}).
Поставьте универсальный установщик Python (universal2) с python.org: он работает и на Intel, и на Apple Silicon.
Системный /usr/bin/python3 и Python из Homebrew обычно одноархитектурные."
done

if [[ " ${BUILD_ARCHS[*]} " == *" x86_64 "* && "$HOST_ARCH" == "arm64" ]]; then
  arch -x86_64 /usr/bin/true 2>/dev/null || die "Для сборки Intel-версии нужен Rosetta. Установите:
  softwareupdate --install-rosetta --agree-to-license"
fi

for f in app.py Posters.spec \
         scripts/mockuper.jsx scripts/excel_filler.py scripts/wb_uploader.py \
         scripts/tiff_jpg_saver.jsx scripts/wb_photo_uploader.py scripts/wb_stocks.py scripts/table_defaults.json scripts/pinterest_uploader.py; do
  [[ -f "$ROOT/$f" ]] || die "Не найден $ROOT/$f"
done

mkdir -p "$ENV_DIR" "$RELEASE"

# ---------- сборка одной архитектуры ----------
build_arch() {
  local ARCH="$1"
  local LABEL
  [[ "$ARCH" == "arm64" ]] && LABEL="AppleSilicon" || LABEL="Intel"

  echo
  echo "=================== $LABEL ($ARCH) ==================="

  local VENV="$ENV_DIR/$ARCH"
  if [[ ! -x "$VENV/bin/python" ]]; then
    arch "-$ARCH" "$PY_BIN" -m venv "$VENV"
  fi
  local PY="$VENV/bin/python"

  local GOT
  GOT="$(arch "-$ARCH" "$PY" -c 'import platform; print(platform.machine())')"
  [[ "$GOT" == "$ARCH" ]] || die "Окружение $VENV работает как $GOT, а не $ARCH. Удалите папку $ENV_DIR и повторите."

  # Tk должен быть 8.6+ (на системном /usr/bin/python3 бывает 8.5 — окно может не открыться)
  local TKV
  TKV="$(arch "-$ARCH" "$PY" -c 'import tkinter; print(tkinter.TkVersion)' 2>/dev/null || echo "нет")"
  echo "Tk: $TKV"
  [[ "$TKV" != "нет" ]] || die "В этом Python нет tkinter. Поставьте Python с python.org (3.12 или 3.13)."
  case "$TKV" in 8.[0-5]*) die "Tk $TKV слишком старый (нужен 8.6+). Поставьте Python с python.org." ;; esac

  arch "-$ARCH" "$PY" -m pip install --upgrade pip
  arch "-$ARCH" "$PY" -m pip install --upgrade -r requirements.txt et_xmlfile

  rm -rf build dist
  POSTERS_ARCH="$ARCH" arch "-$ARCH" "$PY" -m PyInstaller --noconfirm --clean Posters.spec

  [[ -d dist/Posters.app ]] || die "PyInstaller не создал dist/Posters.app"

  # ad-hoc подпись (без Apple Developer; на Apple Silicon без неё приложение не запустится)
  codesign --force --deep --sign - dist/Posters.app || echo "⚠ codesign не удался — приложение может не запуститься на Apple Silicon"
  codesign --verify --deep dist/Posters.app 2>/dev/null && echo "Подпись: ok" || echo "⚠ подпись не прошла проверку"

  local BIN_ARCHS
  BIN_ARCHS="$(lipo -archs dist/Posters.app/Contents/MacOS/Posters 2>/dev/null || echo "?")"
  echo "Архитектура собранного приложения: $BIN_ARCHS"
  [[ "$BIN_ARCHS" == *"$ARCH"* ]] || die "Собранное приложение не содержит $ARCH."

  # Release/<LABEL>/Posters.app + zip (ditto сохраняет ссылки и подпись, обычный zip их ломает)
  rm -rf "$RELEASE/$LABEL" "$RELEASE/Posters-$LABEL.zip"
  mkdir -p "$RELEASE/$LABEL"
  mv dist/Posters.app "$RELEASE/$LABEL/Posters.app"
  ditto -c -k --sequesterRsrc --keepParent "$RELEASE/$LABEL/Posters.app" "$RELEASE/Posters-$LABEL.zip"
  rm -rf build dist

  # копия под ваш Mac — рядом с Mockups, как раньше
  if [[ "$ARCH" == "$HOST_ARCH" ]]; then
    rm -rf "$DEST/Posters.app"
    ditto "$RELEASE/$LABEL/Posters.app" "$DEST/Posters.app"
    echo "Копия для этого Mac: $DEST/Posters.app"
  fi
}

for a in "${BUILD_ARCHS[@]}"; do
  build_arch "$a"
done

[[ -f "$ROOT/Инструкция.md" ]] && cp "$ROOT/Инструкция.md" "$RELEASE/Инструкция.md"

echo
echo "ГОТОВО. Что отправлять:"
ls -1 "$RELEASE" | grep -E "zip$|md$" | sed 's/^/  Release\//'
echo
echo "  Mac с чипом M1/M2/M3/M4 -> Posters-AppleSilicon.zip"
echo "  Mac с процессором Intel -> Posters-Intel.zip"
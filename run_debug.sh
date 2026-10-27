#!/bin/zsh
# Запуск собранного приложения с выводом в Terminal (видны ошибки, которых нет в логе окна).
# Берёт Posters.app рядом с папкой Mockups (его кладёт build_mac.sh).
set -e
cd "$(dirname "$0")"
APP="$(cd .. && pwd)/Posters.app"
[[ -d "$APP" ]] || { echo "Не найдено: $APP — сначала запустите ./build_mac.sh"; exit 1; }
"$APP/Contents/MacOS/Posters"
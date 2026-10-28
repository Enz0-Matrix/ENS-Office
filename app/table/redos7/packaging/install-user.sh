#!/bin/sh
# Установка ENS-Table для РЕД ОС 7 без прав администратора — в домашнюю папку.
# Запуск: sh install-user.sh   Удаление: sh install-user.sh --remove
set -e
here=$(cd "$(dirname "$0")" && pwd)
dest="$HOME/.local/share/ens-table-r7"
apps="$HOME/.local/share/applications"
icons="$HOME/.local/share/icons/hicolor/128x128/apps"
if [ "$1" = "--remove" ]; then
  rm -rf "$dest" "$apps/ens-table-r7.desktop" "$icons/ens-table-r7.png"
  echo "ENS-Table удалена (ваши книги и временные файлы восстановления не тронуты)."
  exit 0
fi
mkdir -p "$dest" "$apps" "$icons"
cp -r "$here/ens-table-r7" "$here/ui" "$dest/"
chmod +x "$dest/ens-table-r7"
cp "$here/ens-table-r7.png" "$icons/"
sed "s|^Exec=ens-table-r7|Exec=$dest/ens-table-r7|; s|^Icon=ens-table-r7|Icon=$icons/ens-table-r7.png|" \
  "$here/ens-table-r7.desktop" > "$apps/ens-table-r7.desktop"
update-desktop-database -q "$apps" 2>/dev/null || true
echo "Готово: ENS-Table появилась в меню «Офис»."
echo "Чтобы .xlsx открывались двойным щелчком: правая кнопка на файле → «Открыть с помощью» → ENS-Table → «Запомнить»."

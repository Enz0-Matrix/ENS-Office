#!/bin/sh
# Установка ENS-Tekst для РЕД ОС 7 без прав администратора — в домашнюю папку.
# Запуск: sh install-user.sh   Удаление: sh install-user.sh --remove
set -e
here=$(cd "$(dirname "$0")" && pwd)
dest="$HOME/.local/share/ens-text-r7"
apps="$HOME/.local/share/applications"
icons="$HOME/.local/share/icons/hicolor/128x128/apps"
if [ "$1" = "--remove" ]; then
  rm -rf "$dest" "$apps/ens-text-r7.desktop" "$icons/ens-text-r7.png"
  echo "ENS-Tekst удалена (ваши документы и временные файлы восстановления не тронуты)."
  exit 0
fi
mkdir -p "$dest" "$apps" "$icons"
cp -r "$here/ens-text-r7" "$here/ui" "$dest/"
chmod +x "$dest/ens-text-r7"
cp "$here/ens-text-r7.png" "$icons/"
sed "s|^Exec=ens-text-r7|Exec=$dest/ens-text-r7|; s|^Icon=ens-text-r7|Icon=$icons/ens-text-r7.png|" \
  "$here/ens-text-r7.desktop" > "$apps/ens-text-r7.desktop"
update-desktop-database -q "$apps" 2>/dev/null || true
echo "Готово: ENS-Tekst появилась в меню «Офис»."
echo "Чтобы .docx открывались двойным щелчком: правая кнопка на файле → «Открыть с помощью» → ENS-Tekst → «Запомнить»."

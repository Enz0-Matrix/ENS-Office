# Окружение разработки ENS-Office (Windows, Git Bash): source tools/env.sh
# Инструменты ставятся в профиль пользователя без прав администратора:
#   Rust (rustup, x86_64-pc-windows-gnu) — ~/.cargo
#   Node.js portable — %LOCALAPPDATA%/Programs/node-portable
#   dlltool из llvm-mingw (нужен крейтам windows-link) — %LOCALAPPDATA%/Programs/ens-dlltool
export PATH="$HOME/.cargo/bin:$(ls -d "$LOCALAPPDATA"/Programs/node-portable/node-v*/ | tail -1):$LOCALAPPDATA/Programs/ens-dlltool:$PATH"

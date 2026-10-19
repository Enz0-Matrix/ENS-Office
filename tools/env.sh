# Окружение разработки ENS-Office (Windows, Git Bash): source tools/env.sh
# Инструменты ставятся в профиль пользователя без прав администратора:
#   Rust (rustup, x86_64-pc-windows-gnu) — ~/.cargo
#   Node.js portable — %LOCALAPPDATA%/Programs/node-portable
#   wasm-bindgen 0.2.126 (готовая сборка) — %LOCALAPPDATA%/Programs/wasm-bindgen-0.2.126
#   dlltool из llvm-mingw (нужен крейтам windows-link) — %LOCALAPPDATA%/Programs/ens-dlltool
export PATH="$HOME/.cargo/bin:$LOCALAPPDATA/Programs/wasm-bindgen-0.2.126:$(ls -d "$LOCALAPPDATA"/Programs/node-portable/node-v*/ | tail -1):$LOCALAPPDATA/Programs/ens-dlltool:$PATH"

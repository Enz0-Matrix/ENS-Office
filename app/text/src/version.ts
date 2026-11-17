// Версия ENS-Text: единый источник — поле version в app/text/package.json (его же читает CI при выпуске).
// tauri.conf.json берёт версию оттуда же ("version": "../package.json"); Cargo.toml (src-tauri, redos7)
// хранят свою копию — их совпадение проверяет тест «версия: один номер…».
import { version } from '../package.json';

export const APP_VERSION: string = version;

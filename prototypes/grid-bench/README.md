# grid-bench — сравнение Univer и своей сетки с IronCalc

Результаты и выводы: [docs/research/proto-01-grid-bench.md](../../docs/research/proto-01-grid-bench.md).

## Подготовка (один раз)
Нужен Node.js 24+. Сторонние файлы в репозиторий не кладём, они собираются локально:

```bash
cd prototypes/grid-bench/univer-build
npm install
npm run build
cp node_modules/@univerjs/preset-sheets-core/lib/index.css ../vendor/univer/
mkdir -p ../vendor/ironcalc/snippets/ironcalc_base-aa721fe913c88060
cd ../vendor/ironcalc
for f in wasm.js wasm_bg.wasm snippets/ironcalc_base-aa721fe913c88060/inline0.js; do curl -sfL -o $f https://cdn.jsdelivr.net/npm/@ironcalc/wasm@0.8.4/$f; done
```

## Запуск
```bash
python -m http.server 8731 --directory prototypes/grid-bench
```
Вручную: `http://localhost:8731/ironcalc.html?rows=10000&lookup=0` (или `univer.html`).

Автоматически (headless Chrome, путь меняется переменной `BENCH_BROWSER`):
```bash
cd prototypes/grid-bench
node run-bench.mjs 1000 10000          # с ВПР
BENCH_LOOKUP=0 node run-bench.mjs 50000 # без ВПР
```

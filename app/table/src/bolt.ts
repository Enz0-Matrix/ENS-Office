// «У. Болт» — необязательный помощник (выключен по умолчанию). Болт-спринтер, скептик
// и саркастичный социофоб: подсказывает по ошибкам формул и больше всего любит, когда
// его кладут. Живёт в углу окна, не трогает данные книги. Рисунок — свой (SVG + CSS).

import { saveSettings, settings } from './settings';

/** Что делать с ошибкой в ячейке: «что это» и «как чинить». */
const ERRORS: [RegExp, string][] = [
  [/^#(DIV\/0|ДЕЛ\/0)!/, 'Деление на ноль. Математика такого не прощает. Проверь делитель или оберни формулу в ЕСЛИОШИБКА.'],
  [/^#(REF|ССЫЛКА)!/, 'Ссылка на то, чего уже нет: строку, столбец или лист удалили. Поправь ссылку в формуле или верни удалённое сочетанием Ctrl+Z.'],
  [/^#(NAME|ИМЯ)\?/, 'Незнакомое имя. Опечатка в названии функции? Или имя ещё не задано? Загляни в «Формулы → Диспетчер имён».'],
  [/^#(VALUE|ЗНАЧ)!/, 'Не тот тип данных: например, текст там, где ждут число. Проверь ячейки, на которые ссылается формула.'],
  [/^#(N\/A|Н\/Д)/, 'Нет данных. ВПР/ПОИСКПОЗ не нашли то, что искали. Проверь искомое значение и тип совпадения.'],
  [/^#(NUM|ЧИСЛО)!/, 'Число не годится: слишком большое, или корень из отрицательного. Проверь аргументы.'],
  [/^#SPILL!/, 'Результату некуда разлиться: справа или снизу что-то мешает. Освободи место.'],
  [/^#CIRC!/, 'Формула ссылается сама на себя, прямо или через другие ячейки. Разорви круг.'],
  [/^#(NULL|ПУСТО)!/, 'Пересечение диапазонов пусто. Скорее всего, пробел там, где нужна запятая или точка с запятой.'],
  [/^#(ERROR|CALC)!/, 'Формула написана неверно. Проверь скобки, запятые и точки с запятой.'],
  [/^#N\/IMPL!/, 'Эта функция пока не поддерживается. Бывает.'],
];

const OPENERS = ['Опять?', 'Ну вот.', 'Я так и знал.', 'Не смотри на меня.', 'Хм.'];
const IDLE = [
  'Я просто болт. Но даже мне видно, что формулы можно упростить.',
  'Если что — я не здесь.',
  'Сохраняйся почаще. Я серьёзно. Мне всё равно, но тебе будет не всё равно.',
  'Ctrl+Z — лучшая функция Excel. И ENS-Table.',
  'Положи меня, пожалуйста. Мне так спокойнее.',
  'Люди — это страшно. Таблицы — нет.',
];
const LAID = ['Ну, положили — и ладно.', 'Вот теперь хорошо.', 'Горизонтально — значит, правильно.', 'Спасибо. Не ожидал.'];
const STOOD = ['Опять стоять…', 'Ладно, постою.', 'Вертикально. Как неудобно.'];

const pick = <T>(a: T[]): T => a[Math.floor(Math.random() * a.length)];

const SVG = `<svg viewBox="0 0 80 120" width="72" height="108" aria-hidden="true">
  <polygon points="14,4 66,4 76,22 66,40 14,40 4,22" fill="#8a96a3" stroke="#4a5560" stroke-width="3" stroke-linejoin="round"/>
  <rect x="26" y="40" width="28" height="70" rx="3" fill="#a7b2bd" stroke="#4a5560" stroke-width="3"/>
  <g stroke="#4a5560" stroke-width="2.5"><line x1="26" y1="54" x2="54" y2="60"/><line x1="26" y1="70" x2="54" y2="76"/><line x1="26" y1="86" x2="54" y2="92"/><line x1="26" y1="102" x2="54" y2="108"/></g>
  <circle cx="28" cy="22" r="7" fill="#fff" stroke="#4a5560" stroke-width="2"/><circle cx="52" cy="22" r="7" fill="#fff" stroke="#4a5560" stroke-width="2"/>
  <circle class="bolt-pupil" cx="30" cy="24" r="3" fill="#222"/><circle class="bolt-pupil" cx="50" cy="24" r="3" fill="#222"/>
  <path d="M19 15 L36 19" stroke="#4a5560" stroke-width="3" stroke-linecap="round"/><path d="M61 15 L44 19" stroke="#4a5560" stroke-width="3" stroke-linecap="round"/>
  <path d="M30 34 Q40 30 50 34" fill="none" stroke="#4a5560" stroke-width="2.5" stroke-linecap="round"/>
</svg>`;

let root: HTMLElement | null = null;
let bubble: HTMLElement;
let lying = false;
let bubbleTimer = 0;
let idleTimer = 0;
let lastKey = '';

function say(text: string, ms = 9000): void {
  if (!root) return;
  bubble.textContent = text;
  bubble.hidden = false;
  // Облачко висит над Болтом и не сдвигает его; у края экрана — снизу или вправо.
  const r = root.getBoundingClientRect();
  bubble.classList.toggle('below', r.top < bubble.offsetHeight + 8);
  bubble.classList.toggle('left', r.right < bubble.offsetWidth + 8);
  clearTimeout(bubbleTimer);
  bubbleTimer = window.setTimeout(() => (bubble.hidden = true), ms);
}

function lay(down: boolean): void {
  lying = down;
  root!.classList.toggle('lying', down);
  say(pick(down ? LAID : STOOD), 5000);
}

function scheduleIdle(): void {
  clearTimeout(idleTimer);
  idleTimer = window.setTimeout(() => {
    if (root && bubble.hidden && !lying) say(pick(IDLE), 7000);
    scheduleIdle();
  }, 150_000 + Math.random() * 150_000);
}

function build(): void {
  root = document.createElement('div');
  root.className = 'bolt';
  root.innerHTML = `<div class="bolt-bubble" hidden></div><div class="bolt-body" title="У. Болт. Перетащите — положите; двойной щелчок — уложить или поставить">${SVG}</div><button class="bolt-x" title="Убрать помощника (вернуть: Файл → Параметры)">✕</button>`;
  bubble = root.querySelector('.bolt-bubble')!;
  document.body.append(root);
  const body = root.querySelector<HTMLElement>('.bolt-body')!;
  root.querySelector('.bolt-x')!.addEventListener('click', () => { settings.helper = false; boltApply(); saveSettings(); });
  body.addEventListener('dblclick', () => lay(!lying));
  body.addEventListener('click', () => { if (!bubble.hidden) bubble.hidden = true; else say(pick(IDLE), 6000); });
  let drag: { dx: number; dy: number; x: number; y: number; moved: boolean } | null = null;
  body.addEventListener('pointerdown', (e) => {
    const r = root!.getBoundingClientRect();
    drag = { dx: e.clientX - r.left, dy: e.clientY - r.top, x: e.clientX, y: e.clientY, moved: false };
    body.setPointerCapture(e.pointerId);
  });
  body.addEventListener('pointermove', (e) => {
    if (!drag) return;
    // Палец при касании чуть дрожит: это ещё не перетаскивание, иначе Болт
    // «убегает» и ложится от каждого нажатия, а двойное касание не ловится.
    if (!drag.moved && Math.hypot(e.clientX - drag.x, e.clientY - drag.y) < 10) return;
    drag.moved = true;
    root!.style.left = `${Math.max(0, Math.min(innerWidth - 80, e.clientX - drag.dx))}px`;
    root!.style.top = `${Math.max(0, Math.min(innerHeight - 100, e.clientY - drag.dy))}px`;
    root!.style.right = root!.style.bottom = 'auto';
  });
  body.addEventListener('pointerup', () => {
    if (drag?.moved) { drag = null; lay(true); } // перетащили и бросили — «положили»
    drag = null;
  });
  scheduleIdle();
  say('Ну, привет. Я У. Болт. Не обращай на меня внимания.', 6000);
}

/** Включить или выключить помощника по настройке. */
export function boltApply(): void {
  if (settings.helper && !root) build();
  else if (!settings.helper && root) { clearTimeout(bubbleTimer); clearTimeout(idleTimer); root.remove(); root = null; lastKey = ''; }
}

/** Выбрана ячейка: если в ней ошибка — прокомментировать (один раз на ячейку). */
export function boltCell(key: string, shown: string): void {
  if (!root || !shown.startsWith('#')) { if (!shown.startsWith('#')) lastKey = ''; return; }
  if (key === lastKey) return;
  lastKey = key;
  const hit = ERRORS.find(([re]) => re.test(shown));
  if (hit) { if (lying) lay(false); say(`${pick(OPENERS)} ${shown} — ${hit[1]}`, 14000); }
}

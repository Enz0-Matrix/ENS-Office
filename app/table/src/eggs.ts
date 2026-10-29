// Пасхалки: безобидные, на данные книги не влияют.
//  1) =ENS() в ячейке — версия и благодарность;
//  2) конфетти при первом сохранении книги на 1000+ строк;
//  3) семь щелчков по логотипу в «Файле» — экран авторов.

import { saveSettings, settings } from './settings';
import { APP_VERSION } from './version';

export const ensMessage = () => `ENS-Office · ENS-Table ${APP_VERSION}. Спасибо, что пользуетесь! Вы делаете этот проект осмысленным.`;

/** Ввод «=ENS()» (в любом регистре, с пробелами) — пасхалка, а не формула. */
export const isEnsEgg = (text: string) => /^=\s*ens\s*\(\s*\)\s*$/i.test(text);

export function confetti(): void {
  if (matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const cv = document.createElement('canvas');
  cv.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;z-index:90;pointer-events:none';
  cv.width = innerWidth; cv.height = innerHeight;
  document.body.append(cv);
  const ctx = cv.getContext('2d')!;
  const colors = ['#e74c3c', '#f1c40f', '#2ecc71', '#3498db', '#9b59b6', '#e67e22'];
  const bits = Array.from({ length: 140 }, () => ({
    x: Math.random() * cv.width, y: -Math.random() * cv.height * 0.5, vx: (Math.random() - 0.5) * 3, vy: 2 + Math.random() * 4,
    s: 4 + Math.random() * 6, a: Math.random() * 6, va: (Math.random() - 0.5) * 0.3, c: colors[Math.floor(Math.random() * colors.length)],
  }));
  const t0 = performance.now();
  const tick = (t: number) => {
    ctx.clearRect(0, 0, cv.width, cv.height);
    for (const b of bits) {
      b.x += b.vx; b.y += b.vy; b.a += b.va;
      ctx.save(); ctx.translate(b.x, b.y); ctx.rotate(b.a); ctx.fillStyle = b.c; ctx.fillRect(-b.s / 2, -b.s / 4, b.s, b.s / 2); ctx.restore();
    }
    if (t - t0 < 3000) requestAnimationFrame(tick); else cv.remove();
  };
  requestAnimationFrame(tick);
}

/** После успешного сохранения: первая книга на 1000+ строк — конфетти (один раз навсегда). */
export function maybeConfetti(rows: number): void {
  if (settings.bigBookCelebrated || rows < 1000) return;
  settings.bigBookCelebrated = true;
  saveSettings();
  confetti();
}

export function showCredits(): void {
  const d = document.createElement('div');
  d.className = 'credits';
  d.innerHTML = `<div class="credits-box"><h2>ENS-Office</h2>
    <p>ENS-Table ${APP_VERSION}</p>
    <p>Идея, постановка задач, испытания: <b>Н. Едренников</b></p>
    <p>Движок таблиц: IronCalc (MIT/Apache-2.0), доработан под русский интерфейс</p>
    <p>Помощник: <b>У. Болт</b> (ему всё равно, но он рад)</p>
    <p class="credits-small">Нажмите в любом месте, чтобы закрыть</p></div>`;
  d.addEventListener('click', () => d.remove());
  document.body.append(d);
  confetti();
}

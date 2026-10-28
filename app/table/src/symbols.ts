// «Символ»: таблица специальных знаков (Вставка → Символ), как в Excel. Щелчок вставляет знак
// в редактируемую ячейку; окно остаётся открытым, можно вставить несколько знаков подряд.

const SETS: [string, string][] = [
  ['Математика', '± × ÷ − ≠ ≈ ≤ ≥ ∞ √ ∛ ∑ ∏ ∫ ∂ ∆ ∇ ∈ ∉ ⊂ ⊃ ∪ ∩ ∧ ∨ ¬ ∀ ∃ ∅ ‰ ‱ ° ′ ″ ² ³ ⁿ ¼ ½ ¾ ⅓ ⅔ ⅛ ⅜ ⅝ ⅞'],
  ['Греческие', 'α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω Α Β Γ Δ Ε Ζ Η Θ Ι Κ Λ Μ Ν Ξ Ο Π Ρ Σ Τ Υ Φ Χ Ψ Ω Ω'],
  ['Валюты', '₽ $ € £ ¥ ¢ ₴ ₸ ₹ ₿ ₩ ₺ ₪ ₫ ฿ ₡ ₦ ₱ ₲ ₵ ₭ ₮ ¤'],
  ['Стрелки', '← ↑ → ↓ ↔ ↕ ↖ ↗ ↘ ↙ ⇐ ⇑ ⇒ ⇓ ⇔ ⇕ ↩ ↪ ↺ ↻ ➔ ➜ ➤ ▲ ▼ ◄ ►'],
  ['Знаки и пунктуация', '© ® ™ § ¶ † ‡ • ‣ … ‹ › « » „ “ ” ‘ ’ ‚ ‐ – — ‽ ¿ ¡ № ℃ ℉ ℓ ℮ ☎ ✉ ✓ ✔ ✗ ✘ ★ ☆ ♥ ♦ ♠ ♣ ☀ ☁ ☂ ☺ ☹'],
  ['Индексы', '⁰ ¹ ² ³ ⁴ ⁵ ⁶ ⁷ ⁸ ⁹ ⁺ ⁻ ⁼ ⁽ ⁾ ₀ ₁ ₂ ₃ ₄ ₅ ₆ ₇ ₈ ₉ ₊ ₋ ₌ ₍ ₎'],
  ['Рамки и фигуры', '─ │ ┌ ┐ └ ┘ ├ ┤ ┬ ┴ ┼ ═ ║ ╔ ╗ ╚ ╝ ■ □ ▪ ▫ ● ○ ◆ ◇ ◉ ◌ ▬ ▭ ▮ ▯'],
  ['Римские цифры', 'Ⅰ Ⅱ Ⅲ Ⅳ Ⅴ Ⅵ Ⅶ Ⅷ Ⅸ Ⅹ Ⅺ Ⅻ ⅰ ⅱ ⅲ ⅳ ⅴ ⅵ ⅶ ⅷ ⅸ ⅹ ⅺ ⅻ'],
];

export function openSymbols(insert: (text: string) => void): void {
  const overlay = document.getElementById('recovery')!;
  let set = 0;
  const draw = () => {
    const chars = [...new Set(SETS[set][1].split(' ').filter(Boolean))];
    overlay.innerHTML = `<div class="modal symbols-modal">
      <h2>Символ</h2>
      <div class="sym-tabs">${SETS.map(([n], i) => `<button data-set="${i}" class="${i === set ? 'on' : ''}">${n}</button>`).join('')}</div>
      <div class="sym-grid">${chars.map((c) => `<button data-ch="${c}" title="U+${c.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}">${c}</button>`).join('')}</div>
      <p class="sym-hint">Щелчок вставляет знак в ячейку, которую вы редактируете (или в выделенную). Окно можно не закрывать.</p>
      <div class="rec-actions"><button data-act="close" class="primary">Закрыть</button></div></div>`;
    overlay.hidden = false;
    const box = overlay.querySelector<HTMLElement>('.modal')!;
    box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') overlay.hidden = true; });
    box.addEventListener('mousedown', (e) => { if ((e.target as HTMLElement).closest('[data-ch]')) e.preventDefault(); });
    box.addEventListener('click', (e) => {
      const t = e.target as HTMLElement;
      const tab = t.closest<HTMLElement>('[data-set]');
      if (tab) { set = Number(tab.dataset.set); draw(); return; }
      const ch = t.closest<HTMLElement>('[data-ch]')?.dataset.ch;
      if (ch) { insert(ch); return; }
      if (t.closest('[data-act="close"]')) overlay.hidden = true;
    });
  };
  draw();
}

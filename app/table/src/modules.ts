// Модули ENS-Table и их уровень лицензии (ADR-008). Лента и команды
// спрашивают available(): недоступная возможность показывается с замком,
// а не исчезает — пользователь видит, что она есть.

export type Tier = 'free' | 'pro';

export interface ModuleInfo {
  id: string;
  name: string;
  tier: Tier;
}

export const MODULES: ModuleInfo[] = [
  { id: 'core', name: 'Таблицы: ввод, формулы, файлы, восстановление', tier: 'free' },
  { id: 'format', name: 'Оформление: шрифты, заливка, границы, объединение', tier: 'free' },
  { id: 'view', name: 'Вид: закрепление областей, фокус ячейки', tier: 'free' },
  // Будущие модули по лицензии для организаций, например:
  // { id: 'links', name: 'Связи между книгами', tier: 'pro' },
  // { id: 'collab', name: 'Совместная работа', tier: 'pro' },
];

/** Лицензия: пока тестовый период — всё открыто. */
export interface License { kind: 'personal' | 'education' | 'commercial' | 'trial'; until?: string }

let license: License = { kind: 'trial' };

export function currentLicense(): License {
  return license;
}

export function setLicense(next: License): void {
  license = next;
}

/** Доступен ли модуль при текущей лицензии. */
export function available(moduleId: string | undefined): boolean {
  if (!moduleId) return true;
  const m = MODULES.find((x) => x.id === moduleId);
  if (!m || m.tier === 'free') return true;
  return license.kind !== 'commercial' || !!license.until; // тестовый период: открыто
}

/** Ссылка «Поддержать проект» (пусто — кнопка скрыта). */
export const DONATE_URL = '';

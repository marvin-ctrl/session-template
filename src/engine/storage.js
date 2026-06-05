import { EMPTY_STATE, STORAGE_KEY } from './config.js';

export function loadStoredState() {
  if (typeof window === 'undefined') return structuredClone(EMPTY_STATE);
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return normalizeStoredState(raw ? JSON.parse(raw) : EMPTY_STATE);
  } catch {
    return structuredClone(EMPTY_STATE);
  }
}

export function saveStoredState(state) {
  if (typeof window === 'undefined') return;
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(normalizeStoredState(state)));
}

export function normalizeStoredState(value) {
  if (!value || typeof value !== 'object') return structuredClone(EMPTY_STATE);
  return {
    ...structuredClone(EMPTY_STATE),
    ...value,
    roster: Array.isArray(value.roster) ? value.roster : [],
    history: Array.isArray(value.history) ? value.history : [],
    currentSession: value.currentSession ?? null,
  };
}

export function exportStateJson(state) {
  return JSON.stringify(normalizeStoredState(state), null, 2);
}

export function importStateJson(json) {
  const parsed = JSON.parse(json);
  return normalizeStoredState(parsed);
}

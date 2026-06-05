import { describe, expect, it } from 'vitest';
import { EMPTY_STATE, SESSION_CONFIG } from './config.js';
import { exportStateJson, importStateJson, normalizeStoredState } from './storage.js';

describe('storage helpers', () => {
  it('normalizes invalid stored shapes back to safe defaults', () => {
    const normalized = normalizeStoredState({
      roster: 'bad',
      history: null,
      currentSession: undefined,
      config: { wrong: true },
    });

    expect(normalized.roster).toEqual([]);
    expect(normalized.history).toEqual([]);
    expect(normalized.currentSession).toBeNull();
    expect(normalized.config).toEqual({ wrong: true });
  });

  it('exports and imports JSON backup state', () => {
    const state = {
      ...EMPTY_STATE,
      currentSession: { id: 'session-1', status: 'active' },
      history: [{ id: 'session-0', status: 'completed' }],
      config: SESSION_CONFIG,
    };

    const imported = importStateJson(exportStateJson(state));
    expect(imported.currentSession.id).toBe('session-1');
    expect(imported.history).toHaveLength(1);
    expect(imported.version).toBe(1);
  });

  it('throws when importing invalid JSON', () => {
    expect(() => importStateJson('{not-json')).toThrow();
  });
});

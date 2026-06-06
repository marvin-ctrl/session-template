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
    expect(normalized.config).toEqual({ ...SESSION_CONFIG, wrong: true });
  });

  it('treats older saved players without availability as active', () => {
    const normalized = normalizeStoredState({
      currentSession: {
        id: 'session-1',
        status: 'active',
        players: [
          { id: 'a', name: 'Ari' },
          { id: 'b', name: 'Blake' },
          { id: 'c', name: 'Casey' },
          { id: 'd', name: 'Devin' },
        ],
        rounds: [],
        currentRound: null,
        upcomingRound: null,
      },
    });

    expect(normalized.currentSession.players.every((player) => player.isActive)).toBe(true);
    expect(normalized.currentSession.upcomingRound.teamAIds).toHaveLength(2);
    expect(normalized.currentSession.upcomingRound.teamBIds).toHaveLength(2);
    expect(normalized.currentSession.upcomingRound.restingIds).toHaveLength(0);
  });

  it('repairs stale lineups that put every active player in rest', () => {
    const normalized = normalizeStoredState({
      currentSession: {
        id: 'session-1',
        status: 'active',
        players: [
          { id: 'a', name: 'Ari', isActive: true },
          { id: 'b', name: 'Blake', isActive: true },
          { id: 'c', name: 'Casey', isActive: true },
          { id: 'd', name: 'Devin', isActive: true },
          { id: 'e', name: 'Emery', isActive: true },
        ],
        rounds: [],
        currentRound: null,
        upcomingRound: {
          id: 'round-bad',
          number: 1,
          eligiblePlayerIds: ['a', 'b', 'c', 'd', 'e'],
          teamAIds: [],
          teamBIds: [],
          restingIds: ['a', 'b', 'c', 'd', 'e'],
        },
      },
    });

    expect(normalized.currentSession.upcomingRound.teamAIds).toHaveLength(2);
    expect(normalized.currentSession.upcomingRound.teamBIds).toHaveLength(2);
    expect(normalized.currentSession.upcomingRound.restingIds).toHaveLength(1);
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

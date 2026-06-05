import { describe, expect, it, vi } from 'vitest';
import {
  calculateSessionStats,
  generateRound,
  getLeaderboard,
  getTeamSize,
  recordResult,
  startSession,
} from './sessionEngine.js';

const names = ['Ari', 'Blake', 'Casey', 'Devin', 'Emery', 'Finn', 'Gray', 'Harper', 'Indy', 'Jules'];

describe('session engine', () => {
  it('maps active player counts to team sizes', () => {
    expect(getTeamSize(4)).toBe(2);
    expect(getTeamSize(5)).toBe(2);
    expect(getTeamSize(6)).toBe(3);
    expect(getTeamSize(7)).toBe(3);
    expect(getTeamSize(8)).toBe(4);
    expect(getTeamSize(9)).toBe(4);
    expect(getTeamSize(10)).toBe(5);
    expect(getTeamSize(15)).toBe(5);
  });

  it('generates a round with the correct players resting', () => {
    const session = startSession(names.slice(0, 7));
    const round = session.upcomingRound;
    expect(round.teamAIds).toHaveLength(3);
    expect(round.teamBIds).toHaveLength(3);
    expect(round.restingIds).toHaveLength(1);
  });

  it('awards points to winners, losers, and draws', () => {
    let session = startSession(names.slice(0, 4));
    session = { ...session, currentRound: session.upcomingRound, upcomingRound: null };
    session = recordResult(session, 'teamA');

    const stats = calculateSessionStats(session.players, session.rounds);
    const teamAPlayer = session.rounds[0].teamAIds[0];
    const teamBPlayer = session.rounds[0].teamBIds[0];

    expect(stats.statsById[teamAPlayer].points).toBe(2);
    expect(stats.statsById[teamAPlayer].wins).toBe(1);
    expect(stats.statsById[teamBPlayer].points).toBe(1);
    expect(stats.statsById[teamBPlayer].losses).toBe(1);
  });

  it('sorts leaderboard by points, wins, draws, then name', () => {
    const players = [
      { id: 'a', name: 'Zara', isActive: true },
      { id: 'b', name: 'Mika', isActive: true },
      { id: 'c', name: 'Ari', isActive: true },
    ];
    const rows = getLeaderboard(players, {
      a: { points: 4, wins: 1, draws: 2 },
      b: { points: 4, wins: 2, draws: 0 },
      c: { points: 4, wins: 1, draws: 2 },
    });

    expect(rows.map((row) => row.name)).toEqual(['Mika', 'Ari', 'Zara']);
  });

  it('prioritizes players who have rested when generating the next round', () => {
    vi.spyOn(Date, 'now').mockReturnValue(1_700_000_000_000);
    const session = startSession(names.slice(0, 5));
    const firstRound = { ...session.upcomingRound, result: 'draw' };
    const generated = generateRound(session.players, [firstRound]);

    expect(generated.ok).toBe(true);
    const selectedNext = new Set([...generated.round.teamAIds, ...generated.round.teamBIds]);
    expect(selectedNext.has(firstRound.restingIds[0])).toBe(true);
    vi.restoreAllMocks();
  });
});

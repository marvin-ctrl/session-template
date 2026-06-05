import { describe, expect, it, vi } from 'vitest';
import {
  calculateSessionStats,
  describeResult,
  formatGameType,
  generateRound,
  getLeaderboard,
  getTeamSize,
  getWinnerIds,
  parseRosterInput,
  recordResult,
  startSession,
} from './sessionEngine.js';

const names = ['Ari', 'Blake', 'Casey', 'Devin', 'Emery', 'Finn', 'Gray', 'Harper', 'Indy', 'Jules', 'Kai', 'Lee', 'Morgan', 'Noa', 'Ocean', 'Pax'];

describe('session engine', () => {
  it('normalizes, splits, and de-duplicates roster input', () => {
    expect(parseRosterInput(' Ari\nBlake,  ari  , Casey  Jones ,,')).toEqual(['Ari', 'Blake', 'Casey Jones']);
  });

  it('rejects sessions below minimum or above maximum roster size', () => {
    expect(() => startSession(names.slice(0, 3))).toThrow('Start with 4-15 players.');
    expect(() => startSession(names.slice(0, 16))).toThrow('Start with 4-15 players.');
  });

  it('maps active player counts to team sizes', () => {
    expect(getTeamSize(4)).toBe(2);
    expect(getTeamSize(5)).toBe(2);
    expect(getTeamSize(6)).toBe(3);
    expect(getTeamSize(7)).toBe(3);
    expect(getTeamSize(8)).toBe(4);
    expect(getTeamSize(9)).toBe(4);
    expect(getTeamSize(10)).toBe(5);
    expect(getTeamSize(15)).toBe(5);
    expect(formatGameType(3)).toBe('1v1');
    expect(formatGameType(10)).toBe('5v5');
  });

  it.each([
    [4, 2, 0],
    [5, 2, 1],
    [6, 3, 0],
    [7, 3, 1],
    [8, 4, 0],
    [9, 4, 1],
    [10, 5, 0],
    [11, 5, 1],
    [15, 5, 5],
  ])('generates correct teams and rests for %i active players', (count, teamSize, restCount) => {
    const session = startSession(names.slice(0, count));
    expect(session.upcomingRound.teamAIds).toHaveLength(teamSize);
    expect(session.upcomingRound.teamBIds).toHaveLength(teamSize);
    expect(session.upcomingRound.restingIds).toHaveLength(restCount);
    expect(session.upcomingRound.eligiblePlayerIds).toHaveLength(count);
  });

  it('uses only active players when generating future rounds', () => {
    const session = startSession(names.slice(0, 6));
    const players = session.players.map((player, index) => (index < 2 ? { ...player, isActive: false } : player));
    const generated = generateRound(players, []);

    expect(generated.ok).toBe(true);
    expect(generated.round.teamAIds).toHaveLength(2);
    expect(generated.round.teamBIds).toHaveLength(2);
    expect(generated.round.eligiblePlayerIds).toHaveLength(4);
  });

  it('blocks round generation below minimum active players and past max rounds', () => {
    const session = startSession(names.slice(0, 4));
    const inactivePlayers = session.players.map((player, index) => (index === 0 ? { ...player, isActive: false } : player));
    expect(generateRound(inactivePlayers, []).ok).toBe(false);

    const completedRounds = Array.from({ length: 8 }, (_, index) => ({
      ...session.upcomingRound,
      id: `round-${index}`,
      number: index + 1,
      result: 'draw',
    }));
    expect(generateRound(session.players, completedRounds).ok).toBe(false);
  });

  it('awards points to Team A winners and Team B losers', () => {
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

  it('awards points to Team B winners and Team A losers', () => {
    let session = startSession(names.slice(0, 4));
    session = { ...session, currentRound: session.upcomingRound, upcomingRound: null };
    session = recordResult(session, 'teamB');

    const stats = calculateSessionStats(session.players, session.rounds);
    const teamAPlayer = session.rounds[0].teamAIds[0];
    const teamBPlayer = session.rounds[0].teamBIds[0];

    expect(stats.statsById[teamBPlayer].points).toBe(2);
    expect(stats.statsById[teamBPlayer].wins).toBe(1);
    expect(stats.statsById[teamAPlayer].points).toBe(1);
    expect(stats.statsById[teamAPlayer].losses).toBe(1);
  });

  it('awards draw points to all playing players and zero points to resting players', () => {
    let session = startSession(names.slice(0, 5));
    session = { ...session, currentRound: session.upcomingRound, upcomingRound: null };
    session = recordResult(session, 'draw');

    const stats = calculateSessionStats(session.players, session.rounds);
    const playedIds = [...session.rounds[0].teamAIds, ...session.rounds[0].teamBIds];
    for (const playerId of playedIds) {
      expect(stats.statsById[playerId].points).toBe(1);
      expect(stats.statsById[playerId].draws).toBe(1);
      expect(stats.statsById[playerId].gamesPlayed).toBe(1);
    }
    const restingPlayer = session.rounds[0].restingIds[0];
    expect(stats.statsById[restingPlayer].points).toBe(0);
    expect(stats.statsById[restingPlayer].rests).toBe(1);
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

  it('returns shared winners when leaders match points and wins', () => {
    const players = [
      { id: 'a', name: 'Ari', isActive: true },
      { id: 'b', name: 'Blake', isActive: true },
      { id: 'c', name: 'Casey', isActive: true },
    ];
    const rows = getLeaderboard(players, {
      a: { points: 4, wins: 1, draws: 2 },
      b: { points: 4, wins: 1, draws: 1 },
      c: { points: 3, wins: 1, draws: 1 },
    });

    expect(getWinnerIds(rows)).toEqual(['a', 'b']);
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

  it('completes the session when recording the max-round result', () => {
    let session = startSession(names.slice(0, 4));
    const priorRounds = Array.from({ length: 7 }, (_, index) => ({
      ...session.upcomingRound,
      id: `round-${index}`,
      number: index + 1,
      result: 'draw',
    }));
    session = {
      ...session,
      rounds: priorRounds,
      currentRound: { ...session.upcomingRound, id: 'final-round', number: 8 },
      upcomingRound: null,
    };

    const completed = recordResult(session, 'teamA');
    expect(completed.status).toBe('completed');
    expect(completed.rounds).toHaveLength(8);
    expect(completed.finishedAt).toBeTruthy();
  });

  it('describes result labels', () => {
    expect(describeResult('teamA')).toBe('Team A won');
    expect(describeResult('teamB')).toBe('Team B won');
    expect(describeResult('draw')).toBe('Draw');
  });
});

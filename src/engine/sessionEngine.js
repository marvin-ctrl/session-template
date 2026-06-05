import { EMPTY_STATE, SESSION_CONFIG } from './config.js';

export function createId(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

export function normalizeName(name) {
  return name.trim().replace(/\s+/g, ' ');
}

export function parseRosterInput(input) {
  const seen = new Set();
  return input
    .split(/[\n,]+/)
    .map(normalizeName)
    .filter(Boolean)
    .filter((name) => {
      const key = name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

export function createPlayer(name) {
  return {
    id: createId('player'),
    name: normalizeName(name),
    isActive: true,
    createdAt: new Date().toISOString(),
  };
}

export function createState() {
  return structuredClone(EMPTY_STATE);
}

export function getTeamSize(activeCount, config = SESSION_CONFIG) {
  return Math.max(0, Math.min(config.maxTeamSize, Math.floor(activeCount / 2)));
}

export function formatGameType(activeCount, config = SESSION_CONFIG) {
  const size = getTeamSize(activeCount, config);
  return size > 0 ? `${size}v${size}` : 'None';
}

export function startSession(rosterNames, config = SESSION_CONFIG) {
  const roster = parseRosterInput(rosterNames.join('\n')).map(createPlayer);
  if (roster.length < config.minPlayers || roster.length > config.maxPlayers) {
    throw new Error(`Start with ${config.minPlayers}-${config.maxPlayers} players.`);
  }

  const session = {
    id: createId('session'),
    status: 'active',
    createdAt: new Date().toISOString(),
    finishedAt: null,
    players: roster,
    rounds: [],
    currentRound: null,
    upcomingRound: null,
    undoSnapshot: null,
  };

  const generated = generateRound(session.players, session.rounds, config);
  return { ...session, upcomingRound: generated.round };
}

export function getActivePlayers(players) {
  return players.filter((player) => player.isActive);
}

export function buildPairKey(a, b) {
  return a < b ? `${a}::${b}` : `${b}::${a}`;
}

export function blankStats() {
  return {
    gamesPlayed: 0,
    points: 0,
    wins: 0,
    draws: 0,
    losses: 0,
    rests: 0,
    currentRestStreak: 0,
    lastPlayedRound: null,
    lastRestRound: null,
    eligibilityCount: 0,
  };
}

export function calculateSessionStats(players, rounds, config = SESSION_CONFIG) {
  const statsById = Object.fromEntries(players.map((player) => [player.id, blankStats()]));
  const teammateCounts = {};
  const opponentCounts = {};

  for (const round of rounds) {
    const eligible = new Set(round.eligiblePlayerIds);
    const teamA = new Set(round.teamAIds);
    const teamB = new Set(round.teamBIds);

    for (const player of players) {
      const stats = statsById[player.id] ?? blankStats();
      statsById[player.id] = stats;

      if (!eligible.has(player.id)) {
        stats.currentRestStreak = 0;
        continue;
      }

      stats.eligibilityCount += 1;
      const played = teamA.has(player.id) || teamB.has(player.id);

      if (!played) {
        stats.rests += 1;
        stats.currentRestStreak += 1;
        stats.lastRestRound = round.number;
        continue;
      }

      stats.gamesPlayed += 1;
      stats.currentRestStreak = 0;
      stats.lastPlayedRound = round.number;

      if (round.result === 'draw') {
        stats.draws += 1;
        stats.points += config.drawPoints;
      } else if ((round.result === 'teamA' && teamA.has(player.id)) || (round.result === 'teamB' && teamB.has(player.id))) {
        stats.wins += 1;
        stats.points += config.winPoints;
      } else {
        stats.losses += 1;
        stats.points += config.lossPoints;
      }
    }

    countTeamPairs(round.teamAIds, teammateCounts);
    countTeamPairs(round.teamBIds, teammateCounts);
    countOpponentPairs(round.teamAIds, round.teamBIds, opponentCounts);
  }

  return {
    statsById,
    teammateCounts,
    opponentCounts,
    leaderboard: getLeaderboard(players, statsById),
  };
}

function countTeamPairs(ids, counts) {
  for (let i = 0; i < ids.length; i += 1) {
    for (let j = i + 1; j < ids.length; j += 1) {
      const key = buildPairKey(ids[i], ids[j]);
      counts[key] = (counts[key] ?? 0) + 1;
    }
  }
}

function countOpponentPairs(teamAIds, teamBIds, counts) {
  for (const a of teamAIds) {
    for (const b of teamBIds) {
      const key = buildPairKey(a, b);
      counts[key] = (counts[key] ?? 0) + 1;
    }
  }
}

export function getLeaderboard(players, statsById) {
  return players
    .map((player) => ({
      playerId: player.id,
      name: player.name,
      isActive: player.isActive,
      ...(statsById[player.id] ?? blankStats()),
    }))
    .sort(compareLeaderboardRows);
}

export function compareLeaderboardRows(a, b) {
  return b.points - a.points || b.wins - a.wins || b.draws - a.draws || a.name.localeCompare(b.name) || a.playerId.localeCompare(b.playerId);
}

export function getWinnerIds(leaderboard) {
  if (leaderboard.length === 0) return [];
  const leader = leaderboard[0];
  return leaderboard.filter((row) => row.points === leader.points && row.wins === leader.wins).map((row) => row.playerId);
}

export function generateRound(players, rounds, config = SESSION_CONFIG) {
  const roundNumber = rounds.length + 1;
  if (roundNumber > config.maxRounds) {
    return { ok: false, error: `The session has reached ${config.maxRounds} rounds.` };
  }

  const activePlayers = getActivePlayers(players);
  if (activePlayers.length < config.minPlayers) {
    return { ok: false, error: `At least ${config.minPlayers} active players are needed.` };
  }

  const teamSize = getTeamSize(activePlayers.length, config);
  const stats = calculateSessionStats(players, rounds, config);
  const ordered = [...activePlayers].sort((a, b) => compareForSelection(a, b, stats.statsById, roundNumber));
  const selected = ordered.slice(0, teamSize * 2);
  const resting = ordered.slice(teamSize * 2);
  const split = chooseTeamSplit(selected.map((player) => player.id), stats);

  return {
    ok: true,
    round: {
      id: createId('round'),
      number: roundNumber,
      eligiblePlayerIds: sortIdsByName(activePlayers.map((player) => player.id), players),
      teamAIds: sortIdsByName(split.teamAIds, players),
      teamBIds: sortIdsByName(split.teamBIds, players),
      restingIds: sortIdsByName(resting.map((player) => player.id), players),
      createdAt: new Date().toISOString(),
      createdBy: 'generator',
    },
  };
}

export function compareForSelection(a, b, statsById, roundNumber) {
  const aStats = statsById[a.id] ?? blankStats();
  const bStats = statsById[b.id] ?? blankStats();
  const aGap = aStats.lastPlayedRound === null ? roundNumber : roundNumber - aStats.lastPlayedRound;
  const bGap = bStats.lastPlayedRound === null ? roundNumber : roundNumber - bStats.lastPlayedRound;

  return (
    aStats.gamesPlayed - bStats.gamesPlayed ||
    bStats.currentRestStreak - aStats.currentRestStreak ||
    bStats.rests - aStats.rests ||
    bGap - aGap ||
    a.name.localeCompare(b.name) ||
    a.id.localeCompare(b.id)
  );
}

export function chooseTeamSplit(selectedIds, stats) {
  const teamSize = selectedIds.length / 2;
  const anchor = selectedIds[0];
  const rest = selectedIds.slice(1);
  let best = null;

  function visit(start, picked) {
    if (picked.length === teamSize - 1) {
      const teamAIds = [anchor, ...picked];
      const teamASet = new Set(teamAIds);
      const teamBIds = selectedIds.filter((id) => !teamASet.has(id));
      const metrics = scoreTeamSplit(teamAIds, teamBIds, stats);
      if (!best || compareSplitMetrics(metrics, best.metrics) < 0) {
        best = { teamAIds, teamBIds, metrics };
      }
      return;
    }

    for (let i = start; i < rest.length; i += 1) {
      visit(i + 1, [...picked, rest[i]]);
    }
  }

  visit(0, []);
  return best ?? { teamAIds: selectedIds.slice(0, teamSize), teamBIds: selectedIds.slice(teamSize), metrics: null };
}

export function scoreTeamSplit(teamAIds, teamBIds, stats) {
  let teammateRepeatPenalty = 0;
  let repeatedTeammatePairs = 0;
  let newTeammatePairs = 0;
  let opponentRepeatPenalty = 0;
  let repeatedOpponentPairs = 0;

  for (const [a, b] of [...getPairs(teamAIds), ...getPairs(teamBIds)]) {
    const count = stats.teammateCounts[buildPairKey(a, b)] ?? 0;
    teammateRepeatPenalty += count;
    if (count > 0) repeatedTeammatePairs += 1;
    else newTeammatePairs += 1;
  }

  for (const a of teamAIds) {
    for (const b of teamBIds) {
      const count = stats.opponentCounts[buildPairKey(a, b)] ?? 0;
      opponentRepeatPenalty += count;
      if (count > 0) repeatedOpponentPairs += 1;
    }
  }

  return {
    teammateRepeatPenalty,
    repeatedTeammatePairs,
    newTeammatePairs,
    opponentRepeatPenalty,
    repeatedOpponentPairs,
    teamKey: `${teamAIds.join('|')}__${teamBIds.join('|')}`,
  };
}

function getPairs(ids) {
  const pairs = [];
  for (let i = 0; i < ids.length; i += 1) {
    for (let j = i + 1; j < ids.length; j += 1) pairs.push([ids[i], ids[j]]);
  }
  return pairs;
}

export function compareSplitMetrics(a, b) {
  return (
    a.teammateRepeatPenalty - b.teammateRepeatPenalty ||
    a.repeatedTeammatePairs - b.repeatedTeammatePairs ||
    b.newTeammatePairs - a.newTeammatePairs ||
    a.opponentRepeatPenalty - b.opponentRepeatPenalty ||
    a.repeatedOpponentPairs - b.repeatedOpponentPairs ||
    a.teamKey.localeCompare(b.teamKey)
  );
}

export function recordResult(session, result, config = SESSION_CONFIG) {
  if (!session.currentRound) return session;
  const completedRound = {
    ...session.currentRound,
    result,
    recordedAt: new Date().toISOString(),
  };
  const rounds = [...session.rounds, completedRound];
  const isComplete = rounds.length >= config.maxRounds;
  return {
    ...session,
    status: isComplete ? 'completed' : 'active',
    rounds,
    currentRound: null,
    upcomingRound: null,
    finishedAt: isComplete ? new Date().toISOString() : session.finishedAt,
    undoSnapshot: {
      status: session.status,
      rounds: session.rounds,
      currentRound: session.currentRound,
      upcomingRound: session.upcomingRound,
      finishedAt: session.finishedAt,
    },
  };
}

export function sortIdsByName(ids, players) {
  const byId = Object.fromEntries(players.map((player) => [player.id, player]));
  return [...ids].sort((a, b) => (byId[a]?.name ?? a).localeCompare(byId[b]?.name ?? b) || a.localeCompare(b));
}

export function describeResult(result) {
  if (result === 'teamA') return 'Team A won';
  if (result === 'teamB') return 'Team B won';
  return 'Draw';
}

export function getPlayerNames(ids, players) {
  const byId = Object.fromEntries(players.map((player) => [player.id, player]));
  return ids.map((id) => byId[id]?.name ?? id);
}

import { EMPTY_STATE, STORAGE_KEY } from './config.js';
import { generateRound, getTeamSize } from './sessionEngine.js';

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
  const config = normalizeConfig(value.config);

  return {
    ...structuredClone(EMPTY_STATE),
    ...value,
    config,
    roster: normalizePlayers(value.roster),
    history: Array.isArray(value.history) ? value.history.map((session) => normalizeSession(session, config)).filter(Boolean) : [],
    currentSession: normalizeSession(value.currentSession, config),
  };
}

export function exportStateJson(state) {
  return JSON.stringify(normalizeStoredState(state), null, 2);
}

export function importStateJson(json) {
  const parsed = JSON.parse(json);
  return normalizeStoredState(parsed);
}

function normalizeConfig(config) {
  return {
    ...EMPTY_STATE.config,
    ...(config && typeof config === 'object' ? config : {}),
  };
}

function normalizeSession(session, config) {
  if (!session || typeof session !== 'object') return null;

  const players = normalizePlayers(session.players);
  const rounds = Array.isArray(session.rounds) ? session.rounds.map(normalizeRound).filter(Boolean) : [];
  let currentRound = normalizeRound(session.currentRound);
  let upcomingRound = normalizeRound(session.upcomingRound);

  if (session.status !== 'completed' && !currentRound && !isRoundUsable(upcomingRound, players, config)) {
    const generated = generateRound(players, rounds, config);
    upcomingRound = generated.ok ? generated.round : null;
  }

  if (currentRound && !isRoundUsable(currentRound, players, config)) {
    currentRound = null;
  }

  if (upcomingRound && !isRoundUsable(upcomingRound, players, config)) {
    const generated = generateRound(players, rounds, config);
    upcomingRound = generated.ok ? generated.round : null;
  }

  return {
    ...session,
    status: session.status === 'completed' ? 'completed' : 'active',
    players,
    rounds,
    currentRound,
    upcomingRound: currentRound ? null : upcomingRound,
    undoSnapshot: session.undoSnapshot ?? null,
  };
}

function normalizePlayers(players) {
  if (!Array.isArray(players)) return [];

  return players
    .filter((player) => player && typeof player === 'object' && typeof player.name === 'string' && player.name.trim())
    .map((player, index) => ({
      ...player,
      id: typeof player.id === 'string' && player.id ? player.id : `player-${index + 1}`,
      name: player.name.trim().replace(/\s+/g, ' '),
      isActive: player.isActive !== false,
    }));
}

function normalizeRound(round) {
  if (!round || typeof round !== 'object') return null;

  return {
    ...round,
    eligiblePlayerIds: normalizeIdList(round.eligiblePlayerIds),
    teamAIds: normalizeIdList(round.teamAIds),
    teamBIds: normalizeIdList(round.teamBIds),
    restingIds: normalizeIdList(round.restingIds),
  };
}

function normalizeIdList(ids) {
  return Array.isArray(ids) ? ids.filter((id) => typeof id === 'string' && id) : [];
}

function isRoundUsable(round, players, config) {
  if (!round) return false;
  const activeIds = new Set(players.filter((player) => player.isActive).map((player) => player.id));
  const activeCount = activeIds.size;
  if (activeCount < config.minPlayers) return false;

  const teamSize = getTeamSize(activeCount, config);
  const playingIds = [...round.teamAIds, ...round.teamBIds];
  const allRoundIds = [...playingIds, ...round.restingIds];
  const uniqueRoundIds = new Set(allRoundIds);

  return (
    round.teamAIds.length === teamSize &&
    round.teamBIds.length === teamSize &&
    playingIds.every((id) => activeIds.has(id)) &&
    allRoundIds.length === uniqueRoundIds.size
  );
}

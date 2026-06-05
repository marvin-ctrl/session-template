export const STORAGE_KEY = 'futsal-coach-session-v1';

export const SESSION_CONFIG = {
  minPlayers: 4,
  maxPlayers: 15,
  maxTeamSize: 5,
  maxRounds: 8,
  gameSeconds: 240,
  breakSeconds: 120,
  winPoints: 2,
  drawPoints: 1,
  lossPoints: 1,
};

export const EMPTY_STATE = {
  version: 1,
  config: SESSION_CONFIG,
  roster: [],
  currentSession: null,
  history: [],
};

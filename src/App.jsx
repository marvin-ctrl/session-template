import { useEffect, useMemo, useReducer, useRef, useState } from 'react';
import {
  Activity,
  Download,
  History,
  Pause,
  Play,
  RotateCcw,
  Save,
  TimerReset,
  Trophy,
  Upload,
  Users,
} from 'lucide-react';
import { SESSION_CONFIG } from './engine/config.js';
import {
  calculateSessionStats,
  createPlayer,
  describeResult,
  formatGameType,
  generateRound,
  getPlayerNames,
  getWinnerIds,
  parseRosterInput,
  recordResult,
  startSession,
} from './engine/sessionEngine.js';
import { exportStateJson, importStateJson, loadStoredState, saveStoredState } from './engine/storage.js';
import { adjustTimerDuration, formatClock } from './engine/timer.js';

const tabs = [
  ['current', 'Current', Activity],
  ['players', 'Players', Users],
  ['leaderboard', 'Table', Trophy],
  ['history', 'History', History],
  ['backup', 'Backup', Save],
];

function appReducer(state, action) {
  switch (action.type) {
    case 'START_SESSION': {
      const session = startSession(action.names, state.config);
      return { ...state, roster: session.players, currentSession: session };
    }
    case 'RESET_TO_SETUP':
      return { ...state, currentSession: null };
    case 'GENERATE_ROUND': {
      const session = state.currentSession;
      if (!session || session.currentRound || session.upcomingRound) return state;
      const generated = generateRound(session.players, session.rounds, state.config);
      return generated.ok ? { ...state, currentSession: { ...session, upcomingRound: generated.round } } : state;
    }
    case 'START_ROUND': {
      const session = state.currentSession;
      if (!session?.upcomingRound || session.currentRound) return state;
      return { ...state, currentSession: { ...session, currentRound: session.upcomingRound, upcomingRound: null } };
    }
    case 'RECORD_RESULT': {
      const session = recordResult(state.currentSession, action.result, state.config);
      const history = session.status === 'completed' ? upsertHistory(state.history, session) : state.history;
      return { ...state, currentSession: session, history };
    }
    case 'UNDO_RESULT': {
      const session = state.currentSession;
      if (!session?.undoSnapshot) return state;
      return { ...state, currentSession: { ...session, ...session.undoSnapshot, undoSnapshot: null } };
    }
    case 'FINISH_SESSION': {
      const session = {
        ...state.currentSession,
        status: 'completed',
        currentRound: null,
        upcomingRound: null,
        finishedAt: state.currentSession.finishedAt ?? new Date().toISOString(),
      };
      return { ...state, currentSession: session, history: upsertHistory(state.history, session) };
    }
    case 'ADD_PLAYER': {
      const name = action.name.trim();
      const session = state.currentSession;
      if (!name || !session || session.players.length >= state.config.maxPlayers) return state;
      if (session.players.some((player) => player.name.toLowerCase() === name.toLowerCase())) return state;
      const player = createPlayer(name);
      return { ...state, currentSession: { ...session, players: [...session.players, player] } };
    }
    case 'TOGGLE_PLAYER': {
      const session = state.currentSession;
      if (!session) return state;
      const players = session.players.map((player) => (player.id === action.playerId ? { ...player, isActive: !player.isActive } : player));
      return { ...state, currentSession: { ...session, players } };
    }
    case 'IMPORT_STATE':
      return action.state;
    default:
      return state;
  }
}

function upsertHistory(history, session) {
  const saved = {
    ...session,
    currentRound: null,
    upcomingRound: null,
    undoSnapshot: null,
    status: 'completed',
    finishedAt: session.finishedAt ?? new Date().toISOString(),
  };
  return [saved, ...history.filter((item) => item.id !== session.id)].slice(0, 60);
}

export default function App() {
  const [state, dispatch] = useReducer(appReducer, undefined, loadStoredState);
  const [activeTab, setActiveTab] = useState(() => (loadStoredState().currentSession ? 'current' : 'setup'));
  const session = state.currentSession;
  const stats = useMemo(() => (session ? calculateSessionStats(session.players, session.rounds, state.config) : null), [session, state.config]);

  useEffect(() => {
    saveStoredState(state);
  }, [state]);

  const startNewSession = () => {
    if (!session || window.confirm('Start a new session and keep previous completed history?')) {
      dispatch({ type: 'RESET_TO_SETUP' });
      setActiveTab('setup');
    }
  };

  return (
    <div className="app-shell">
      <header className="top-bar">
        <div>
          <span>Mobile courtside</span>
          <strong>Futsal Session Coach</strong>
        </div>
        {session ? (
          <button className="ghost-button" onClick={startNewSession}>
            New
          </button>
        ) : null}
      </header>

      <main>
        {!session || activeTab === 'setup' ? (
          <SetupScreen
            maxPlayers={state.config.maxPlayers}
            minPlayers={state.config.minPlayers}
            onStart={(names) => {
              dispatch({ type: 'START_SESSION', names });
              setActiveTab('current');
            }}
          />
        ) : null}
        {session && activeTab === 'current' ? <CurrentRound session={session} stats={stats} dispatch={dispatch} /> : null}
        {session && activeTab === 'players' ? <PlayersScreen session={session} stats={stats} dispatch={dispatch} /> : null}
        {session && activeTab === 'leaderboard' ? <LeaderboardScreen stats={stats} players={session.players} session={session} /> : null}
        {activeTab === 'history' ? <HistoryScreen history={state.history} /> : null}
        {activeTab === 'backup' ? <BackupScreen state={state} dispatch={dispatch} /> : null}
      </main>

      {session ? (
        <nav className="bottom-nav" aria-label="Primary">
          {tabs.map(([id, label, Icon]) => (
            <button key={id} className={activeTab === id ? 'active' : ''} onClick={() => setActiveTab(id)}>
              <Icon size={18} />
              <span>{label}</span>
            </button>
          ))}
        </nav>
      ) : null}
    </div>
  );
}

function SetupScreen({ minPlayers, maxPlayers, onStart }) {
  const [input, setInput] = useState('');
  const names = parseRosterInput(input);
  const ready = names.length >= minPlayers && names.length <= maxPlayers;

  return (
    <section className="setup-layout">
      <div className="intro-panel">
        <p className="eyebrow">Session Setup</p>
        <h1>Paste the roster, generate round one, run the court.</h1>
        <p>Works offline after install. Sessions stay on this device unless exported.</p>
        <div className="metric-row">
          <Metric label="Players" value={`${minPlayers}-${maxPlayers}`} />
          <Metric label="Games" value="2v2-5v5" />
          <Metric label="Scoring" value="2/1/1" />
        </div>
      </div>
      <div className="work-panel">
        <label htmlFor="roster">Roster import</label>
        <textarea
          id="roster"
          value={input}
          onChange={(event) => setInput(event.target.value)}
          rows={12}
          placeholder={'Ari\nBlake\nCasey\nDevin'}
        />
        <div className="setup-footer">
          <span className={ready ? 'ready-label ready' : 'ready-label'}>{names.length} ready</span>
          <button disabled={!ready} onClick={() => onStart(names)}>
            <Play size={18} />
            Start Session
          </button>
        </div>
        {names.length > maxPlayers ? <p className="error-text">Remove {names.length - maxPlayers} players to stay under the session cap.</p> : null}
        {names.length > 0 && names.length < minPlayers ? <p className="muted-text">Add {minPlayers - names.length} more players to start.</p> : null}
        <div className="chip-list">
          {names.map((name) => (
            <span key={name}>{name}</span>
          ))}
        </div>
      </div>
    </section>
  );
}

function CurrentRound({ session, stats, dispatch }) {
  const activeCount = session.players.filter((player) => player.isActive).length;
  const latestRound = session.rounds.at(-1);
  const visibleRound = session.currentRound ?? session.upcomingRound;
  const canGenerate = session.status === 'active' && !session.currentRound && !session.upcomingRound && session.rounds.length < SESSION_CONFIG.maxRounds && activeCount >= SESSION_CONFIG.minPlayers;

  if (!visibleRound) {
    return (
      <section className="stack">
        <StatusPanel
          title={latestRound ? `Round ${latestRound.number} saved` : 'Session ready'}
          subtitle={`${session.rounds.length}/${SESSION_CONFIG.maxRounds} rounds logged`}
          value={canGenerate ? 'Generate lineup' : 'Waiting'}
        />
        {latestRound ? <RoundSummary title="Last result" round={latestRound} players={session.players} /> : null}
        <button className="primary-action" disabled={!canGenerate} onClick={() => dispatch({ type: 'GENERATE_ROUND' })}>
          Generate Next Round
        </button>
        <button className="secondary-action" disabled={!session.undoSnapshot} onClick={() => dispatch({ type: 'UNDO_RESULT' })}>
          <RotateCcw size={18} />
          Undo Last Result
        </button>
        {session.rounds.length > 0 ? (
          <button className="secondary-action" onClick={() => dispatch({ type: 'FINISH_SESSION' })}>
            Finish Session
          </button>
        ) : null}
      </section>
    );
  }

  const isUpcoming = Boolean(session.upcomingRound);
  return (
    <section className="stack">
      <StatusPanel
        title={isUpcoming ? `Round ${visibleRound.number} lineup` : `Round ${visibleRound.number} live`}
        subtitle={`${formatGameType(activeCount)} · ${visibleRound.restingIds.length} resting`}
        value={isUpcoming ? 'Ready' : 'In play'}
      />
      <TeamGrid round={visibleRound} players={session.players} />
      {isUpcoming ? (
        <button className="primary-action" onClick={() => dispatch({ type: 'START_ROUND' })}>
          <Play size={18} />
          Start Round
        </button>
      ) : (
        <>
          <GameTimer key={visibleRound.id} seconds={SESSION_CONFIG.gameSeconds} />
          <div className="result-dock">
            <button onClick={() => dispatch({ type: 'RECORD_RESULT', result: 'teamA' })}>Team A Won</button>
            <button onClick={() => dispatch({ type: 'RECORD_RESULT', result: 'draw' })}>Draw</button>
            <button onClick={() => dispatch({ type: 'RECORD_RESULT', result: 'teamB' })}>Team B Won</button>
          </div>
        </>
      )}
      <LeaderboardMini rows={stats.leaderboard.slice(0, 3)} />
      <button className="secondary-action" disabled={session.currentRound} onClick={() => dispatch({ type: 'FINISH_SESSION' })}>
        Finish Session
      </button>
    </section>
  );
}

function GameTimer({ seconds }) {
  const [durationSeconds, setDurationSeconds] = useState(seconds);
  const [remainingSeconds, setRemainingSeconds] = useState(seconds);
  const [isRunning, setIsRunning] = useState(true);
  const elapsedPercent = ((durationSeconds - remainingSeconds) / durationSeconds) * 100;
  const isFinished = remainingSeconds === 0;

  useEffect(() => {
    if (!isRunning || isFinished) return undefined;
    const interval = window.setInterval(() => {
      setRemainingSeconds((current) => {
        if (current <= 1) {
          setIsRunning(false);
          return 0;
        }
        return current - 1;
      });
    }, 1000);
    return () => window.clearInterval(interval);
  }, [isRunning, isFinished]);

  const resetTimer = () => {
    setRemainingSeconds(durationSeconds);
    setIsRunning(false);
  };

  const adjustDuration = (changeSeconds) => {
    setRemainingSeconds((currentRemaining) => {
      const adjusted = adjustTimerDuration(durationSeconds, currentRemaining, changeSeconds, isFinished);
      setDurationSeconds(adjusted.durationSeconds);
      return adjusted.remainingSeconds;
    });
  };

  return (
    <section className={isFinished ? 'timer-panel timer-panel-finished' : 'timer-panel'}>
      <div className="timer-panel-header">
        <div>
          <p className="eyebrow">Game Timer</p>
          <strong>{formatClock(remainingSeconds)}</strong>
        </div>
        <span>{isFinished ? 'Time' : isRunning ? 'Running' : 'Paused'}</span>
      </div>
      <div className="timer-track" aria-hidden="true">
        <span style={{ width: `${elapsedPercent}%` }} />
      </div>
      <div className="timer-adjust" aria-label="Adjust timer duration">
        <button onClick={() => adjustDuration(-30)} disabled={durationSeconds <= 60}>
          -30s
        </button>
        <strong>{formatClock(durationSeconds)}</strong>
        <button onClick={() => adjustDuration(30)} disabled={durationSeconds >= 900}>
          +30s
        </button>
      </div>
      <div className="timer-actions">
        <button onClick={() => setIsRunning((current) => !current)} disabled={isFinished}>
          {isRunning ? <Pause size={18} /> : <Play size={18} />}
          {isRunning ? 'Pause' : 'Start'}
        </button>
        <button className="secondary-action" onClick={resetTimer}>
          <TimerReset size={18} />
          Reset
        </button>
      </div>
    </section>
  );
}

function PlayersScreen({ session, stats, dispatch }) {
  const [name, setName] = useState('');
  const lowerNames = new Set(session.players.map((player) => player.name.toLowerCase()));
  const canAdd = name.trim() && !lowerNames.has(name.trim().toLowerCase()) && session.players.length < SESSION_CONFIG.maxPlayers;

  return (
    <section className="stack">
      <StatusPanel title="Player List" subtitle="Availability applies to future generated rounds." value={`${session.players.length} total`} />
      <div className="inline-add">
        <input value={name} onChange={(event) => setName(event.target.value)} placeholder="Late arrival" />
        <button
          disabled={!canAdd}
          onClick={() => {
            dispatch({ type: 'ADD_PLAYER', name });
            setName('');
          }}
        >
          Add
        </button>
      </div>
      <div className="player-list">
        {stats.leaderboard.map((row) => (
          <div key={row.playerId} className="player-row">
            <div>
              <strong>{row.name}</strong>
              <span>
                {row.points} pts · {row.gamesPlayed} GP · {row.rests} rests
              </span>
            </div>
            <button className={row.isActive ? 'toggle active' : 'toggle'} onClick={() => dispatch({ type: 'TOGGLE_PLAYER', playerId: row.playerId })}>
              {row.isActive ? 'Active' : 'Inactive'}
            </button>
          </div>
        ))}
      </div>
    </section>
  );
}

function LeaderboardScreen({ stats, players, session }) {
  const winnerIds = getWinnerIds(stats.leaderboard);
  const winnerNames = getPlayerNames(winnerIds, players).join(', ');

  return (
    <section className="stack">
      <StatusPanel title={session.status === 'completed' ? 'Final Leaderboard' : 'Live Leaderboard'} subtitle={winnerNames || 'No results yet'} value={`${session.rounds.length} rounds`} />
      <div className="leader-list">
        {stats.leaderboard.map((row, index) => (
          <div key={row.playerId} className="leader-row">
            <span className="rank">{index + 1}</span>
            <div>
              <strong>{row.name}</strong>
              <span>
                {row.gamesPlayed} GP · {row.wins}-{row.draws}-{row.losses} · {row.rests} rests
              </span>
            </div>
            <strong className="points">{row.points}</strong>
          </div>
        ))}
      </div>
    </section>
  );
}

function HistoryScreen({ history }) {
  return (
    <section className="stack">
      <StatusPanel title="Session History" subtitle="Completed sessions stored on this device." value={history.length} />
      {history.length === 0 ? <p className="muted-text">No completed sessions yet.</p> : null}
      {history.map((session) => {
        const stats = calculateSessionStats(session.players, session.rounds);
        const winnerNames = getPlayerNames(getWinnerIds(stats.leaderboard), session.players).join(', ');
        return (
          <article className="history-item" key={session.id}>
            <div>
              <strong>{winnerNames || 'No winner'}</strong>
              <span>{new Date(session.finishedAt ?? session.createdAt).toLocaleDateString()}</span>
            </div>
            <span>
              {session.rounds.length} rounds · {stats.leaderboard[0]?.points ?? 0} pts
            </span>
          </article>
        );
      })}
    </section>
  );
}

function BackupScreen({ state, dispatch }) {
  const [text, setText] = useState('');
  const fileRef = useRef(null);

  const downloadBackup = () => {
    const blob = new Blob([exportStateJson(state)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `futsal-session-coach-${new Date().toISOString().slice(0, 10)}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const importBackup = () => {
    try {
      dispatch({ type: 'IMPORT_STATE', state: importStateJson(text) });
      setText('');
      alert('Backup imported.');
    } catch {
      alert('That JSON backup could not be imported.');
    }
  };

  return (
    <section className="stack">
      <StatusPanel title="Backup" subtitle="Export or restore all local sessions." value="JSON" />
      <button className="primary-action" onClick={downloadBackup}>
        <Download size={18} />
        Export Backup
      </button>
      <input
        ref={fileRef}
        type="file"
        accept="application/json"
        hidden
        onChange={async (event) => {
          const file = event.target.files?.[0];
          if (file) setText(await file.text());
        }}
      />
      <button className="secondary-action" onClick={() => fileRef.current?.click()}>
        <Upload size={18} />
        Choose JSON File
      </button>
      <textarea value={text} onChange={(event) => setText(event.target.value)} rows={8} placeholder="Paste exported JSON here" />
      <button disabled={!text.trim()} onClick={importBackup}>
        Import Backup
      </button>
    </section>
  );
}

function StatusPanel({ title, subtitle, value }) {
  return (
    <div className="status-panel">
      <div>
        <p className="eyebrow">{title}</p>
        <h1>{subtitle}</h1>
      </div>
      <strong>{value}</strong>
    </div>
  );
}

function TeamGrid({ round, players }) {
  return (
    <div className="team-grid">
      <TeamColumn title="Team A" names={getPlayerNames(round.teamAIds, players)} tone="a" />
      <TeamColumn title="Team B" names={getPlayerNames(round.teamBIds, players)} tone="b" />
      <TeamColumn title="Rest" names={getPlayerNames(round.restingIds, players)} tone="rest" />
    </div>
  );
}

function TeamColumn({ title, names, tone }) {
  return (
    <section className={`team-column ${tone}`}>
      <h2>{title}</h2>
      {names.length === 0 ? <span className="muted-text">None</span> : null}
      {names.map((name) => (
        <span key={name}>{name}</span>
      ))}
    </section>
  );
}

function RoundSummary({ title, round, players }) {
  return (
    <article className="round-summary">
      <div>
        <p className="eyebrow">{title}</p>
        <strong>{describeResult(round.result)}</strong>
      </div>
      <TeamGrid round={round} players={players} />
    </article>
  );
}

function LeaderboardMini({ rows }) {
  return (
    <div className="mini-table">
      <p className="eyebrow">Top three</p>
      {rows.map((row, index) => (
        <div key={row.playerId}>
          <span>{index + 1}</span>
          <strong>{row.name}</strong>
          <b>{row.points}</b>
        </div>
      ))}
    </div>
  );
}

function Metric({ label, value }) {
  return (
    <span className="metric">
      <small>{label}</small>
      <strong>{value}</strong>
    </span>
  );
}

// @vitest-environment jsdom
import '@testing-library/jest-dom/vitest';
import { render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import App from './App.jsx';

const sevenPlayerRoster = ['Ari', 'Blake', 'Casey', 'Devin', 'Emery', 'Finn', 'Gray'].join('\n');

describe('App courtside flow', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    const store = new Map();
    Object.defineProperty(window, 'localStorage', {
      configurable: true,
      value: {
        getItem: (key) => store.get(key) ?? null,
        setItem: (key, value) => store.set(key, String(value)),
        removeItem: (key) => store.delete(key),
        clear: () => store.clear(),
      },
    });
    window.localStorage.clear();
    vi.spyOn(window, 'alert').mockImplementation(() => {});
    vi.spyOn(window, 'confirm').mockReturnValue(true);
  });

  it('keeps Start Session disabled until the roster reaches the minimum', async () => {
    const user = userEvent.setup();
    render(<App />);

    const startButton = screen.getByRole('button', { name: /start session/i });
    expect(startButton).toBeDisabled();

    await user.type(screen.getByLabelText(/roster import/i), 'Ari\nBlake\nCasey');
    expect(screen.getByText(/add 1 more player/i)).toBeInTheDocument();
    expect(startButton).toBeDisabled();
  });

  it('runs setup, live round, adjustable timer, scoring, backup, and history screens', async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.type(screen.getByLabelText(/roster import/i), sevenPlayerRoster);
    await user.click(screen.getByRole('button', { name: /start session/i }));

    expect(screen.getByText(/round 1 lineup/i)).toBeInTheDocument();
    expect(screen.getByText(/3v3 · 1 resting/i)).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /start round/i }));
    expect(screen.getByText(/game timer/i)).toBeInTheDocument();
    expect(screen.getAllByText('4:00').length).toBeGreaterThan(0);

    await user.click(screen.getByRole('button', { name: /\+30s/i }));
    expect(screen.getAllByText('4:30').length).toBeGreaterThan(0);

    await user.click(screen.getByRole('button', { name: /-30s/i }));
    expect(screen.getAllByText('4:00').length).toBeGreaterThan(0);

    await user.click(screen.getByRole('button', { name: /pause/i }));
    expect(screen.getByRole('button', { name: /^start$/i })).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /reset/i }));
    expect(screen.getAllByText('4:00').length).toBeGreaterThan(0);

    await user.click(screen.getByRole('button', { name: /team a won/i }));
    expect(screen.getByText(/round 1 saved/i)).toBeInTheDocument();
    expect(screen.getByText(/team a won/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /undo last result/i })).toBeEnabled();
    expect(screen.getByRole('button', { name: /finish session/i })).toBeEnabled();

    await user.click(screen.getByRole('button', { name: /table/i }));
    expect(screen.getByText(/live leaderboard/i)).toBeInTheDocument();
    expect(screen.getAllByText(/1-0-0/).length).toBeGreaterThan(0);

    await user.click(screen.getByRole('button', { name: /players/i }));
    expect(screen.getByText(/player list/i)).toBeInTheDocument();
    expect(screen.getAllByRole('button', { name: /active/i }).length).toBeGreaterThan(0);

    await user.click(screen.getByRole('button', { name: /backup/i }));
    expect(screen.getByRole('button', { name: /export backup/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /choose json file/i })).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /import backup/i })).toBeDisabled();

    await user.click(screen.getByRole('button', { name: /current/i }));
    await user.click(screen.getByRole('button', { name: /finish session/i }));
    await user.click(screen.getByRole('button', { name: /history/i }));

    expect(screen.getByText(/session history/i)).toBeInTheDocument();
    expect(screen.getByText(/1 rounds · 2 pts/i)).toBeInTheDocument();

    const stored = JSON.parse(window.localStorage.getItem('futsal-coach-session-v1'));
    expect(stored.history).toHaveLength(1);
    expect(stored.currentSession.status).toBe('completed');
  });

  it('can toggle player availability from the Players screen', async () => {
    const user = userEvent.setup();
    render(<App />);

    await user.type(screen.getByLabelText(/roster import/i), sevenPlayerRoster);
    await user.click(screen.getByRole('button', { name: /start session/i }));
    await user.click(screen.getByRole('button', { name: /players/i }));

    const ariRow = screen.getByText('Ari').closest('.player-row');
    await user.click(within(ariRow).getByRole('button', { name: /active/i }));

    expect(within(ariRow).getByRole('button', { name: /inactive/i })).toBeInTheDocument();
  });
});

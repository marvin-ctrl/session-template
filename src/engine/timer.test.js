import { describe, expect, it } from 'vitest';
import { adjustTimerDuration, clampTimerDuration, formatClock } from './timer.js';

describe('timer helpers', () => {
  it('formats clock values', () => {
    expect(formatClock(0)).toBe('0:00');
    expect(formatClock(59)).toBe('0:59');
    expect(formatClock(240)).toBe('4:00');
    expect(formatClock(615)).toBe('10:15');
  });

  it('clamps timer duration between 1 and 15 minutes', () => {
    expect(clampTimerDuration(30)).toBe(60);
    expect(clampTimerDuration(240)).toBe(240);
    expect(clampTimerDuration(1200)).toBe(900);
  });

  it('adjusts untouched timers to the new duration', () => {
    expect(adjustTimerDuration(240, 240, 30)).toEqual({
      durationSeconds: 270,
      remainingSeconds: 270,
    });
  });

  it('adjusts running timers while preserving elapsed time', () => {
    expect(adjustTimerDuration(240, 180, 30)).toEqual({
      durationSeconds: 270,
      remainingSeconds: 210,
    });
    expect(adjustTimerDuration(240, 180, -30)).toEqual({
      durationSeconds: 210,
      remainingSeconds: 150,
    });
  });

  it('resets finished timers to the adjusted duration', () => {
    expect(adjustTimerDuration(240, 0, 30, true)).toEqual({
      durationSeconds: 270,
      remainingSeconds: 270,
    });
  });
});

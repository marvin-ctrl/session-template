export const MIN_TIMER_SECONDS = 60;
export const MAX_TIMER_SECONDS = 900;

export function formatClock(totalSeconds) {
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${seconds.toString().padStart(2, '0')}`;
}

export function clampTimerDuration(seconds) {
  return Math.max(MIN_TIMER_SECONDS, Math.min(MAX_TIMER_SECONDS, seconds));
}

export function adjustTimerDuration(currentDuration, currentRemaining, changeSeconds, isFinished = false) {
  const nextDuration = clampTimerDuration(currentDuration + changeSeconds);
  const nextRemaining =
    currentRemaining === currentDuration || isFinished
      ? nextDuration
      : Math.max(0, Math.min(nextDuration, currentRemaining + changeSeconds));

  return { durationSeconds: nextDuration, remainingSeconds: nextRemaining };
}

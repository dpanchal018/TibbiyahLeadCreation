export type QaMilestone = {
  index: number;
  total: number;
  name: string;
  outcome?: 'ok' | 'skipped' | 'failed';
};

export function logQaMilestone(step: QaMilestone): void {
  const { index, total, name } = step;
  const tail =
    step.outcome === 'skipped'
      ? 'skipped'
      : step.outcome === 'failed'
        ? 'FAILED'
        : 'OK';
  console.log(`SF ${index}/${total} ${name} — ${tail}`);
}

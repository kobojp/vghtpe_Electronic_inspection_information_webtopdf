export function monthCount(start: string, end: string): number {
  const pattern = /^(\d{4})-(0[1-9]|1[0-2])$/;
  if (!pattern.test(start) || !pattern.test(end)) return 0;
  const [sy, sm] = start.split('-').map(Number);
  const [ey, em] = end.split('-').map(Number);
  if (sy < 2022 || ey < 2022) return 0;
  return Math.max(0, (ey - sy) * 12 + em - sm + 1);
}
export function fileName(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}

export function mergePdfPaths(existing: string[], incoming: string[]): string[] {
  const seen = new Set<string>();
  return [...existing, ...incoming].filter(path => {
    const key = path.replaceAll('\\', '/').toLocaleLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

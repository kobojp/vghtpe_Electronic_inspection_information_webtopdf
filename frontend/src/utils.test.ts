import { describe, expect, it } from 'vitest';
import { fileName, mergePdfPaths, monthCount } from './utils';

describe('download month range', () => {
  it('includes both ends across years', () => expect(monthCount('2022-01', '2023-12')).toBe(24));
  it('accepts a single month', () => expect(monthCount('2024-02', '2024-02')).toBe(1));
  it.each([['2021-12', '2022-01'], ['2026-13', '2026-14'], ['2026-1', '2026-02'], ['2026-03', '2026-02'], ['', '2026-02']])('rejects invalid range %s to %s', (start, end) => expect(monthCount(start, end)).toBe(0));
});
describe('PDF display names', () => {
  it('handles Windows paths', () => expect(fileName('C:\\報表\\2026-01\\消防.pdf')).toBe('消防.pdf'));
  it('handles slash separators', () => expect(fileName('/reports/消防.pdf')).toBe('消防.pdf'));
});

describe('PDF file selection', () => {
  it('merges selected and dropped files while preserving their original paths', () => {
    expect(mergePdfPaths(['C:/reports/first.pdf'], ['C:/reports/second.pdf', 'C:/reports/first.pdf']))
      .toEqual(['C:/reports/first.pdf', 'C:/reports/second.pdf']);
  });
  it('deduplicates Windows path case and separators', () => {
    expect(mergePdfPaths(['C:\\Reports\\First.PDF'], ['c:/reports/first.pdf'])).toEqual(['C:\\Reports\\First.PDF']);
  });
  it('keeps identical names from different folders', () => {
    expect(mergePdfPaths([], ['C:/January/report.pdf', 'C:/February/report.pdf'])).toHaveLength(2);
  });
});

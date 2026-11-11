/**
 * Conformance harness. Every `conformance/*.json` file is a list of cases
 * `{ f, e, src, grid?, skip? }` run against a fresh engine:
 *   f     the formula
 *   e     the value Excel gives: number, string, boolean, an error code such
 *         as "#N/A", null (blank), or rows of those for an array result
 *   src   where e comes from: "excel" (computed by Excel itself), "docs" (a
 *         worked example from Microsoft documentation) or "expected" (the
 *         author's knowledge of Excel, not yet checked against it)
 *   grid  cells added to / overriding the default grid, by A1 reference
 *   skip  a reason this case is a known gap (reported, not run)
 * Numbers match within 1e-9 relative (a zero expectation: within 1e-12).
 */
/// <reference types="vite/client" />
import { describe, it, expect } from 'vitest';
import { FormulaEngine, RangeValue } from '../formula-engine.js';
import { type GridData, cellKey, refToCoord } from '../../types.js';

/** The grid every case sees unless it says otherwise. */
export const DEFAULT_GRID: Record<string, string> = {
  A1: '1', A2: '2', A3: '3', A4: '4', A5: '5',
  B1: 'a', B2: 'b', B3: 'c',
  C1: '10', C2: '20', C3: '30',
  D1: '=1/0',
  E1: 'apple', E2: 'Banana', E3: 'cherry',
};

type Case = { f: string; e: unknown; src: 'excel' | 'docs' | 'expected'; grid?: Record<string, string>; skip?: string };

function makeData(entries: Record<string, string>): GridData {
  const data: GridData = new Map();
  for (const [ref, raw] of Object.entries(entries)) {
    const { row, col } = refToCoord(ref);
    const isNum = raw.trim() !== '' && !isNaN(Number(raw));
    data.set(cellKey(row, col), { rawValue: raw, displayValue: raw, type: isNum ? 'number' : 'text' });
  }
  return data;
}

function plain(value: unknown): unknown {
  if (value instanceof RangeValue) {
    return Array.from({ length: value.rows }, (_, r) => value.getRow(r).map(plain));
  }
  return value === undefined ? null : value;
}

function same(actual: unknown, expected: unknown): boolean {
  if (Array.isArray(expected)) {
    return (
      Array.isArray(actual) &&
      actual.length === expected.length &&
      expected.every((row, r) => (row as unknown[]).length === (actual[r] as unknown[]).length &&
        (row as unknown[]).every((x, c) => same((actual[r] as unknown[])[c], x)))
    );
  }
  if (typeof expected === 'number' && typeof actual === 'number') {
    // Relative for anything nonzero (a missing tail like 1e-45 must not pass as 0);
    // a zero expectation allows only floating-point noise
    if (expected === 0) return Math.abs(actual) <= 1e-12;
    return Math.abs(actual - expected) <= 1e-9 * Math.abs(expected);
  }
  if (typeof expected === 'string' && typeof actual === 'string') return actual === expected;
  return actual === expected;
}

const fixtures = import.meta.glob<Case[]>('./conformance/*.json', { eager: true, import: 'default' });
const files = Object.keys(fixtures).sort();
const seen = new Set<string>();

for (const file of files) {
  const cases = fixtures[file];
  describe(`conformance: ${file.replace('./conformance/', '')}`, () => {
    it('has well-formed, unique cases', () => {
      for (const c of cases) {
        expect(typeof c.f, JSON.stringify(c)).toBe('string');
        expect(c.f.startsWith('='), c.f).toBe(true);
        expect('e' in c, c.f).toBe(true);
        expect(['excel', 'docs', 'expected'], c.f).toContain(c.src);
        const id = c.f + '|' + JSON.stringify(c.grid ?? {});
        expect(seen.has(id), `duplicate case ${c.f}`).toBe(false);
        seen.add(id);
      }
    });
    for (const c of cases) {
      const run = c.skip ? it.skip : it;
      run(`${c.f}${c.skip ? ' — ' + c.skip : ''}`, () => {
        const engine = new FormulaEngine();
        engine.setData(makeData({ ...DEFAULT_GRID, ...c.grid }));
        const actual = plain(engine.evaluateValue(c.f));
        expect(same(actual, c.e), `${c.f}: expected ${JSON.stringify(c.e)}, got ${JSON.stringify(actual)}`).toBe(true);
      });
    }
  });
}

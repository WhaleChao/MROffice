/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PDF page's text as a column to read on a phone (Reflow): the page's text
 * runs, as the reader's text layer reads them (pdf.js text content), put
 * back into lines, columns, paragraphs and headings, set at the reader's
 * size rather than the page's. Nothing of the file changes.
 *
 * - A line is the runs on one baseline; runs further apart than a few spaces
 *   on that baseline are two columns' lines, read column by column when the
 *   page keeps the split at one place.
 * - A paragraph ends at a gap wider than the lines' own spacing, a change of
 *   size, or an indented first line after a line that ends a sentence.
 * - A paragraph set larger than the page's body text is a heading.
 * - Lines are joined with a space, without one between CJK characters, and a
 *   word hyphenated at a line's end is joined whole.
 */

export interface ReflowTextRun {
  readonly str: string;
  /** pdf.js text item transform: [a, b, c, d, x, y] in page units, y up. */
  readonly transform: readonly number[];
  readonly width: number;
  readonly height?: number;
}

export interface ReflowBlock {
  readonly kind: "heading" | "paragraph";
  /** 1-3 for a heading, by how much larger than the body it is set. */
  readonly level?: 1 | 2 | 3;
  readonly text: string;
}

interface Run { text: string; x: number; y: number; end: number; size: number }
interface Line { x: number; end: number; y: number; size: number; text: string }

const CJK = /[⺀-鿿가-힯豈-﫿＀-￯　-〿]/u;
const SENTENCE_END = /[.!?:;。！？：；」』)\]）]$/u;

function runOf(item: ReflowTextRun): Run | null {
  const text = typeof item.str === "string" ? item.str : "";
  const t = item.transform;
  if (!text || !Array.isArray(t) || t.length < 6) return null;
  const size = Math.hypot(Number(t[2]) || 0, Number(t[3]) || 0) || Number(item.height) || 0;
  const x = Number(t[4]);
  const y = Number(t[5]);
  if (!(size > 0) || !Number.isFinite(x) || !Number.isFinite(y)) return null;
  return { text, x, y, end: x + Math.max(0, Number(item.width) || 0), size };
}

function joinRuns(runs: readonly Run[]): string {
  let text = "";
  let previous: Run | null = null;
  for (const run of runs) {
    if (previous) {
      const gap = run.x - previous.end;
      const spaced = /\s$/u.test(text) || /^\s/u.test(run.text);
      if (!spaced && gap > Math.min(previous.size, run.size) * 0.2 &&
        !(CJK.test(text.slice(-1)) && CJK.test(run.text[0] ?? ""))) text += " ";
    }
    text += run.text;
    previous = run;
  }
  return text.replace(/\s+/gu, " ").trim();
}

function lineOf(runs: readonly Run[]): Line | null {
  const text = joinRuns(runs);
  if (!text) return null;
  const size = Math.max(...runs.map((run) => run.size));
  return { x: runs[0].x, end: Math.max(...runs.map((run) => run.end)), y: runs[0].y, size, text };
}

/** Runs on one baseline, split where a gap is wider than a column gutter. */
function baselineSegments(runs: Run[]): Line[] {
  runs.sort((a, b) => a.x - b.x);
  const out: Line[] = [];
  let current: Run[] = [];
  for (const run of runs) {
    const last = current[current.length - 1];
    if (last && run.x - last.end > Math.max(last.size, run.size) * 2.5) {
      const line = lineOf(current);
      if (line) out.push(line);
      current = [];
    }
    current.push(run);
  }
  const line = lineOf(current);
  if (line) out.push(line);
  return out;
}

function median(values: readonly number[]): number {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

/** Lines in reading order: one column top to bottom, or the left column before the right. */
function readingLines(runs: readonly Run[]): Line[] {
  const byBaseline: Run[][] = [];
  for (const run of [...runs].sort((a, b) => b.y - a.y || a.x - b.x)) {
    const group = byBaseline[byBaseline.length - 1];
    if (group && Math.abs(group[0].y - run.y) <= Math.min(group[0].size, run.size) * 0.5) group.push(run);
    else byBaseline.push([run]);
  }
  const rows = byBaseline.map(baselineSegments);
  const split = rows.filter((row) => row.length === 2);
  // A two-column page: most lines that split, split at about the same place.
  if (split.length >= 3 && split.length >= rows.filter((row) => row.length > 0).length * 0.3) {
    const gutter = median(split.map((row) => (row[0].end + row[1].x) / 2));
    const tolerance = median(split.map((row) => row[0].size)) * 3;
    if (split.filter((row) => Math.abs((row[0].end + row[1].x) / 2 - gutter) <= tolerance).length >= split.length * 0.7) {
      const left: Line[] = [];
      const right: Line[] = [];
      const full: { at: number; lines: Line[] }[] = [];
      for (const row of rows) {
        for (const line of row) {
          if (line.end <= gutter + tolerance / 2) left.push(line);
          else if (line.x >= gutter - tolerance / 2) right.push(line);
          else full.push({ at: line.y, lines: [line] });
        }
      }
      // Lines across both columns (a title above them, a footer below)
      // stay where they stand: before the columns when above both, after
      // them when below.
      const top = Math.max(...[...left, ...right].map((line) => line.y));
      return [
        ...full.filter((entry) => entry.at > top).flatMap((entry) => entry.lines),
        ...left, ...right,
        ...full.filter((entry) => entry.at <= top).flatMap((entry) => entry.lines),
      ];
    }
  }
  return rows.flat();
}

function joinLines(previous: string, next: string): string {
  if (/[A-Za-zÀ-ɏ]-$/u.test(previous) && /^[a-zß-ÿ]/u.test(next)) return previous.slice(0, -1) + next;
  if (CJK.test(previous.slice(-1)) && CJK.test(next[0] ?? "")) return previous + next;
  return `${previous} ${next}`;
}

/** The page's text as headings and paragraphs, in reading order. */
export function reflowPageBlocks(items: readonly ReflowTextRun[]): ReflowBlock[] {
  const runs = items.map(runOf).filter((run): run is Run => run !== null && run.text.trim() !== "");
  if (!runs.length) return [];
  const lines = readingLines(runs);
  // The body size: the size most of the page's characters are set in.
  const bySize = new Map<number, number>();
  for (const line of lines) {
    const size = Math.round(line.size * 2) / 2;
    bySize.set(size, (bySize.get(size) ?? 0) + line.text.length);
  }
  const body = [...bySize].sort((a, b) => b[1] - a[1])[0][0];
  const paragraphs: { size: number; text: string; x: number; lastY: number; lastSize: number }[] = [];
  for (const line of lines) {
    const current = paragraphs[paragraphs.length - 1];
    const gap = current ? current.lastY - line.y : 0;
    const spacing = Math.max(current?.lastSize ?? 0, line.size);
    const sameSize = current && Math.abs(current.size - line.size) <= Math.max(current.size, line.size) * 0.15;
    const continues = current && sameSize && gap > 0 && gap <= spacing * 1.75 &&
      !(line.x - current.x > line.size * 1.5 && SENTENCE_END.test(current.text));
    if (continues && current) {
      current.text = joinLines(current.text, line.text);
      current.lastY = line.y;
      current.lastSize = line.size;
    } else {
      paragraphs.push({ size: line.size, text: line.text, x: line.x, lastY: line.y, lastSize: line.size });
    }
  }
  return paragraphs.map(({ size, text }) => {
    const ratio = size / body;
    if (ratio >= 1.25 && text.length <= 200) {
      return { kind: "heading", level: ratio >= 1.8 ? 1 : ratio >= 1.45 ? 2 : 3, text } as const;
    }
    return { kind: "paragraph", text } as const;
  });
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Canonical form used to decide whether two OMML models are the same
 * equation for MRAI, whose stored source is LaTeX. Differences that neither
 * LaTeX nor Word's linear format can express are folded away:
 *
 * - run splitting, run properties without visual effect (m:lit, roman),
 *   italic as the default for letters, and spacing characters;
 * - the extent of an n-ary operand and of a function argument (LaTeX writes
 *   `\sum_i a_i` and `\sin x` without one);
 * - a prime written as a superscript or as a following character;
 * - brackets written as run text versus a delimiter object (Word builds the
 *   former into the latter);
 * - Word's cases (a brace and an equation array) versus a left-aligned
 *   matrix in a brace (the exporter's `cases`).
 */

import type { MathList, MathNode, MathRunStyle } from "./ast";
import {
  CLOSERS,
  OPENERS,
  SELF_DELIMITERS,
  defaultLimitLocation,
  isItalicByDefault,
  letterlikeLetter,
} from "./characters";
import { combiningAccent } from "./serialize";

type Atom = { a: string; s: string };
type Canonical = Atom | { k: string; [key: string]: unknown };

const CHARACTER_FOLDS: Readonly<Record<string, string>> = {
  "−": "-",
  "'": "′",
  "″": "′′",
  "‴": "′′′",
  "⁗": "′′′′",
  "〈": "⟨",
  "〉": "⟩",
  "∗": "*",
};

function styleKey(character: string, style: MathRunStyle): string {
  if (style.nor) return "n";
  const scr = style.scr === "roman" ? undefined : style.scr;
  const letter = isItalicByDefault(character);
  const digit = /^[0-9]$/.test(character);
  if (!letter && !digit) return "";
  const bold = style.sty === "b" || style.sty === "bi";
  if (digit) {
    if (scr === "double-struck" || scr === "monospace") return scr;
    if (scr === "sans-serif") return bold ? "sans:b" : "sans";
    return bold ? "b" : "";
  }
  if (scr === "script" || scr === "fraktur") return `${scr}${bold ? ":b" : ""}`;
  if (scr === "double-struck" || scr === "monospace") return scr;
  const sty = style.sty ?? "i";
  return scr === "sans-serif" ? `sans:${sty}` : sty;
}

function isPrimeText(list: MathList): boolean {
  return (
    list.length > 0 &&
    list.every((node) => node.kind === "text" && !node.style.nor && /^['′″‴⁗]+$/u.test(node.text))
  );
}

function atoms(list: MathList, out: Canonical[]): void {
  for (const node of list) {
    for (const item of canonicalNode(node)) out.push(item);
  }
}

function canonicalList(list: MathList): Canonical[] {
  const flat: Canonical[] = [];
  atoms(list, flat);
  return mergeAtoms(wrapBrackets(flat));
}

function canonicalNode(node: MathNode): Canonical[] {
  switch (node.kind) {
    case "text": {
      const out: Canonical[] = [];
      for (const raw of Array.from(node.text)) {
        if (/^[\p{Zs}​\t]$/u.test(raw)) continue;
        const letterlike = node.style.nor || node.style.scr ? null : letterlikeLetter(raw);
        if (letterlike) {
          const style: MathRunStyle = { sty: letterlike.sty ?? node.style.sty, scr: letterlike.scr };
          out.push({ a: letterlike.character, s: styleKey(letterlike.character, style) });
          continue;
        }
        const folded = node.style.nor ? raw : (CHARACTER_FOLDS[raw] ?? raw);
        for (const character of Array.from(folded)) {
          out.push({ a: character, s: styleKey(character, node.style) });
        }
      }
      return out;
    }
    case "f":
      return [{ k: "f", type: node.type, num: canonicalList(node.num), den: canonicalList(node.den) }];
    case "sSup":
      if (isPrimeText(node.sup)) return [...flatten(node.e), ...flatten(node.sup)];
      return [{ k: "sSup", e: canonicalList(node.e), sup: canonicalList(node.sup) }];
    case "sSub":
      return [{ k: "sSub", e: canonicalList(node.e), sub: canonicalList(node.sub) }];
    case "sSubSup":
      return [{ k: "sSubSup", e: canonicalList(node.e), sub: canonicalList(node.sub), sup: canonicalList(node.sup) }];
    case "sPre":
      return [{ k: "sPre", e: canonicalList(node.e), sub: canonicalList(node.sub), sup: canonicalList(node.sup) }];
    case "rad":
      return [{ k: "rad", deg: canonicalList(node.deg), e: canonicalList(node.e) }];
    case "nary":
      return [
        {
          k: "nary",
          chr: node.chr,
          limLoc: node.limLoc ?? defaultLimitLocation(node.chr),
          sub: canonicalList(node.sub ?? []),
          sup: canonicalList(node.sup ?? []),
        },
        ...flatten(node.e),
      ];
    case "d": {
      const only = node.args.length === 1 && node.args[0].length === 1 ? node.args[0][0] : null;
      if (node.beg === "{" && node.end === "" && only?.kind === "eqArr") {
        const rows = only.rows.map(splitAlignment);
        if (rows.every((cells) => cells.length <= 2)) {
          const width = Math.max(...rows.map((cells) => cells.length));
          return [
            {
              k: "cases",
              rows: rows.map((cells) =>
                [...cells, ...Array.from({ length: width - cells.length }, () => [] as MathList)].map(canonicalList),
              ),
            },
          ];
        }
      }
      const args = node.args.map(canonicalList);
      const cases = casesMatrix(node.beg, node.end, args);
      if (cases) return [cases];
      return [{ k: "d", beg: CHARACTER_FOLDS[node.beg] ?? node.beg, end: CHARACTER_FOLDS[node.end] ?? node.end, sep: node.sep, args }];
    }
    case "m": {
      const columns = Math.max(0, ...node.rows.map((row) => row.length));
      return [
        {
          k: "m",
          rows: node.rows.map((row) => row.map(canonicalList)),
          jc: Array.from({ length: columns }, (_, index) => node.columnJc[index] ?? "center"),
        },
      ];
    }
    case "eqArr":
      return [{ k: "eqArr", rows: node.rows.map(canonicalList) }];
    case "func":
      return [...flatten(node.name), ...flatten(node.e)];
    case "limLow":
    case "limUpp":
      return [{ k: node.kind, e: canonicalList(node.e), lim: canonicalList(node.lim) }];
    case "acc":
      return [{ k: "acc", chr: combiningAccent(node.chr) ?? node.chr, e: canonicalList(node.e) }];
    case "bar":
      return [{ k: "bar", pos: node.pos, e: canonicalList(node.e) }];
    case "groupChr":
      return [{ k: "groupChr", chr: node.chr, pos: node.pos, vertJc: node.vertJc, e: canonicalList(node.e) }];
    case "borderBox":
      return [{ k: "borderBox", flags: [...node.flags].sort(), e: canonicalList(node.e) }];
    case "box":
      return [{ k: "box", e: canonicalList(node.e) }];
    case "phant":
      return [
        {
          k: "phant",
          show: node.show,
          zeroWid: node.zeroWid,
          zeroAsc: node.zeroAsc,
          zeroDesc: node.zeroDesc,
          transp: node.transp,
          e: canonicalList(node.e),
        },
      ];
  }
}

/** Splice a list into its parent (atoms stay unmerged until the parent merges). */
function flatten(list: MathList): Canonical[] {
  const out: Canonical[] = [];
  atoms(list, out);
  return out;
}

const isAtom = (item: Canonical): item is Atom => "a" in item;

/** The exporter's cases: a left-aligned matrix of one or two columns in a brace. */
function casesMatrix(beg: string, end: string, args: Canonical[][]): Canonical | null {
  if (beg !== "{" || end !== "" || args.length !== 1 || args[0].length !== 1) return null;
  const only = args[0][0];
  if (isAtom(only) || only.k !== "m") return null;
  const jc = only.jc as string[];
  if (jc.length === 0 || jc.length > 2 || jc.some((value) => value !== "left")) return null;
  return { k: "cases", rows: only.rows };
}

/** An equation-array row split at its `&` alignment marks. */
function splitAlignment(row: MathList): MathList[] {
  const cells: MathList[] = [[]];
  for (const node of row) {
    if (node.kind !== "text" || node.style.nor || node.style.lit || !node.text.includes("&")) {
      cells[cells.length - 1].push(node);
      continue;
    }
    node.text.split("&").forEach((part, index) => {
      if (index > 0) cells.push([]);
      if (part) cells[cells.length - 1].push({ ...node, text: part });
    });
  }
  return cells;
}

/** Matched bracket characters written as text become delimiter objects. */
function wrapBrackets(items: Canonical[]): Canonical[] {
  const stack: { character: string; index: number; self: boolean; content: boolean }[] = [];
  const pairs = new Map<number, number>();
  items.forEach((item, index) => {
    const touch = () => {
      const top = stack.at(-1);
      if (top) top.content = true;
    };
    if (!isAtom(item) || item.s === "n") {
      touch();
      return;
    }
    const character = item.a;
    if (OPENERS.has(character)) {
      touch();
      stack.push({ character, index, self: false, content: false });
    } else if (CLOSERS.has(character)) {
      while (stack.length && stack[stack.length - 1].self) stack.pop();
      const open = stack.pop();
      if (open) {
        pairs.set(open.index, index);
        touch();
      }
    } else if (SELF_DELIMITERS.has(character)) {
      const top = stack.at(-1);
      if (top && top.self && top.character === character && top.content) {
        stack.pop();
        pairs.set(top.index, index);
        touch();
      } else {
        touch();
        stack.push({ character, index, self: true, content: false });
      }
    } else touch();
  });
  if (pairs.size === 0) return items;
  const build = (from: number, to: number): Canonical[] => {
    const out: Canonical[] = [];
    for (let index = from; index < to; index += 1) {
      const close = pairs.get(index);
      if (close === undefined) {
        out.push(items[index]);
        continue;
      }
      const open = items[index] as Atom;
      const end = items[close] as Atom;
      out.push({
        k: "d",
        beg: CHARACTER_FOLDS[open.a] ?? open.a,
        end: CHARACTER_FOLDS[end.a] ?? end.a,
        sep: "|",
        args: [mergeAtoms(build(index + 1, close))],
      });
      index = close;
    }
    return out;
  };
  return build(0, items.length);
}

function mergeAtoms(items: Canonical[]): Canonical[] {
  const out: Canonical[] = [];
  for (const item of items) {
    const last = out.at(-1);
    if (isAtom(item) && last && !isAtom(last) && last.k === "t" && last.s === item.s) {
      last.t = `${String(last.t)}${item.a}`;
      continue;
    }
    out.push(isAtom(item) ? ({ k: "t", t: item.a, s: item.s } as Canonical) : item);
  }
  return out;
}

/** A string identity for the equation; equal strings mean the same equation. */
export function canonicalMath(list: MathList): string {
  return JSON.stringify(canonicalList(list));
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type MergeNode = { type?: string; attrs?: Record<string, unknown> | null; [key: string]: unknown };

/** Past this many changed tokens on one side the merge is not tried. */
const MAX_EDIT_DISTANCE = 2_000;

/** A value as its author wrote it: null, false and `source` left out, a
 * mark list's order ignored (canonical-block-merge's `authored`). */
function authored(value: unknown, key?: string): string {
  if (Array.isArray(value)) {
    const items = value.map((item) => authored(item));
    return `[${(key === "marks" ? items.sort() : items).join(",")}]`;
  }
  if (value && typeof value === "object") {
    const entries = Object.entries(value as Record<string, unknown>)
      .filter(([name, item]) => item !== undefined && item !== null && item !== false && name !== "source")
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
    return `{${entries.map(([name, item]) => `${JSON.stringify(name)}:${authored(item, name)}`).join(",")}}`;
  }
  return JSON.stringify(value ?? null);
}

type Token =
  | { kind: "c"; key: string; ch: string; marks: unknown }
  | { kind: "o"; key: string; node: MergeNode }
  | { kind: "x"; key: string }
  | { kind: "a"; key: string; node: MergeNode };

const CLOSE: Token = { kind: "x", key: "x" };

function tokenize(content: unknown, out: Token[] = []): Token[] {
  if (!Array.isArray(content)) return out;
  for (const item of content) {
    if (!item || typeof item !== "object") continue;
    const node = item as MergeNode;
    if (node.type === "text" && typeof node.text === "string") {
      const marksKey = authored(node.marks ?? null, "marks");
      for (const ch of Array.from(node.text)) out.push({ kind: "c", key: `c${marksKey}${ch}`, ch, marks: node.marks });
    } else if (Array.isArray(node.content)) {
      const { content: inner, ...shell } = node;
      out.push({ kind: "o", key: `o${authored(shell)}`, node: shell });
      tokenize(inner, out);
      out.push(CLOSE);
    } else {
      out.push({ kind: "a", key: `a${authored(node)}`, node });
    }
  }
  return out;
}

/**
 * For each token of `a`, the index of the token of `b` it is matched with in
 * a longest common subsequence (or -1). Null past MAX_EDIT_DISTANCE.
 */
function matchTokens(a: readonly { key: string }[], b: readonly { key: string }[]): Int32Array | null {
  const n = a.length, m = b.length;
  const match = new Int32Array(n).fill(-1);
  let start = 0;
  while (start < n && start < m && a[start].key === b[start].key) { match[start] = start; start += 1; }
  let endA = n, endB = m;
  while (endA > start && endB > start && a[endA - 1].key === b[endB - 1].key) {
    endA -= 1; endB -= 1;
    match[endA] = endB;
  }
  const N = endA - start, M = endB - start;
  if (N === 0 || M === 0) return match;
  // Myers' O(ND) greedy search; each step's frontier (k in -d..d) is kept to walk back.
  const max = N + M;
  const limit = Math.min(max, MAX_EDIT_DISTANCE);
  if (limit < Math.abs(N - M)) return null;
  const offset = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  let found = -1;
  for (let d = 0; d <= limit; d += 1) {
    // trace[d]: the frontier before step d, at k + d + 1.
    trace.push(v.slice(offset - d - 1, offset + d + 2));
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1]) ? v[offset + k + 1] : v[offset + k - 1] + 1;
      let y = x - k;
      while (x < N && y < M && a[start + x].key === b[start + y].key) { x += 1; y += 1; }
      v[offset + k] = x;
      if (x >= N && y >= M) { found = d; break; }
    }
    if (found >= 0) break;
  }
  if (found < 0) return null;
  // Walk back through the frontiers, marking the diagonal (matched) runs.
  let x = N, y = M;
  for (let d = found; d > 0; d -= 1) {
    const frontier = trace[d];
    const at = (k: number) => frontier[k + d + 1];
    const k = x - y;
    const prevK = k === -d || (k !== d && at(k - 1) < at(k + 1)) ? k + 1 : k - 1;
    const prevX = at(prevK);
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) { x -= 1; y -= 1; match[start + x] = start + y; }
    x = prevX; y = prevY;
  }
  while (x > 0 && y > 0) { x -= 1; y -= 1; match[start + x] = start + y; }
  return match;
}

/** What `side` put in before each token of the base (gap i: before base[i];
 * gap n: at the end), given the match of base to side. */
function insertionsByGap(baseCount: number, side: readonly Token[], match: Int32Array): Token[][] {
  const gaps: Token[][] = Array.from({ length: baseCount + 1 }, () => []);
  const baseAt = new Int32Array(side.length).fill(-1);
  for (let i = 0; i < baseCount; i += 1) if (match[i] >= 0) baseAt[match[i]] = i;
  let gap = 0;
  for (let j = 0; j < side.length; j += 1) {
    if (baseAt[j] >= 0) { gap = baseAt[j] + 1; continue; }
    gaps[gap].push(side[j]);
  }
  return gaps;
}

function keysOf(tokens: readonly Token[]): string {
  return tokens.map((token) => token.key).join("\u0000");
}

/** Both sides put something at one place: once when one holds the other at
 * its start or end (the same words typed, or one side's typing carried on),
 * else the room's and then this view's. */
function bothInserted(remote: Token[], local: Token[]): Token[] {
  if (remote.length === 0) return local;
  if (local.length === 0) return remote;
  const r = keysOf(remote), l = keysOf(local);
  if (r === l) return remote;
  const [long, short, longTokens] = r.length >= l.length ? [r, l, remote] : [l, r, local];
  if (long.startsWith(short) || long.endsWith(short)) return longTokens;
  return [...remote, ...local];
}

/** Whether one side took out the start or end of a node (a cell, a row, a
 * paragraph joined to the next) that the other side changed anything inside:
 * the other's words would land outside any node that can hold them. */
function structureCutUnderOtherEdit(
  base: readonly Token[],
  toLocal: Int32Array, localGaps: readonly Token[][],
  toRemote: Int32Array, remoteGaps: readonly Token[][],
): boolean {
  const n = base.length;
  const changes = (match: Int32Array, gaps: readonly Token[][]) => {
    // sum[i]: changes at base tokens < i and gaps <= i... kept as two prefix sums.
    const deleted = new Int32Array(n + 1), inserted = new Int32Array(n + 2);
    for (let i = 0; i < n; i += 1) deleted[i + 1] = deleted[i] + (match[i] < 0 ? 1 : 0);
    for (let g = 0; g <= n; g += 1) inserted[g + 1] = inserted[g] + (gaps[g].length > 0 ? 1 : 0);
    /** Any change inside the node base[open..close]: its tokens, the gaps within. */
    return (open: number, close: number) =>
      deleted[close + 1] - deleted[open] > 0 || inserted[close + 1] - inserted[open + 1] > 0;
  };
  const localInside = changes(toLocal, localGaps);
  const remoteInside = changes(toRemote, remoteGaps);
  const stack: number[] = [];
  for (let i = 0; i < n; i += 1) {
    const kind = base[i].kind;
    if (kind === "o") { stack.push(i); continue; }
    if (kind !== "x") continue;
    const open = stack.pop();
    if (open === undefined) return true;
    const localCut = toLocal[open] < 0 || toLocal[i] < 0;
    const remoteCut = toRemote[open] < 0 || toRemote[i] < 0;
    if ((localCut && remoteInside(open, i)) || (remoteCut && localInside(open, i))) return true;
  }
  return false;
}

function rebuild(tokens: readonly Token[]): unknown[] | null {
  const root: unknown[] = [];
  const stack: Array<{ node: MergeNode; content: unknown[] }> = [];
  let marksKeyOfLast: string | null = null;
  const contentNow = () => (stack.length > 0 ? stack[stack.length - 1].content : root);
  for (const token of tokens) {
    if (token.kind === "c") {
      const content = contentNow();
      const marksKey = authored(token.marks ?? null, "marks");
      const last = content[content.length - 1] as { type?: string; text?: string } | undefined;
      if (last && last.type === "text" && marksKeyOfLast === marksKey) {
        last.text += token.ch;
      } else {
        content.push(token.marks !== undefined ? { type: "text", text: token.ch, marks: structuredClone(token.marks) } : { type: "text", text: token.ch });
      }
      marksKeyOfLast = marksKey;
      continue;
    }
    marksKeyOfLast = null;
    if (token.kind === "a") {
      contentNow().push(structuredClone(token.node));
    } else if (token.kind === "o") {
      stack.push({ node: structuredClone(token.node), content: [] });
    } else {
      const open = stack.pop();
      if (!open) return null;
      contentNow().push({ ...open.node, content: open.content });
    }
  }
  return stack.length === 0 ? root : null;
}

function sameValue(left: unknown, right: unknown): boolean {
  return authored(left) === authored(right);
}

/** The block's own fields (its type, attrs, marks) merged key by key: each
 * side's change, the room's where both changed one. */
function mergeShell(base: MergeNode, local: MergeNode, remote: MergeNode): MergeNode {
  const out: MergeNode = {};
  const keys = new Set([...Object.keys(base), ...Object.keys(local), ...Object.keys(remote)]);
  keys.delete("content");
  for (const key of keys) {
    if (key === "attrs") continue;
    const value = sameValue(local[key], base[key]) ? remote[key] : sameValue(remote[key], base[key]) ? local[key] : remote[key];
    if (value !== undefined) out[key] = value;
  }
  const attrsOf = (node: MergeNode) => (node.attrs && typeof node.attrs === "object" ? node.attrs : {});
  const [b, l, r] = [attrsOf(base), attrsOf(local), attrsOf(remote)];
  if (base.attrs !== undefined || local.attrs !== undefined || remote.attrs !== undefined) {
    const attrs: Record<string, unknown> = {};
    for (const key of new Set([...Object.keys(b), ...Object.keys(l), ...Object.keys(r)])) {
      const value = sameValue(l[key], b[key]) ? r[key] : sameValue(r[key], b[key]) ? l[key] : r[key];
      if (value !== undefined) attrs[key] = value;
    }
    out.attrs = attrs;
  }
  return out;
}

/**
 * `local` (this view's) and `remote` (the room's or the store's) versions of
 * a block, both changed from `base`, as one block holding both changes; null
 * when they cannot be read as one (the caller keeps both).
 */
/** For each item of `a`, the index of the item of `b` it is matched with in a
 * longest common run of the two (by `key`), or -1; null past the edit bound. */
export function matchKeyed(a: readonly { key: string }[], b: readonly { key: string }[]): Int32Array | null {
  return matchTokens(a, b);
}

export function mergeBlockText(base: MergeNode, local: MergeNode, remote: MergeNode): MergeNode | null {
  const [baseTokens, localTokens, remoteTokens] = [tokenize(base.content), tokenize(local.content), tokenize(remote.content)];
  const toLocal = matchTokens(baseTokens, localTokens);
  const toRemote = matchTokens(baseTokens, remoteTokens);
  if (!toLocal || !toRemote) return null;
  const localGaps = insertionsByGap(baseTokens.length, localTokens, toLocal);
  const remoteGaps = insertionsByGap(baseTokens.length, remoteTokens, toRemote);
  if (structureCutUnderOtherEdit(baseTokens, toLocal, localGaps, toRemote, remoteGaps)) return null;
  const merged: Token[] = [];
  for (let i = 0; i <= baseTokens.length; i += 1) {
    merged.push(...bothInserted(remoteGaps[i], localGaps[i]));
    if (i < baseTokens.length && toLocal[i] >= 0 && toRemote[i] >= 0) merged.push(remoteTokens[toRemote[i]]);
  }
  const content = rebuild(merged);
  if (!content) return null;
  const shell = mergeShell(base, local, remote);
  const hadContent = Array.isArray(base.content) || Array.isArray(local.content) || Array.isArray(remote.content);
  return hadContent ? { ...shell, content } : shell;
}

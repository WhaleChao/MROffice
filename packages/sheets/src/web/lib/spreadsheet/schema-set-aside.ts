/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A stored value the editor could not read, set aside and saved back. */
export interface SpreadsheetSetAsideValue {
  /** The sheet the value belongs to, or null for a workbook-level value. */
  readonly sheetId: string | null;
  /** Its path inside that sheet (or the workbook), never through an array. */
  readonly path: readonly string[];
  readonly value: unknown;
  /**
   * A whole element of one of the sheet's lists set aside (a drawing object
   * the editor cannot read): the list at ``path``, the element's own id and
   * the place it held. It goes back by its id; an element without one (a
   * conditional-format rule, a validation) goes back at the place it held,
   * unless the list holds an equal element again.
   */
  readonly element?: { readonly id: string | null; readonly index: number };
  





  readonly cell?: { readonly value: unknown; readonly formula: unknown };
}

/** A value the schema refused, where it was in the content parsed. */
export interface SchemaRemovedValue {
  path: ReadonlyArray<string | number>;
  value: unknown;
  element?: { id: string | null; index: number };
}

/** The issues of a schema refusal (zod's, read structurally). */
export interface SchemaIssueLike {
  code: string;
  path: ReadonlyArray<PropertyKey>;
  keys?: readonly string[];
}

type Json = Record<string, unknown>;

function sameJson(a: unknown, b: unknown): boolean {
  try {
    return JSON.stringify(a) === JSON.stringify(b);
  } catch {
    return false;
  }
}

/**
 * ``content`` with every set-aside value put back where the content holds
 * nothing: a value the user has since replaced (or whose cell or sheet is
 * gone) stays as the user left it.
 */
export function restoreSetAsideValues<T extends Record<string, unknown>>(
  content: T,
  setAside: readonly SpreadsheetSetAsideValue[] | null | undefined,
): T {
  if (!setAside?.length) return content;
  let result: Json = content;
  let sheets: unknown[] | null = null;
  for (const entry of setAside) {
    if (entry.element) {
      // A whole list element: back into its sheet's list at its place, unless
      // the list holds that element again (by its id, or equal when it has none).
      const list = (sheets ?? (Array.isArray(result.sheets) ? result.sheets : null)) as unknown[] | null;
      const index = list?.findIndex(
        (sheet) => sheet && typeof sheet === "object" && (sheet as { id?: unknown }).id === entry.sheetId,
      ) ?? -1;
      const key = entry.path[0];
      if (!list || index < 0 || entry.path.length !== 1 || !key) continue;
      const sheet = list[index] as Json;
      const elements = Array.isArray(sheet[key]) ? (sheet[key] as unknown[]) : [];
      const id = entry.element.id;
      const present = id !== null
        ? elements.some((element) => element && typeof element === "object" && (element as { id?: unknown }).id === id)
        : elements.some((element) => sameJson(element, entry.value));
      if (present) continue;
      const next = [...elements];
      next.splice(Math.min(entry.element.index, next.length), 0, entry.value);
      sheets = sheets ?? [...list];
      sheets[index] = { ...sheet, [key]: next };
      continue;
    }
    let root: Json | null = null;
    let replace: ((next: Json) => void) | null = null;
    if (entry.sheetId === null) {
      root = result;
      replace = (next) => {
        result = next;
      };
    } else {
      const list = (sheets ?? (Array.isArray(result.sheets) ? result.sheets : null)) as unknown[] | null;
      const index = list?.findIndex(
        (sheet) => sheet && typeof sheet === "object" && (sheet as { id?: unknown }).id === entry.sheetId,
      ) ?? -1;
      if (!list || index < 0) continue;
      root = list[index] as Json;
      replace = (next) => {
        sheets = sheets ?? [...list];
        sheets[index] = next;
      };
    }
    if (!root || !entry.path.length) continue;
    // Copy along the path, so the caller's objects are never written to.
    const chain: Json[] = [root];
    let missingParent = false;
    for (const step of entry.path.slice(0, -1)) {
      const next = chain[chain.length - 1][step];
      if (!next || typeof next !== "object" || Array.isArray(next)) {
        missingParent = true;
        break;
      }
      chain.push(next as Json);
    }
    const last = entry.path[entry.path.length - 1];
    if (missingParent || last in chain[chain.length - 1]) continue;
    if (entry.cell && (chain[2]?.value !== entry.cell.value || chain[2]?.formula !== entry.cell.formula)) continue;
    let rebuilt: Json = { ...chain[chain.length - 1], [last]: entry.value };
    for (let depth = chain.length - 2; depth >= 0; depth -= 1) {
      rebuilt = { ...chain[depth], [entry.path[depth]]: rebuilt };
    }
    replace(rebuilt);
  }
  if (sheets) result = { ...result, sheets };
  return result as T;
}

/**
 * The removed values anchored where a save can put them back: a sheet's by
 * the sheet's id, a workbook value by its key, and never through an array
 * (an index moves with every insert). Null when one cannot be anchored.
 */
export function anchorSetAsideValues(
  content: Json,
  removed: ReadonlyArray<SchemaRemovedValue>,
): SpreadsheetSetAsideValue[] | null {
  const sheets = Array.isArray(content.sheets) ? content.sheets : [];
  const anchored: SpreadsheetSetAsideValue[] = [];
  for (const { path, value, element } of removed) {
    if (element) {
      const sheet = typeof path[1] === "number" ? sheets[path[1]] : undefined;
      const id = sheet && typeof sheet === "object" ? (sheet as { id?: unknown }).id : undefined;
      if (path[0] !== "sheets" || typeof id !== "string" || !id || path.length !== 3 || typeof path[2] !== "string") return null;
      anchored.push({ sheetId: id, path: [path[2]], value, element });
      continue;
    }
    let sheetId: string | null = null;
    let rest: ReadonlyArray<string | number> = path;
    if (path[0] === "sheets") {
      const sheet = typeof path[1] === "number" ? sheets[path[1]] : undefined;
      const id = sheet && typeof sheet === "object" ? (sheet as { id?: unknown }).id : undefined;
      if (typeof id !== "string" || !id || path.length < 3) return null;
      sheetId = id;
      rest = path.slice(2);
    }
    if (!rest.length || rest.some((step) => typeof step !== "string")) return null;
    const heldBy = sheetId !== null && rest[0] === "cells" && rest.length >= 3
      ? ((sheets[path[1] as number] as Json).cells as Json | undefined)?.[rest[1] as string]
      : undefined;
    anchored.push({
      sheetId,
      path: rest as string[],
      value,
      ...(heldBy && typeof heldBy === "object"
        ? { cell: { value: (heldBy as Json).value, formula: (heldBy as Json).formula } }
        : {}),
    });
  }
  return anchored;
}

/**
 * The content with the values its schema issues point at removed, or null when
 * an issue cannot be answered by removing something (the root itself, or an
 * element of a list that is not one of a sheet's own lists).
 */
export function withoutSchemaIssues(
  content: Json,
  issues: ReadonlyArray<SchemaIssueLike>,
  removed?: Array<SchemaRemovedValue>,
): Json | null {
  const clone = structuredClone(content);
  // An issue inside an element of one of a sheet's own lists (a drawing
  // object, a conditional-format rule, a validation: `sheets.N.<list>.M...`)
  // sets that whole element aside: removing one key inside it can leave an
  // element the schema still refuses as a whole, and the element's place in
  // its list moves with every insert, so it goes back by its id when it has
  // one (else at the place it held). Taken out after the other issues, so
  // their indices still hold.
  const elementsAside = new Map<unknown[], Set<number>>();
  const elementOf = (path: ReadonlyArray<string | number>) => {
    if (path[0] !== "sheets" || typeof path[1] !== "number" || typeof path[2] !== "string" || typeof path[3] !== "number") return null;
    const sheet = (clone.sheets as unknown[] | undefined)?.[path[1]];
    const list = sheet && typeof sheet === "object" ? (sheet as Json)[path[2]] : undefined;
    const element = Array.isArray(list) ? list[path[3]] : undefined;
    if (!Array.isArray(list) || element === undefined) return null;
    const id = element && typeof element === "object" ? (element as { id?: unknown }).id : undefined;
    return { list, index: path[3], id: typeof id === "string" && id ? id : null };
  };
  const at = (path: ReadonlyArray<string | number>): unknown => {
    let node: unknown = clone;
    for (const step of path) {
      if (!node || typeof node !== "object") return undefined;
      node = (node as Record<string | number, unknown>)[step];
    }
    return node;
  };
  for (const issue of issues) {
    let path = issue.path.filter((step): step is string | number => typeof step !== "symbol");
    // A cell that a value set aside left still refused (a cell stripped of
    // the `value` the schema refused is not a cell at all) is set aside
    // whole, with what was already set aside from it put back into it: one
    // such cell left the whole workbook read-only. The cell goes back on save
    // unless the author typed into it.
    if (
      removed &&
      path[0] === "sheets" && typeof path[1] === "number" && path[2] === "cells" && typeof path[3] === "string"
    ) {
      const cellPath = path.slice(0, 4);
      const within = (entry: SchemaRemovedValue) =>
        !entry.element && entry.path.length > 4 && cellPath.every((step, index) => entry.path[index] === step);
      const earlier = removed.filter(within);
      const cell = at(cellPath);
      if (earlier.length && cell && typeof cell === "object" && !Array.isArray(cell)) {
        const whole: Json = structuredClone(cell as Json);
        for (const entry of earlier) {
          if (entry.path.length === 5) whole[entry.path[4] as string] = entry.value;
        }
        if (earlier.every((entry) => entry.path.length === 5)) {
          removed.splice(0, removed.length, ...removed.filter((entry) => !within(entry)));
          path = cellPath;
          const holder = at(path.slice(0, -1)) as Json;
          removed.push({ path: [...path], value: whole });
          delete holder[path[3] as string];
          continue;
        }
      }
    }
    const element = elementOf(path);
    if (element) {
      const aside = elementsAside.get(element.list) ?? new Set<number>();
      if (!aside.has(element.index)) {
        aside.add(element.index);
        removed?.push({
          path: path.slice(0, 3),
          value: structuredClone(element.list[element.index]),
          element: { id: element.id, index: element.index },
        });
      }
      elementsAside.set(element.list, aside);
      continue;
    }
    if (issue.code === "unrecognized_keys") {
      const holder = at(path);
      if (!holder || typeof holder !== "object" || Array.isArray(holder)) return null;
      for (const key of issue.keys ?? []) {
        if (key in (holder as Json)) {
          removed?.push({ path: [...path, key], value: (holder as Json)[key] });
        }
        delete (holder as Json)[key];
      }
      continue;
    }
    if (path.length === 0) return null;
    const holder = at(path.slice(0, -1));
    if (!holder || typeof holder !== "object" || Array.isArray(holder)) return null;
    const key = path[path.length - 1];
    if (key in (holder as Record<string | number, unknown>)) {
      removed?.push({ path: [...path], value: (holder as Record<string | number, unknown>)[key] });
    }
    delete (holder as Record<string | number, unknown>)[key];
  }
  for (const [list, aside] of elementsAside) {
    const kept = list.filter((_, index) => !aside.has(index));
    list.splice(0, list.length, ...kept);
  }
  return clone;
}

/** What parsing with values set aside gave: the parse, and what was set
 * aside (null when nothing was, or what was could not be placed back). */
export interface SetAsideParse<P> {
  readonly parsed: P;
  /** The content the parse read (``content`` itself when nothing was set aside). */
  readonly content: Json | null;
  readonly setAside: SpreadsheetSetAsideValue[] | null;
}

type SafeParse = { success: boolean; error?: { issues: ReadonlyArray<SchemaIssueLike> } | undefined };

/**
 * ``parse(content)``; when it refuses, the refused values are set aside and the
 * rest parsed again (a few rounds: removing a value can surface an issue it
 * was hiding). ``setAside`` is non-null only when the parse succeeded and every
 * value set aside can be put back by a save -- the condition for the workbook
 * to stay editable.
 */
export function parseSettingAside<P extends SafeParse>(
  content: Json,
  parse: (candidate: Json) => P,
  rounds = 3,
): SetAsideParse<P> {
  let parsed = parse(content);
  if (parsed.success) return { parsed, content, setAside: null };
  let candidate: Json | null = content;
  const removed: SchemaRemovedValue[] = [];
  for (let round = 0; round < rounds && candidate && !parsed.success; round += 1) {
    candidate = withoutSchemaIssues(candidate, parsed.error?.issues ?? [], removed);
    if (candidate) parsed = parse(candidate);
  }
  const setAside = parsed.success ? anchorSetAsideValues(content, removed) : null;
  return { parsed, content: parsed.success ? candidate : null, setAside };
}

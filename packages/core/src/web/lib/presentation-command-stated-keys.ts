/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PRESENTATION_LEVEL_STYLE_KEYS, type PresentationLevelStyleKey } from "./presentation-runs";

export const COMMAND_STATED_KEYS_FIELD = "commandStatedKeys";

type LooseRecord = Record<string, unknown>;

function asRecord(value: unknown): LooseRecord | undefined {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as LooseRecord) : undefined;
}

/**
 * The level keys `style` names as whole-object command formats, as a set in the canonical order (a co-edited
 * array may merge with repeats or in any order): only keys the style still states.
 */
export function readCommandStatedKeys(style: unknown): PresentationLevelStyleKey[] {
  const record = asRecord(style);
  const raw = record?.[COMMAND_STATED_KEYS_FIELD];
  if (!record || !Array.isArray(raw)) return [];
  const named = new Set(raw.filter((key): key is string => typeof key === "string"));
  return PRESENTATION_LEVEL_STYLE_KEYS.filter((key) => named.has(key) && record[key] !== undefined && record[key] !== null);
}

/**
 * `style` with `add` named and `remove` no longer named as whole-object command formats, written in its regular
 * form (canonical order, no repeats, only keys the style states; the field gone when none is left).
 */
export function withCommandStatedKeys(
  style: LooseRecord,
  change: { add?: readonly string[]; remove?: readonly string[] },
): LooseRecord {
  const named = new Set<string>(readCommandStatedKeysLoose(style));
  for (const key of change.add ?? []) named.add(key);
  for (const key of change.remove ?? []) named.delete(key);
  const regular = PRESENTATION_LEVEL_STYLE_KEYS.filter(
    (key) => named.has(key) && style[key] !== undefined && style[key] !== null,
  );
  const next = { ...style };
  if (regular.length) next[COMMAND_STATED_KEYS_FIELD] = regular;
  else delete next[COMMAND_STATED_KEYS_FIELD];
  return next;
}

function readCommandStatedKeysLoose(style: LooseRecord): string[] {
  const raw = style[COMMAND_STATED_KEYS_FIELD];
  return Array.isArray(raw) ? raw.filter((key): key is string => typeof key === "string") : [];
}

function runTakingBox(run: unknown, box: LooseRecord, keys: readonly PresentationLevelStyleKey[]): unknown {
  const record = asRecord(run);
  const style = asRecord(record?.style);
  const inherited = Array.isArray(style?.inheritedKeys) ? (style!.inheritedKeys as unknown[]) : [];
  if (!record || !style || !inherited.length) return run;
  let next: LooseRecord | undefined;
  for (const key of keys) {
    if (!inherited.includes(key)) continue;
    const value = box[key];
    const sameFace = key !== "fontFamily" || (style.themeFontRole ?? undefined) === (box.themeFontRole ?? undefined);
    if (JSON.stringify(style[key]) === JSON.stringify(value) && sameFace) continue;
    next ??= { ...style };
    next[key] = value;
    if (key === "fontFamily") {
      if (box.themeFontRole === "heading" || box.themeFontRole === "body") next.themeFontRole = box.themeFontRole;
      else delete next.themeFontRole;
    }
    if (key === "color") {
      // The box's colour is the command's, not a list style's or a Shape Style's.
      delete next.colorFromList;
      delete next.recipeColor;
    }
  }
  return next ? { ...record, style: next } : run;
}

function runsTakingBox(runs: unknown, box: LooseRecord, keys: readonly PresentationLevelStyleKey[]): unknown {
  if (!Array.isArray(runs)) return runs;
  let changed = false;
  const next = runs.map((run) => {
    const taken = runTakingBox(run, box, keys);
    if (taken !== run) changed = true;
    return taken;
  });
  return changed ? next : runs;
}

/**
 * `block` with every run that leaves a whole-object command's key to its list level drawing the box's value
 * for it -- the run still leaves it to its level (its `inheritedKeys` stay), the value it caches is the box's.
 * The block itself when no run changes.
 */
export function blockWithCommandStatedRuns<T>(block: T): T {
  const record = asRecord(block);
  const box = asRecord(record?.style);
  if (!record || !box) return block;
  const keys = readCommandStatedKeys(box);
  if (!keys.length) return block;
  let next: LooseRecord | undefined;
  const runs = runsTakingBox(record.runs, box, keys);
  if (runs !== record.runs) next = { ...record, runs };
  if (Array.isArray(record.paragraphs)) {
    let changed = false;
    const paragraphs = (record.paragraphs as unknown[]).map((paragraph) => {
      const row = asRecord(paragraph);
      if (!row) return paragraph;
      const taken = runsTakingBox(row.runs, box, keys);
      if (taken === row.runs) return paragraph;
      changed = true;
      return { ...row, runs: taken };
    });
    if (changed) next = { ...(next ?? record), paragraphs };
  }
  return (next ?? block) as T;
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the page's held edits (AutoSave off: changed on screen, written by
 * Save) supersede in the PDF the reader shows: the annotation objects whose
 * new version is drawn by the page's own layers until Save, the PDF's copy
 * hidden meanwhile (lib/pdf/form-appearance-storage.ts
 * syncPdfHeldAnnotationStorage). Word and WPS show an unsaved change and
 * nothing else; showing the old copy beside it, or under it, is what made a
 * moved shape or retyped box look doubled.
 *
 * Read from the page content's annotation rows: each row the held content
 * changes (a new or edited row; a delete is a row too) names its object by
 * the wire directive it carries -- an MRAI object by its `/NM`
 * (`MRNote:<token>`), a source object by its pdf.js reference -- and its
 * reader object by the row id.
 */
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";

export type PdfHeldAnnotationFamily =
  | "shape"
  | "signature"
  | "stamp"
  | "link"
  | "ink"
  | "measurement"
  | "textBox"
  | "note";

export interface PdfHeldAnnotationTargets {
  /** `/NM` names of the MRAI annotations a held edit replaces. */
  readonly names: ReadonlySet<string>;
  /** pdf.js references of the source annotations a held edit replaces. */
  readonly refs: ReadonlySet<string>;
  /** Pages those rows are on (0: a row that names none). */
  readonly pages: ReadonlySet<number>;
  /** Reader ids of the held objects, by family. */
  readonly ids: Readonly<Record<PdfHeldAnnotationFamily, ReadonlySet<string>>>;
  /** Reader ids of the objects the held edits made (no saved row). */
  readonly created?: ReadonlySet<string>;
  /** Equal for equal targets (what the reader repaints by). */
  readonly key: string;
  /** Written by Save: still hidden on the document shown until the saved
   * revision replaces it (lifting them first showed the old copies again
   * for that moment), and nothing on any later document. */
  readonly retired?: boolean;
}

/** `targets` once Save has written them (see `retired`). */
export function retirePdfHeldAnnotationTargets(
  targets: PdfHeldAnnotationTargets | null,
): PdfHeldAnnotationTargets | null {
  return targets && !targets.retired ? Object.freeze({ ...targets, retired: true }) : targets;
}

const ROW_FAMILIES: ReadonlyArray<readonly [string, PdfHeldAnnotationFamily]> = [
  ["__mrnote_native_shape__", "shape"],
  ["__mrnote_native_signature__", "signature"],
  ["__mrnote_native_stamp__", "stamp"],
  ["__mrnote_native_link__", "link"],
  ["__mrnote_native_ink__", "ink"],
  ["__mrnote_native_measurement__", "measurement"],
  ["__mrnote_native_text_box__", "textBox"],
];

const DIRECTIVE = /^\[\[MRNOTE_NATIVE_([A-Z0-9_]+?)_V\d+\|[a-z_-]+\|([^|\]\n]+)/;

type Row = { id?: unknown; text?: unknown; page?: unknown };

function rowsOf(content: Record<string, unknown> | null | undefined): Row[] {
  const rows = content?.annotations;
  return Array.isArray(rows) ? rows.filter((row): row is Row => Boolean(row) && typeof row === "object") : [];
}

/** The targets of what `held` changes against `base`; null when nothing. */
export function pdfHeldAnnotationTargets(
  base: Record<string, unknown> | null | undefined,
  held: Record<string, unknown> | null | undefined,
): PdfHeldAnnotationTargets | null {
  const before = new Map<string, Row>();
  for (const row of rowsOf(base)) if (typeof row.id === "string") before.set(row.id, row);
  const names = new Set<string>();
  const refs = new Set<string>();
  const pages = new Set<number>();
  const ids = Object.fromEntries(
    ["shape", "signature", "stamp", "link", "ink", "measurement", "textBox", "note"].map((family) => [family, new Set<string>()]),
  ) as Record<PdfHeldAnnotationFamily, Set<string>>;
  const created = new Set<string>();
  for (const row of rowsOf(held)) {
    if (typeof row.id !== "string" || typeof row.text !== "string") continue;
    const prior = before.get(row.id);
    if (prior && pageDocumentJsonEquals(prior, row)) continue;
    const directive = DIRECTIVE.exec(row.text);
    if (!directive) continue;
    const [, kind, target] = directive;
    if (kind === "TEXT" || kind === "MARKUP" || kind.endsWith("_SOURCE")) {
      refs.add(target);
    } else {
      // An object of MRAI's own the held edits made: its reader row too.
      if (!prior) {
        const madeFamily = ROW_FAMILIES.find(([prefix]) => (row.id as string).startsWith(prefix));
        created.add(madeFamily ? (row.id as string).slice(madeFamily[0].length) : (row.id as string));
      }
      names.add(`MRNote:${target}`);
      // A proofing mark's replacement text is a child of it.
      if (kind === "PROOFING") names.add(`MRNote:${target}:replace-mark`);
    }
    pages.add(typeof row.page === "number" && Number.isFinite(row.page) ? row.page : 0);
    const family = ROW_FAMILIES.find(([prefix]) => row.id!.toString().startsWith(prefix));
    if (family) ids[family[1]].add((row.id as string).slice(family[0].length));
    else if (kind === "NOTE" || kind === "MARKUP_CREATE" || kind === "PROOFING") ids.note.add(row.id);
  }
  if (names.size === 0 && refs.size === 0) return null;
  const key = JSON.stringify([[...names].sort(), [...refs].sort(),
    Object.entries(ids).map(([family, set]) => [family, [...set].sort()])]);
  return Object.freeze({ names, refs, pages, ids, created, key });
}

/** Targets named directly (a held text box by its id), not read from rows. */
export function pdfHeldAnnotationTargetsOf(input: Readonly<{
  names?: Iterable<string>;
  refs?: Iterable<string>;
  pages?: Iterable<number>;
  ids?: Partial<Record<PdfHeldAnnotationFamily, Iterable<string>>>;
}>): PdfHeldAnnotationTargets | null {
  const names = new Set(input.names ?? []);
  const refs = new Set(input.refs ?? []);
  if (names.size === 0 && refs.size === 0) return null;
  const ids = Object.fromEntries(
    ["shape", "signature", "stamp", "link", "ink", "measurement", "textBox", "note"].map((family) =>
      [family, new Set(input.ids?.[family as PdfHeldAnnotationFamily] ?? [])]),
  ) as Record<PdfHeldAnnotationFamily, Set<string>>;
  const key = JSON.stringify([[...names].sort(), [...refs].sort(),
    Object.entries(ids).map(([family, set]) => [family, [...set].sort()])]);
  return Object.freeze({ names, refs, pages: new Set(input.pages ?? []), ids, key });
}

/** Two sets of targets as one. A part Save has written (retired) goes with
 * one still held: only the held part stays hidden then. */
export function mergePdfHeldAnnotationTargets(
  left: PdfHeldAnnotationTargets | null,
  right: PdfHeldAnnotationTargets | null,
): PdfHeldAnnotationTargets | null {
  if (!left || !right) return left ?? right;
  if (Boolean(left.retired) !== Boolean(right.retired)) return left.retired ? right : left;
  const merged = pdfHeldAnnotationTargetsOf({
    names: [...left.names, ...right.names],
    refs: [...left.refs, ...right.refs],
    pages: [...left.pages, ...right.pages],
    ids: Object.fromEntries(Object.keys(left.ids).map((family) => [family,
      [...left.ids[family as PdfHeldAnnotationFamily], ...right.ids[family as PdfHeldAnnotationFamily]]])),
  });
  return merged && left.retired ? Object.freeze({ ...merged, retired: true }) : merged;
}

/**
 * The annotation rows a loaded page shows while AutoSave-off edits are held
 * (a peer's save read in, the working copy's save at Save): the rows the
 * held edits changed or added are theirs, the rows they took away stay
 * gone, and every other row -- a peer's new comment among them -- is the
 * loaded page's. Laid over whole, the held array hid what the peer added and
 * the next Save wrote the page without it.
 */
export function withHeldAnnotationRows(
  loadedRows: unknown,
  baseRows: unknown,
  heldRows: unknown,
): unknown[] {
  const base = new Map<string, Row>();
  for (const row of rowsOf({ annotations: baseRows })) if (typeof row.id === "string") base.set(row.id, row);
  const held = new Map<string, Row>();
  const heldOrder: Row[] = [];
  for (const row of rowsOf({ annotations: heldRows })) {
    if (typeof row.id !== "string") { heldOrder.push(row); continue; }
    held.set(row.id, row); heldOrder.push(row);
  }
  const changed = new Set<string>();
  for (const [id, row] of held) {
    const saved = base.get(id);
    if (!saved || !pageDocumentJsonEquals(saved, row)) changed.add(id);
  }
  const removed = new Set<string>();
  for (const id of base.keys()) if (!held.has(id)) removed.add(id);
  const out: unknown[] = [];
  const seen = new Set<string>();
  for (const row of rowsOf({ annotations: loadedRows })) {
    const id = typeof row.id === "string" ? row.id : undefined;
    if (id !== undefined && removed.has(id)) continue;
    if (id !== undefined && changed.has(id)) { if (!seen.has(id)) { out.push(held.get(id)); seen.add(id); } continue; }
    if (id !== undefined) seen.add(id);
    out.push(row);
  }
  // Rows the held edits added (or changed) that the loaded page lacks.
  for (const row of heldOrder) {
    const id = typeof row.id === "string" ? row.id : undefined;
    if (id === undefined) { out.push(row); continue; }
    if (!seen.has(id) && changed.has(id)) { out.push(row); seen.add(id); }
  }
  return out;
}

/**
 * The annotation rows a write that is not Save sends (a text box written
 * now, AutoSave off): the rows the held edits changed go back to their saved
 * version, every other row stays as the write has it. One array holds every
 * family's rows, so holding it whole lost the write's own row.
 */
export function withoutHeldAnnotationRows(
  rows: ReadonlyArray<unknown>,
  baseRows: unknown,
  heldRows: unknown,
): unknown[] {
  const base = new Map<string, Row>();
  for (const row of rowsOf({ annotations: baseRows })) if (typeof row.id === "string") base.set(row.id, row);
  const heldIds = new Set<string>();
  const changed = new Set<string>();
  for (const row of rowsOf({ annotations: heldRows })) {
    if (typeof row.id !== "string") continue;
    heldIds.add(row.id);
    const saved = base.get(row.id);
    if (!saved || !pageDocumentJsonEquals(saved, row)) changed.add(row.id);
  }
  // A saved row the held edits took away is theirs too.
  for (const id of base.keys()) if (!heldIds.has(id)) changed.add(id);
  const out: unknown[] = [];
  const seen = new Set<string>();
  for (const row of rows) {
    const id = row && typeof row === "object" ? (row as Row).id : undefined;
    if (typeof id === "string" && changed.has(id)) {
      const saved = base.get(id);
      if (saved && !seen.has(id)) { out.push(saved); seen.add(id); }
      continue;
    }
    if (typeof id === "string") seen.add(id);
    out.push(row);
  }
  for (const id of changed) {
    const saved = base.get(id);
    if (saved && !seen.has(id)) out.push(saved);
  }
  return out;
}

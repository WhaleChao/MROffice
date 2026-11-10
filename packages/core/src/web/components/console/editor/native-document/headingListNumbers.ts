/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The numbers a Word numbering gives the headings it numbers ("Heading 1"
 * bound to a multilevel list: 1 / 1.1 / 1.1.1, 第一章 …). An imported
 * numbered heading keeps its numbering on `attrs.source.listNumbering`
 * (`numId`, `level` and the numbering's own levels, `definition`); the editor
 * never holds such a heading in a list. Word counts every paragraph of one
 * numbering instance together, in document order, each level restarting
 * after the levels above it; the heading shows its level's text ("%1.%2")
 * and then its suffix (a tab, a space or nothing).
 *
 * One counter walk serves the heading's own drawn number
 * (headingNumberingExtension) and the table of contents' rows
 * (tableOfContents.ts), so both read the same number.
 */

import { HOLDS_HEADING } from "../../../../lib/editor/sheet-facts";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";

import { bulletGlyph, normalizeBulletStyle } from "./bulletListStyle";

import {
  ListCounterState,
  listMarkerText,
  normalizeListDefinition,
  type ListDefinition,
  type ListLevelSuffix,
} from "./listLevelDefinition";
import { fragmentCarries } from "./paged/fragmentRoles";

export interface HeadingListNumber {
  readonly label: string;
  readonly suffix: ListLevelSuffix;
}

const definitions = new WeakMap<object, ListDefinition | null>();

function definitionOf(raw: unknown): ListDefinition | null {
  if (!raw || typeof raw !== "object") return null;
  if (definitions.has(raw)) return definitions.get(raw) ?? null;
  const definition = normalizeListDefinition(raw);
  definitions.set(raw, definition);
  return definition;
}

/**
 * A numberer for headings visited in document order: each call counts one
 * heading of a Word numbering and returns its number, or null for a heading
 * no numbering numbers.
 */
export function createHeadingListNumberer(): (
  attrs: Record<string, unknown> | null | undefined,
) => HeadingListNumber | null {
  // One count per `chain`: Word counts the numbering instances of one
  // abstract definition as one list, and an instance's `w:startOverride`
  // starts that count again at its first item at the level (corpus R09-D04:
  // "1" for a restarting instance's Heading 1, then "2" for the next one of
  // another instance). A heading without a chain counts by its own instance.
  const counts = new Map<string, ListCounterState>();
  const restarted = new Set<string>();
  return (attrs) => {
    const source = attrs?.source;
    if (!source || typeof source !== "object") return null;
    const numbering = (source as { listNumbering?: unknown }).listNumbering;
    if (!numbering || typeof numbering !== "object") return null;
    const { type, numId, level, definition: rawDefinition, chain, restartLevels } = numbering as Record<string, unknown>;
    if (type !== "ordered" || typeof numId !== "number" || !Number.isInteger(numId)) return null;
    const safeLevel = typeof level === "number" && Number.isInteger(level) ? Math.min(Math.max(level, 0), 8) : 0;
    const definition = definitionOf(rawDefinition);
    if (!definition) return null;
    const key = typeof chain === "string" && /^[an][0-9]{1,10}$/.test(chain) ? chain : `n${numId}`;
    let state = counts.get(key);
    if (!state) {
      state = new ListCounterState();
      counts.set(key, state);
    }
    const restartKey = `${numId}:${safeLevel}`;
    if (
      Array.isArray(restartLevels) &&
      restartLevels.includes(safeLevel) &&
      !restarted.has(restartKey)
    ) {
      restarted.add(restartKey);
      state.seedLevel(safeLevel, definition.levels[safeLevel].start);
    }
    const values = state.advance(definition, safeLevel);
    const label = listMarkerText(definition, safeLevel, values);
    if (!label) return null;
    return { label, suffix: definition.levels[safeLevel]?.suffix ?? "tab" };
  };
}

const byDoc = new WeakMap<ProseMirrorNode, ReadonlyMap<number, HeadingListNumber>>();

/** Every Word-numbered heading's number in `doc`, by the heading's position. */
export function headingListNumbers(doc: ProseMirrorNode): ReadonlyMap<number, HeadingListNumber> {
  const known = byDoc.get(doc);
  if (known) return known;
  const numbers = new Map<number, HeadingListNumber>();
  const numberer = createHeadingListNumberer();
  HOLDS_HEADING.descendants(doc, (node, pos) => {
    if (node.type.name !== "heading") return true;
    // The rest of a heading a page boundary cut is the same heading.
    if (!fragmentCarries("listLabel", node.attrs)) return false;
    const number = numberer(node.attrs);
    if (number) numbers.set(pos, number);
    return false;
  });
  byDoc.set(doc, numbers);
  return numbers;
}


export interface HeadingListBullet extends HeadingListNumber {
  /** The bullet's own font (a symbol font's glyph), or null for the heading's. */
  readonly font: string | null;
}

const bulletsByDoc = new WeakMap<ProseMirrorNode, ReadonlyMap<number, HeadingListBullet>>();

/**
 * Every bulleted heading's bullet in `doc`, by the heading's position: an
 * imported heading a Word bullet numbering formats (`source.listNumbering`
 * of type "bullet", with the level's glyph and font or its gallery style).
 * Word draws the bullet in the heading's hanging indent and tabs to its text
 * (corpus R92-D13's Heading 3 / Heading 4 bullets); a table of contents
 * lists no bullet, so they are kept apart from the numbers.
 */
export function headingListBullets(doc: ProseMirrorNode): ReadonlyMap<number, HeadingListBullet> {
  const known = bulletsByDoc.get(doc);
  if (known) return known;
  const bullets = new Map<number, HeadingListBullet>();
  HOLDS_HEADING.descendants(doc, (node, pos) => {
    if (node.type.name !== "heading") return true;
    if (!fragmentCarries("listLabel", node.attrs)) return false;
    const source = node.attrs?.source;
    const numbering = source && typeof source === "object"
      ? (source as { listNumbering?: unknown }).listNumbering
      : null;
    if (!numbering || typeof numbering !== "object") return false;
    const { type, glyph, font, bulletStyle } = numbering as Record<string, unknown>;
    if (type !== "bullet") return false;
    const ownGlyph = typeof glyph === "string" && glyph.length >= 1 && glyph.length <= 4 && !/[\u0000-\u001f]/u.test(glyph)
      ? glyph
      : null;
    if (!ownGlyph && typeof bulletStyle !== "string") return false;
    const label = ownGlyph ?? bulletGlyph(normalizeBulletStyle(bulletStyle));
    const face = ownGlyph && typeof font === "string" && /^[^"\\;{}<>]{1,64}$/.test(font.trim()) ? font.trim() : null;
    bullets.set(pos, { label, suffix: "tab", font: face });
    return false;
  });
  bulletsByDoc.set(doc, bullets);
  return bullets;
}

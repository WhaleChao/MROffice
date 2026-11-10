/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design > Document Formatting > Save as a New Style Set: Word saves the
 * document's Quick Styles as a template in the user's QuickStyles folder
 * and lists it under "Custom" in the style-set gallery, where applying it
 * restyles a document like a built-in set; right-click > Delete removes it.
 *
 * A document's Quick Styles in MRAI are always one of Word's style sets
 * with a Paragraph Spacing choice (the styles themselves are Word's), so a
 * saved set is that pair under its name: applying it gives a document
 * exactly the styles the saved one had. The library is the user's, kept by
 * the browser (as Word's folder is the computer's).
 */

import { useEffect, useState } from "react";

import { documentParagraphSpacingPreset } from "./document-paragraph-spacing";
import type { DocumentDesign } from "./document-design";
import { isWordStyleSetId, type WordStyleSetId } from "./office-word-style-sets";

export interface DocumentCustomStyleSet {
  readonly id: string;
  readonly name: string;
  readonly styleSet: WordStyleSetId;
  readonly paragraphSpacing: DocumentDesign["paragraphSpacing"];
}

const STORAGE_KEY = "mrai.docs.customStyleSets";
const CHANGE_EVENT = "mrai:docs-custom-style-sets";
const MAX_SETS = 50;
const MAX_NAME = 120;

function normalizeName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const name = value.replace(/[\u0000-\u001f\u007f]+/g, "").trim().slice(0, MAX_NAME);
  return name || null;
}

function normalizeEntry(value: unknown): DocumentCustomStyleSet | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const name = normalizeName(record.name);
  const id = typeof record.id === "string" && /^[a-z0-9-]{1,64}$/.test(record.id) ? record.id : null;
  if (!name || !id || !isWordStyleSetId(record.styleSet)) return null;
  const preset = documentParagraphSpacingPreset(record.paragraphSpacing);
  return Object.freeze({
    id,
    name,
    styleSet: record.styleSet,
    paragraphSpacing: preset && preset.id !== "default" ? preset.id : null,
  });
}

/** The user's saved style sets, oldest first (Word lists its Custom folder by name). */
export function readCustomStyleSets(): DocumentCustomStyleSet[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = JSON.parse(window.localStorage.getItem(STORAGE_KEY) ?? "[]");
    if (!Array.isArray(raw)) return [];
    return raw
      .map(normalizeEntry)
      .filter((entry): entry is DocumentCustomStyleSet => entry !== null)
      .slice(0, MAX_SETS)
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch {
    return [];
  }
}

function write(sets: readonly DocumentCustomStyleSet[]): boolean {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(sets.slice(0, MAX_SETS)));
    window.dispatchEvent(new Event(CHANGE_EVENT));
    return true;
  } catch {
    return false;
  }
}

/**
 * Save as a New Style Set: the document's style set and Paragraph Spacing
 * under `name`. A set of the same name is replaced, as saving over a file
 * of that name is. Null when the name is empty or the library cannot be kept.
 */
export function saveCustomStyleSet(rawName: string, design: DocumentDesign): DocumentCustomStyleSet | null {
  if (typeof window === "undefined") return null;
  const name = normalizeName(rawName);
  if (!name || !isWordStyleSetId(design.styleSet)) return null;
  const others = readCustomStyleSets().filter((entry) => entry.name.toLowerCase() !== name.toLowerCase());
  if (others.length >= MAX_SETS) return null;
  const entry = Object.freeze({
    id: `set-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
    name,
    styleSet: design.styleSet,
    paragraphSpacing: design.paragraphSpacing,
  });
  return write([...others, entry]) ? entry : null;
}

/** Right-click > Delete on a Custom set. */
export function deleteCustomStyleSet(id: string): boolean {
  if (typeof window === "undefined") return false;
  const sets = readCustomStyleSets();
  const next = sets.filter((entry) => entry.id !== id);
  return next.length !== sets.length && write(next);
}

/** The design a document takes from a saved set: its style set and spacing. */
export function documentDesignWithCustomStyleSet(
  design: DocumentDesign,
  entry: DocumentCustomStyleSet,
): DocumentDesign {
  return Object.freeze({ ...design, styleSet: entry.styleSet, paragraphSpacing: entry.paragraphSpacing });
}

/** Whether a document's styles are a saved set's (the gallery marks it). */
export function customStyleSetMatches(design: DocumentDesign, entry: DocumentCustomStyleSet): boolean {
  return design.styleSet === entry.styleSet && design.paragraphSpacing === entry.paragraphSpacing;
}

/** The saved sets, redrawn when this or another tab of the browser changes them. */
export function useCustomStyleSets(): DocumentCustomStyleSet[] {
  const [sets, setSets] = useState<DocumentCustomStyleSet[]>([]);
  useEffect(() => {
    const read = () => setSets(readCustomStyleSets());
    read();
    const onStorage = (event: StorageEvent) => {
      if (event.key === STORAGE_KEY) read();
    };
    window.addEventListener(CHANGE_EVENT, read);
    window.addEventListener("storage", onStorage);
    return () => {
      window.removeEventListener(CHANGE_EVENT, read);
      window.removeEventListener("storage", onStorage);
    };
  }, []);
  return sets;
}

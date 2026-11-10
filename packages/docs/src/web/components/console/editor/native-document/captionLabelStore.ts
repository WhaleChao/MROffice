/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The viewer's own caption labels (Word's New Label... / Delete Label).
 *
 * Word keeps the custom label list with the user, not in the document: a
 * label made once is offered for every later caption, and deleting it only
 * removes it from the list — captions already written keep their text. The
 * labels a document uses come from its captions (captionLabelOptions); this
 * store adds the labels the viewer made and remembers the ones they deleted.
 *
 * Browser storage is a convenience: every read and write is guarded, and the
 * dialog works without it (labels then last for the session's document use).
 */

import { sanitizeCaptionLabelName } from "@mrnote/core/web/components/console/editor/native-document/captions.ts";

const STORAGE_KEY = "mrai.document.captionLabels.v1";
const MAX_STORED_LABELS = 200;

export interface CaptionLabelStore {
  saved: string[];
  deleted: string[];
}

function cleanList(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  const names: string[] = [];
  const seen = new Set<string>();
  for (const raw of value) {
    const name = sanitizeCaptionLabelName(raw);
    const key = name.toLowerCase();
    if (!name || seen.has(key)) continue;
    seen.add(key);
    names.push(name);
    if (names.length >= MAX_STORED_LABELS) break;
  }
  return names;
}

export function readCaptionLabelStore(): CaptionLabelStore {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : null;
    if (parsed && typeof parsed === "object") {
      const candidate = parsed as { saved?: unknown; deleted?: unknown };
      return {
        saved: cleanList(candidate.saved),
        deleted: cleanList(candidate.deleted),
      };
    }
  } catch {
    // No storage (private window, blocked site data): start empty.
  }
  return { saved: [], deleted: [] };
}

function writeCaptionLabelStore(store: CaptionLabelStore): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
  } catch {
    // Not persisted; the caller keeps the list in memory.
  }
}

const without = (names: readonly string[], name: string) =>
  names.filter((entry) => entry.toLowerCase() !== name.toLowerCase());

/** New Label: remember the name and offer it again if it was deleted. */
export function saveCaptionLabel(
  store: CaptionLabelStore,
  rawName: string,
): CaptionLabelStore {
  const name = sanitizeCaptionLabelName(rawName);
  if (!name) return store;
  const next = {
    saved: cleanList([...without(store.saved, name), name]),
    deleted: without(store.deleted, name),
  };
  writeCaptionLabelStore(next);
  return next;
}

/** Delete Label: stop offering the name; captions that use it are kept. */
export function deleteCaptionLabel(
  store: CaptionLabelStore,
  rawName: string,
): CaptionLabelStore {
  const name = sanitizeCaptionLabelName(rawName);
  if (!name) return store;
  const next = {
    saved: without(store.saved, name),
    deleted: cleanList([...without(store.deleted, name), name]),
  };
  writeCaptionLabelStore(next);
  return next;
}

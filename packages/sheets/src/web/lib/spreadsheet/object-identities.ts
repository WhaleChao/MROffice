/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  freeWorkbookObjectId,
  officeObjectHasImportIdentity,
  withProvenanceOutOfSource,
} from "@mrnote/core/web/lib/office-object-provenance.ts";

type JsonRecord = Record<string, unknown>;

function isRecord(value: unknown): value is JsonRecord {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function objectId(item: JsonRecord): string | null {
  return typeof item.id === "string" && item.id ? item.id : null;
}

/** The fields by which an object names another of its sheet (a connector's
 *  glued ends). */
const OBJECT_REFERENCE_FIELDS = ["startObjectId", "endObjectId"] as const;

















export function normalizeSpreadsheetObjectIdentities<C extends JsonRecord | undefined>(content: C): C {
  if (!content || !Array.isArray(content.sheets)) return content;
  let changed = false;
  // Copy-on-write: only the sheets and lists that change are copied.
  const sheets = (content.sheets as unknown[]).map((sheet) => {
    if (!isRecord(sheet)) return sheet;
    let next: JsonRecord = sheet;
    for (const key of ["objects", "charts"] as const) {
      const items = sheet[key];
      if (!Array.isArray(items)) continue;
      let moved = false;
      const repaired = items.map((item) => {
        if (!isRecord(item)) return item;
        const out = withProvenanceOutOfSource(item);
        if (out !== item) moved = true;
        return out;
      });
      if (moved) {
        next = next === sheet ? { ...sheet } : next;
        next[key] = repaired;
        changed = true;
      }
    }
    return next;
  });

  const taken = new Set<string>();
  for (const sheet of sheets) {
    if (!isRecord(sheet) || !Array.isArray(sheet.objects)) continue;
    for (const item of sheet.objects) {
      const id = isRecord(item) ? objectId(item) : null;
      if (id && officeObjectHasImportIdentity(item as JsonRecord)) taken.add(id);
    }
  }
  for (const key of ["charts", "objects"] as const) {
    sheets.forEach((sheet, sheetIndex) => {
      if (!isRecord(sheet) || !Array.isArray(sheet[key])) return;
      const sheetId = String(sheet.id || "sheet");
      const renamed = new Map<string, string>();
      const items = (sheet[key] as unknown[]).map((item) => {
        if (!isRecord(item)) return item;
        const id = objectId(item);
        if (id === null || officeObjectHasImportIdentity(item)) return item;
        const fresh = freeWorkbookObjectId(id, sheetId, taken);
        taken.add(fresh);
        if (fresh === id) return item;
        if (!renamed.has(id)) renamed.set(id, fresh);
        return { ...item, id: fresh };
      });
      if (renamed.size === 0) return;
      changed = true;
      sheets[sheetIndex] = {
        ...sheet,
        [key]: key === "objects" ? followRenamedObjects(items, renamed) : items,
      };
    });
  }
  return changed ? ({ ...content, sheets } as C) : content;
}

/** A connector names the object it is glued to by id: it follows the object
 *  renamed, unless another object of the sheet still holds the old id. */
function followRenamedObjects(objects: unknown[], renamed: ReadonlyMap<string, string>): unknown[] {
  const held = new Set(objects.map((item) => (isRecord(item) ? objectId(item) : null)));
  return objects.map((item) => {
    if (!isRecord(item)) return item;
    let next = item;
    for (const field of OBJECT_REFERENCE_FIELDS) {
      const target = item[field];
      if (typeof target === "string" && renamed.has(target) && !held.has(target)) {
        next = { ...next, [field]: renamed.get(target) };
      }
    }
    return next;
  });
}

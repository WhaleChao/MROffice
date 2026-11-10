/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The repairs an import made to open a damaged Office file (the server's
 * `office_package_repair`: page `content_json.importRepairs`, resource
 * `metadata_json.import_repairs`), in the shape of a spreadsheet repair log:
 * "Removed Part: /xl/styles.xml (XML error ...)".
 */
export type ImportRepairAction = "removed_part" | "replaced_part" | "removed_records" | "repaired_records";

export interface ImportRepair {
  part: string;
  action: ImportRepairAction;
  detail: string;
}

const ACTIONS = new Set<ImportRepairAction>(["removed_part", "replaced_part", "removed_records", "repaired_records"]);
const MAX_REPAIRS = 64;

export function readImportRepairs(value: unknown): ImportRepair[] {
  if (!Array.isArray(value)) return [];
  const repairs: ImportRepair[] = [];
  for (const entry of value.slice(0, MAX_REPAIRS)) {
    if (!entry || typeof entry !== "object") continue;
    const { part, action, detail } = entry as Record<string, unknown>;
    if (typeof action !== "string" || !ACTIONS.has(action as ImportRepairAction)) continue;
    repairs.push({
      part: typeof part === "string" ? part.slice(0, 300) : "",
      action: action as ImportRepairAction,
      detail: typeof detail === "string" ? detail.slice(0, 600) : "",
    });
  }
  return repairs;
}

/** One line of the log: "<Action>: <part> (<detail>)". */
export function importRepairLine(repair: ImportRepair, actionLabel: string): string {
  const where = repair.part ? ` ${repair.part}` : "";
  const why = repair.detail ? ` (${repair.detail})` : "";
  return `${actionLabel}:${where}${why}`;
}

/** Identifies one report (a later import with other repairs shows again). */
export function importRepairsFingerprint(repairs: readonly ImportRepair[]): string {
  const text = JSON.stringify(repairs);
  let hash = 2_166_136_261;
  for (let index = 0; index < text.length; index += 1) hash = Math.imul(hash ^ text.charCodeAt(index), 16_777_619);
  return `${repairs.length}:${(hash >>> 0).toString(16)}`;
}

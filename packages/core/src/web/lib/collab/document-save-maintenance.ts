/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Schema } from "@tiptap/pm/model";
import { jsonEquals, type JsonValue } from "./json-crdt";

const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const equal = (a: unknown, b: unknown) => jsonEquals(a as JsonValue, b as JsonValue);

/** Permit only schema materialization of the same live original. Generated
 * paragraph/heading identities may fill an absent slot, never replace a peer's
 * identity. Text, marks, order, existing attributes and opaque source metadata
 * remain exact. This is validation for an explicit save, not an author writer. */
export function isDocumentSaveMaintenance(original: unknown, requested: unknown, schema: Schema): boolean {
  const markMaintenance = (before: unknown, after: unknown): boolean => {
    if (equal(before, after)) return true;
    if (!record(before) || !record(after) || before.type !== after.type || typeof after.type !== "string") return false;
    const spec = schema.marks[after.type]?.spec;
    if (!spec) return false;
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      if (key !== "attrs") { if (!equal(before[key], after[key])) return false; continue; }
      const left = before.attrs ?? {}, right = after.attrs ?? {};
      if (!record(left) || !record(right)) return false;
      for (const attr of new Set([...Object.keys(left), ...Object.keys(right)])) {
        if (equal(left[attr], right[attr])) continue;
        if (Object.hasOwn(left, attr) || !Object.hasOwn(right, attr) || !spec.attrs ||
            !Object.hasOwn(spec.attrs, attr)) return false;
        const definition = spec.attrs[attr];
        if (!Object.hasOwn(definition, "default") || !equal(definition.default, right[attr])) return false;
      }
    }
    return true;
  };
  const visit = (before: unknown, after: unknown): boolean => {
    if (equal(before, after)) return true;
    if (!record(before) || !record(after) || before.type !== after.type ||
        typeof after.type !== "string") return false;
    const spec = schema.nodes[after.type]?.spec;
    if (!spec) return false;
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      if (key === "attrs") {
        const left = before.attrs ?? {}, right = after.attrs ?? {};
        if (!record(left) || !record(right)) return false;
        for (const attr of new Set([...Object.keys(left), ...Object.keys(right)])) {
          if (equal(left[attr], right[attr])) continue;
          if (Object.hasOwn(left, attr) || !Object.hasOwn(right, attr) || !spec.attrs ||
              !Object.hasOwn(spec.attrs, attr)) return false;
          const definition = spec.attrs[attr];
          if (Object.hasOwn(definition, "default") && equal(definition.default, right[attr])) continue;
          if (after.type !== "doc" && (attr === "block_id" || attr === "anchor_id") &&
              definition.default == null && typeof right[attr] === "string" && right[attr].length > 0) continue;
          return false;
        }
      } else if (key === "marks") {
        // A mark materialised by the schema is the same mark: a link written
        // with only its href gains target/rel/title defaults in the view, and
        // that must not stop an explicit save (Copilot could not start on such
        
        // The view orders a text's marks by the schema's mark order, not as
        // the file listed them ([bold, italic, fontSize] comes back as
        // [fontSize, bold, italic]): the same marks, so compared by type
        
        if (!Array.isArray(before.marks) || !Array.isArray(after.marks) ||
            before.marks.length !== after.marks.length) return false;
        const byType = (marks: unknown[]) => [...marks].sort((left, right) => {
          const a = record(left) ? String(left.type) : "", b = record(right) ? String(right.type) : "";
          return a < b ? -1 : a > b ? 1 : 0;
        });
        const left = byType(before.marks), right = byType(after.marks);
        if (!left.every((mark, index) => markMaintenance(mark, right[index]))) return false;
      } else if (key === "content") {
        if (!Array.isArray(before.content) || !Array.isArray(after.content) || before.content.length !== after.content.length ||
            !before.content.every((child, index) => visit(child, (after.content as unknown[])[index]))) return false;
      } else if (!equal(before[key], after[key])) return false;
    }
    return true;
  };
  return visit(original, requested);
}

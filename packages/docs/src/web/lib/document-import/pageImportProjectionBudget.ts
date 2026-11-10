/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The server's batch ceilings (`PAGE_IMPORT_BATCH_MAX_PROJECTED_*`): one
// document page's own node and text budget.
export const PAGE_IMPORT_MAX_PROJECTED_OBJECTS = 1_200_000;
export const PAGE_IMPORT_MAX_PROJECTED_STRING_CHARS = 16_000_000;

export type PageImportProjectionBudgetReason = "object_nodes" | "string_chars";

export class PageImportProjectionBudgetError extends Error {
  constructor(readonly reason: PageImportProjectionBudgetReason) {
    super(
      reason === "object_nodes"
        ? "The import contains too many projected objects."
        : "The import contains too much projected text.",
    );
    this.name = "PageImportProjectionBudgetError";
  }
}

/**
 * Mirrors the server's whole-manifest JSON-tree budget: plain objects count as
 * nodes, and string values count by UTF-16 length. Arrays and object keys are
 * traversed but do not themselves consume the object-node budget.
 */
export class PageImportProjectionBudget {
  private objectNodes = 0;
  private stringChars = 0;

  reserveObjects(count = 1): void {
    if (!Number.isSafeInteger(count) || count < 0) {
      throw new TypeError(
        "Projection object reservation must be non-negative.",
      );
    }
    this.objectNodes += count;
    if (this.objectNodes > PAGE_IMPORT_MAX_PROJECTED_OBJECTS) {
      throw new PageImportProjectionBudgetError("object_nodes");
    }
  }

  reserveString(value: string): void {
    this.stringChars += value.length;
    if (this.stringChars > PAGE_IMPORT_MAX_PROJECTED_STRING_CHARS) {
      throw new PageImportProjectionBudgetError("string_chars");
    }
  }

  consumeValue(value: unknown): void {
    const pending: unknown[] = [value];
    const ancestors = new WeakSet<object>();
    while (pending.length > 0) {
      const current = pending.pop();
      if (typeof current === "string") {
        this.reserveString(current);
        continue;
      }
      if (!current || typeof current !== "object") continue;
      if (ancestors.has(current)) {
        throw new TypeError("Import projection contains a circular value.");
      }
      // Values accepted here are immutable import artifacts. A global seen set
      // also prevents a shared reference from being expanded repeatedly before
      // canonical JSON turns it into independent values.
      ancestors.add(current);
      if (Array.isArray(current)) {
        for (let index = current.length - 1; index >= 0; index -= 1) {
          pending.push(current[index]);
        }
        continue;
      }
      this.reserveObjects();
      const record = current as Record<string, unknown>;
      const values = Object.values(record);
      for (let index = values.length - 1; index >= 0; index -= 1) {
        pending.push(values[index]);
      }
    }
  }
}

export function assertPageImportProjectionBudget(value: unknown): void {
  new PageImportProjectionBudget().consumeValue(value);
}

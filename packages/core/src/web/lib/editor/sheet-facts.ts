/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a sheet holds, kept with the sheet.
 *
 * A key changes one sheet of a document; every other sheet is the node it was.
 * Plugin after plugin walks the whole document on each change to find what few
 * documents hold and few sheets of those (a table of a kind, a heading, a note
 * reference): at 2,747 sheets each walk is 20,000 blocks, and there are a dozen.
 * A fact is asked once of each sheet node, so a key reads it again only for the
 * sheet it changed, and a walk that wants only such nodes visits only the sheets
 * that hold one, at the positions it would have had.
 *
 * `holds` must be true of every node the walk would act on; it may be true of
 * more. A top-level child that is no sheet (a document not paged) is treated the
 * same way, one block at a time.
 */
import type { Node as PmNode } from "@tiptap/pm/model";

type Visit = (node: PmNode, pos: number, parent: PmNode | null, index: number) => boolean | void;

export class SheetFact {
  private readonly bySheet = new WeakMap<PmNode, boolean>();
  private readonly byDoc = new WeakMap<PmNode, boolean>();

  constructor(private readonly holds: (node: PmNode) => boolean) {}

  /** Whether this top-level child, or anything in it, is such a node. */
  sheet(sheet: PmNode): boolean {
    const known = this.bySheet.get(sheet);
    if (known !== undefined) return known;
    let found = this.holds(sheet);
    if (!found) {
      sheet.descendants((node) => {
        if (found) return false;
        if (this.holds(node)) found = true;
        return !found;
      });
    }
    this.bySheet.set(sheet, found);
    return found;
  }

  /** Whether anything in the document is such a node. */
  inDoc(doc: PmNode): boolean {
    const known = this.byDoc.get(doc);
    if (known !== undefined) return known;
    let found = false;
    for (let index = 0; index < doc.childCount && !found; index += 1) found = this.sheet(doc.child(index));
    this.byDoc.set(doc, found);
    return found;
  }

  /** `doc.descendants(visit)`, the sheets that hold no such node left out. */
  descendants(doc: PmNode, visit: Visit): void {
    doc.forEach((sheet, offset, index) => {
      if (!this.sheet(sheet)) return;
      if (visit(sheet, offset, doc, index) === false) return;
      sheet.descendants((node, pos, parent, at) => visit(node, offset + 1 + pos, parent, at));
    });
  }
}

/** Asked by many plugins: one answer per sheet node between them. */
export const HOLDS_TABLE = new SheetFact((node) => node.type.name === "table");
export const HOLDS_HEADING = new SheetFact((node) => node.type.name === "heading");

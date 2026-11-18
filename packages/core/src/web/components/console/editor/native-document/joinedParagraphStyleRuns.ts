/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Extension } from "@tiptap/core";
import type { Node as ProseMirrorNode, Schema, Slice } from "@tiptap/pm/model";
import { Plugin, PluginKey, type EditorState, type Transaction } from "@tiptap/pm/state";
import { ReplaceStep } from "@tiptap/pm/transform";


















type RunLook = {
  /** undefined: a theme face, not traded. */
  family: string | null | undefined;
  sizePt: number | null;
  color: string | null;
  bold: boolean | null;
  italic: boolean | null;
  /** The style's underline ("single", "double", ...); null for none. */
  underline?: string | null;
  /** The style's Superscript / Subscript (`w:vertAlign`); null for the baseline. */
  vertAlign?: "superscript" | "subscript" | null;
};

const record = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;

const bareFamily = (value: unknown) =>
  typeof value === "string" ? value.split(",")[0].trim().replace(/^['"]|['"]$/g, "").toLowerCase() : "";

const pointsOf = (size: unknown): number | null => {
  const value = typeof size === "string" ? Number.parseFloat(size) : Number.NaN;
  if (!Number.isFinite(value)) return null;
  return typeof size === "string" && size.trim().endsWith("pt") ? value : (value * 3) / 4;
};

const px = (points: number) => `${Math.round(((points * 4) / 3) * 10000) / 10000}px`;

/** A paragraph style's run look through its basedOn chain (the nearest style
 *  stating a property decides); null when the chain cannot be read. A theme
 *  face (no w:ascii) leaves the face unknown, which is then not traded. */
export function paragraphStyleRunLook(docAttrs: Record<string, unknown>, styleId: string): RunLook | null {
  const definitions = record(docAttrs.styleDefinitions);
  const rows = Array.isArray(definitions?.styles) ? (definitions!.styles as unknown[]) : [];
  const styles = new Map<string, Record<string, unknown>>();
  for (const row of rows) {
    const style = record(row);
    if (style && typeof style.styleId === "string" && !styles.has(style.styleId)) styles.set(style.styleId, style);
  }
  const chain: Record<string, unknown>[] = [];
  const seen = new Set<string>();
  let current: string | null = styleId;
  while (current) {
    if (seen.has(current) || seen.size >= 64) return null;
    seen.add(current);
    const style = styles.get(current);
    if (!style || style.type !== "paragraph") return null;
    chain.unshift(style);
    current = typeof style.basedOn === "string" ? style.basedOn : null;
  }
  const look: RunLook = { family: null, sizePt: null, color: null, bold: null, italic: null, underline: null, vertAlign: null };
  let themedFace = false;
  for (const style of chain) {
    const rPr = record(style.rPr) ?? {};
    const fonts = record(rPr.fonts);
    if (fonts) {
      const face = typeof fonts.ascii === "string" ? fonts.ascii : typeof fonts.hAnsi === "string" ? fonts.hAnsi : null;
      if (typeof fonts.asciiTheme === "string" || typeof fonts.hAnsiTheme === "string") themedFace = true;
      else if (face) {
        look.family = face;
        themedFace = false;
      }
    }
    if (typeof rPr.fontSizeHalfPoints === "number") look.sizePt = rPr.fontSizeHalfPoints / 2;
    const color = record(rPr.color)?.val;
    if (typeof color === "string" && /^[0-9a-f]{6}$/i.test(color)) look.color = `#${color.toLowerCase()}`;
    if (rPr.bold === true || rPr.bold === false) look.bold = rPr.bold;
    if (rPr.italic === true || rPr.italic === false) look.italic = rPr.italic;
    // Word's `wave` is the editor's `wavy` (underlineStyles.ts).
    if (typeof rPr.underline === "string") {
      look.underline = rPr.underline === "none" ? null : rPr.underline === "wave" ? "wavy" : rPr.underline;
    }
    if (rPr.vertAlign === "superscript" || rPr.vertAlign === "subscript") look.vertAlign = rPr.vertAlign;
    else if (rPr.vertAlign === "baseline") look.vertAlign = null;
  }
  if (themedFace) look.family = undefined;
  if (look.sizePt === null) {
    // The document's default size, else Word's own 10 pt: the page states a
    // size on every imported run (corpus R391-D05: a style stating none).
    const half = record(record(definitions?.docDefaults)?.rPr)?.fontSizeHalfPoints;
    look.sizePt = typeof half === "number" ? half / 2 : 10;
  }
  return look;
}

/** The default paragraph style's id, for a paragraph naming none. */
function defaultParagraphStyleId(docAttrs: Record<string, unknown>): string | null {
  const rows = record(docAttrs.styleDefinitions)?.styles;
  if (!Array.isArray(rows)) return null;
  const found = rows.map(record).find((style) => style?.type === "paragraph" && style.default === true);
  return typeof found?.styleId === "string" ? found.styleId : null;
}

function paragraphStyleId(node: ProseMirrorNode, docAttrs: Record<string, unknown>): string | null {
  const own = (node.attrs as { docxStyleId?: unknown }).docxStyleId;
  return typeof own === "string" && own ? own : defaultParagraphStyleId(docAttrs);
}

/** Trade the old style's look for the new one's on the text in [from, to). */
export function restyleJoinedRuns(
  tr: Transaction,
  schema: Schema,
  from: number,
  to: number,
  before: RunLook,
  after: RunLook,
): void {
  const edits: Array<() => void> = [];
  tr.doc.nodesBetween(from, to, (child, position) => {
    if (!child.isText) return true;
    const start = Math.max(from, position);
    const end = Math.min(to, position + child.nodeSize);
    if (start >= end) return false;
    const mark = (name: string) => child.marks.find((item) => item.type.name === name);
    const trade = (
      markName: string,
      old: string | number | null | undefined,
      next: string | number | null | undefined,
      read: (attrs: Record<string, unknown>) => string | number | null,
      make: (value: string | number) => Record<string, unknown>,
    ) => {
      const type = schema.marks[markName];
      // Unknown on either side (a theme face) or the same: nothing to trade.
      if (!type || old === undefined || next === undefined || old === next) return;
      const current = mark(markName);
      const value = current ? read(current.attrs) : null;
      // A run stating nothing shows its paragraph style's look.
      const fromStyle = current ? value === old : true;
      if (!fromStyle) return;
      edits.push(() => {
        tr.removeMark(start, end, type);
        if (next !== null) tr.addMark(start, end, type.create(make(next)));
      });
    };
    const face = (family: string | null | undefined) => (typeof family === "string" ? bareFamily(family) : family);
    trade("fontFamily", face(before.family), face(after.family),
      (attrs) => bareFamily(attrs.family) || null,
      () => ({ family: `"${after.family}"` }));
    trade("fontSize", before.sizePt, after.sizePt,
      (attrs) => {
        const points = pointsOf(attrs.size);
        return points === null ? null : Math.round(points * 2) / 2;
      },
      (points) => ({ size: px(points as number) }));
    trade("textColor", before.color, after.color,
      (attrs) => (typeof attrs.color === "string" ? attrs.color.toLowerCase() : null),
      (color) => ({ color }));
    for (const name of ["bold", "italic"] as const) {
      const type = schema.marks[name];
      const old = before[name] === true;
      const next = after[name] === true;
      if (!type || old === next) continue;
      const current = mark(name);
      // A direct choice (value true / false) is the run's own.
      if (current && typeof current.attrs.value === "boolean") continue;
      if (old && current) edits.push(() => tr.removeMark(start, end, type));
      if (next && !current) edits.push(() => tr.addMark(start, end, type.create()));
    }
    // The style's underline, as the import states it: a bare mark for its
    // single line, else the line's name (corpus R455-D04 E6: a Subtitle line
    // joined onto a Title underlined by its style draws the Title's
    // underline in Word, and the file it saves says so).
    const underlineType = schema.marks.underline;
    const oldLine = before.underline ?? null;
    const nextLine = after.underline ?? null;
    if (underlineType && oldLine !== nextLine) {
      const current = mark("underline");
      const value = current ? String(current.attrs.underlineStyle ?? "single") : null;
      // No mark under a style that underlines is the run's own "no line".
      if (current ? value === oldLine : oldLine === null) {
        edits.push(() => {
          tr.removeMark(start, end, underlineType);
          if (nextLine !== null) {
            tr.addMark(start, end, underlineType.create(nextLine === "single" ? {} : { underlineStyle: nextLine }));
          }
        });
      }
    }
    // The style's Superscript / Subscript: Word applies a paragraph style's
    // w:vertAlign to the words in it that state none, so words joined onto a
    // paragraph whose style raises its text are raised (corpus R470-D53 E6:
    // an affiliation line backspaced onto an "Authors" line whose style is
    // superscript; Word, and the file the save writes, raise them). A run
    // whose raise differs from its old style's is its own and stays.
    const oldRaise = before.vertAlign ?? null;
    const nextRaise = after.vertAlign ?? null;
    if (oldRaise !== nextRaise && schema.marks.superscript && schema.marks.subscript) {
      const current = mark("superscript") ? "superscript" : mark("subscript") ? "subscript" : null;
      if (current === oldRaise) {
        edits.push(() => {
          if (oldRaise) tr.removeMark(start, end, schema.marks[oldRaise]);
          if (nextRaise) tr.addMark(start, end, schema.marks[nextRaise].create());
        });
      }
    }
    return false;
  });
  for (const edit of edits) edit();
}

type JoinedTail = { from: number; to: number; before: string; after: string };

/** The joins the transactions made: the second paragraph's words, where they
 *  now stand, and the two styles. A step joins when it deletes from inside
 *  one text block into another and puts no inline content of its own there. */
function joinedTails(transactions: readonly Transaction[], newState: EditorState): JoinedTail[] {
  const docAttrs = newState.doc.attrs as Record<string, unknown>;
  const maps = transactions.flatMap((tr) => tr.mapping.maps);
  const found: Array<JoinedTail & { step: number }> = [];
  let index = 0;
  for (const tr of transactions) {
    tr.steps.forEach((step, own) => {
      const at = index + own;
      if (!(step instanceof ReplaceStep)) return;
      const { from, to, slice } = step as unknown as { from: number; to: number; slice: Slice };
      if (from >= to) return;
      // The document before this step (Transform.docs), not re-applied.
      const doc = tr.docs[own];
      if (!doc || to > doc.content.size) return;
      let inline = false;
      slice.content.descendants((node) => {
        if (node.isInline) inline = true;
        return !inline;
      });
      if (inline) return;
      const $from = doc.resolve(from);
      const $to = doc.resolve(to);
      const first = $from.parent;
      const second = $to.parent;
      if (first === second || !first.isTextblock || !second.isTextblock || $to.parentOffset >= second.content.size) return;
      const before = paragraphStyleId(second, docAttrs);
      const after = paragraphStyleId(first, docAttrs);
      if (before && after && before !== after) found.push({ from: to, to: $to.end(), before, after, step: at });
    });
    index += tr.steps.length;
  }
  return found.map(({ from, to, before, after, step }) => {
    for (const map of maps.slice(step)) {
      from = map.map(from, 1);
      to = map.map(to, -1);
    }
    return { from, to, before, after };
  });
}

const joinedParagraphStyleRunsKey = new PluginKey("joinedParagraphStyleRuns");

export const JoinedParagraphStyleRuns = Extension.create({
  name: "joinedParagraphStyleRuns",
  addProseMirrorPlugins() {
    const schema = this.editor.schema;
    return [
      new Plugin({
        key: joinedParagraphStyleRunsKey,
        appendTransaction(transactions, _oldState, newState) {
          // Another client's change (the room) and Undo/Redo already hold
          // their runs as they were written.
          if (
            !transactions.some((tr) => tr.docChanged) ||
            transactions.some(
              (tr) =>
                tr.getMeta("y-sync$") !== undefined ||
                tr.getMeta("history$") !== undefined ||
                tr.getMeta(joinedParagraphStyleRunsKey) !== undefined,
            )
          ) {
            return null;
          }
          const tails = joinedTails(transactions, newState);
          if (!tails.length) return null;
          const docAttrs = newState.doc.attrs as Record<string, unknown>;
          const tr = newState.tr;
          for (const tail of tails) {
            if (tail.from >= tail.to) continue;
            const before = paragraphStyleRunLook(docAttrs, tail.before);
            const after = paragraphStyleRunLook(docAttrs, tail.after);
            if (!before || !after) continue;
            restyleJoinedRuns(tr, schema, tail.from, tail.to, before, after);
          }
          if (!tr.docChanged) return null;
          tr.setMeta(joinedParagraphStyleRunsKey, true);
          return tr;
        },
      }),
    ];
  },
});

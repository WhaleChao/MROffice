/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet's side of the shared SmartArt Design tab.
 *
 * The tab's LOGIC is host-neutral and lives in `lib/office-smartart/design-ops`;
 * what is worksheet-specific is the PROJECTION. A sheet diagram is its members
 * sharing one `groupId` (docs/office-graphics/SMARTART-IMPORT.md §3b), and
 * `lib/spreadsheet/smart-art.ts` projects an engine scene onto those members:
 * EMU become CSS pixels relative to the diagram's frame, a shape's paint moves
 * into `object.shape`, and an item's LABEL is merged into the shape of the same
 * outline item instead of staying a separate text node.
 *
 * That projection is one-way, so every Design-tab command needs its inverse:
 * read the outline and the colours back off sheet objects, hand them to
 * `design-ops`, and project the rebuilt scene again. This module is exactly
 * that inverse — no geometry, no style and no outline rule is re-derived here,
 * and nothing in it knows about React, anchors or `replaceActiveSheet` (the
 * host owns the transaction). The rebuild itself lays the diagram out and runs
 * on the server (`./smartart-rebuild`, through the engine host).
 */

import { pxToPt } from "../office-units";
import {
  DEFAULT_SMARTART_STYLE,
  isSmartArtHierBranch,
  MAX_SMARTART_OUTLINE_LEVEL,
  readSmartArtBlockMarker,
  type OfficeSmartArtHostNode,
  type SmartArtHierBranch,
  type SmartArtDirection,
  type SmartArtKind,
  type SmartArtStyle,
  type SmartArtNodePaint,
} from "../office-smartart/model";
import {
  smartArtConnectorColor,
  type OfficeSmartArtFrame,
  smartArtPaletteNode,
  smartArtVisibleFill,
} from "../office-smartart/design-model";
import {
  MAX_SMARTART_OUTLINE_ITEMS,
  MAX_SMARTART_OUTLINE_TEXT_LENGTH,
  type SmartArtOutlineItem,
} from "../office-smartart/outline";
import {
  DEFAULT_SMARTART_STYLE_ID,
  isSmartArtStyleId,
  type SmartArtStyleId,
} from "../office-smartart/styles-model";

import { normalizeOfficeShapeTextFont } from "../office-shape-text-font";

import {
  SPREADSHEET_SMARTART_FONT_PT,
  type SpreadsheetSmartArtMember,
} from "./smart-art";
import { normalizeShapeStyle } from "./shapes";

/** EMU per CSS pixel (96 dpi) — the sheet object layer's own scale. */
const EMU_PER_CSS_PIXEL = 9525;

/** Points → the px the sheet shape renderer speaks, and back. */
const PT_PER_PX = pxToPt(1);

function shapeStyleOf(node: unknown): Record<string, unknown> | undefined {
  const shape = (node as Record<string, unknown> | null)?.shape;
  return shape && typeof shape === "object" && !Array.isArray(shape)
    ? (shape as Record<string, unknown>)
    : undefined;
}

function stringField(
  record: Record<string, unknown> | undefined,
  key: string,
): string | undefined {
  const value = record?.[key];
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

/**
 * The diagram's outline, read back off its sheet members.
 *
 * The inverse of `spreadsheetSmartArtMembers`' label merge: an item's text is
 * its own `label` member when the projection had to keep one (a label with no
 * shape of its own, e.g. a picture layout's caption), and otherwise the text
 * merged into the `node` shape of that item. An item whose text was empty
 * carries none at all, which reads back as the empty string it was.
 *
 * Refuses (null) on anything it cannot prove is one whole outline — the same
 * bar `readSmartArtOutline` holds the slide model to: contiguous indices from
 * 0, unique item ids, a first item at level 0 and no level jumping by more
 * than one.
 */
export function readSheetSmartArtOutline(
  members: readonly OfficeSmartArtHostNode[],
): SmartArtOutlineItem[] | null {
  const byIndex = new Map<
    number,
    { id: string; level: number; text: string; fromLabel: boolean; assistant: boolean; hierBranch?: SmartArtHierBranch }
  >();
  // Each org chart box's Layout, on its node (Office's hierBranch).
  const branches = new Map<number, SmartArtHierBranch>();
  // Titles (Name and Title): the words of the members that name an item.
  const titles = new Map<number, string>();
  for (const member of members) {
    const marker = readSmartArtBlockMarker(member);
    const titled = (member as Record<string, unknown>).smartArtTitleIndex;
    if (marker?.role === "connector" && typeof titled === "number") {
      const words = stringField(shapeStyleOf(member), "text") ?? "";
      if (words.length > 0 && words.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH) titles.set(titled, words);
      continue;
    }
    if (
      !marker ||
      (marker.role !== "node" && marker.role !== "label") ||
      marker.itemId === null ||
      marker.itemIndex === null ||
      marker.itemIndex < 0 ||
      marker.level === null ||
      marker.level < 0 ||
      marker.level > MAX_SMARTART_OUTLINE_LEVEL
    ) {
      continue;
    }
    const text = stringField(shapeStyleOf(member), "text") ?? "";
    if (text.length > MAX_SMARTART_OUTLINE_TEXT_LENGTH) return null;
    const branch = (member as Record<string, unknown>).smartArtHierBranch;
    if (marker.role === "node" && isSmartArtHierBranch(branch)) branches.set(marker.itemIndex, branch);
    const existing = byIndex.get(marker.itemIndex);
    const fromLabel = marker.role === "label";
    // A label member is the authority for the text where the projection kept
    // one; the node shape answers for every ordinary layout.
    if (existing && (!fromLabel || existing.fromLabel)) {
      if (existing.id !== marker.itemId || existing.level !== marker.level) {
        return null;
      }
      continue;
    }
    byIndex.set(marker.itemIndex, {
      id: marker.itemId,
      level: marker.level,
      text,
      fromLabel,
      assistant:
        (member as Record<string, unknown>).smartArtAssistant === true ||
        existing?.assistant === true,
    });
  }
  const rows = [...byIndex.entries()].sort(([left], [right]) => left - right);
  if (rows.length === 0 || rows.length > MAX_SMARTART_OUTLINE_ITEMS) return null;
  const ids = new Set<string>();
  for (let index = 0; index < rows.length; index += 1) {
    const [position, row] = rows[index];
    if (
      position !== index ||
      ids.has(row.id) ||
      (index === 0 && row.level !== 0) ||
      (index > 0 && row.level > rows[index - 1][1].level + 1)
    ) {
      return null;
    }
    ids.add(row.id);
  }
  return rows.map(([position, row]) => ({
    id: row.id,
    text: row.text,
    level: row.level,
    ...(row.assistant && row.level > 0 ? { assistant: true as const } : {}),
    ...(titles.has(position) ? { title: titles.get(position)! } : {}),
    ...(branches.has(position) ? { hierBranch: branches.get(position)! } : {}),
  }));
}

/**
 * The colour set the diagram currently wears, read off its OWN members — the
 * sheet's `readSmartArtGroupColors`, reading `object.shape` instead of the
 * top-level paint a slide block carries. A rebuild keeps the author's colours
 * whatever quick style is applied.
 */
export function readSheetSmartArtColors(
  members: readonly OfficeSmartArtHostNode[],
): SmartArtStyle {
  const byRole = (role: "node" | "connector") =>
    shapeStyleOf(
      members.find((member) => readSmartArtBlockMarker(member)?.role === role),
    );
  const node = shapeStyleOf(
    smartArtPaletteNode(
      members.filter(
        (member) =>
          readSmartArtBlockMarker(member)?.role === "node" &&
          (member as unknown as Record<string, unknown>)
            .smartArtPicturePlaceholder !== true,
      ),
      (member) =>
        smartArtVisibleFill(shapeStyleOf(member)?.fill, shapeStyleOf(member)?.opacity),
      members.filter(
        (member) =>
          readSmartArtBlockMarker(member)?.role !== "label" &&
          (member as unknown as Record<string, unknown>)
            .smartArtPicturePlaceholder !== true,
      ),
    ),
  );
  const connector = byRole("connector");
  const labelledMembers = members.filter((member) => {
    const role = readSmartArtBlockMarker(member)?.role;
    return (
      (role === "label" || role === "node") &&
      stringField(shapeStyleOf(member), "textColor") !== undefined
    );
  });
  // One shape the author re-sized or re-coloured by hand is that shape's own
  // formatting, not the diagram's: the diagram's type is read off a member
  // that still wears the engine's, and only when every member was formatted
  // (the whole graphic) is the author's value the diagram's.
  const engineMember = labelledMembers.find(
    (member) =>
      (member as unknown as Record<string, unknown>).smartArtAuthorText !== true &&
      (member as unknown as Record<string, unknown>).smartArtAuthorTextSize !== true,
  );
  const labelled = shapeStyleOf(engineMember ?? labelledMembers[0]);
  const textSize = labelled?.textSize;
  // An engine size is the FITTED one: every label shrinks together when one
  // is crowded (`spreadsheetSmartArtMembers`). Read back as the diagram's
  // size, it would never grow again once that label is short; the rebuild
  // starts from the sheet's own size at least, and the fit takes it down.
  const engineFitted = engineMember !== undefined;
  return {
    nodeFill:
      smartArtVisibleFill(stringField(node, "fill"), node?.opacity) ??
      DEFAULT_SMARTART_STYLE.nodeFill,
    nodeStroke: stringField(node, "stroke") ?? DEFAULT_SMARTART_STYLE.nodeStroke,
    connector: smartArtConnectorColor(
      stringField(connector, "fill"),
      stringField(connector, "stroke"),
      DEFAULT_SMARTART_STYLE.connector,
    ),
    textColor:
      stringField(labelled, "textColor") ?? DEFAULT_SMARTART_STYLE.textColor,
    fontSize:
      typeof textSize === "number" && textSize > 0
        ? engineFitted
          ? Math.max(
              SPREADSHEET_SMARTART_FONT_PT,
              Math.round(textSize * PT_PER_PX),
            )
          : Math.round(textSize * PT_PER_PX)
        : DEFAULT_SMARTART_STYLE.fontSize,
  };
}

/**
 * The quick style the diagram's members agree on, so the gallery can check it.
 *
 * The sheet keeps the style marker with the rest of the style's PAINT, inside
 * `object.shape` (`lib/spreadsheet/smart-art.ts`), not beside `smartArtKind` the
 * way a slide block does — so this is where the worksheet looks. Members that
 * disagree report none, rather than letting an arbitrary member decide.
 */
export function readSheetSmartArtStyleId(
  members: readonly OfficeSmartArtHostNode[],
): SmartArtStyleId | null {
  const ids = new Set(
    members.map((member) => {
      const value = shapeStyleOf(member)?.smartArtStyle;
      return isSmartArtStyleId(value) ? value : null;
    }),
  );
  const [only] = [...ids];
  return ids.size === 1 ? only : null;
}

export interface SheetSmartArtRebuildInput {
  /** The workbook's / document's body face (`+mn-lt`), which sets how tall the labels' lines stand. */
  bodyFace?: string | null;
  /** The face the workbook theme gives Chinese words (`SmartArtColorTheme.eastAsianFace`). */
  eastAsianFace?: string | null;
  /** Change Layout's pick: an aligned Office alias, or `null` for the catalog layout. */
  layoutUrn?: string | null;
  /** The diagram's current members; their paint seeds the rebuilt nodes. */
  members: readonly OfficeSmartArtHostNode[];
  /** The diagram id every rebuilt member keeps. */
  groupId: string;
  /** The layout to build — the current one for an outline or style edit. */
  kind: SmartArtKind;
  /** The diagram's box on the sheet, in CSS pixels. */
  frame: OfficeSmartArtFrame;
  /** The outline the rebuilt diagram carries, in order. */
  outline: readonly SmartArtOutlineItem[];
  /** Colours; defaults to the members' own ({@link readSheetSmartArtColors}). */
  colors?: SmartArtStyle;
  /** Omitted rebuilds at the engine default, exactly as a layout swap does. */
  styleId?: SmartArtStyleId;
  /**
   * Which way the rebuilt diagram runs; omitted keeps the members' own
   * (`smartArtDirection` on each object), so only Right to Left changes it.
   */
  direction?: SmartArtDirection;
  /** Each item's own fill and line (a Change Colors variation). */
  nodePaint?: readonly (SmartArtNodePaint | null)[];
  /** Whether the rebuild may change the SET of items (add / delete). */
  allowItemChanges?: boolean;
  /**
   * Whether `colors` are the AUTHOR'S CHOICE — true for Change Colors, false
   * for Reset Graphic, omitted when the command is not a recolour. See
   * `withSmartArtAuthorPaint`; without it a UNIFORM recolour reaches Office
   * without its colours until the file has been exported and reopened once.
   */
  authorPaint?: boolean;
  /**
   * Change Colors' variation (Office's id) the rebuilt graphic wears; `null`
   * is Office's default, omitted keeps the members' own (`smartArtColors`).
   */
  colorsId?: string | null;
  /** The workbook theme a variation is resolved against (its accents). */
  colorTheme?: { accents?: readonly string[] } | null;
  newId?: () => string;
}

/** One rebuilt diagram: members in the projection's own frame-relative pixels. */
export interface SheetSmartArtRebuild {
  members: SpreadsheetSmartArtMember[];
  /** The quick style every rebuilt member now carries. */
  styleId: SmartArtStyleId;
}

/**
 * Carry the author's TEXT formatting across a rebuild, the text twin of
 * `preserveSmartArtAuthorPaint`.
 *
 * Every structural command (Add / Delete / Promote / Demote, Change Layout,
 * a quick style, Change Colors) rebuilds the members from the engine, which
 * knows nothing of a font the author picked on Home, so without this every
 * such command silently put the label back in the default face. Office keeps
 * text formatting through all of those ("most of your customizations will be
 * kept"); a SmartArt Style only changes fill, line and effects.
 *
 * Keyed on the durable item id, like the pictures and the paint claim: a
 * relayout moves positions and indices, the item id survives, and an item
 * that is gone takes its formatting with it. The face, bold, italic,
 * underline and strikethrough are always the author's (the engine never sets
 * them); the size only where the member says so (`smartArtAuthorTextSize`)
 * and the colour only where it says so (`smartArtAuthorText`), so an
 * engine-fitted size keeps being re-fitted and an authored one does not.
 */
export function preserveSheetSmartArtAuthorText(
  previous: readonly unknown[],
  rebuilt: readonly SpreadsheetSmartArtMember[],
): SpreadsheetSmartArtMember[] {
  const byItem = new Map<string, Record<string, unknown>>();
  for (const member of previous) {
    const record = member as Record<string, unknown> | null;
    if (!record || typeof record !== "object") continue;
    const role = record.smartArtRole;
    const itemId = record.smartArtItemId;
    if ((role !== "node" && role !== "label") || typeof itemId !== "string") {
      continue;
    }
    const shape = normalizeShapeStyle(record.shape);
    const authoredColour = record.smartArtAuthorText === true;
    const authoredSize = record.smartArtAuthorTextSize === true;
    // A line spacing every paragraph of the words states (a converted
    // point's `a:lnSpc`, or Format Shape's on the sheet) is the author's.
    const spacings = new Set(shape.textLayout?.paragraphs.map((paragraph) => paragraph.lineSpacing));
    const spacing = spacings.size === 1 ? [...spacings][0] : undefined;
    const carried: Record<string, unknown> = {
      ...normalizeOfficeShapeTextFont(shape),
      ...(typeof spacing === "number" ? { lineSpacing: spacing } : {}),
      ...(authoredSize && shape.textSize !== undefined
        ? { textSize: shape.textSize }
        : {}),
      ...(authoredColour && shape.textColor !== undefined
        ? { textColor: shape.textColor }
        : {}),
    };
    if (Object.keys(carried).length === 0) continue;
    // The member that carries the item's words speaks for it.
    if (byItem.has(itemId) && !shape.text) continue;
    byItem.set(itemId, {
      ...carried,
      ...(authoredColour ? { authoredColour: true } : {}),
      ...(authoredSize ? { authoredSize: true } : {}),
    });
  }
  if (byItem.size === 0) return [...rebuilt];
  return rebuilt.map((member) => {
    const carried =
      (member.smartArtRole === "node" || member.smartArtRole === "label") &&
      typeof member.smartArtItemId === "string"
        ? byItem.get(member.smartArtItemId)
        : undefined;
    if (!carried) return member;
    const { authoredColour, authoredSize, lineSpacing, ...font } = carried;
    const lines = typeof member.shape.text === "string" ? Math.min(member.shape.text.split("\n").length, 200) : 0;
    return {
      ...member,
      shape: normalizeShapeStyle({
        ...member.shape,
        ...font,
        ...(typeof lineSpacing === "number" && lines > 0
          ? { textLayout: { paragraphs: Array.from({ length: lines }, () => ({ level: 0, bullet: "none", lineSpacing })) } }
          : {}),
      }),
      ...(authoredColour === true ? { smartArtAuthorText: true as const } : {}),
      ...(authoredSize === true ? { smartArtAuthorTextSize: true as const } : {}),
    };
  });
}

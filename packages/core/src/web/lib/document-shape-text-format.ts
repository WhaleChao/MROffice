/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx, pxToPtRounded } from "./office-units";
import {
  DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_FAMILY,
  DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_SIZE_PT,
  documentShapeDefaultAttrs,
  documentShapeFormatPatch,
  normalizeDocumentShapeAttrs,
  type DocumentShapeAttrs,
  type DocumentShapeAttrsInput,
  type DocumentShapeFormatRequest,
  type DocumentShapeGroupMember,
} from "./document-shapes";
import {
  applySmartArtTextFormat,
  summarizeSmartArtTextFormat,
  type SmartArtNodeTextFormat,
  type SmartArtNodeTextFormatPatch,
  type SmartArtTextFormatAdapter,
  type SmartArtTextFormatCommand,
  type SmartArtTextFormatSummary,
  stepOfficeFontSizePt,
} from "./office-smartart/text-format";
import type { TextRun } from "./presentation-runs";
import { officeTextPaintWithFontColor, type OfficeTextPaint } from "./office-text-paint";
import { documentSmartArtLabelFontPx } from "./office-smartart/document-shapes-model";
import {
  documentShapeRunFontSizePt,
  documentShapeTextRunsWithout,
  type DocumentShapeRunStyleKey,
  normalizeDocumentShapeTextRuns,
} from "./document-shape-text-runs";

/** One shape whose words Home formats: the shape itself, or a group member. */
export interface DocumentShapeTextLeaf {
  readonly memberId?: string;
  readonly leaf: DocumentShapeAttrs;
}

const MEMBER_IDENTITY_KEYS = new Set([
  "memberId",
  "drawingId",
  "kind",
  "connectorPreset",
  "startMemberId",
  "startConnectionSite",
  "endMemberId",
  "endConnectionSite",
]);

function memberLeaf(member: DocumentShapeGroupMember): DocumentShapeAttrs {
  const leafFields = Object.fromEntries(
    Object.entries(member).filter(([key]) => !MEMBER_IDENTITY_KEYS.has(key)),
  );
  return normalizeDocumentShapeAttrs({
    ...documentShapeDefaultAttrs(member.preset),
    ...leafFields,
  });
}

/**
 * The leaves Home formats for `value`: every member of a group or SmartArt
 * graphic that can hold words (a picture member carries none), the one
 * member named by `memberId`, or the shape itself. Empty for an ink stroke.
 */
export function documentShapeTextLeaves(
  value: DocumentShapeAttrsInput,
  memberId?: string | null,
  themeBodyFace?: string,
): DocumentShapeTextLeaf[] {
  const attrs = normalizeDocumentShapeAttrs(value);
  if (attrs.inkTool) return [];
  const members = attrs.groupMembers;
  if (members?.length) {
    // A SmartArt graphic draws the words of every shape whose size nobody
    // set at ONE fitted size (`documentSmartArtLabelFontPx`); Home reads and
    // steps that drawn size, as Word's size box does, not the stored default.
    const fittedPx = documentSmartArtLabelFontPx(attrs);
    const fittedPt = fittedPx === null ? null : pxToPtRounded(fittedPx, 1);
    return members
      .filter(
        (member) =>
          member.kind !== "picture" && (!memberId || member.memberId === memberId),
      )
      .map((member) => {
        const leaf = memberLeaf(member);
        const fitted =
          fittedPt !== null &&
          member.text.trim().length > 0 &&
          !member.textBody &&
          member.smartArtAuthorTextSize !== true &&
          !(
            typeof member.textFontSizePt === "number" &&
            member.textFontSizePt !== DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_SIZE_PT
          );
        // …and the face the graphic draws words with none of their own in:
        // the document theme's body font, as Word's Font box shows it.
        const themed =
          themeBodyFace !== undefined &&
          typeof attrs.smartArtKind === "string" &&
          !member.textBody &&
          member.smartArtAuthorTextFont !== true &&
          (member.textFontFamily === undefined ||
            member.textFontFamily === DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_FAMILY);
        return {
          memberId: member.memberId,
          leaf: {
            ...leaf,
            ...(fitted ? { textFontSizePt: fittedPt } : {}),
            ...(themed ? { textFontFamily: themeBodyFace } : {}),
          },
        };
      });
  }
  return [{ leaf: attrs }];
}

const TEXT_KEYS = [
  "textFontFamily",
  "textFontSizePt",
  "textBold",
  "textItalic",
  "textUnderline",
  "textStrike",
  "textFill",
  "textHorizontalAlign",
  "textBody",
] as const;

const DOCUMENT_SHAPE_TEXT_ADAPTER: SmartArtTextFormatAdapter<DocumentShapeTextLeaf> = {
  read({ leaf }): SmartArtNodeTextFormat {
    return {
      fontFamily: leaf.textFontFamily,
      fontSizePt: leaf.textFontSizePt,
      bold: leaf.textBold,
      italic: leaf.textItalic,
      ...(leaf.textUnderline === "none" ? {} : { underline: leaf.textUnderline }),
      strike: leaf.textStrike,
      ...(leaf.textFill === "none" ? {} : { color: leaf.textFill }),
      align: leaf.textHorizontalAlign,
    };
  },
  write(node, patch: SmartArtNodeTextFormatPatch) {
    const { leaf } = node;
    const request: Record<string, unknown> = {};
    if (Object.hasOwn(patch, "fontFamily")) {
      // The shape stores a font NAME (Word's `w:rFonts`); "none" is the default face.
      request.textFontFamily =
        patch.fontFamily ??
        normalizeDocumentShapeAttrs({ ...leaf, textFontFamily: undefined }).textFontFamily;
    }
    if (Object.hasOwn(patch, "fontSizePt") && patch.fontSizePt !== undefined) {
      request.textFontSizePt = patch.fontSizePt;
    }
    if (Object.hasOwn(patch, "bold")) request.textBold = patch.bold === true;
    if (Object.hasOwn(patch, "italic")) request.textItalic = patch.italic === true;
    if (Object.hasOwn(patch, "strike")) request.textStrike = patch.strike === true;
    if (Object.hasOwn(patch, "underline")) request.textUnderline = patch.underline ?? "none";
    if (Object.hasOwn(patch, "color")) {
      // Automatic: the colour the shape inherits (its default ink).
      request.textFill =
        patch.color ??
        normalizeDocumentShapeAttrs({ ...leaf, textFill: undefined }).textFill;
    }
    if (Object.hasOwn(patch, "align") && patch.align) {
      request.textHorizontalAlign = patch.align;
    }
    const applied = documentShapeFormatPatch(leaf, request as DocumentShapeFormatRequest);
    if (!applied) return node;
    return { ...node, leaf: normalizeDocumentShapeAttrs({ ...leaf, ...applied }) };
  },
};

/** What Home shows for the selected shape's (or graphic's) words. */
export function summarizeDocumentShapeText(
  value: DocumentShapeAttrsInput,
  memberId?: string | null,
  themeBodyFace?: string,
): SmartArtTextFormatSummary | null {
  const leaves = documentShapeTextLeaves(value, memberId, themeBodyFace);
  const summary = summarizeSmartArtTextFormat(runLeaves(leaves), DOCUMENT_SHAPE_TEXT_ADAPTER);
  if (!summary) return summary;
  // An imported body's paragraphs carry their own `w:jc`, drawn as such (a
  // paragraph that states none draws at the start); the shape's one scalar
  // is not what the words show. Home reads the paragraphs, "mixed" where
  // they differ, so Center on a left+right body centres both rather than
  
  const paragraphAligns = leaves.flatMap((node) =>
    (node.leaf.textBody?.paragraphs ?? []).map((paragraph) => paragraph.align ?? "left"),
  );
  if (paragraphAligns.length === 0) return summary;
  const first = paragraphAligns[0]!;
  const mixed = summary.align.mixed || paragraphAligns.some((align) => align !== first);
  return { ...summary, align: { value: mixed ? undefined : first, mixed } };
}

/**
 * A shape whose words carry their own formatting reads as each run would
 * draw (the run's field, else the shape's): Home shows "mixed" where they
 * differ, as Word does with the shape's border selected.
 */
function runLeaves(leaves: DocumentShapeTextLeaf[]): DocumentShapeTextLeaf[] {
  return leaves.flatMap((node) => {
    const runs = node.leaf.textRuns;
    if (!runs?.length) return [node];
    return runs.map((run) => {
      const style = run.style ?? {};
      const size = documentShapeRunFontSizePt(style);
      return {
        ...(node.memberId ? { memberId: node.memberId } : {}),
        leaf: {
          ...node.leaf,
          ...(style.bold !== undefined ? { textBold: style.bold } : {}),
          ...(style.italic !== undefined ? { textItalic: style.italic } : {}),
          ...(style.strike !== undefined ? { textStrike: style.strike } : {}),
          ...(style.underline !== undefined
            ? { textUnderline: style.underline ? (node.leaf.textUnderline === "double" ? "double" : "single") : "none" }
            : {}),
          ...(style.color ? { textFill: style.color } : {}),
          ...(size !== undefined ? { textFontSizePt: size } : {}),
          ...(style.fontFamily ? { textFontFamily: style.fontFamily } : {}),
        },
      };
    });
  });
}

/** The format fields a whole-shape command decides (toggle: its own field). */
const FORMAT_FIELDS_OF: Record<SmartArtTextFormatCommand["kind"], readonly string[]> = {
  family: ["fontFamily"],
  variant: ["fontFamily", "bold", "italic"],
  size: ["fontSizePt"],
  step: ["fontSizePt"],
  toggle: [],
  underline: ["underline"],
  color: ["color"],
  align: ["align"],
};

/** The run fields a whole-shape command sets for every word. */
function runKeysOf(command: SmartArtTextFormatCommand): DocumentShapeRunStyleKey[] {
  switch (command.kind) {
    case "toggle":
      return [command.field];
    case "underline":
      return ["underline", "underlineStyle"];
    case "size":
    case "step":
      return ["fontSize"];
    case "family":
      return ["fontFamily"];
    case "variant":
      return ["fontFamily", "bold", "italic"];
    case "color":
      return ["color"];
    default:
      return [];
  }
}

/**
 * A group member's runs after a whole-member command: the command's fields
 * leave the runs (the member now states them for every word), except that
 * Grow/Shrink Font steps every run's own size, as it does a shape's.
 */
function memberRunsAfterCommand(
  member: DocumentShapeGroupMember,
  command: SmartArtTextFormatCommand,
): readonly TextRun[] | undefined {
  const runs = member.textRuns;
  if (!runs?.length) return runs;
  if (command.kind === "step") {
    let changed = false;
    const stepped = runs.map((run) => {
      const px = run.style?.fontSize;
      if (typeof px !== "number") return run;
      const next = stepOfficeFontSizePt(Math.round(((px * 3) / 4) * 10) / 10, command.direction);
      if (next === null) return run;
      changed = true;
      return { text: run.text, style: { ...run.style, fontSize: ptToPx(next) } };
    });
    return changed ? normalizeDocumentShapeTextRuns(stepped, member.text) : runs;
  }
  const keys = runKeysOf(command);
  if (!keys.length) return runs;
  return documentShapeTextRunsWithout(runs, member.text, keys);
}

/**
 * One Home command over the shape's words (or every member's): the attrs
 * patch to commit, or `null` when nothing changes. Font Color on WordArt
 * words is their fill: the shape's (a member's) paint and a run's own paint
 * take the colour, outline and effects kept (`officeTextPaintWithFontColor`);
 * it wrote `textFill` beside the paint, which went on drawing the words.
 */
export function documentShapeTextCommandPatch(
  value: DocumentShapeAttrsInput,
  command: SmartArtTextFormatCommand,
  memberId?: string | null,
): Readonly<Partial<DocumentShapeAttrs>> | null {
  const patch = documentShapeTextCommandPatchOfFields(value, command, memberId);
  if (command.kind !== "color") return patch;
  const attrs = normalizeDocumentShapeAttrs(value);
  const color = command.color ?? undefined;
  const recolor = (paint: OfficeTextPaint, automaticFill?: string) =>
    officeTextPaintWithFontColor(paint, color, { automaticFill });
  const runsRecolored = (runs: readonly TextRun[] | null | undefined): TextRun[] | null => {
    if (!runs?.some((run) => run.style?.textPaint)) return null;
    return runs.map((run) => {
      if (!run.style?.textPaint) return run;
      const { textPaint, ...rest } = run.style;
      const paint = recolor(textPaint);
      const style = paint ? { ...rest, textPaint: paint } : rest;
      return Object.keys(style).length ? { ...run, style } : { text: run.text };
    });
  };
  const written = (patch ?? {}) as Record<string, unknown>;
  const next: Record<string, unknown> = { ...written };
  if (!attrs.groupMembers?.length) {
    // Automatic: the colour the plain patch gave the words, under an outline.
    const plainFill = typeof written.textFill === "string" && written.textFill !== "none" ? written.textFill : undefined;
    if (attrs.textPaint) next.textPaint = recolor(attrs.textPaint, plainFill) ?? null;
    const runs = runsRecolored(
      Object.hasOwn(written, "textRuns") ? (written.textRuns as TextRun[] | null) : attrs.textRuns,
    );
    if (runs) next.textRuns = normalizeDocumentShapeAttrs({ ...attrs, textRuns: runs }).textRuns ?? null;
  } else {
    const members = (written.groupMembers as readonly DocumentShapeGroupMember[] | undefined) ?? attrs.groupMembers;
    let changed = false;
    const groupMembers = members.map((member) => {
      if (memberId && member.memberId !== memberId) return member;
      const updated: Record<string, unknown> = { ...member };
      if (member.textPaint) {
        const paint = recolor(member.textPaint, member.textFill && member.textFill !== "none" ? member.textFill : undefined);
        if (paint) updated.textPaint = paint;
        else delete updated.textPaint;
      }
      const runs = runsRecolored(member.textRuns);
      if (runs) updated.textRuns = runs;
      if (JSON.stringify(updated) === JSON.stringify(member)) return member;
      changed = true;
      return updated as unknown as DocumentShapeGroupMember;
    });
    if (changed) {
      const normalized = normalizeDocumentShapeAttrs({ ...attrs, groupMembers }).groupMembers;
      if (normalized?.length === attrs.groupMembers.length) next.groupMembers = normalized;
    }
  }
  const differs = Object.keys(next).some(
    (key) => JSON.stringify(next[key] ?? null) !== JSON.stringify((attrs as unknown as Record<string, unknown>)[key] ?? null),
  );
  return differs ? (next as Partial<DocumentShapeAttrs>) : patch;
}

/**
 * The face picked in the Font box (or "no font") is what a SmartArt label already draws
 * (Arial, the model's default and "no face of its own"): nothing about the
 * words changes, but from now on the face is the author's, stated in the
 * file, as Word writes a face picked on Home.
 */
function smartArtFaceStated(
  attrs: DocumentShapeAttrs,
  leaves: readonly DocumentShapeTextLeaf[],
  command: SmartArtTextFormatCommand,
): Readonly<Partial<DocumentShapeAttrs>> | null {
  if (command.kind !== "family" && command.kind !== "variant") return null;
  if (!attrs.groupMembers?.length) return null;
  // "No font" gives a label that already draws the default back to the theme.
  const stated = command.family !== null;
  const targeted = new Set(leaves.map((leaf) => leaf.memberId));
  let changed = false;
  const groupMembers = attrs.groupMembers.map((member) => {
    if (
      !targeted.has(member.memberId) ||
      member.smartArtRole === undefined ||
      member.smartArtRole === "connector" ||
      (member.smartArtAuthorTextFont === true) === stated
    ) {
      return member;
    }
    changed = true;
    if (stated) return { ...member, smartArtAuthorTextFont: true as const };
    const { smartArtAuthorTextFont: _cleared, ...rest } = member;
    return rest as DocumentShapeGroupMember;
  });
  return changed ? { groupMembers } : null;
}

function documentShapeTextCommandPatchOfFields(
  value: DocumentShapeAttrsInput,
  command: SmartArtTextFormatCommand,
  memberId?: string | null,
): Readonly<Partial<DocumentShapeAttrs>> | null {
  const attrs = normalizeDocumentShapeAttrs(value);
  const leaves = documentShapeTextLeaves(attrs, memberId);
  if (leaves.length === 0) return null;
  if (!attrs.groupMembers?.length && attrs.textRuns?.length) {
    // The shape's border is selected: the command formats ALL of its words
    // (Word). It is decided on what the runs show ("mixed" turns Bold on),
    // written on the shape, and the runs no longer override that field.
    const shown = runLeaves(leaves);
    const decided = applySmartArtTextFormat(shown, command, DOCUMENT_SHAPE_TEXT_ADAPTER);
    if (!decided) return null;
    const base = leaves[0];
    let next: DocumentShapeAttrs;
    let runs: TextRun[] | undefined;
    if (command.kind === "step") {
      // Grow/Shrink Font steps every run's own size (Word): the shape's
      // size steps, and each run that has its own size steps that.
      next = applySmartArtTextFormat([base], command, DOCUMENT_SHAPE_TEXT_ADAPTER)?.[0].leaf ?? base.leaf;
      runs = attrs.textRuns.map((run, index) =>
        run.style?.fontSize === undefined
          ? run
          : {
              text: run.text,
              style: { ...run.style, fontSize: ptToPx(decided[index].leaf.textFontSizePt) },
            },
      );
    } else {
      // The value the command decided (the same for every word), written
      // on the shape; the runs stop overriding that field.
      const shownFormat = DOCUMENT_SHAPE_TEXT_ADAPTER.read(decided[0]) as Record<string, unknown>;
      const fields = FORMAT_FIELDS_OF[command.kind];
      const picked: Record<string, unknown> = {};
      for (const field of command.kind === "toggle" ? [command.field] : fields) {
        picked[field] = shownFormat[field];
      }
      next = DOCUMENT_SHAPE_TEXT_ADAPTER.write(base, picked as SmartArtNodeTextFormatPatch).leaf;
      runs = documentShapeTextRunsWithout(attrs.textRuns, attrs.text, runKeysOf(command));
    }
    const patch: Record<string, unknown> = {};
    for (const key of TEXT_KEYS) {
      if (JSON.stringify(attrs[key] ?? null) !== JSON.stringify(next[key] ?? null)) {
        patch[key] = next[key] ?? null;
      }
    }
    const normalizedRuns = runs?.length
      ? normalizeDocumentShapeAttrs({ ...attrs, textRuns: runs }).textRuns
      : undefined;
    if (JSON.stringify(normalizedRuns ?? null) !== JSON.stringify(attrs.textRuns ?? null)) {
      patch.textRuns = normalizedRuns ?? null;
    }
    return Object.keys(patch).length ? (patch as Partial<DocumentShapeAttrs>) : null;
  }
  const written = applySmartArtTextFormat(leaves, command, DOCUMENT_SHAPE_TEXT_ADAPTER);
  if (!written) return smartArtFaceStated(attrs, leaves, command);
  if (!attrs.groupMembers?.length) {
    const next = written[0].leaf;
    const patch: Record<string, unknown> = {};
    for (const key of TEXT_KEYS) {
      if (JSON.stringify(attrs[key] ?? null) !== JSON.stringify(next[key] ?? null)) {
        patch[key] = next[key] ?? null;
      }
    }
    return Object.keys(patch).length ? (patch as Partial<DocumentShapeAttrs>) : null;
  }
  const byMember = new Map(written.map((node) => [node.memberId, node.leaf]));
  const before = new Map(leaves.map((node) => [node.memberId, node.leaf]));
  const groupMembers = attrs.groupMembers.map((member) => {
    const next = byMember.get(member.memberId);
    const previous = before.get(member.memberId);
    if (!next || !previous || next === previous) return member;
    const updated: Record<string, unknown> = { ...member };
    for (const key of TEXT_KEYS) {
      if (JSON.stringify(previous[key] ?? null) === JSON.stringify(next[key] ?? null)) continue;
      if (next[key] === undefined || next[key] === null) delete updated[key];
      else updated[key] = next[key];
    }
    // The member's words formatted in part: the command decides the field
    // for ALL of them (Word, the shape's border selected), so its runs stop
    // overriding that field; Grow/Shrink steps each run's own size instead.
    // A size set or stepped by hand is the author's on a SmartArt shape, at
    // any value (14 pt included, which alone would read as never set).
    if (
      (command.kind === "size" || command.kind === "step") &&
      member.smartArtRole !== undefined &&
      member.smartArtRole !== "connector" &&
      JSON.stringify(previous.textFontSizePt ?? null) !== JSON.stringify(next.textFontSizePt ?? null)
    ) {
      updated.smartArtAuthorTextSize = true;
    }
    // A face chosen by hand is the author's on a SmartArt shape at any value
    // (Arial included, the shape model's default face, which alone reads as
    // "no face of its own" and is then drawn and exported in the theme's
    // body font); the Font box's "no font" gives the label back to the theme.
    if (
      (command.kind === "family" || command.kind === "variant") &&
      member.smartArtRole !== undefined &&
      member.smartArtRole !== "connector"
    ) {
      // A face picked in the Font box is stated, as Word writes it, even the
      // one the label already showed.
      if (command.family === null) delete updated.smartArtAuthorTextFont;
      else updated.smartArtAuthorTextFont = true;
    }
    const runs = memberRunsAfterCommand(member, command);
    if (runs !== member.textRuns) {
      if (runs?.length) updated.textRuns = runs;
      else delete updated.textRuns;
    }
    return updated as unknown as DocumentShapeGroupMember;
  });
  const next = normalizeDocumentShapeAttrs({ ...attrs, groupMembers });
  return next.groupMembers?.length === attrs.groupMembers.length
    ? { groupMembers: next.groupMembers }
    : null;
}

/**
 * Text Highlight Color on a selected shape whose words are runs (a label, a
 * group's members, a SmartArt graphic's nodes): every run of the shape -- of
 * one member, when a member is selected -- takes the colour, or loses it
 * (`null`, No Color). `toggle` is the face's click: words that all wear the
 * colour lose it. A shape whose words are a story is highlighted in the
 * document itself, not here. `null` when nothing would change.
 */
export function documentShapeHighlightPatch(
  value: DocumentShapeAttrsInput,
  color: string | null,
  options: { memberId?: string | null; toggle?: boolean } = {},
): Readonly<Partial<DocumentShapeAttrs>> | null {
  const attrs = normalizeDocumentShapeAttrs(value);
  const runsOf = (text: string, runs: readonly TextRun[] | null | undefined): TextRun[] =>
    runs?.length ? [...runs] : text ? [{ text }] : [];
  const targets: TextRun[][] = attrs.groupMembers?.length
    ? attrs.groupMembers
        .filter((member) => !options.memberId || member.memberId === options.memberId)
        .map((member) => runsOf(member.text, member.textRuns))
    : [runsOf(attrs.text, attrs.textRuns)];
  const worded = targets.flat().filter((run) => run.text.length > 0);
  if (worded.length === 0) return null;
  const off = color === null || (options.toggle === true && worded.every((run) => run.style?.highlight === color));
  const recolor = (runs: TextRun[]): TextRun[] =>
    runs.map((run) => {
      const style: NonNullable<TextRun["style"]> = { ...run.style };
      if (off) delete style.highlight;
      else style.highlight = color!;
      return Object.keys(style).length ? { ...run, style } : { text: run.text };
    });
  let next: Partial<DocumentShapeAttrs>;
  if (attrs.groupMembers?.length) {
    const groupMembers = attrs.groupMembers.map((member) =>
      options.memberId && member.memberId !== options.memberId
        ? member
        : { ...member, textRuns: recolor(runsOf(member.text, member.textRuns)) },
    );
    next = { groupMembers: normalizeDocumentShapeAttrs({ ...attrs, groupMembers }).groupMembers };
  } else {
    next = { textRuns: normalizeDocumentShapeAttrs({ ...attrs, textRuns: recolor(runsOf(attrs.text, attrs.textRuns)) }).textRuns ?? null };
  }
  const changed = Object.entries(next).some(
    ([key, item]) => JSON.stringify(item ?? null) !== JSON.stringify((attrs as unknown as Record<string, unknown>)[key] ?? null),
  );
  return changed ? next : null;
}

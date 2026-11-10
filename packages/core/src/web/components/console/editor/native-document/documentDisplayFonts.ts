/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { patchTopLevelDecorations } from "./topLevelDecorationPatch";
import {documentSyntheticBoldAdvanceDecorations} from "./documentSyntheticBoldAdvanceProjection";
import { readDocumentNumberingLayout } from "./documentNumberingLayout";
import {prepareDocumentSyntheticBoldAdvanceMeasurement,documentSyntheticBoldAdvanceCandidateFace,documentSyntheticBoldAdvanceFontStateSnapshot,documentSyntheticBoldAdvanceFontStatesEqual,type SyntheticBoldAdvanceMeasurement} from "./documentSyntheticBoldAdvance";
import {loadOfficeData} from "../../../../lib/office-data/store";
import {prepareDocumentSyntheticBoldLineBoxStyles, resolveDocumentSyntheticBoldLineBox, syntheticBoldLineBoxStyleKey} from "./documentSyntheticBoldLineMetrics";
import {warmDocumentFontCharacters} from "./documentFontCharacters";
import {prepareLocalWordLineMetricFace} from "./documentLocalWordLineMetrics";
import {prepareDocumentMonoLineBoxStyles,resolveDocumentMonoLineBox,monoLineBoxStyleKey} from "./documentMonoLineMetrics";
import { writeDocumentLayoutStatus } from "./documentLayoutStatus";
import { Extension, type Editor } from "@tiptap/core";
import { Plugin, PluginKey, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import { effectiveBoldAt } from "./boldExtension";
import { normalizeDocxBreakRunMarks } from "./documentHardBreakRunFormat";
import { documentFontScriptRanges, documentFontDescriptor, fontFamilyWithCjkRenderFallback } from "../extensions/TextStyling";
import { normalizeOfficeFontFace, officeFontFaceFamily, type OfficeFontFace } from "../../../office-ribbon/officeFontFace";
import { isEastAsianFontCodePoint } from "../../../office-ribbon/officeFontScriptSlots";
import { officePrimaryFontFamily } from "../../../office-ribbon/officeTypographyCatalog";
import { officeFontSubstitution } from "../../../../lib/office-fonts/substitutes";
import { retainOfficeFontAssets, registerLeasedFontAsset, OFFICE_FONT_ASSETS_EVENT } from "../../../office-ribbon/officeFontAssets";
import { getPageFontBindings } from "../../../../lib/notebook-sdk";
import { captureSavedPageHotCacheScope, isSavedPageHotCacheScopeCurrent } from "../../../../lib/saved-page-hot-cache";
import {readWordFontSourceContext, resolveDocumentWordLineBox, prepareDocumentWordLineBoxStyles, documentWordLineBoxStyleKey,
  WORD_LINE_METRICS_TARGET, WORD_LINE_FRACTION_PROPERTY, type WordFontSourceContext, type WordLineMetricsTarget} from "./documentWordLineMetrics";

export const DOCUMENT_DISPLAY_FONT_BINDINGS_EVENT = "mrnote:document-display-font-bindings-changed";
export type DocumentDisplayFontFailure = "timeout" | "receipt" | "font" | "request" |
  "font-characters" | "font-characters-timeout" | "font-characters-source" |
  "font-characters-unstable" | "font-characters-budget";
export function notifyDocumentDisplayFontBindings(pageId: string): void {
  window.dispatchEvent(new CustomEvent(DOCUMENT_DISPLAY_FONT_BINDINGS_EVENT, {detail: {pageId}}));
}

type Role = "regular" | "bold" | "italic" | "boldItalic";
export interface DocumentDisplayFontPlan {
  faces: OfficeFontFace[];
  byNameRole: Map<string, OfficeFontFace>;
  sourceContext?: WordFontSourceContext | null;
  lineMetricsTarget?: WordLineMetricsTarget;
  lineBoxStyles?: Map<string,string>;
  syntheticBoldLineBoxStyles?: Map<string,string>;
  monoLineBoxStyles?: Map<string,string>;
  syntheticBoldAdvanceMeasurements?: Map<OfficeFontFace,SyntheticBoldAdvanceMeasurement>;
  localMetricFaces?: OfficeFontFace[];
  /** Declared family (lowercased) -> the face Word draws its East Asian
   * characters in instead: an embedded face whose font-table entry declares
   * no East Asian character set is not used for them (recorded at import). */
  eastAsianSubstitutes?: Map<string, string>;
  /** Lower-case family -> the faces Word sets it in where it is not
   * installed: the font table's alternate name, then Word's own pick by
   * charset and family (recorded at import,
   * document_word_font_context.read_docx_missing_font_substitutes). */
  missingFontSubstitutes?: Map<string, readonly string[]>;
}
/** How long a document waits for its font plan before it is laid out without it. */
const DISPLAY_FONT_PLAN_DEADLINE_MS = 15000;
const emptyPlan = (): DocumentDisplayFontPlan => ({faces: [], byNameRole: new Map()});
const keyFor = (name: string, role: string) => `${name.toLowerCase()}\u0000${role}`;

export function readDocumentDisplayFontPlan(value: unknown, expected: {pageId: string; notebookId: string; version: number}): DocumentDisplayFontPlan {
  if (!value || typeof value !== "object") throw new Error("document_display_fonts_invalid_receipt");
  const raw = value as Record<string, unknown>;
  if (raw.page_id !== expected.pageId || raw.notebook_id !== expected.notebookId || raw.content_version !== expected.version ||
      !Array.isArray(raw.font_faces) || raw.font_faces.length > 64 || !Array.isArray(raw.verified_font_aliases) || raw.verified_font_aliases.length > 4160) throw new Error("document_display_fonts_changed_source");
  const faces: OfficeFontFace[] = [], files = new Map<string, OfficeFontFace>();
  for (const rawFace of raw.font_faces) {
    const face = normalizeOfficeFontFace(rawFace);
    if (!face?.asset || face.family !== face.asset.family || face.asset.files.some(file => !file.attachmentId)) throw new Error("document_display_fonts_invalid_identity");
    faces.push(face);
    for (const file of face.asset.files) {
      const key = keyFor(face.asset.family, file.role);
      if (files.has(key) && files.get(key)!.asset!.files.find(peer => peer.role === file.role)?.sha256 !== file.sha256) throw new Error("document_display_fonts_ambiguous_identity");
      files.set(key, face);
    }
  }
  const byNameRole = new Map<string, OfficeFontFace>();
  for (const entry of raw.verified_font_aliases) {
    if (!entry || typeof entry !== "object") throw new Error("document_display_fonts_invalid_alias");
    const alias = entry as Record<string, unknown>;
    if (Object.keys(alias).sort().join() !== "family,name,role,sha256" || typeof alias.name !== "string" || !alias.name.trim() || alias.name.length > 256 || /[\x00-\x1f\x7f]/.test(alias.name) || typeof alias.family !== "string" || typeof alias.role !== "string" || !["regular","bold","italic","boldItalic"].includes(alias.role)) throw new Error("document_display_fonts_invalid_alias");
    const face = files.get(keyFor(alias.family, alias.role));
    if (!face || face.asset!.files.find(file => file.role === alias.role)?.sha256 !== alias.sha256) throw new Error("document_display_fonts_unverified_alias");
    const key = keyFor(alias.name, alias.role), previous = byNameRole.get(key);
    if (previous && officeFontFaceFamily(previous) !== officeFontFaceFamily(face)) throw new Error("document_display_fonts_ambiguous_alias");
    byNameRole.set(key, face);
  }
  const eastAsianSubstitutes = new Map<string, string>();
  if (Array.isArray(raw.east_asian_substitutes)) {
    for (const entry of raw.east_asian_substitutes.slice(0, 64)) {
      if (!entry || typeof entry !== "object") continue;
      const {family, substitute} = entry as Record<string, unknown>;
      if (typeof family !== "string" || typeof substitute !== "string" || !substitute.trim()
          || substitute.length > 64 || /[\x00-\x1f\x7f<>;{}\\"']/.test(substitute)) continue;
      eastAsianSubstitutes.set(family.toLowerCase(), substitute);
    }
  }
  const missingFontSubstitutes = new Map<string, readonly string[]>();
  if (Array.isArray(raw.missing_font_substitutes)) {
    for (const entry of raw.missing_font_substitutes.slice(0, 256)) {
      if (!entry || typeof entry !== "object") continue;
      const {family, faces} = entry as Record<string, unknown>;
      if (typeof family !== "string" || !family.trim() || family.length > 128 || !Array.isArray(faces)) continue;
      const clean = faces.slice(0, 4).filter((face): face is string => typeof face === "string" && Boolean(face.trim())
        && face.length <= 64 && !/[\x00-\x1f\x7f<>;{}\\"']/.test(face));
      if (clean.length) missingFontSubstitutes.set(family.toLowerCase(), clean);
    }
  }
  return {faces, byNameRole, sourceContext: readWordFontSourceContext(raw.source_font_context),
    ...(eastAsianSubstitutes.size ? {eastAsianSubstitutes} : {}),
    ...(missingFontSubstitutes.size ? {missingFontSubstitutes} : {})};
}

/** The alias families (lowercased) a verified plan actually paints the given
 * declared Word families under. A bound face renders as `MRAI-Font-<sha>`, so a
 * caller reading a computed font stack can no longer see the declared name. */
export function documentDisplayFontAliasFamilies(
  plan: DocumentDisplayFontPlan | null | undefined,
  families: Iterable<string>,
): Set<string> {
  const aliases = new Set<string>();
  if (!plan?.byNameRole.size) return aliases;
  for (const family of families) {
    for (const role of ["regular", "bold", "italic", "boldItalic"] as const) {
      const face = plan.byNameRole.get(keyFor(family, role));
      if (face) aliases.add(officeFontFaceFamily(face).toLowerCase());
    }
  }
  return aliases;
}

/** The alias families (lowercased) a verified plan paints East Asian faces
 * under. East Asian is the FACE's property (Word reads it from the font's
 * code pages), so a bound `MRAI-Font-<sha>` keeps it from the Word family it
 * was verified for, or from its asset's family: the recorded table cannot
 * know the sha of a subset the page itself produced (an exported Noto Sans
 * CJK SC), and the painted stack no longer shows the declared name. */
export function documentDisplayFontEastAsianAliasFamilies(
  plan: DocumentDisplayFontPlan | null | undefined,
  isEastAsian: (family: string) => boolean,
): Set<string> {
  const aliases = new Set<string>();
  if (!plan?.byNameRole.size) return aliases;
  for (const [key, face] of plan.byNameRole) {
    const name = key.slice(0, key.indexOf("\u0000"));
    if (isEastAsian(name) || (face.asset !== undefined && isEastAsian(face.asset.family))) {
      aliases.add(officeFontFaceFamily(face).toLowerCase());
    }
  }
  return aliases;
}

/** Atom formatting lives in its canonical carrier; never recreate authored marks. */
function displayFontRanges(state: EditorState, start = 0, end = state.doc.content.size) {
  const ranges = documentFontScriptRanges(state, start, end).map(range => ({...range, marks: range.node.marks, isBreak: false}));
  state.doc.nodesBetween(start, end, (node, from) => {
    if (node.type.name !== "hardBreak") return;
    let formats;
    try { formats = normalizeDocxBreakRunMarks(node.attrs.docxRunMarks); }
    catch { return; }
    if (!formats) return;
    const marks = formats.flatMap(mark => {
      const type = mark.type && state.schema.marks[mark.type];
      return type ? [type.create(mark.attrs)] : [];
    });
    const primary = marks.find(mark => mark.type.name === "fontFamily");
    const cjk = marks.find(mark => mark.type.name === "fontCjkFamily");
    const selected = primary ?? cjk;
    if (!selected) return;
    ranges.push({from, to: from + node.nodeSize, node, slot: "primary", marks, isBreak: true,
      value: documentFontDescriptor(selected.attrs.family, selected.attrs.face)});
  });
  return ranges;
}

export interface DocumentDisplayFontNeed { family: string; bold?: boolean; italic?: boolean; boldItalic?: boolean }
export function unresolvedDocumentDisplayFontNeeds(state: EditorState, plan?: DocumentDisplayFontPlan | null): DocumentDisplayFontNeed[] {
  const needs = new Map<string, DocumentDisplayFontNeed>();
  for (const range of displayFontRanges(state)) {
    if (range.value?.face) continue;
    const family = officePrimaryFontFamily(range.value?.family ?? String(state.doc.attrs.defaultFontFamily ?? ""));
    if (!family || ["serif", "sans-serif", "monospace", "system-ui", "ui-monospace"].includes(family.toLowerCase())) continue;
    const bold = effectiveBoldAt(state, range.from, range.marks);
    const italic = range.marks.some(mark => mark.type.name === "italic" && mark.attrs.value !== false);
    const role: Role = bold && italic ? "boldItalic" : bold ? "bold" : italic ? "italic" : "regular";
    if (plan?.byNameRole.has(keyFor(family, role))) continue;
    const need = needs.get(family) ?? {family};
    if (role !== "regular") need[role] = true;
    needs.set(family, need);
  }
  state.doc.descendants(block => {
    if (!block.isTextblock) return;
    const layout = readDocumentNumberingLayout(block.attrs);
    const family = layout && officePrimaryFontFamily(layout.markerRun.fonts.ascii);
    if (!layout || !family) return;
    const { bold, italic } = layout.markerRun;
    const role: Role = bold && italic ? "boldItalic" : bold ? "bold" : italic ? "italic" : "regular";
    if (plan?.byNameRole.has(keyFor(family, role))) return;
    const need = needs.get(family) ?? { family };
    if (role !== "regular") need[role] = true;
    needs.set(family, need);
  });
  return [...needs.values()].sort((a,b)=>a.family.localeCompare(b.family));
}

export const documentDisplayFontPluginKey = new PluginKey<{plan: DocumentDisplayFontPlan; decorations: DecorationSet}>("documentDisplayFonts");

type DisplayFontRange = ReturnType<typeof displayFontRanges>[number] & {eastAsianSubstitute?: string};

/** The installed faces that are Word's East Asian substitutes on each
 * platform, ahead of the served open face: Word's 宋体 is SimSun on Windows
 * and Songti on a Mac, and a stack that named only the served face drew the
 * system's sans while that face was still unloaded. */
const EAST_ASIAN_SUBSTITUTE_SYSTEM_FACES: Readonly<Record<string, readonly string[]>> = {
  "宋体": ["SimSun", '"Songti SC"', "STSong"],
  simsun: ["SimSun", '"Songti SC"', "STSong"],
  "新細明體": ["PMingLiU", '"Songti TC"', '"LiSong Pro"'],
  pmingliu: ["PMingLiU", '"Songti TC"', '"LiSong Pro"'],
  "ｍｓ 明朝": ['"MS Mincho"', '"Hiragino Mincho ProN"'],
  "ms mincho": ['"MS Mincho"', '"Hiragino Mincho ProN"'],
  "바탕": ["Batang", "AppleMyungjo"],
  batang: ["Batang", "AppleMyungjo"],
};

function eastAsianSubstituteStack(substitute: string): string {
  const served = fontFamilyWithCjkRenderFallback(`"${substitute}"`) ?? `"${substitute}"`;
  const system = EAST_ASIAN_SUBSTITUTE_SYSTEM_FACES[substitute.toLowerCase()] ?? [];
  return [`"${substitute}"`, ...system, served].join(", ");
}

/** A range whose family Word does not use for East Asian characters, split
 * into its East Asian characters (painted in the recorded substitute) and the
 * rest. A run that names one family for every slot carries a single font mark,
 * so its East Asian characters reach here in the primary range; Word routes
 * them by character (`isEastAsianFontCodePoint`, its slot rule), quotes and
 * Latin punctuation staying in the family. */
function withEastAsianSubstitutes(state: EditorState, plan: DocumentDisplayFontPlan, range: DisplayFontRange): DisplayFontRange[] {
  if (range.isBreak || range.value?.face || !plan.eastAsianSubstitutes?.size) return [range];
  const family = officePrimaryFontFamily(range.value?.family ?? String(state.doc.attrs.defaultFontFamily ?? ""));
  const substitute = plan.eastAsianSubstitutes.get(family.toLowerCase());
  if (!substitute) return [range];
  if (range.slot === "eastAsia") return [{...range, eastAsianSubstitute: substitute}];
  const text = state.doc.textBetween(range.from, range.to, "\ufffc", "\ufffc");
  if (text.length !== range.to - range.from) return [range];
  const pieces: DisplayFontRange[] = [];
  let start = 0, eastAsian: boolean | null = null;
  for (let index = 0; index < text.length;) {
    const cp = text.codePointAt(index)!;
    const width = cp > 0xffff ? 2 : 1;
    const isEastAsian = isEastAsianFontCodePoint(cp);
    if (eastAsian !== null && isEastAsian !== eastAsian) {
      pieces.push({...range, from: range.from + start, to: range.from + index, ...(eastAsian ? {eastAsianSubstitute: substitute} : {})});
      start = index;
    }
    eastAsian = isEastAsian;
    index += width;
  }
  if (eastAsian !== null) pieces.push({...range, from: range.from + start, to: range.to, ...(eastAsian ? {eastAsianSubstitute: substitute} : {})});
  return pieces;
}

/**
 * The stack a family the page cannot draw is painted with: the family itself
 * (a machine that has it keeps it), then the faces Word sets it in where it is
 * missing -- not the browser's default face. A family the product serves
 * (lib/office-fonts) or the page binds is not one of these. Null otherwise.
 */
const quotedFontName = (name: string) => `"${name.replace(/["\\]/g, "")}"`;
export function documentMissingFontStack(plan: DocumentDisplayFontPlan, family: string): string | null {
  const faces = plan.missingFontSubstitutes?.get(family.toLowerCase());
  if (!faces?.length || officeFontSubstitution(family)?.substitutes.length) return null;
  if (["serif", "sans-serif", "monospace", "system-ui", "ui-monospace"].includes(family.toLowerCase())) return null;
  return fontFamilyWithCjkRenderFallback([quotedFontName(family), ...faces.map(quotedFontName)].join(", "));
}

/** The stack a bound face paints with: the face, then the run's own family. A
 * subset (`w:subsetted`, `saveSubsetFonts`) holds only the characters its
 * author typed: one typed here falls back alone, to the family's own faces
 * and not to the browser's default, and the same stack lays it out. */
export function documentBoundFaceStack(face: OfficeFontFace, family: string): string {
  const own = fontFamilyWithCjkRenderFallback(quotedFontName(family));
  return `"${officeFontFaceFamily(face)}"${own ? `, ${own}` : ""}`;
}

function decorations(state: EditorState, plan: DocumentDisplayFontPlan): DecorationSet {
  if (!plan.byNameRole.size && !plan.missingFontSubstitutes?.size) return DecorationSet.empty;
  return DecorationSet.create(state.doc, [...rangeDecorations(state, plan, 0, state.doc.content.size),...documentSyntheticBoldAdvanceDecorations(state,plan)]);
}

/** The display-font decorations of the text between `start` and `end`. */
function rangeDecorations(state: EditorState, plan: DocumentDisplayFontPlan, start: number, end: number): Decoration[] {
  if (!plan.byNameRole.size && !plan.missingFontSubstitutes?.size) return [];
  return displayFontRanges(state, start, end).flatMap(range => withEastAsianSubstitutes(state, plan, range)).flatMap(range => {
    if (range.value?.face) return [];
    const family = officePrimaryFontFamily(range.value?.family ?? String(state.doc.attrs.defaultFontFamily ?? ""));
    const bold = effectiveBoldAt(state, range.from, range.marks);
    const italic = range.marks.some(mark => mark.type.name === "italic" && mark.attrs.value !== false);
    const role: Role = bold && italic ? "boldItalic" : bold ? "bold" : italic ? "italic" : "regular";
    const substitute = range.eastAsianSubstitute;
    if (substitute) {
      // Word sets these characters in the theme's East Asian face, not in the
      // embedded family the run names (measured in Word 16 for Mac).
      const stack = eastAsianSubstituteStack(substitute);
      return [Decoration.inline(range.from, range.to, {
        class: "document-display-font",
        "data-document-display-font": "true",
        "data-document-word-line-metrics": "unverified",
        "data-document-east-asian-substitute": substitute,
        style: `--document-display-font-family:${stack};`,
      }, {inclusiveStart: false, inclusiveEnd: false})];
    }
    if (!plan.byNameRole.has(keyFor(family, "regular")) && !plan.byNameRole.has(keyFor(family, role))) {
      // A face neither installed nor served, which Word itself replaces:
      // painted in Word's replacement rather than the browser's default.
      const stack = documentMissingFontStack(plan, family);
      if (!stack) return [];
      return [Decoration.inline(range.from, range.to, {
        class: "document-display-font",
        "data-document-display-font": "true",
        "data-document-word-line-metrics": "unverified",
        "data-document-missing-font": family,
        style: `--document-display-font-family:${stack};`,
      }, {inclusiveStart: false, inclusiveEnd: false})];
    }
    let face = plan.byNameRole.get(keyFor(family, role));
    let lineStyle:string|undefined, syntheticBold=false, exactSynthetic=false;
    const parent = state.doc.resolve(range.from).parent;
    const lineMultiplier=Number(parent.attrs.lineHeight ?? state.doc.attrs.defaultLineHeight);
    if (!face && role !== "regular" && !range.isBreak && parent.attrs.lineHeightPt != null &&
        parent.attrs.lineHeightRule !== "atLeast" && plan.byNameRole.has(keyFor(family, "regular"))) {
      // An EXACT line is its stated height whatever face sets it, so a bold
      // or italic run of a family that ships only its regular face is that
      // face, emboldened or slanted by the engine -- what Word draws (a
      // converter's "MRAI Sans" table headers) -- with no line box at stake.
      face = plan.byNameRole.get(keyFor(family, "regular"));
      exactSynthetic = true;
    }
    if (!face && role === "bold" && !range.isBreak) {
      const regular = plan.byNameRole.get(keyFor(family, "regular"));
      const box = regular ? resolveDocumentSyntheticBoldLineBox({state,from:range.from,marks:range.marks,
        face:regular,role,text:range.node.textContent,sourceContext:plan.sourceContext ?? null,target:plan.lineMetricsTarget}) : null;
      lineStyle = box ? plan.syntheticBoldLineBoxStyles?.get(syntheticBoldLineBoxStyleKey(box)) : undefined;
      // An absent bold file is not a generic fallback. Only the independently
      // bounded synthetic paragraph and its current 700 probe can bind here.
      if (!regular || !lineStyle) return [];
      face = regular;
      syntheticBold = true;
    }
    if (!face) return [];
    if (!syntheticBold && !exactSynthetic) {
      const lineBox=range.isBreak ? null : resolveDocumentWordLineBox({state,from:range.from,marks:range.marks,
        face,role,text:range.node.textContent,sourceContext:plan.sourceContext ?? null,target:plan.lineMetricsTarget});
      lineStyle=lineBox ? plan.lineBoxStyles?.get(documentWordLineBoxStyleKey(lineBox,lineMultiplier)) : undefined;
      if (!lineStyle && !range.isBreak) {
        const mono=resolveDocumentMonoLineBox({state,from:range.from,marks:range.marks,face,role,
          text:range.node.textContent,sourceContext:plan.sourceContext??null,target:plan.lineMetricsTarget,
          regularFaces:plan.byNameRole,regularLineBoxStyles:plan.lineBoxStyles});
        lineStyle=mono?plan.monoLineBoxStyles?.get(monoLineBoxStyleKey(mono)):undefined;
      }
    }
    const attrs = {
      class: `document-display-font${lineStyle ? " document-word-linebox" : ""}`,
      "data-document-display-font": "true",
      "data-document-word-line-metrics": lineStyle ? "native-control" : "unverified",
      ...(syntheticBold ? {"data-document-word-line-role":"synthetic-bold-advance"} : {}),
      "data-document-word-font-proof": face.asset ? "attachment-sha256" : "local-metrics",
      "data-office-font-face": JSON.stringify(face),
      // A custom property cannot be reparsed as an authored font-family mark.
      style: `--document-display-font-family:${documentBoundFaceStack(face, family)};${lineStyle ?? ""}`,
    };
    if (!range.isBreak && lineStyle?.includes(`${WORD_LINE_FRACTION_PROPERTY}:`)) {
      // Explicit nodeName makes these two distinct inline boxes in PM's outer
      // decoration stack. PM does not keep same-range decorations in order once
      // another inline decoration splits the run (removeOverlap reverses the
      // split pieces), so either span may end up outside. Both carry the
      // measured variables and the stylesheet assigns the ascent-carrier role to
      // the outer box and the missing-descent role to the inner one by nesting.
      const spec={inclusiveStart:false,inclusiveEnd:false,documentWordLineBox:true};
      return [
        // Merge into the existing source/script leaf before adding wrappers.
        // An inherited family on the outer native box cannot override a
        // font-family explicitly rendered on that inner source leaf.
        Decoration.inline(range.from,range.to,{class:"document-display-font-leaf"},{inclusiveStart:false,inclusiveEnd:false}),
        Decoration.inline(range.from,range.to,{nodeName:"span",class:"document-word-line-fraction",style:attrs.style},spec),
        Decoration.inline(range.from,range.to,{...attrs,nodeName:"span"},spec),
      ];
    }
    return [range.isBreak ? Decoration.node(range.from, range.to, attrs)
      : Decoration.inline(range.from, range.to, attrs, {inclusiveStart: false, inclusiveEnd: false, documentWordLineBox: Boolean(lineStyle)})];
  });
}

export function createDocumentDisplayFontPlugin(): Plugin {
  return new Plugin({
    key: documentDisplayFontPluginKey,
    state: {
      init: () => ({plan: emptyPlan(), decorations: DecorationSet.empty}),
      apply(transaction, previous, oldState, state) {
        const plan = transaction.getMeta(documentDisplayFontPluginKey) as DocumentDisplayFontPlan | undefined;
        if (plan) return {plan, decorations: decorations(state, plan)};
        if (!transaction.docChanged) return previous;
        // An edit builds again only the top-level blocks (a paged document's
        // pages) it replaced: the whole document was walked on every
        // keystroke (900 ms of 40 on the 272-page book). The synthetic-bold
        // advance, a document-wide pass when it applies at all, still
        // rebuilds everything.
        if (documentSyntheticBoldAdvanceDecorations(state, previous.plan).length) {
          return {plan: previous.plan, decorations: decorations(state, previous.plan)};
        }
        return {
          plan: previous.plan,
          decorations: patchTopLevelDecorations(oldState.doc, previous.decorations, state.doc,
            (from, to) => rangeDecorations(state, previous.plan, from, to), {readsDocAttrs: true}),
        };
      },
    },
    props: {decorations: state => documentDisplayFontPluginKey.getState(state)?.decorations},
    view(view){
      const doc=view.dom.ownerDocument;
      let disposed=false,queued=false;
      const capture=(plan:DocumentDisplayFontPlan)=>new Map([...plan.faces,...(plan.localMetricFaces??[])].map(face=>
        [face,documentSyntheticBoldAdvanceFontStateSnapshot(doc,officeFontFaceFamily(face))] as const));
      let previous=capture(documentDisplayFontPluginKey.getState(view.state)?.plan??emptyPlan());
      const onFonts=()=>{
        if(queued||disposed)return;
        queued=true;
        queueMicrotask(()=>{
          queued=false;
          if(disposed||view.isDestroyed||view.dom.hasAttribute("data-docs-display-fonts-pending"))return;
          const plan=documentDisplayFontPluginKey.getState(view.state)?.plan;
          if(!plan)return;
          const current=capture(plan);
          const changed=[...current].filter(([face,snapshot])=>!previous.has(face)||
            !documentSyntheticBoldAdvanceFontStatesEqual(previous.get(face)!,snapshot));
          if(!changed.length&&current.size===previous.size)return;
          previous=current;
          const tokens=new Map(plan.syntheticBoldAdvanceMeasurements??[]);
          for(const [face]of changed){
            const token=prepareDocumentSyntheticBoldAdvanceMeasurement(doc,face);
            if(token)tokens.set(face,token);else tokens.delete(face);
          }
          for(const face of tokens.keys())if(!current.has(face))tokens.delete(face);
          view.dispatch(view.state.tr.setMeta(documentDisplayFontPluginKey,{...plan,syntheticBoldAdvanceMeasurements:tokens}).setMeta("addToHistory",false));
          view.dom.dispatchEvent(new CustomEvent("mrn-page-rebalance",{bubbles:true,detail:{reason:"document-display-fonts"}}));
        });
      };
      for(const event of ["loading","loadingdone","loadingerror"])doc.fonts?.addEventListener?.(event,onFonts);
      doc.addEventListener(OFFICE_FONT_ASSETS_EVENT,onFonts);
      return {
        update(current,oldState){
          const plan=documentDisplayFontPluginKey.getState(current.state)?.plan;
          if(plan&&plan!==documentDisplayFontPluginKey.getState(oldState)?.plan)previous=capture(plan);
        },
        destroy(){disposed=true;for(const event of ["loading","loadingdone","loadingerror"])doc.fonts?.removeEventListener?.(event,onFonts);doc.removeEventListener(OFFICE_FONT_ASSETS_EVENT,onFonts);},
      };
    },
  });
}

/** Keep the view plugin installed for the editor lifetime. Reconfiguring the
 * plugin list rebuilds every viewport/scroll observer in a large document. */
export const DocumentDisplayFonts = Extension.create({
  name: "documentDisplayFonts",
  addProseMirrorPlugins: () => [createDocumentDisplayFontPlugin()],
});

export function mountDocumentDisplayFonts(editor: Editor, pageId: string, notebookId: string, onPlan?: (plan: DocumentDisplayFontPlan | null, failure?: DocumentDisplayFontFailure) => void, lineMetricsTarget?: WordLineMetricsTarget) {
  return mountDocumentDisplayFontSource(editor, {
    read: (version, signal) => getPageFontBindings(pageId, {signal}, version),
    parse: (response, version) => readDocumentDisplayFontPlan(response, {pageId, notebookId, version}),
  }, onPlan, lineMetricsTarget);
}

/** Read-only rendering source. Its parser binds the response to the exact
 * authorized revision; the shared lifecycle owns byte checks and font leases. */
export interface DocumentDisplayFontSource {
  read: (version: number, signal: AbortSignal) => Promise<unknown>;
  parse: (response: unknown, version: number) => DocumentDisplayFontPlan;
}

/** Word families a DOCX commonly declares without embedding, and whose local
 * face this build can identify by its own recorded metrics. */
const LOCAL_WORD_LINE_METRIC_FAMILIES = ['Arial Unicode MS', 'Aptos Mono'];

export function mountDocumentDisplayFontSource(editor: Editor, source: DocumentDisplayFontSource, onPlan?: (plan: DocumentDisplayFontPlan | null, failure?: DocumentDisplayFontFailure) => void, lineMetricsTarget?: WordLineMetricsTarget) {
  const root = editor.view.dom, doc = root.ownerDocument;
  if (!documentDisplayFontPluginKey.getState(editor.state)) throw new Error("document_display_font_plugin_missing");
  let release: () => void = () => {}, active = true, accepted = false, generation = 0, pending: AbortController | null = null;
  // The gate the plan holds the page with (data-docs-display-fonts-pending)
  // closes for one deadline in all, from the first read. Every refresh used
  // to close it again for its own 15 s, and a refresh comes with every new
  // version, every automatic retry and every bindings event: on a server
  // answering 503 while the settle saved versions, a book stayed unlaid out
  // for five minutes. Past the deadline a read goes on in the background and
  // a plan that lands is applied and the page laid out again.
  let gateOpenedAt: number | null = null, gateTimer: ReturnType<typeof setTimeout> | null = null;
  const nudge = () => root.dispatchEvent(new CustomEvent("mrn-page-rebalance", {bubbles: true, detail: {reason: "document-display-fonts"}}));
  const status = (value: string | null) => {
    writeDocumentLayoutStatus(root, "data-document-display-fonts", value, editor);
    writeDocumentLayoutStatus(root, "data-docs-display-fonts-pending", value === "loading" ? "true" : null, editor);
  };
  const apply = (plan: DocumentDisplayFontPlan) => {
    if (!editor.isDestroyed) editor.view.dispatch(editor.state.tr.setMeta(documentDisplayFontPluginKey, plan).setMeta("addToHistory", false));
  };
  const declaresLocalWordLineMetricFamily = (family: string) =>
    displayFontRanges(editor.state).some(range =>
      officePrimaryFontFamily(range.value?.family ?? '').toLowerCase() === family.toLowerCase());
  const withPreparedLineBoxes = (plan: DocumentDisplayFontPlan) => {
    const faces = [...plan.faces, ...(plan.localMetricFaces ?? [])];
    plan.lineMetricsTarget = lineMetricsTarget;
    plan.lineBoxStyles = new Map(faces.flatMap(face => [...prepareDocumentWordLineBoxStyles(doc, face)]));
    plan.syntheticBoldLineBoxStyles = new Map(faces.flatMap(face => [...prepareDocumentSyntheticBoldLineBoxStyles(doc, face)]));
    plan.monoLineBoxStyles = new Map(faces.flatMap(face => [...prepareDocumentMonoLineBoxStyles(doc, face)]));
    return plan;
  };
  // A family the package does not carry is painted by the reader's own local
  // face whether or not the source plan arrives; only Word's line box travels
  // with the plan. Losing it moved every page break in the document, so a page
  // that cannot read its plan still measures the face it is actually painting.
  // The probe is local, cached per document and family, and binds nothing it
  // cannot identify by its own recorded metrics — on a machine without the
  // face this is exactly the empty plan it replaces.
  const prepareLocalLineMetricPlan = async (): Promise<DocumentDisplayFontPlan | null> => {
    if (lineMetricsTarget !== WORD_LINE_METRICS_TARGET) return null;
    const plan = emptyPlan();
    for (const family of LOCAL_WORD_LINE_METRIC_FAMILIES) {
      if (!declaresLocalWordLineMetricFamily(family)) continue;
      const local = await prepareLocalWordLineMetricFace(doc, family);
      if (editor.isDestroyed) return null;
      if (local) {
        (plan.localMetricFaces ??= []).push(local);
        plan.byNameRole.set(keyFor(local.family, 'regular'), local);
      }
    }
    return plan.localMetricFaces?.length ? withPreparedLineBoxes(plan) : null;
  };
  return {
    async refresh(version: number, force = false) {
      if (!active || editor.isDestroyed || (accepted && !force)) return;
      const ownGeneration = ++generation, scope = captureSavedPageHotCacheScope();
      pending?.abort(); const controller = new AbortController(); pending = controller;
      let acquired: (() => void) | undefined;
      let preparedPlan: DocumentDisplayFontPlan | undefined;
      let localLineMetricPlan: DocumentDisplayFontPlan | null = null;
      const releasePending = () => { acquired?.(); acquired = undefined; };
      const fail = (reason: DocumentDisplayFontFailure) => {
        if (active && !editor.isDestroyed && ownGeneration === generation) {
          // The plan is genuinely unavailable and still says so, retry control
          // included. What survives is only the line box of a face this page
          // paints anyway, so a failed read cannot repaginate the document.
          accepted = false; apply(localLineMetricPlan ?? emptyPlan()); release(); release = () => {};
          onPlan?.(null, reason); status("unavailable"); nudge();
        }
      };
      const scopeCurrent = () => active && !editor.isDestroyed && ownGeneration === generation && isSavedPageHotCacheScopeCurrent(scope);
      const current = () => scopeCurrent() && !controller.signal.aborted;
      const preparedCurrent = () => preparedPlan && scopeCurrent() &&
        documentDisplayFontPluginKey.getState(editor.state)?.plan===preparedPlan;
      const retainPreparedForRetry = (reason: DocumentDisplayFontFailure) => {
        release(); release=acquired ?? (()=>{}); acquired=undefined; accepted=false;
        onPlan?.(null,reason); status("unavailable"); nudge();
      };
      let phase: DocumentDisplayFontFailure = "request";
      let deadlineExpired=false;
      controller.signal.addEventListener("abort", () => {
        const reason=phase==="font-characters" ? "font-characters-timeout" : "timeout";
        // Only our explicit deadline can retain validated paint. A caller's
        // cancellation or a lost account/editor/plan never borrows this path.
        if(deadlineExpired && phase==="font-characters" && preparedCurrent()) retainPreparedForRetry(reason);
        else { releasePending(); fail(reason); }
      }, {once: true});
      const timeout = setTimeout(() => {deadlineExpired=true;controller.abort();}, DISPLAY_FONT_PLAN_DEADLINE_MS);
      const now = Date.now();
      gateOpenedAt ??= now;
      const gateLeftMs = gateOpenedAt + DISPLAY_FONT_PLAN_DEADLINE_MS - now;
      if (gateLeftMs > 0) {
        status("loading"); nudge();
        // A later refresh holds only for what is left of the first deadline.
        if (gateTimer !== null) clearTimeout(gateTimer);
        gateTimer = setTimeout(() => {
          gateTimer = null;
          if (active && !editor.isDestroyed && root.getAttribute("data-document-display-fonts") === "loading") { status("unavailable"); nudge(); }
        }, gateLeftMs);
      }
      // Alongside the read, never in front of it: the source plan is what the
      // page is waiting for, and the local probe only has to be ready by the
      // time a failure would otherwise drop the document's line metrics.
      const localLineMetrics = prepareLocalLineMetricPlan().then(
        plan => {localLineMetricPlan = plan;}, () => {localLineMetricPlan = null;});
      try {
        const response = await source.read(version, controller.signal);
        phase = "receipt";
        const plan = source.parse(response, version);
        if (!current()) return;
        phase = "font";
        const assets = plan.faces.map(face => face.asset!);
        acquired = retainOfficeFontAssets(assets, doc);
        await Promise.all(assets.map(asset => registerLeasedFontAsset(asset, doc)));
        if (!current()) return;
        for (const face of plan.faces) for (const file of face.asset!.files) {
          const found = [...doc.fonts].some(font => font.family.replace(/^["']|["']$/g, "") === officeFontFaceFamily(face) && font.status === "loaded" &&
            font.weight === (file.role === "bold" || file.role === "boldItalic" ? "700" : "400") && font.style === (file.role === "italic" || file.role === "boldItalic" ? "italic" : "normal"));
          if (!found) throw new Error("document_display_fonts_unavailable");
        }
        // An ordinary DOCX need not embed a font already installed locally.
        // Bind only the native PostScript face whose independently recorded
        // glyph advances and font metrics match; never infer from its family name.
        for (const family of LOCAL_WORD_LINE_METRIC_FAMILIES) {
          if (!(lineMetricsTarget===WORD_LINE_METRICS_TARGET && plan.sourceContext?.context.fontActivation==='external-only' &&
            declaresLocalWordLineMetricFamily(family) && !plan.byNameRole.has(keyFor(family,'regular')))) continue;
          const local=await prepareLocalWordLineMetricFace(doc,family);
          if (!current()) return;
          if(local){(plan.localMetricFaces??=[]).push(local);plan.byNameRole.set(keyFor(local.family,'regular'),local);}
        }
        withPreparedLineBoxes(plan);
        // Apply font aliases while the loading gate still holds pagination.
        // FontFaceSet.ready alone covers only already requested subsets; every
        // DOM character must participate before publishing a stable font plan.
        apply(plan);
        preparedPlan=plan;
        phase="font-characters";
        await warmDocumentFontCharacters(root,{signal:controller.signal,current:()=>
          current() && documentDisplayFontPluginKey.getState(editor.state)?.plan===plan});
        if (!current()) return;
        // Word's measured bold advances are served (lib/office-data/store.ts):
        // a plan with the reference face waits for them here, behind the
        // font gate, so the projection is in the first layout.
        if([...plan.faces,...(plan.localMetricFaces ?? [])].some(documentSyntheticBoldAdvanceCandidateFace)){
          await loadOfficeData("word-synthetic-bold-advance");
          if (!current()) return;
        }
        plan.syntheticBoldAdvanceMeasurements=new Map();
        for(const face of [...plan.faces,...(plan.localMetricFaces ?? [])]){
          const measurement=prepareDocumentSyntheticBoldAdvanceMeasurement(doc,face);
          if(measurement)plan.syntheticBoldAdvanceMeasurements.set(face,measurement);
        }
        if(plan.syntheticBoldAdvanceMeasurements.size)apply(plan);
        release(); release = acquired; acquired = undefined; accepted = true;
        onPlan?.(plan); status(plan.byNameRole.size ? "available" : "unbound"); nudge();
      } catch (error) {
        await localLineMetrics;
        if (!controller.signal.aborted) {
          // Static stage codes are exposed through the existing retry control;
          // never surface font strings, source text, URLs, or account details.
          const codes:Record<string,DocumentDisplayFontFailure>={
            document_font_character_source_changed:"font-characters-source",
            document_font_character_source_unstable:"font-characters-unstable",
            document_font_character_style_budget:"font-characters-budget",
          };
          const reason=phase==="font-characters" && error instanceof Error ? codes[error.message] ?? phase : phase;
          if (phase==="font-characters" && reason!=="font-characters-source" && preparedPlan &&
              current() && documentDisplayFontPluginKey.getState(editor.state)?.plan===preparedPlan) {
            // CSS subset warmup is not the authority for these already checked
            // source faces. Keep their paint/leases while explicitly exposing
            // incomplete warmup for retry; never claim a settled font plan.
            retainPreparedForRetry(reason);
          } else fail(reason);
        }
      }
      finally { clearTimeout(timeout); releasePending(); }
    },
    destroy() {
      if (!active) return;
      active = false; generation++; pending?.abort(); release(); status(null);
      if (gateTimer !== null) { clearTimeout(gateTimer); gateTimer = null; }
      if (!editor.isDestroyed) apply(emptyPlan());
    },
  };
}

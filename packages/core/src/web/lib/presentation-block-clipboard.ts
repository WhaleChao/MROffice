/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { CSS_PX_PER_POINT } from "./presentation/text-units";
import { POINT_TEXT_UNIT, presentationTextUsesPoints } from "./presentation/text-units-block";
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import { presentationBlockSchema } from "./native-page-schemas";
import { isPictureObjectUrl } from "./page-picture-objects";
import { readPresentationNativeThemePaint } from "./presentation-native-theme-paint";
import { isValidInlineImageSource } from "./inline-image-source";
import { isObjectPlaceholder } from "./presentation-object-degraded";
import { readNativeListIndents } from "./native-drawing-geometry";
import { cssFontFamilyList } from "./css-font-family";

/**
 * Pure block-clipboard logic for the presentation editor canvas (PowerPoint / WPS
 * parity: Copy / Paste / Cut / Duplicate / Select-All of positioned blocks).
 *
 * The component owns the React state (the clipboard ref, the selection set, the
 * deck commit paths); this module owns the two pure, side-effect-free pieces that
 * are worth testing in isolation:
 *
 *  1. {@link reidAndOffsetBlocks} — given a deep snapshot of one or more blocks,
 *     produce INDEPENDENT clones: every block gets a fresh id, intra-clipboard
 *     references (group id, connector start/end shape anchors) are remapped onto
 *     the new ids (a ref pointing OUTSIDE the snapshot is dropped so a half-copied
 *     connector never dangles onto a source-slide shape), and the geometry is
 *     nudged by a fixed offset and clamped on-canvas. This is the shared kernel
 *     behind both Paste and Duplicate (Duplicate is just a Paste of the live
 *     selection without a clipboard round-trip).
 *
 *  2. {@link selectAllBlockIds} — the id set of every POSITIONED (geometry-bearing)
 *     block on a slide, the target of Cmd/Ctrl+A (the multi-select set only ever
 *     holds positioned blocks, mirroring the marquee / align contract).
 *
 * No React or DOM state: ids and group ids are injected so tests can assert the
 * exact remapping. The system-envelope admission edge deliberately reuses the
 * exported persisted presentation-block schema rather than maintaining a
 * weaker parallel block validator.
 *
 * All geometry values are EMU, the same coordinate space as the rest of the
 * presentation stack (`presentation-drag` / `presentation-align` / `-marquee`).
 */

/**
 * A block as this module consumes it: an opaque record that carries at least a
 * string `id`. We deliberately stay structural (not the full `PresentationBlock`
 * union) so the kernel is dependency-free and works on the JSON-cloned snapshots
 * the clipboard holds. Geometry / group / connector fields are read by key.
 */
export type ClipboardBlock = Record<string, unknown> & { id: string };

/** The four truthful Paste modes exposed by the PowerPoint Home ribbon. */
export type PresentationPasteMode =
  | "keep-source"
  | "destination-theme"
  | "picture"
  | "text-only";

/** Chromium's opt-in Web Custom Format name for MRAI presentation blocks. */
export const PRESENTATION_BLOCK_CLIPBOARD_MIME =
  "web application/vnd.mrai.presentation-blocks+json";

/** Blob media type carried behind {@link PRESENTATION_BLOCK_CLIPBOARD_MIME}. */
export const PRESENTATION_BLOCK_CLIPBOARD_BLOB_TYPE =
  "application/vnd.mrai.presentation-blocks+json";

export const PRESENTATION_BLOCK_CLIPBOARD_VERSION = 1 as const;
const PRESENTATION_BLOCK_CLIPBOARD_KIND = "mrai.presentation.blocks" as const;
const MAX_PRESENTATION_CLIPBOARD_BLOCKS = 200;
export const PRESENTATION_BLOCK_CLIPBOARD_MAX_BYTES = 16 * 1024 * 1024;
const MAX_PRESENTATION_CLIPBOARD_DEPTH = 40;
const MAX_PRESENTATION_CLIPBOARD_NODES = 50_000;

const SYSTEM_ASSET_CUSTODY_KEYS = [
  "sourceKind",
  "assetId",
  "mimeType",
  "pixelWidth",
  "pixelHeight",
  "sizeBytes",
  "hasAlpha",
  "exifOrientation",
] as const;

export interface PresentationBlockClipboardEnvelopeV1 {
  kind: typeof PRESENTATION_BLOCK_CLIPBOARD_KIND;
  version: typeof PRESENTATION_BLOCK_CLIPBOARD_VERSION;
  blocks: ClipboardBlock[];
}

const PRESENTATION_PASTE_MODES = new Set<PresentationPasteMode>([
  "keep-source",
  "destination-theme",
  "picture",
  "text-only",
]);

const SUPPORTED_CLIPBOARD_BLOCK_TYPES = new Set([
  "text",
  "image",
  "table",
  "chart",
  "shape",
  "embed",
]);

const DESTINATION_THEME_KEYS = [
  "fill",
  "fillTo",
  "gradientAngle",
  "stroke",
  "outlineColor",
  "borderColor",
  "color",
  "fontFamily",
  "highlight",
] as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function isSupportedClipboardBlock(value: unknown): value is ClipboardBlock {
  if (!isRecord(value)) return false;
  return (
    typeof value.id === "string" &&
    value.id.length > 0 &&
    typeof value.type === "string" &&
    SUPPORTED_CLIPBOARD_BLOCK_TYPES.has(value.type)
  );
}

function hasOwn(record: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(record, key);
}

function hasClipboardCustodyReceipt(block: ClipboardBlock): boolean {
  return (
    hasOwn(block, "source") ||
    hasOwn(block, "sourceOrder") ||
    hasOwn(block, "zIndex")
  );
}

function hasSystemAssetCustody(block: ClipboardBlock): boolean {
  return SYSTEM_ASSET_CUSTODY_KEYS.some((key) => hasOwn(block, key));
}

function stripSystemAssetCustody(block: ClipboardBlock): void {
  for (const key of SYSTEM_ASSET_CUSTODY_KEYS) delete block[key];
}

function hasPositionedClipboardGeometry(block: ClipboardBlock): boolean {
  const left = block.left;
  const top = block.top;
  const width = block.width;
  const height = block.height;
  return (
    typeof left === "number" &&
    Number.isFinite(left) &&
    typeof top === "number" &&
    Number.isFinite(top) &&
    typeof width === "number" &&
    Number.isFinite(width) &&
    width > 0 &&
    typeof height === "number" &&
    Number.isFinite(height) &&
    height > 0
  );
}

function hasSelfContainedSystemImage(block: ClipboardBlock): boolean {
  if (block.type !== "image") return true;
  const src = block.src;
  // A picture stored beside its deck travels by its content address: the
  // server copies it into the deck it is pasted into only for a reader of
  // the deck it came from (lib/page-picture-objects).
  if (isPictureObjectUrl(src)) return true;
  return (
    typeof src === "string" &&
    src.toLowerCase().startsWith("data:image/") &&
    isValidInlineImageSource(src)
  );
}

function clipboardUtf8ByteLength(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function clipboardRunText(value: unknown): string {
  if (!Array.isArray(value)) return "";
  return value
    .map((run) =>
      isRecord(run) && typeof run.text === "string" ? run.text : "",
    )
    .join("");
}

/**
 * Extract the plain textual payload PowerPoint exposes for Paste Text Only.
 * The traversal is deliberately bounded to the presentation model's known
 * text carriers; unknown object fields never become executable HTML or text.
 */
export function presentationClipboardBlockPlainText(
  block: ClipboardBlock,
): string | null {
  if (typeof block.text === "string") return block.text;
  if (Array.isArray(block.paragraphs)) {
    return block.paragraphs
      .map((value) => {
        if (!isRecord(value)) return "";
        if (typeof value.text === "string") return value.text;
        return clipboardRunText(value.runs);
      })
      .join("\n");
  }
  if (Array.isArray(block.rows)) {
    return block.rows
      .map((row) =>
        Array.isArray(row)
          ? row
              .map((cell) =>
                isRecord(cell) && typeof cell.text === "string"
                  ? cell.text
                  : "",
              )
              .join("\t")
          : "",
      )
      .join("\n");
  }
  if (typeof block.title === "string") return block.title;
  return null;
}

/** Standard `text/plain` representation published beside the MRAI envelope. */
export function serializePresentationClipboardPlainText(
  source: ReadonlyArray<ClipboardBlock> | null | undefined,
): string | null {
  if (!canPreparePresentationPaste(source, "keep-source")) return null;
  const values = source.map(presentationClipboardBlockPlainText);
  if (values.some((value) => value === null)) return null;
  const text = (values as string[]).join("\n");
  return text.trim().length > 0 ? text : null;
}

function escapeClipboardHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function clipboardStyleCss(source: Record<string, unknown>): string {
  const css: string[] = [];
  if (source.bold === true) css.push("font-weight:bold");
  if (source.italic === true) css.push("font-style:italic");
  const decorations = [source.underline === true ? "underline" : "", source.strike === true ? "line-through" : ""]
    .filter(Boolean)
    .join(" ");
  if (decorations) css.push(`text-decoration:${decorations}`);
  if (typeof source.fontSize === "number" && Number.isFinite(source.fontSize) && source.fontSize > 0) {
    css.push(`font-size:${source.fontSize}pt`);
  }
  if (typeof source.color === "string" && /^#[0-9a-f]{6}$/i.test(source.color)) css.push(`color:${source.color}`);
  if (typeof source.fontFamily === "string" && /^[\w .-]{1,80}$/.test(source.fontFamily)) {
    css.push(`font-family:${cssFontFamilyList(source.fontFamily)}`);
  }
  return css.join(";");
}

function clipboardRunsHtml(runs: unknown, fallback: string): string {
  if (!Array.isArray(runs) || runs.length === 0) return escapeClipboardHtml(fallback);
  return runs
    .map((run) => {
      if (!isRecord(run) || typeof run.text !== "string") return "";
      const style = clipboardStyleCss(run);
      const text = escapeClipboardHtml(run.text);
      return style ? `<span style="${style}">${text}</span>` : text;
    })
    .join("");
}

/**
 * A table's `text/html`: its rows and cells with the character and cell
 * formatting a web page or Office can take (bold, italic, underline, size,
 * colour, fill), not the tab-separated lines wrapped in `<p>` that lost
 * the table (PPT-B255).
 */
function presentationClipboardTableHtml(block: ClipboardBlock): string | null {
  if (!Array.isArray(block.rows)) return null;
  const rows = block.rows
    .map((row) => {
      if (!Array.isArray(row)) return "";
      const cells = row
        .map((cell) => {
          if (!isRecord(cell)) return "<td></td>";
          const style = [
            clipboardStyleCss(cell),
            typeof cell.backgroundColor === "string" && /^#[0-9a-f]{6}$/i.test(cell.backgroundColor)
              ? `background-color:${cell.backgroundColor}`
              : "",
            typeof cell.align === "string" && ["left", "center", "right", "justify"].includes(cell.align)
              ? `text-align:${cell.align}`
              : "",
          ]
            .filter(Boolean)
            .join(";");
          const attributes = [
            typeof cell.colspan === "number" && cell.colspan > 1 ? ` colspan="${cell.colspan}"` : "",
            typeof cell.rowspan === "number" && cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : "",
            style ? ` style="${style}"` : "",
          ].join("");
          const text = typeof cell.text === "string" ? cell.text : "";
          return `<td${attributes}>${clipboardRunsHtml(cell.runs, text)}</td>`;
        })
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");
  return `<table><tbody>${rows}</tbody></table>`;
}

/** Safe, inert `text/html` companion for interoperability with Office/Web. */
export function serializePresentationClipboardHtml(
  source: ReadonlyArray<ClipboardBlock> | null | undefined,
): string | null {
  const text = serializePresentationClipboardPlainText(source);
  if (text === null) return null;
  if (source && source.length === 1 && source[0].type === "table") {
    const table = presentationClipboardTableHtml(source[0]);
    if (table) return table;
  }
  return text
    .split("\n")
    .map((line) => `<p>${escapeClipboardHtml(line)}</p>`)
    .join("");
}

/** Convert a bounded standard `text/html` clipboard payload to inert text. */
export function presentationClipboardPlainTextFromHtml(
  html: unknown,
): string | null {
  if (typeof html !== "string" || html.length === 0 || html.length > 1_000_000) {
    return null;
  }
  if (typeof DOMParser === "undefined") return null;
  try {
    const document = new DOMParser().parseFromString(html, "text/html");
    for (const element of document.querySelectorAll(
      "script,style,template,iframe,object,embed",
    )) {
      element.remove();
    }
    // Preformatted text keeps its spaces and tabs (HTML white-space: pre):
    // they are held as private-use marks through the collapse below and
    // restored after it (PPT-B409: four spaces in a <pre> became one).
    const PRESERVED_SPACE = "\uE000";
    const PRESERVED_TAB = "\uE001";
    const preservesWhitespace = (element: Element | null): boolean => {
      for (let current = element; current; current = current.parentElement) {
        if (current.localName === "pre" || current.localName === "textarea") return true;
        const space = (current as HTMLElement).style?.whiteSpace;
        if (space) return /^(pre|pre-wrap|break-spaces)$/.test(space);
      }
      return false;
    };
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (preservesWhitespace(node.parentElement) && node.nodeValue) {
        node.nodeValue = node.nodeValue
          .replace(/ /g, PRESERVED_SPACE)
          .replace(/\t/g, PRESERVED_TAB);
      }
    }
    for (const lineBreak of document.querySelectorAll("br")) {
      lineBreak.replaceWith(document.createTextNode("\n"));
    }
    for (const block of document.querySelectorAll(
      "p,div,li,pre,blockquote,h1,h2,h3,h4,h5,h6,tr",
    )) {
      block.append(document.createTextNode("\n"));
    }
    for (const cell of document.querySelectorAll("th,td")) {
      cell.append(document.createTextNode("\t"));
    }
    const text = (document.body.textContent ?? "")
      .replace(/\r\n?/g, "\n")
      .replace(/[\t\f\v ]+/g, " ")
      .replace(/ *\n */g, "\n")
      .trim()
      .replace(/\uE000/g, " ")
      .replace(/\uE001/g, "\t");
    return text.trim().length > 0 && text.length <= 100_000 ? text : null;
  } catch {
    return null;
  }
}

/**
 * Validate every member with the persisted presentation union and retain only
 * the schema-parsed value. System payloads additionally refuse every
 * workspace/deck asset-custody field. All clipboard blocks must carry the full
 * positioned-object geometry contract before a caller can cast or commit them.
 */
function validateClipboardBlocks(
  source: ReadonlyArray<unknown>,
  options: { forbidSystemAssetCustody: boolean },
): ClipboardBlock[] | null {
  if (
    source.length === 0 ||
    source.length > MAX_PRESENTATION_CLIPBOARD_BLOCKS
  ) {
    return null;
  }
  const blocks: ClipboardBlock[] = [];
  const ids = new Set<string>();
  for (const candidate of source) {
    if (!isRecord(candidate)) return null;
    if (
      hasClipboardCustodyReceipt(candidate as ClipboardBlock) ||
      (options.forbidSystemAssetCustody &&
        hasSystemAssetCustody(candidate as ClipboardBlock))
    ) {
      return null;
    }
    const result = presentationBlockSchema.safeParse(candidate);
    if (!result.success) return null;
    const block = result.data as ClipboardBlock;
    if (
      !hasPositionedClipboardGeometry(block) ||
      ids.has(block.id) ||
      (options.forbidSystemAssetCustody &&
        (hasSystemAssetCustody(block) || !hasSelfContainedSystemImage(block)))
    ) {
      return null;
    }
    ids.add(block.id);
    blocks.push(block);
  }
  return canPreparePresentationPaste(blocks, "keep-source") ? blocks : null;
}

/**
 * Bound the untrusted JSON tree before it reaches the paste kernel. Full block
 * semantics remain owned by the presentation model; this guard proves that the
 * versioned envelope is plain JSON, finite in depth/node count, and cannot
 * carry prototype keys. Unknown envelope keys are rejected separately.
 */
function isBoundedClipboardJson(value: unknown): boolean {
  let nodes = 0;
  const visit = (member: unknown, depth: number): boolean => {
    nodes += 1;
    if (
      nodes > MAX_PRESENTATION_CLIPBOARD_NODES ||
      depth > MAX_PRESENTATION_CLIPBOARD_DEPTH
    ) {
      return false;
    }
    if (
      member === null ||
      typeof member === "boolean" ||
      typeof member === "string"
    ) {
      return true;
    }
    if (typeof member === "number") return Number.isFinite(member);
    if (Array.isArray(member)) {
      return member.every((child) => visit(child, depth + 1));
    }
    if (!isRecord(member)) return false;
    for (const [key, child] of Object.entries(member)) {
      if (key === "__proto__" || key === "prototype" || key === "constructor") {
        return false;
      }
      if (!visit(child, depth + 1)) return false;
    }
    return true;
  };
  return visit(value, 0);
}

/**
 * Whether an authenticated in-editor snapshot can execute a paste mode.
 *
 * Validation is all-or-nothing: duplicate ids, unknown block kinds, or a
 * mixed payload whose members have no textual projection for Text Only refuses
 * the complete operation. This is
 * deliberately stricter than silently dropping an unsupported member.
 */
export function canPreparePresentationPaste(
  source: ReadonlyArray<unknown> | null | undefined,
  mode: PresentationPasteMode,
): source is ReadonlyArray<ClipboardBlock> {
  if (!PRESENTATION_PASTE_MODES.has(mode)) return false;
  if (!Array.isArray(source) || source.length === 0) return false;
  const ids = new Set<string>();
  for (const block of source) {
    if (!isSupportedClipboardBlock(block) || ids.has(block.id)) return false;
    // The frame of an object the page could not display (TRIAGE 152) is no
    // object: copied, it would become a dashed box of the user's own.
    if (isObjectPlaceholder(block as { source?: unknown })) return false;
    if (
      mode === "text-only" &&
      presentationClipboardBlockPlainText(block) === null
    ) {
      return false;
    }
    ids.add(block.id);
  }
  return true;
}

/** A size in points as the design px an authored object states it in. */
function pointsAsDesignPx(value: unknown): unknown {
  return typeof value === "number" && Number.isFinite(value)
    ? Math.round(value * CSS_PX_PER_POINT * 1000) / 1000
    : value;
}

function materializeRunSizes(runs: unknown): void {
  if (!Array.isArray(runs)) return;
  for (const run of runs) {
    if (isRecord(run) && isRecord(run.style) && "fontSize" in run.style) {
      run.style.fontSize = pointsAsDesignPx(run.style.fontSize);
    }
  }
}









function keepPointFontSizes(block: ClipboardBlock): void {
  const record = block as Record<string, unknown>;
  if (record.type === "text" || record.type === "shape") {
    record.textFontSizeUnit = POINT_TEXT_UNIT;
    return;
  }
  materializePointFontSizes(block);
}

/** An object's point sizes restated as the design px an older authored object states. */
function materializePointFontSizes(block: ClipboardBlock): void {
  const record = block as Record<string, unknown>;
  if (isRecord(record.style) && "fontSize" in record.style) {
    record.style.fontSize = pointsAsDesignPx(record.style.fontSize);
  }
  materializeRunSizes(record.runs);
  if (Array.isArray(record.paragraphs)) {
    for (const paragraph of record.paragraphs) {
      if (isRecord(paragraph)) materializeRunSizes(paragraph.runs);
    }
  }
}

function materializeNativeThemePaint(block: ClipboardBlock): void {
  const paint = readPresentationNativeThemePaint(block);
  if (!paint) return;
  const record = block as Record<string, unknown>;
  for (const key of ["fill", "fillTo", "fillGradient", "gradientAngle", "stroke", "strokeWidth"] as const) {
    if (record[key] === undefined && paint[key] !== undefined) record[key] = paint[key];
  }
}

/**
 * Keep Source Formatting keeps a copied box's list indentation: each row of
 * an imported box takes its level's indent from the list style the package
 * gave the box (`source.nativeListIndents`, its `a:lstStyle` with what it
 * inherits), which goes with the box's identity. The rows that state none
 * are given their level's as their own (`marL` / `indent`, which the writer
 * writes as the row's `a:pPr`), so the copy's bullets and hanging text stand
 * where the original's do, on the canvas and in the file.
 */
function keepListIndents(block: ClipboardBlock): void {
  const record = block as Record<string, unknown>;
  if (!isRecord(record.source) || !Array.isArray(record.paragraphs)) return;
  const ladder = readNativeListIndents(record.source.nativeListIndents);
  if (!ladder) return;
  for (const paragraph of record.paragraphs) {
    if (!isRecord(paragraph) || paragraph.marL !== undefined || paragraph.indent !== undefined) continue;
    const level = typeof paragraph.level === "number" && Number.isInteger(paragraph.level)
      ? Math.max(0, Math.min(8, paragraph.level))
      : 0;
    const rung = ladder[level];
    if (!rung) continue;
    paragraph.marL = rung.marL;
    paragraph.indent = rung.indent;
  }
}

const PLAIN_PAINT = new Set([undefined, null, "none", "transparent"]);











function placeholderWordsAsTextBox(block: ClipboardBlock): ClipboardBlock {
  const record = block as Record<string, unknown>;
  if (record.type !== "shape") return block;
  const source = isRecord(record.source) ? record.source : null;
  const placeholder = [record.placeholder, record.placeholderType, record.role, source?.placeholderType]
    .some((value) => typeof value === "string" && value.trim() !== "");
  if (!placeholder) return block;
  if (record.shape !== undefined && record.shape !== "rectangle") return block;
  if (!PLAIN_PAINT.has(record.fill as string | null | undefined) || !PLAIN_PAINT.has(record.stroke as string | null | undefined)) return block;
  for (const key of ["fillGradient", "fillImage", "fillPattern", "fillTo", "customPath", "shapeEffects", "shadow", "glow", "softEdge", "bevel", "reflection", "shape3d"]) {
    if (record[key] !== undefined && record[key] !== null && record[key] !== false) return block;
  }
  if (typeof record.text !== "string" || record.text.trim() === "") return block;
  const { shape: _shape, fill: _fill, stroke: _stroke, strokeWidth: _strokeWidth, strokeDash: _strokeDash, ...rest } = record;
  void _shape; void _fill; void _stroke; void _strokeWidth; void _strokeDash;
  return { ...rest, type: "text", kind: "textbox" } as unknown as ClipboardBlock;
}

function stripImportedIdentity(block: ClipboardBlock): void {
  // `source` is the original-package patchback receipt. A pasted block is a
  // fresh object, never a second authority for the source XML node. Imported
  // order/z-index receipts are also discarded: append order makes the pasted
  // selection the new topmost group while reidAndOffsetBlocks preserves its
  // internal order. Authenticated same-editor memory keeps asset custody; the
  // separate system serializer removes it before bytes leave the workspace.
  // A copy of a file's placeholder is a plain object that looks the same,
  // as Duplicate writes it (`pptx/placeholder_copy.py`): a slide keeps one
  // placeholder per layout slot, and Layout / Reset / New Slide stay open on
  // the slide it lands on.
  if (hasOwn(block, "source")) {
    delete block.placeholder;
    delete block.placeholderType;
    delete block.placeholderIndex;
    delete block.role;
  }
  delete block.source;
  delete block.sourceOrder;
  delete block.zIndex;
}

/**
 * No Fill and No Outline are not colours the destination theme restates:
 * Use Destination Theme keeps them, as PowerPoint does, and only a stated
 * colour or face goes (PPT-B184: a No Fill shape came in solid light blue).
 */
function stripDestinationThemeKeys(value: unknown): void {
  if (!isRecord(value)) return;
  for (const key of DESTINATION_THEME_KEYS) {
    if (value[key] === "none" || value[key] === "transparent") continue;
    delete value[key];
  }
}

function stripDestinationThemeRuns(runs: unknown): void {
  if (!Array.isArray(runs)) return;
  for (const run of runs) if (isRecord(run)) stripDestinationThemeKeys(run.style);
}

/**
 * Use Destination Theme: the object's own paint and font overrides go, at
 * the levels that state them -- the object, its text style, its runs, its
 * table cells -- so it takes the destination's. It used to strip those keys
 * from EVERY nested object, which emptied structures that need them (a
 * WordArt paint lost its fill, a gradient its stop colours) and the pasted
 * WordArt crashed the editor (PPT-B122). Structured paints are kept whole.
 */
function applyDestinationTheme(block: ClipboardBlock): void {
  const record = block as Record<string, unknown>;
  stripDestinationThemeKeys(record);
  stripDestinationThemeKeys(record.style);
  stripDestinationThemeRuns(record.runs);
  if (Array.isArray(record.paragraphs)) {
    for (const paragraph of record.paragraphs) {
      if (isRecord(paragraph)) stripDestinationThemeRuns(paragraph.runs);
    }
  }
  if (Array.isArray(record.rows)) {
    for (const row of record.rows) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        if (!isRecord(cell)) continue;
        stripDestinationThemeKeys(cell);
        stripDestinationThemeRuns(cell.runs);
      }
    }
  }
}

/**
 * Deep-clone and normalize a clipboard snapshot for one Paste command.
 *
 * The returned blocks have no original-package source receipt. Keep Source
 * otherwise retains authored appearance; Destination Theme removes explicit
 * theme paint/font overrides; Picture keeps the safe source snapshot for the
 * caller's one-shot raster transaction; Text Only delegates construction of a
 * plain text box to the caller. The
 * source snapshot is never mutated.
 */
export function preparePresentationPasteBlocks<T extends ClipboardBlock>(
  source: ReadonlyArray<T> | null | undefined,
  mode: PresentationPasteMode,
  toTextOnlyBlock?: (block: T, plainText: string) => T,
  options: {
    /** Whether an object's text sizes are points (an imported object's). */
    readonly textUsesPoints?: (block: T) => boolean;
  } = {},
): T[] | null {
  if (!canPreparePresentationPaste(source, mode)) return null;
  let clones: T[];
  try {
    clones = JSON.parse(JSON.stringify(source)) as T[];
  } catch {
    return null;
  }
  if (!canPreparePresentationPaste(clones, mode)) return null;

  // Every path strips the identity that made an object's sizes points, so
  // every path keeps them: Copy's own snapshot (`serializePresentationBlockClipboard`)
  // named no reader, and a 27 pt imported box pasted from it drew and saved
  // at 27 px (20.25 pt).
  const textUsesPoints = options.textUsesPoints ?? ((block: T) => presentationTextUsesPoints(block as unknown as PresentationBlock));
  for (let index = 0; index < clones.length; index += 1) {
    let block = clones[index];
    if (textUsesPoints(block)) keepPointFontSizes(block);
    // Keep Source Formatting keeps what the original DREW: the theme paint its
    // native identity carried (a theme gradient) becomes the copy's own paint
    // before the identity goes, else the copy fell back to the default solid
    // light blue (PPT-B377). Use Destination Theme takes the destination's.
    if (mode === "keep-source") {
      materializeNativeThemePaint(block);
      keepListIndents(block);
    }
    if (mode === "keep-source" || mode === "destination-theme") {
      block = placeholderWordsAsTextBox(block) as T;
      clones[index] = block;
    }
    stripImportedIdentity(block);
    if (mode === "destination-theme") applyDestinationTheme(block);
    if (mode === "text-only") {
      const plainText = presentationClipboardBlockPlainText(block);
      if (!toTextOnlyBlock || plainText === null) return null;
      try {
        block = toTextOnlyBlock(block, plainText);
      } catch {
        return null;
      }
      if (!isSupportedClipboardBlock(block) || block.type !== "text") {
        return null;
      }
      stripImportedIdentity(block);
      clones[index] = block;
    }
  }
  return canPreparePresentationPaste(clones, mode) ? clones : null;
}

/**
 * Serialize an authenticated live selection into the sole supported clipboard
 * envelope. The serialized snapshot is independent of live state and has every
 * original-package/order custody receipt removed before it can leave the
 * editor. The system envelope additionally removes the complete asset custody
 * tuple while retaining a schema-bounded inline `src`; groups and connector
 * endpoints survive for the existing paste remap kernel.
 */
export function serializePresentationBlockClipboard(
  source: ReadonlyArray<ClipboardBlock> | null | undefined,
): string | null {
  const blocks = preparePresentationPasteBlocks(source, "keep-source");
  if (!blocks || blocks.length > MAX_PRESENTATION_CLIPBOARD_BLOCKS) return null;
  const validated = validateClipboardBlocks(blocks, {
    forbidSystemAssetCustody: false,
  });
  if (!validated) return null;
  for (const block of validated) stripSystemAssetCustody(block);
  const systemBlocks = validateClipboardBlocks(validated, {
    forbidSystemAssetCustody: true,
  });
  if (!systemBlocks) return null;
  const envelope: PresentationBlockClipboardEnvelopeV1 = {
    kind: PRESENTATION_BLOCK_CLIPBOARD_KIND,
    version: PRESENTATION_BLOCK_CLIPBOARD_VERSION,
    blocks: systemBlocks,
  };
  if (!isBoundedClipboardJson(envelope)) return null;
  try {
    const serialized = JSON.stringify(envelope);
    return clipboardUtf8ByteLength(serialized) <=
      PRESENTATION_BLOCK_CLIPBOARD_MAX_BYTES
      ? serialized
      : null;
  } catch {
    return null;
  }
}

/**
 * Parse a Web Custom Format payload. Admission is all-or-nothing: exact root
 * keys/kind/version, bounded plain JSON, supported unique blocks, and no source
 * custody receipts. The returned array is a new deep snapshot ready for the
 * normal paste planner; malformed/future envelopes never fall through as an
 * editable partial selection.
 */
export function parsePresentationBlockClipboard(
  serialized: string | null | undefined,
): ClipboardBlock[] | null {
  if (
    typeof serialized !== "string" ||
    serialized.length === 0 ||
    clipboardUtf8ByteLength(serialized) >
      PRESENTATION_BLOCK_CLIPBOARD_MAX_BYTES
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(serialized);
  } catch {
    return null;
  }
  if (!isRecord(parsed) || !isBoundedClipboardJson(parsed)) return null;
  const rootKeys = Object.keys(parsed).sort();
  if (
    rootKeys.length !== 3 ||
    rootKeys[0] !== "blocks" ||
    rootKeys[1] !== "kind" ||
    rootKeys[2] !== "version" ||
    parsed.kind !== PRESENTATION_BLOCK_CLIPBOARD_KIND ||
    parsed.version !== PRESENTATION_BLOCK_CLIPBOARD_VERSION ||
    !Array.isArray(parsed.blocks)
  ) {
    return null;
  }
  const blocks = validateClipboardBlocks(parsed.blocks, {
    forbidSystemAssetCustody: true,
  });
  return blocks
    ? preparePresentationPasteBlocks(blocks, "keep-source")
    : null;
}







function isSourcelessClipboardMedia(block: ClipboardBlock): boolean {
  if (block.type !== "embed") return false;
  if (block.kind !== "video" && block.kind !== "audio") return false;
  const url = block.url;
  return !(typeof url === "string" && url.trim() !== "");
}

/**
 * Split a parsed system-clipboard selection into what can be pasted and the
 * count of video/audio objects that lost their clip on the way. A group left
 * with a single member is ungrouped.
 */
export function withoutSourcelessClipboardMedia<T extends ClipboardBlock>(
  blocks: ReadonlyArray<T>,
): { readonly blocks: T[]; readonly droppedMedia: number } {
  const kept = blocks.filter((block) => !isSourcelessClipboardMedia(block));
  const droppedMedia = blocks.length - kept.length;
  if (droppedMedia === 0) return { blocks: kept, droppedMedia };
  const groupSizes = new Map<string, number>();
  for (const block of kept) {
    const gid = block.groupId;
    if (typeof gid === "string" && gid) {
      groupSizes.set(gid, (groupSizes.get(gid) ?? 0) + 1);
    }
  }
  const result = kept.map((block) => {
    const gid = block.groupId;
    if (typeof gid !== "string" || !gid || (groupSizes.get(gid) ?? 0) > 1) {
      return block;
    }
    const ungrouped: ClipboardBlock = { ...block };
    delete ungrouped.groupId;
    return ungrouped as T;
  });
  return { blocks: result, droppedMedia };
}

/** Options for {@link reidAndOffsetBlocks}. */
export interface ReidOffsetOptions {
  /** EMU nudge applied to `left`/`top` so a clone is visibly offset from its source. */
  offsetEmu: number;
  /** A different `top` nudge (EMU); `offsetEmu` when omitted. Duplicate's
   *  smart offset moves right and down by different amounts (PPT-E006). */
  offsetYEmu?: number;
  /** Slide width (EMU); a clone's `left` is clamped so the block stays on-canvas. */
  maxLeft: number;
  /** Slide height (EMU); a clone's `top` is clamped so the block stays on-canvas. */
  maxTop: number;
  /** Fresh-id factory (component injects `crypto.randomUUID`). */
  makeId: () => string;
  /** Fresh group-id factory (component injects `newGroupId`). */
  makeGroupId: () => string;
}

/** Read a finite number off a record by key, else `null`. */
function finiteNumber(record: Record<string, unknown>, key: string): number | null {
  const value = record[key];
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/**
 * Read a non-empty string group id off a block (the `groupId` field), else
 * `undefined`. Mirrors `blockGroupId` from `presentation-crop-group` but inlined
 * here to keep this kernel free of that module's wider surface.
 */
function blockGroupIdOf(record: Record<string, unknown>): string | undefined {
  const value = record["groupId"];
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

/**
 * Clone a snapshot of blocks into INDEPENDENT copies with fresh ids, remapped
 * intra-snapshot references, and an offset+clamped geometry. The exact behaviour
 * (matching the editor's right-click Copy/Paste):
 *
 *  - **Fresh ids.** Every block's `id` is replaced via `makeId()`; an old→new map
 *    is built first so references below can be remapped onto the new ids.
 *  - **Group remap.** Each distinct source `groupId` maps to ONE fresh id (via
 *    `makeGroupId()`), so a multi-block group stays a single group on the clones
 *    and never collides with the source group. Ungrouped blocks stay ungrouped.
 *  - **Connector anchors.** `startShapeId` / `endShapeId` that point at another
 *    block IN the snapshot are remapped onto that block's new id; a ref pointing
 *    OUTSIDE the snapshot is DELETED (key removed, not set to `undefined`) so a
 *    partially-copied connector doesn't dangle onto a source-slide shape.
 *  - **Offset + clamp.** `left`/`top` (when present and finite) are nudged by
 *    `offsetEmu` and clamped to `[0, max - size]` so the clone lands on-canvas,
 *    near the source. A block missing `left`/`top` (a non-positioned block) is
 *    left where it is.
 *
 * Pure: the input array and its blocks are NOT mutated — the caller passes a
 * deep (JSON) clone, and we operate on those clone objects in place and return
 * them. Deterministic given deterministic `makeId` / `makeGroupId`.
 */
export function reidAndOffsetBlocks<T extends ClipboardBlock>(
  clones: ReadonlyArray<T>,
  options: ReidOffsetOptions,
): T[] {
  const { offsetEmu, maxLeft, maxTop, makeId, makeGroupId } = options;
  const offsetYEmu = options.offsetYEmu ?? offsetEmu;
  // old id -> new id, for connector-anchor remapping. Built up-front across ALL
  // blocks so a connector listed before its endpoint still remaps correctly.
  const idMap = new Map<string, string>();
  for (const block of clones) {
    if (typeof block.id === "string") idMap.set(block.id, makeId());
  }
  // ONE offset for the whole set, clamped so the set's union lands on the
  // canvas: clamping each clone on its own pinned an edge member while the
  // others moved, squeezing a duplicated group's inner spacing (PPT-B035).
  // For a single clone this is exactly the old per-block clamp.
  const commonDelta = (
    startKey: "left" | "top",
    sizeKey: "width" | "height",
    max: number,
    offset: number,
  ): number => {
    let low = Infinity;
    let high = -Infinity;
    for (const block of clones) {
      const record = block as Record<string, unknown>;
      const start = finiteNumber(record, startKey);
      if (start === null) continue;
      low = Math.min(low, start);
      high = Math.max(high, start + (finiteNumber(record, sizeKey) ?? 0));
    }
    if (!Number.isFinite(low)) return offset;
    return Math.max(-low, Math.min(offset, max - high));
  };
  const dx = commonDelta("left", "width", maxLeft, offsetEmu);
  const dy = commonDelta("top", "height", maxTop, offsetYEmu);
  // old groupId -> one fresh groupId, so a multi-block group stays one group.
  const groupMap = new Map<string, string>();
  for (const block of clones) {
    const record = block as Record<string, unknown>;
    // Fresh id (fall back to a new id if somehow unmapped — never reuse the old).
    record.id = idMap.get(block.id) ?? makeId();
    // Remap group id to a fresh, stable id per source group.
    const freshGroupId = (gid: string): string => {
      let mapped = groupMap.get(gid);
      if (!mapped) {
        mapped = makeGroupId();
        groupMap.set(gid, mapped);
      }
      return mapped;
    };
    const gid = blockGroupIdOf(record);
    if (gid) {
      record.groupId = freshGroupId(gid);
      // The groups nested inside it come along, each as one fresh group.
      if (Array.isArray(record.groupPath)) {
        const path = record.groupPath.filter(
          (entry): entry is string => typeof entry === "string" && entry.length > 0,
        );
        if (path.length) record.groupPath = path.map(freshGroupId);
        else delete record.groupPath;
      }
    } else {
      delete record.groupPath;
    }
    // Remap connector anchors onto pasted endpoints; drop a ref outside the set.
    for (const key of ["startShapeId", "endShapeId"] as const) {
      const ref = record[key];
      if (typeof ref === "string") {
        const mappedRef = idMap.get(ref);
        if (mappedRef) record[key] = mappedRef;
        else delete record[key];
      }
    }
    // An action button copied with its player plays the pasted player.
    const action = record.clickAction as { targetBlockId?: unknown } | undefined;
    if (
      action &&
      typeof action === "object" &&
      typeof action.targetBlockId === "string" &&
      idMap.has(action.targetBlockId)
    ) {
      record.clickAction = { ...action, targetBlockId: idMap.get(action.targetBlockId) };
    }
    // Offset by the set's common delta so the clones land on-canvas, near the
    // source, keeping their layout.
    const left = finiteNumber(record, "left");
    const top = finiteNumber(record, "top");
    if (left !== null) record.left = left + dx;
    if (top !== null) record.top = top + dy;
  }
  return clones as T[];
}

/**
 * The ids of every POSITIONED (geometry-bearing) block on a slide, in slide
 * order — the target set for Cmd/Ctrl+A (Select All). A block is "positioned"
 * when it carries a finite `left` OR `top` (same rule as the editor's
 * `hasLayoutGeometry`), since the multi-select set only ever holds blocks with a
 * rect (align / distribute / nudge act on geometry). Non-positioned body blocks
 * are excluded so selecting all and then nudging / aligning stays well-defined.
 *
 * Pure: no DOM, inputs never mutated. A `null`/empty list yields an empty set.
 */
export function selectAllBlockIds(
  blocks: ReadonlyArray<ClipboardBlock | null | undefined> | null | undefined,
): Set<string> {
  const ids = new Set<string>();
  if (!Array.isArray(blocks)) return ids;
  for (const block of blocks) {
    if (!block || typeof block.id !== "string") continue;
    const record = block as Record<string, unknown>;
    if (finiteNumber(record, "left") !== null || finiteNumber(record, "top") !== null) {
      ids.add(block.id);
    }
  }
  return ids;
}

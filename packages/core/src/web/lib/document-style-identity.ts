/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { DocumentBuiltinStyleId } from "../components/console/editor/documentStyleGallery";

type StyleRecord = Record<string, unknown>;
const record = (value: unknown): StyleRecord | null =>
  value && typeof value === "object" && !Array.isArray(value)
    ? (value as StyleRecord)
    : null;
const token = (value: unknown): string | null =>
  typeof value === "string" &&
  value.trim().length > 0 &&
  value.length <= 256 &&
  !/[\u0000-\u001f]/.test(value)
    ? value.trim()
    : null;

/** OOXML IDs are opaque; a localized display name can identify the built-in. */
export function wordHeadingLevelFromToken(value: unknown): number | null {
  const match = token(value)?.match(/^(?:heading|标题|標題)[\s_]*([1-6])$/i);
  return match ? Number(match[1]) : null;
}

export function wordBuiltinStyleFromToken(
  value: unknown,
): DocumentBuiltinStyleId | null {
  const heading = wordHeadingLevelFromToken(value);
  if (heading) return `heading${heading}` as DocumentBuiltinStyleId;
  const name = token(value)
    ?.replace(/[\s_]+/g, "")
    .toLowerCase();
  const aliases: Record<string, DocumentBuiltinStyleId> = {
    normal: "normal",
    正文: "normal",
    nospacing: "noSpacing",
    无间隔: "noSpacing",
    無間隔: "noSpacing",
    title: "title",
    标题: "title",
    標題: "title",
    subtitle: "subtitle",
    副标题: "subtitle",
    副標題: "subtitle",
    quote: "quote",
    引用: "quote",
    intensequote: "intenseQuote",
    明显引用: "intenseQuote",
    明顯引用: "intenseQuote",
    listparagraph: "listParagraph",
    列表段落: "listParagraph",
    subtleemphasis: "subtleEmphasis",
    轻微着重: "subtleEmphasis",
    輕微著重: "subtleEmphasis",
    emphasis: "emphasis",
    着重: "emphasis",
    著重: "emphasis",
    intenseemphasis: "intenseEmphasis",
    明显着重: "intenseEmphasis",
    strong: "strong",
    强调: "strong",
    強調: "strong",
    subtlereference: "subtleReference",
    轻微参考: "subtleReference",
    intensereference: "intenseReference",
    明显参考: "intenseReference",
    booktitle: "bookTitle",
    书籍标题: "bookTitle",
    書籍標題: "bookTitle",
  };
  return name ? (aliases[name] ?? null) : null;
}

/** Duplicated IDs cannot establish a unique source style identity. */
function indexStyles(values: unknown): Map<string, StyleRecord> {
  const result = new Map<string, StyleRecord>(),
    duplicates = new Set<string>();
  if (!Array.isArray(values) || values.length > 512) return result;
  for (const value of values) {
    const style = record(value),
      id = token(style?.styleId);
    if (!style || !id || duplicates.has(id)) continue;
    if (result.has(id)) {
      result.delete(id);
      duplicates.add(id);
    } else result.set(id, style);
  }
  return result;
}

/** A paragraph's direct outline level belongs to TOC inclusion, not its style. */
export function wordStyleHeadingLevels(
  values: unknown,
): Map<string, number | null> {
  const styles = indexStyles(values),
    resolved = new Map<string, number | null>();
  for (const id of styles.keys()) {
    let current: string | null = id,
      level = wordHeadingLevelFromToken(id);
    const visited = new Set<string>();
    while (current && !visited.has(current)) {
      visited.add(current);
      const style = styles.get(current);
      if (!style || style.type !== "paragraph") break;
      level ??=
        wordHeadingLevelFromToken(style.styleId) ??
        wordHeadingLevelFromToken(style.name);
      const outline = record(style.pPr)?.outlineLvl;
      if (
        typeof outline === "number" &&
        Number.isInteger(outline) &&
        outline >= 0 &&
        outline <= 9
      ) {
        level = outline <= 5 ? outline + 1 : null;
        break;
      }
      current = token(style.basedOn);
    }
    if (styles.get(id)?.type === "paragraph") resolved.set(id, level);
  }
  return resolved;
}

export interface ImportedDocumentStyleIdentity {
  styleId: string;
  name: string;
  builtin: DocumentBuiltinStyleId | null;
}
const styleIndexes = new WeakMap<object, Map<string, StyleRecord>>();
/** Read identity only: recognizing a style must never reapply its default paint. */
export function importedDocumentStyleIdentity(
  definitions: unknown,
  styleId: unknown,
  kind: "paragraph" | "character",
): ImportedDocumentStyleIdentity | null {
  const id = token(styleId);
  if (!id) return null;
  const container = record(definitions);
  let styles = container ? styleIndexes.get(container) : undefined;
  if (!styles) {
    styles = indexStyles(container?.version === 1 ? container.styles : []);
    if (container) styleIndexes.set(container, styles);
  }
  const style = styles.get(id);
  if (style && style.type !== kind) return null;
  const name = token(style?.name) ?? id;
  // A custom style keeps its own name even when it inherits from Heading 1.
  const builtin =
    style?.custom === true || (container?.version === 1 && !style)
      ? null
      : (wordBuiltinStyleFromToken(name) ?? wordBuiltinStyleFromToken(id));
  const paragraph =
    builtin &&
    [
      "normal",
      "noSpacing",
      "title",
      "subtitle",
      "quote",
      "intenseQuote",
      "listParagraph",
      "heading1",
      "heading2",
      "heading3",
      "heading4",
      "heading5",
      "heading6",
    ].includes(builtin);
  return {
    styleId: id,
    name,
    builtin:
      builtin && (kind === "paragraph") === Boolean(paragraph) ? builtin : null,
  };
}

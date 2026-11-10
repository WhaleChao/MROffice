/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/core";
import type { Mark, Node as ProseMirrorNode } from "@tiptap/pm/model";

/**
 * Simplified <-> Traditional Chinese conversion (简繁转换) for the native
 * document editor — the Word/WPS Writer "Review → 简转繁 / 繁转简" feature.
 *
 * This is a pure EDITOR TRANSFORM: it rewrites the text of the affected text
 * nodes in place through OpenCC and changes nothing else. Marks (bold, links,
 * font, …) and block structure are preserved because every text node is
 * replaced by a fresh text node carrying that node's own marks, and only text
 * leaves are touched (inline atoms such as footnotes / inline-math / citations
 * carry no characters and are skipped). There is no schema or serialization
 * change, so a document that never uses the feature serializes byte-for-byte
 * unchanged.
 *
 * DEFER: per-region phrasing variants (zh-TW 軟體 vs zh-HK 軟件, etc.). We use
 * OpenCC's standard Simplified(cn) <-> Traditional(tw) character mapping, which
 * is the neutral default for s2t/t2s.
 */
export type ChineseConversionDirection = "s2t" | "t2s";

/** OpenCC converter signature: a pure (text) => convertedText mapping. */
export type ChineseConverter = (text: string) => string;

// OpenCC's `from`/`to` locale codes for the two standard directions. `cn` is
// Simplified (Mainland); `tw` is Traditional (Taiwan, no regional phrase
// substitution beyond the standard character set — see DEFER above).
const CONVERSION_OPTIONS: Record<
  ChineseConversionDirection,
  { from: string; to: string }
> = {
  s2t: { from: "cn", to: "tw" },
  t2s: { from: "tw", to: "cn" },
};

// Minimal shape of the part of opencc-js we depend on. The package exposes
// `Converter` both as a named export and on its default export; we accept
// either so the lazy import is resilient to ESM/CJS interop.
interface OpenCCModuleLike {
  Converter?: (options: { from: string; to: string }) => ChineseConverter;
  default?: {
    Converter?: (options: { from: string; to: string }) => ChineseConverter;
  };
}

function resolveConverterFactory(
  opencc: OpenCCModuleLike,
): ((options: { from: string; to: string }) => ChineseConverter) | null {
  if (typeof opencc.Converter === "function") return opencc.Converter;
  if (typeof opencc.default?.Converter === "function") {
    return opencc.default.Converter;
  }
  return null;
}

// opencc-js ships a sizeable dictionary bundle, so it is loaded lazily on first
// use via a dynamic import (code-split) and the built converters are cached.
// `loadConverter` lets tests inject a stub instead of importing the real module.
export type ChineseConverterLoader = (
  direction: ChineseConversionDirection,
) => Promise<ChineseConverter>;

const converterCache = new Map<
  ChineseConversionDirection,
  Promise<ChineseConverter>
>();
/** The converters whose dictionaries have finished loading. */
const loadedConverters = new Map<ChineseConversionDirection, ChineseConverter>();

/**
 * The converter for a direction when its dictionary is already loaded, else
 * null: what a synchronous caller (the AI's command host, which loads the
 * dictionaries before it runs a batch) converts with.
 */
export function loadedChineseConverter(
  direction: ChineseConversionDirection,
): ChineseConverter | null {
  return loadedConverters.get(direction) ?? null;
}

const defaultLoader: ChineseConverterLoader = async (direction) => {
  const opencc = (await import("opencc-js")) as OpenCCModuleLike;
  const factory = resolveConverterFactory(opencc);
  if (!factory) {
    throw new Error("opencc-js: Converter factory not found");
  }
  return factory(CONVERSION_OPTIONS[direction]);
};

/**
 * Resolve (and cache) the OpenCC converter for a direction. The promise is
 * memoized per direction so the dictionary is parsed at most once. A failed
 * load is not cached, so a later attempt can retry.
 */
export function getChineseConverter(
  direction: ChineseConversionDirection,
  loader: ChineseConverterLoader = defaultLoader,
): Promise<ChineseConverter> {
  const cached = converterCache.get(direction);
  if (cached) return cached;
  const pending = loader(direction).then(
    (converter) => {
      loadedConverters.set(direction, converter);
      return converter;
    },
    (error) => {
      converterCache.delete(direction);
      throw error;
    },
  );
  converterCache.set(direction, pending);
  return pending;
}

/** Test-only: drop the memoized converters so each case starts clean. */
export function __resetChineseConverterCache(): void {
  converterCache.clear();
  loadedConverters.clear();
}

interface TextNodeSpan {
  from: number;
  to: number;
  text: string;
  marks: readonly Mark[];
}

/**
 * Collect the text leaves inside [from, to) as {from, to, text, marks} spans.
 * Only `isText` nodes contribute (inline atoms have no text and are left
 * untouched). Positions are absolute document positions usable directly with
 * `tr.replaceWith`.
 */
function collectTextSpans(
  doc: ProseMirrorNode,
  from: number,
  to: number,
): TextNodeSpan[] {
  const spans: TextNodeSpan[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText || typeof node.text !== "string" || node.text.length === 0) {
      return true;
    }
    // Clamp to the requested range so a partially-selected text node only has
    // its selected slice converted (e.g. a half-highlighted word).
    const nodeFrom = Math.max(pos, from);
    const nodeTo = Math.min(pos + node.nodeSize, to);
    if (nodeTo <= nodeFrom) return false;
    const sliceStart = nodeFrom - pos;
    const sliceEnd = nodeTo - pos;
    spans.push({
      from: nodeFrom,
      to: nodeTo,
      text: node.text.slice(sliceStart, sliceEnd),
      marks: node.marks,
    });
    return false;
  });
  return spans;
}

/**
 * Convert the text of the editor's current selection (or the whole document
 * when the selection is empty/collapsed) through `convert`, in a SINGLE
 * transaction. Each affected text node is replaced by a fresh text node with
 * the converted text and the original node's marks, so formatting and document
 * structure are preserved — only character content changes.
 *
 * Spans are replaced back-to-front so a length change (rare, but OpenCC can map
 * 1->N characters) never invalidates an earlier span's position. Returns true
 * when at least one character actually changed (and a transaction was
 * dispatched), false otherwise — so callers can no-op cleanly on, e.g., a
 * selection of Latin text.
 */
export function applyChineseConversionToSelection(
  editor: Editor,
  convert: ChineseConverter,
): boolean {
  if (editor.isDestroyed || !editor.isEditable) return false;
  const { state } = editor;
  const { selection, doc, schema } = state;
  const useWholeDoc = selection.empty;
  const from = useWholeDoc ? 0 : selection.from;
  const to = useWholeDoc ? doc.content.size : selection.to;
  if (to <= from) return false;

  const spans = collectTextSpans(doc, from, to);
  if (spans.length === 0) return false;

  let transaction = state.tr;
  let changed = false;
  // Back-to-front: replacing a later span first keeps every earlier span's
  // absolute positions valid even if the replacement length differs.
  for (let index = spans.length - 1; index >= 0; index -= 1) {
    const span = spans[index];
    const converted = convert(span.text);
    if (converted === span.text) continue;
    const replacement =
      converted.length > 0 ? schema.text(converted, span.marks) : null;
    transaction = replacement
      ? transaction.replaceWith(span.from, span.to, replacement)
      : transaction.delete(span.from, span.to);
    changed = true;
  }

  if (!changed) return false;
  editor.view.dispatch(transaction.scrollIntoView());
  return true;
}

/**
 * Convenience entry used by the toolbar: lazily load the converter for the
 * direction, then apply it to the editor selection. Resolves to whether any
 * text changed. Loading failures reject and are surfaced to the caller (which
 * simply leaves the document untouched).
 */
export async function convertSelectionChinese(
  editor: Editor,
  direction: ChineseConversionDirection,
  {
    loader = defaultLoader,
    mayApply = () => true,
  }: {
    loader?: ChineseConverterLoader;
    mayApply?: () => boolean;
  } = {},
): Promise<boolean> {
  if (editor.isDestroyed || !editor.isEditable || !mayApply()) return false;
  const invocationDocument = editor.state.doc;
  const invocationSelection = editor.state.selection;
  const convert = await getChineseConverter(direction, loader);
  if (
    editor.isDestroyed ||
    !editor.isEditable ||
    !mayApply() ||
    editor.state.doc !== invocationDocument ||
    !editor.state.selection.eq(invocationSelection)
  ) {
    return false;
  }
  return applyChineseConversionToSelection(editor, convert);
}

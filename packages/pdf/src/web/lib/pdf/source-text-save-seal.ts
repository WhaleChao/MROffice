/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { sha256PdfBytes } from "./engine/native/transaction-error";
import { copyRememberedPdfSourceBytes } from "./source-bytes-memo";
import {
  MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS,
  cachedPdfSourceTextBoxAppearanceSeals,
  pdfSourceTextBindingSealedForSave,
  readPdfSourceTextBoxAppearanceSeals,
  readPdfSourceTextBoxBinding,
  type PdfSourceTextBoxBinding,
} from "./source-text-box";

type Content = Record<string, unknown>;

const LAYOUT_PREFIX = "[[MRNOTE_NATIVE_TEXT_BOX_V5|";
const SOURCE_PREFIX = "[[MRNOTE_NATIVE_TEXT_BOX_V4|";

interface SourceRow {
  readonly binding: PdfSourceTextBoxBinding;
  /** The row restores the source text (a box filled or deleted). */
  readonly restore: boolean;
  readonly tokenStart: number;
  readonly tokenEnd: number;
}

function bytesToBase64Url(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return globalThis.btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/u, "");
}

/** A managed text-box row's source binding, where the row carries one: the
 * binding token follows the row plan's (V5) when there is one. */
function readSourceRow(text: unknown): SourceRow | null {
  if (typeof text !== "string") return null;
  let tokenStart: number;
  if (text.startsWith(SOURCE_PREFIX)) {
    tokenStart = SOURCE_PREFIX.length;
  } else if (text.startsWith(LAYOUT_PREFIX)) {
    const layoutEnd = text.indexOf("|", LAYOUT_PREFIX.length);
    if (layoutEnd < 0 || !text.startsWith(SOURCE_PREFIX.slice(2), layoutEnd + 1)) return null;
    tokenStart = layoutEnd + 1 + SOURCE_PREFIX.length - 2;
  } else {
    return null;
  }
  const tokenEnd = text.indexOf("|", tokenStart);
  if (tokenEnd <= tokenStart || tokenEnd - tokenStart > MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS) return null;
  const token = text.slice(tokenStart, tokenEnd);
  if (!/^[A-Za-z0-9_-]+$/u.test(token)) return null;
  try {
    const payload = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(
      Uint8Array.from(globalThis.atob(token.replaceAll("-", "+").replaceAll("_", "/")), (ch) => ch.charCodeAt(0)),
    )) as { restore?: unknown } | null;
    const binding = readPdfSourceTextBoxBinding(payload);
    return binding ? { binding, restore: payload?.restore === true, tokenStart, tokenEnd } : null;
  } catch {
    return null;
  }
}

/** The row with its binding stating `seal`; the row itself when it carries
 * no binding, states that seal already, or cannot state it. */
function sealedSourceRowText(text: string, row: SourceRow, seal: string | null): string {
  const sealed = pdfSourceTextBindingSealedForSave(row.binding, seal);
  if (sealed.expectedAppearance === row.binding.expectedAppearance) return text;
  // As the row's own writer spells it (nativeTextBoxDirective).
  const token = bytesToBase64Url(new TextEncoder().encode(JSON.stringify({
    ...sealed, ...(row.restore ? { restore: true } : {}),
  })));
  if (token.length > MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS) return text;
  return `${text.slice(0, row.tokenStart)}${token}${text.slice(row.tokenEnd)}`;
}

function rowsOf(content: Content | null | undefined): readonly unknown[] {
  return Array.isArray(content?.annotations) ? content.annotations as unknown[] : [];
}

function rowIdAndText(row: unknown): { id: string; text: string } | null {
  if (!row || typeof row !== "object" || Array.isArray(row)) return null;
  const { id, text } = row as { id?: unknown; text?: unknown };
  return typeof id === "string" && id && typeof text === "string" ? { id, text } : null;
}

export interface PdfSourceTextSaveRows {
  /** The page as the file the save is written on was written: a row that
   * reads there as it reads in the save is not written again (the server
   * leaves its box as it is), and keeps the seal it was written with. */
  readonly written: Content | null | undefined;
  /** The page's last write was refused: its JSON may be ahead of the file,
   * so every row may be one the file does not hold yet. */
  readonly every: boolean;
}

function sourceRowsWritten(content: Content, rows: PdfSourceTextSaveRows): Map<number, SourceRow> {
  // (Most pages hold no such box: nothing is read of them but each row's first characters.)
  const candidates = new Map<number, { id: string; text: string }>();
  rowsOf(content).forEach((row, index) => {
    const read = rowIdAndText(row);
    if (read && (read.text.startsWith(SOURCE_PREFIX) || read.text.startsWith(LAYOUT_PREFIX))) candidates.set(index, read);
  });
  const found = new Map<number, SourceRow>();
  if (!candidates.size) return found;
  const written = new Map<string, string>();
  if (!rows.every) {
    for (const row of rowsOf(rows.written)) {
      const read = rowIdAndText(row);
      if (read) written.set(read.id, read.text);
    }
  }
  for (const [index, read] of candidates) {
    if (written.get(read.id) === read.text) continue;
    const source = readSourceRow(read.text);
    if (source) found.set(index, source);
  }
  return found;
}

/** The bindings of the source-line boxes a save writes: those whose seals
 * the save has to read from the file it is written on. */
export function pdfSourceTextBindingsWrittenBySave(
  content: Content,
  rows: PdfSourceTextSaveRows,
): PdfSourceTextBoxBinding[] {
  return [...sourceRowsWritten(content, rows).values()].map((row) => row.binding);
}

/**
 * `content` as the save sends it: every source-line box's row the save
 * writes states the seal `seals` holds for its box (the file's; null where
 * the file holds no such box yet), and so does the box's own entry. A box
 * `seals` does not name is left as it is. `content` itself where nothing
 * changes.
 */
export function sealPdfSourceTextRowsForSave(
  content: Content,
  rows: PdfSourceTextSaveRows,
  seals: ReadonlyMap<string, string | null>,
): Content {
  const sealedRows = new Map<number, unknown>();
  const sealedBoxes = new Map<string, { page: number; seal: string | null }>();
  for (const [index, source] of sourceRowsWritten(content, rows)) {
    if (!seals.has(source.binding.boxId)) continue;
    const seal = seals.get(source.binding.boxId) ?? null;
    const row = rowsOf(content)[index] as { text: string };
    const text = sealedSourceRowText(row.text, source, seal);
    if (text === row.text) continue;
    sealedRows.set(index, { ...row, text });
    sealedBoxes.set(source.binding.boxId, { page: source.binding.page, seal });
  }
  if (!sealedRows.size) return content;
  const next: Content = { ...content, annotations: rowsOf(content).map((row, index) => sealedRows.get(index) ?? row) };
  for (const key of ["textBoxEdits", "text_box_edits"] as const) {
    const edits = content[key];
    if (!Array.isArray(edits)) continue;
    next[key] = edits.map((edit: unknown) => {
      if (!edit || typeof edit !== "object" || Array.isArray(edit)) return edit;
      const box = edit as { id?: unknown; page?: unknown; sourceText?: unknown };
      const sealedBox = typeof box.id === "string" ? sealedBoxes.get(box.id) : undefined;
      const binding = sealedBox ? readPdfSourceTextBoxBinding(box.sourceText) : null;
      if (!sealedBox || !binding || Number(box.page) + 1 !== sealedBox.page) return edit;
      const sealed = pdfSourceTextBindingSealedForSave(binding, sealedBox.seal);
      return sealed.expectedAppearance === binding.expectedAppearance ? edit
        : { ...box, sourceText: { ...(box.sourceText as object), expectedAppearance: sealed.expectedAppearance } };
    });
  }
  return next;
}

/**
 * The seals the file `sourceSha256` names holds for `bindings`' boxes (null:
 * no such box in it yet): from the source's seal index where it was read off
 * the main thread already, else from the bytes the reader remembers or
 * `download` fetches -- hashed here, so only bytes of exactly that source
 * are read. Throws as the direct read does (a page gone, an ambiguous box, a
 * malformed seal) and when the bytes are not that source.
 */
export async function readPdfSourceTextSealsForSave(input: {
  sourceSha256: string;
  bindings: readonly PdfSourceTextBoxBinding[];
  download: () => Promise<{ bytes: Uint8Array }>;
  /** Told once the index has answered: whether it held the source. */
  onIndexRead?: (cached: boolean) => void;
}): Promise<Map<string, string | null>> {
  const cached = await cachedPdfSourceTextBoxAppearanceSeals(input.sourceSha256, input.bindings);
  input.onIndexRead?.(Boolean(cached));
  if (cached) return cached;
  const remembered = copyRememberedPdfSourceBytes(input.sourceSha256);
  let loaded = remembered ? { bytes: remembered } : await input.download();
  if (remembered && await sha256PdfBytes(remembered) !== input.sourceSha256) {
    // Not the bytes it claims to be: fetch the real ones.
    remembered.fill(0);
    loaded = await input.download();
  }
  try {
    if (await sha256PdfBytes(loaded.bytes) !== input.sourceSha256) throw new Error("The PDF source changed before saving.");
    return await readPdfSourceTextBoxAppearanceSeals(loaded.bytes, input.bindings);
  } finally {
    loaded.bytes.fill(0);
  }
}

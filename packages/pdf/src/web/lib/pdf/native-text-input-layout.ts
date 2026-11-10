/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Display-only soft line breaks. None of these values grants PDF write authority. */
export interface PdfNativeTextInputLayout {
  readonly text: string;
  readonly value: string;
  readonly breaks: readonly Readonly<{ offset: number; text: string }>[];
  readonly lineHeightPoints: number;
}

/** Read baseline spacing, never infer it from a box padded for glyph ink. */
export function pdfNativeTextInputLineHeight(
  items: readonly Readonly<{ transform: readonly number[] }>[], lineCount: number,
): number | null {
  const first = items[0]?.transform;
  if (!first || lineCount < 2) return null;
  const magnitude = Math.hypot(first[0], first[1]);
  if (!(magnitude > 0)) return null;
  const blocks: number[] = [];
  for (const { transform } of items) {
    const block = (-first[1] * transform[4] + first[0] * transform[5]) / magnitude;
    if (!Number.isFinite(block)) return null;
    if (!blocks.length || Math.abs(block - blocks.at(-1)!) > .05) blocks.push(block);
  }
  if (blocks.length !== lineCount) return null;
  const step = blocks[1] - blocks[0];
  if (!step || blocks.slice(1).some((block, i) => Math.abs(block - blocks[i] - step) > .05)) return null;
  return Math.abs(step);
}

export function createPdfNativeTextInputLayout(
  text: string,
  allLines: readonly string[],
  lineHeightPoints: number,
): PdfNativeTextInputLayout | null {
  // The rows a deletion across lines emptied at the end of a block show
  
  // row opened under the joined text). Rows that hold the text's own
  // trailing spaces stay.
  let count = allLines.length;
  while (count > 1 && allLines[count - 1] === "" && !/\s$/.test(text)) count--;
  const lines = count === allLines.length ? allLines : allLines.slice(0, count);
  if (lines.length < 2 || lines.length > 128 || text.length > 2000 ||
    !Number.isFinite(lineHeightPoints) || lineHeightPoints <= 0) return null;
  let cursor = 0;
  let value = "";
  const breaks: Array<{ offset: number; text: string }> = [];
  for (const [index, line] of lines.entries()) {
    if (index) {
      const start = cursor;
      // The source summary may own a space between visual lines, or no
      // separator at all (CJK). Preserve precisely those canonical bytes.
      while (!text.startsWith(line, cursor) && text[cursor] === " ") cursor++;
      breaks.push({ offset: value.length, text: text.slice(start, cursor) });
      value += "\n";
    }
    if (!text.startsWith(line, cursor)) return null;
    value += line;
    cursor += line.length;
  }
  if (cursor !== text.length) return null;
  return { text, value, breaks, lineHeightPoints };
}

export function pdfNativeTextInputCanonicalOffset(layout: PdfNativeTextInputLayout, offset: number): number {
  let result = Math.max(0, Math.min(layout.value.length, offset));
  for (const line of layout.breaks) if (line.offset < offset) result += line.text.length - 1;
  return result;
}

export function pdfNativeTextInputDisplayOffset(layout: PdfNativeTextInputLayout, offset: number): number {
  let delta = 0;
  for (const line of layout.breaks) {
    const start = line.offset + delta;
    if (offset < start) break;
    if (offset === start) return line.offset + (line.text.length === 0 ? 1 : 0);
    if (offset < start + line.text.length) return line.offset;
    delta += line.text.length - 1;
  }
  return Math.max(0, Math.min(layout.value.length, offset - delta));
}

/** Retain only synthetic breaks outside the user's actual edit. A new newline
 * typed by the user is real content; a surviving display break is not. */
export function editPdfNativeTextInputLayout(
  current: PdfNativeTextInputLayout,
  value: string,
  selection?: Readonly<{ start: number; end: number }>,
): PdfNativeTextInputLayout {
  let suffix = 0;
  while (suffix < current.value.length && suffix < value.length &&
    current.value[current.value.length - suffix - 1] === value[value.length - suffix - 1]) suffix++;
  let prefix = 0;
  while (prefix < current.value.length - suffix && prefix < value.length - suffix && current.value[prefix] === value[prefix]) prefix++;
  if (selection && selection.start >= 0 && selection.end >= selection.start && selection.end <= current.value.length) {
    const nextEnd = value.length - (current.value.length - selection.end);
    if (nextEnd >= selection.start &&
      current.value.slice(0, selection.start) === value.slice(0, selection.start) &&
      current.value.slice(selection.end) === value.slice(nextEnd)) {
      // beforeinput disambiguates repeated identical rows. String diff alone
      // cannot tell appending a third "foo" from prepending one.
      prefix = selection.start;
      suffix = current.value.length - selection.end;
    }
  }
  const shift = value.length - current.value.length;
  const breaks = current.breaks.flatMap(line => line.offset < prefix ? [line]
    : line.offset >= current.value.length - suffix ? [{ ...line, offset: line.offset + shift }] : []);
  let cursor = 0;
  let text = "";
  for (const line of breaks) {
    text += value.slice(cursor, line.offset) + line.text;
    cursor = line.offset + 1;
  }
  text += value.slice(cursor);
  return { ...current, text, value, breaks };
}

const previewLayouts = new WeakMap<HTMLCanvasElement, {
  sourceKey: string; width: number; height: number; layout: PdfNativeTextInputLayout;
}>();

export function rememberPdfNativeTextInputLayout(
  canvas: HTMLCanvasElement, sourceKey: string, layout: PdfNativeTextInputLayout,
): void {
  previewLayouts.set(canvas, { sourceKey, width: canvas.width, height: canvas.height, layout });
}

export function readPdfNativeTextInputLayout(
  canvas: HTMLCanvasElement, sourceKey: string, text: string,
): PdfNativeTextInputLayout | null {
  const entry = previewLayouts.get(canvas);
  // A canvas reused by another render must be registered again. Reading a
  // preview consumes its layout epoch, even when the caller's scope is stale.
  previewLayouts.delete(canvas);
  return entry && entry.sourceKey === sourceKey && entry.layout.text === text &&
    canvas.width > 0 && canvas.height > 0 && entry.width === canvas.width && entry.height === canvas.height
    ? entry.layout : null;
}

/** The input's text with its rows' breaks made places to wrap, offset for
 * offset: the space the row ended in, or, where no space was (CJK), a break
 * that draws nothing. */
export function pdfNativeTextInputLayoutJoinedRows(layout: PdfNativeTextInputLayout): string {
  let row = 0;
  return layout.value.replace(/\n/gu, () => (layout.breaks[row++]?.text.length ? " " : "\u200b"));
}

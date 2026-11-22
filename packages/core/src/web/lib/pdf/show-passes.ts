/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The passes of a show, read back out of what the reader reads.
 *
 * The editor's bold of a Type3 letter is the letter's show drawn twice more
 * right after it, inside a marked-content span that says the passes are no
 * text of their own -- once a little along the line, once where the show
 * stands (engine/native/standing-state-edit.ts, mraiNativeStandingOverstrike).
 * pdf.js reads glyphs and no /ActualText: it lists the two passes as text
 * items of their own, so every reader of its text content read a bold word
 * three times -- the text layer a selection copies from, the find, a
 * highlight's words, a page's text taken for an import or for the assistant.
 *
 * They are read back here, where pdf.js is loaded (pdfjs-runtime.ts): the
 * text content every page hands out is the one it handed out before the
 * bold, item for item -- the passes' two items taken out, the show's own item
 * given the line end the second pass was read with. Nothing that reads a page
 * through the loader sees a pass, whichever of the page's readings it asks
 * for (`getTextContent`, `streamTextContent`, with marked content or not).
 *
 * Only the editor's own passes are: the two items must stand between the
 * marks of one marked-content span tagged /Span that names no MCID, right
 * after the show, and be that show drawn again by the rule below. pdf.js
 * hands out those marks only where a reading asks for marked content, and
 * ends a text item at them whether it is asked or not -- so the loader asks
 * pdf.js for them on every reading, and takes them out again of a reading
 * that did not ask: the items of that reading are the ones pdf.js would have
 * handed it. Three items that only look like a show and its passes (another
 * producer's bold by overprinting) are left as they are read. (Of the first
 * four pages of the 1040 files of pdf-walk-1000, 3557 pages nobody edited,
 * the rule without the marks changed none either; with them it cannot.)
 *
 * This is the one place they are read back: the editor's box builder, the
 * find, a selection and every other reader are handed items by a page that
 * came through the loader, and read no pass. (A caller that loads pdf.js
 * itself -- a test, a probe -- reads the passes as pdf.js lists them.)
 */

/** What the rule reads of a reader item. */
export type PdfReaderTextItem = Readonly<{
  str: string;
  fontName: string;
  transform: readonly number[];
  width: number;
  height: number;
  hasEOL?: boolean;
}>;

/** How far apart two readings of one place may stand (the box builder's PDFJS_SPLIT_SPACE_TOLERANCE). */
const PLACE_TOLERANCE = 0.05;
/** How far along its line a pass may stand from its show, as a share of the item's size. */
const PASS_MAX_EM = 0.1;

/**
 * Whether `drawn`, `along` and `over` -- three reader items one after
 * another -- are a show and its two passes, and what the show's item is with
 * them read back into it: its text with whatever the third item reads after
 * the show's own, its width to where the third ends, the third's line end.
 * Null where they are not.
 *
 * Kept to the shape of such passes: one font and glyph matrix on one
 * baseline, upright; the second holds the text the first ends with and the
 * third begins with, stands along the line from the third by no more than a
 * pass over a drawing does (a tenth of an em), and ends that much past the
 * first; the third starts where the second's text starts in the first.
 */
export function pdfReaderShowAndPasses(
  drawn: PdfReaderTextItem,
  along: PdfReaderTextItem | null | undefined,
  over: PdfReaderTextItem | null | undefined,
): Readonly<{ str: string; width: number; hasEOL: boolean }> | null {
  if (!along || !over || drawn.hasEOL || along.hasEOL || !/\S/u.test(along.str) ||
    !(drawn.height > 0) || !(along.height > 0) || !(over.height > 0) ||
    along.fontName !== drawn.fontName || over.fontName !== drawn.fontName ||
    !(drawn.transform[0] > 0) || Math.abs(drawn.transform[1]) > PLACE_TOLERANCE ||
    !drawn.str.endsWith(along.str) || !over.str.startsWith(along.str)) return null;
  const em = Math.hypot(along.transform[2], along.transform[3]);
  const offset = along.transform[4] - over.transform[4];
  const end = drawn.transform[4] + drawn.width;
  if (!(em > 0) || !(offset > 0) || offset > em * PASS_MAX_EM ||
    [along, over].some((pass) => [0, 1, 2, 3, 5].some((component) =>
      Math.abs(pass.transform[component] - drawn.transform[component]) > PLACE_TOLERANCE)) ||
    Math.abs(along.transform[4] + along.width - offset - end) > PLACE_TOLERANCE ||
    Math.abs(over.transform[4] + along.width - end) > PLACE_TOLERANCE ||
    over.width < along.width - PLACE_TOLERANCE) return null;
  return Object.freeze({
    str: drawn.str + over.str.slice(along.str.length),
    width: over.transform[4] + over.width - drawn.transform[4],
    hasEOL: over.hasEOL === true,
  });
}

/** A text item of pdf.js's text content, as the rule reads one; null for a
 * marked-content entry, a line-end entry of no text, or anything else. */
function textItemOf(entry: unknown): PdfReaderTextItem | null {
  const item = entry as Partial<PdfReaderTextItem> | null;
  return item && typeof item === "object" && typeof item.str === "string" && item.str !== "" &&
    typeof item.fontName === "string" && Array.isArray(item.transform) && item.transform.length === 6 &&
    typeof item.width === "number" && typeof item.height === "number" ? item as PdfReaderTextItem : null;
}

const markedContentType = (entry: unknown): string | null => {
  const type = (entry as { type?: unknown } | null)?.type;
  return typeof type === "string" && typeof (entry as { str?: unknown }).str !== "string" ? type : null;
};

/** The mark that opens the passes' span: marked content with properties, tagged /Span, no MCID. */
const opensPassSpan = (entry: unknown): boolean => {
  const mark = entry as { tag?: unknown; id?: unknown } | null;
  return markedContentType(entry) === "beginMarkedContentProps" && mark?.tag === "Span" && (mark.id === null || mark.id === undefined);
};

/** How many entries a show and its passes are, as pdf.js lists them with
 * marked content: the show, the span's opening mark, the two passes, its
 * closing mark. */
const SHOW_AND_PASSES = 5;

/** The show at `index` with its passes read back into it, where the four
 * entries after it are its passes between their span's marks; else null. */
function showWithPassesAt(items: readonly unknown[], index: number): ReturnType<typeof pdfReaderShowAndPasses> {
  if (index + SHOW_AND_PASSES > items.length) return null;
  const drawn = textItemOf(items[index]);
  return drawn && opensPassSpan(items[index + 1]) && markedContentType(items[index + 4]) === "endMarkedContent"
    ? pdfReaderShowAndPasses(drawn, textItemOf(items[index + 2]), textItemOf(items[index + 3])) : null;
}

/**
 * The items of a page's text content with the passes of its shows taken out:
 * what the page read before a weight drew them. `items` as pdf.js lists them
 * with marked content (the passes are known by their span's marks, which go
 * with them; a list read without marked content holds none to know them by,
 * and is handed back as it is). The same array where the page holds no pass.
 */
export function pdfTextItemsWithoutShowPasses<Item>(items: readonly Item[]): readonly Item[] {
  let out: Item[] | null = null;
  for (let index = 0; index < items.length; index += 1) {
    const whole = showWithPassesAt(items, index);
    if (!whole) {
      out?.push(items[index]);
      continue;
    }
    out ??= items.slice(0, index);
    out.push({ ...(items[index] as object), ...whole } as Item);
    index += SHOW_AND_PASSES - 1;
  }
  return out ?? items;
}

/** `items` without the marks of marked content: what pdf.js lists for a reading that did not ask for them. */
function withoutMarks<Item>(items: readonly Item[]): readonly Item[] {
  return items.some((entry) => markedContentType(entry) !== null) ? items.filter((entry) => markedContentType(entry) === null) : items;
}

type TextContentLike = { items?: unknown; styles?: unknown; lang?: unknown };

/** A page's text content, read with marked content, with the passes of its shows taken out (pdfTextItemsWithoutShowPasses). */
export function pdfTextContentWithoutShowPasses<Content extends TextContentLike>(content: Content): Content {
  if (!content || !Array.isArray(content.items)) return content;
  const items = pdfTextItemsWithoutShowPasses(content.items as readonly unknown[]);
  return items === content.items ? content : { ...content, items };
}

/**
 * Where the entries of `read` that cannot be told yet begin: a show standing
 * so near the end that its passes, if it has any, are not all read -- itself
 * and at most three entries after it (its span's opening mark, the passes).
 * `read.length` where every entry is told.
 */
function undecidedFrom(read: readonly unknown[]): number {
  for (let index = Math.max(0, read.length - (SHOW_AND_PASSES - 1)); index < read.length; index += 1) {
    if (!textItemOf(read[index])) continue;
    const after = read.length - index - 1;
    if ((after < 1 || opensPassSpan(read[index + 1])) && (after < 2 || textItemOf(read[index + 2])) &&
      (after < 3 || textItemOf(read[index + 3]))) return index;
  }
  return read.length;
}

/**
 * A page's stream of text content with the passes taken out, handed on chunk
 * by chunk as pdf.js sends it: a reader of the stream (the text layer builds
 * a span for every item of a chunk as it arrives) is given a page's first
 * items when pdf.js has them, not when it has read the whole page. A show
 * and its passes may stand across two chunks, so the last entries of a chunk
 * that may be a show waiting for its passes -- four at most -- go with the
 * next one (the fonts of theirs went with the chunk they were read in).
 * `source` is read with marked content; `marks`: whether the reading asked
 * for it (else the marks are taken out again).
 *
 * The stream is pdf.js's own to its reader in everything else. It reads one
 * chunk of pdf.js's for each chunk asked of it and keeps none ahead (pdf.js's
 * stream does its own queueing); a reader that cancels cancels pdf.js's
 * stream with its reason, and waits for that as it would have -- the worker's
 * task ends when the reader says so; and once the stream has closed, failed
 * or been cancelled, nothing more is put on it. (Read whole and handed on as
 * one chunk, as this was first written, the text layer drew nothing until the
 * page was read to its end, and a cancel met a read still under way.)
 */
function streamWithoutShowPasses(source: ReadableStream<TextContentLike>, marks: boolean): ReadableStream<TextContentLike> {
  const reader = source.getReader();
  let held: readonly unknown[] = [];
  /** Closed, failed or cancelled: the stream takes nothing more. */
  let over = false;
  const handed = (items: readonly unknown[]) => {
    const kept = pdfTextItemsWithoutShowPasses(items);
    return (marks ? kept : withoutMarks(kept)) as unknown[];
  };
  return new ReadableStream<TextContentLike>({
    async pull(controller) {
      try {
        for (;;) {
          const { value, done } = await reader.read();
          if (over) return;
          if (done) {
            over = true;
            // Nothing follows what was kept back: it is as it was read.
            if (held.length) controller.enqueue({ items: handed(held), styles: Object.create(null), lang: null });
            held = [];
            controller.close();
            return;
          }
          if (!value) continue;
          const items = Array.isArray(value.items) ? value.items as readonly unknown[] : [];
          const read = held.length ? [...held, ...items] : items;
          const told = undecidedFrom(read);
          held = told < read.length ? read.slice(told) : [];
          controller.enqueue({ ...value, items: handed(told < read.length ? read.slice(0, told) : read) });
          return;
        }
      } catch (error) {
        if (over) return;
        over = true;
        held = [];
        controller.error(error);
      }
    },
    cancel(reason) {
      over = true;
      held = [];
      return reader.cancel(reason);
    },
  }, {
    // (Nothing read ahead of its reader: one chunk of pdf.js's for each asked of it.)
    highWaterMark: 0,
  });
}

const READS_SHOW_PASSES_BACK = Symbol.for("mrai.pdfjs.reads-show-passes-back");

type Patched = { [READS_SHOW_PASSES_BACK]?: true };

/** A page's prototype, reading its text content with the passes taken out
 * (`getTextContent` reads `streamTextContent`). Once. */
function readShowPassesBackOnPages(page: unknown): void {
  const prototype = page && typeof page === "object" ? Object.getPrototypeOf(page) as (Patched & {
    streamTextContent?: (this: unknown, params?: unknown) => ReadableStream<TextContentLike>;
  }) | null : null;
  if (!prototype || prototype === Object.prototype || prototype[READS_SHOW_PASSES_BACK] ||
    typeof prototype.streamTextContent !== "function" || typeof ReadableStream === "undefined") return;
  const streamTextContent = prototype.streamTextContent;
  prototype.streamTextContent = function streamTextContentWithoutShowPasses(this: unknown, params?: unknown) {
    // (With the marks of marked content, whatever the reading asks: the passes are known by their span's.)
    const asked = params && typeof params === "object" ? params as { includeMarkedContent?: unknown } : {};
    const source = streamTextContent.call(this, { ...asked, includeMarkedContent: true });
    return source && typeof (source as { getReader?: unknown }).getReader === "function"
      ? streamWithoutShowPasses(source, asked.includeMarkedContent === true) : source;
  };
  prototype[READS_SHOW_PASSES_BACK] = true;
}

/**
 * pdf.js, reading every page's text content with the passes of its shows
 * taken out: its `getDocument` hands out documents whose pages do. (The
 * page class is not exported, so it is reached from the first page a
 * document hands out, before the caller that asked for the page has it.)
 */
export function pdfJsReadingShowPassesBack<Api extends { getDocument?: unknown }>(pdfjs: Api): Api {
  const getDocument = pdfjs.getDocument as ((...args: unknown[]) => unknown) | undefined;
  if (typeof getDocument !== "function") return pdfjs;
  const wrapped = (...args: unknown[]) => {
    const task = getDocument(...args) as { promise?: PromiseLike<unknown> } | null;
    const opened = task?.promise;
    if (opened && typeof opened.then === "function") {
      void Promise.resolve(opened).then((document) => {
        const prototype = document && typeof document === "object" ? Object.getPrototypeOf(document) as (Patched & {
          getPage?: (this: unknown, pageNumber: number) => Promise<unknown>;
        }) | null : null;
        if (!prototype || prototype === Object.prototype || prototype[READS_SHOW_PASSES_BACK] || typeof prototype.getPage !== "function") return;
        const getPage = prototype.getPage;
        prototype.getPage = function getPageReadingShowPassesBack(this: unknown, pageNumber: number) {
          const page = getPage.call(this, pageNumber);
          return page && typeof (page as { then?: unknown }).then === "function"
            ? page.then((opened) => { readShowPassesBackOnPages(opened); return opened; }) : page;
        };
        prototype[READS_SHOW_PASSES_BACK] = true;
      }, () => { /* a document that does not open hands out no page */ });
    }
    return task;
  };
  return { ...pdfjs, getDocument: wrapped as Api["getDocument"] };
}

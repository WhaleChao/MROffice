/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { appendMraiNativeObjectRevisions } from "./engine/native/content-stream-transaction";
import { planPdfToUnicodeRepair } from "./engine/native/to-unicode-repair";
import { sha256PdfBytes } from "./engine/native/transaction-error";
import { pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

type PdfLib = typeof import("pdf-lib");
type PdfDict = import("pdf-lib").PDFDict;
type PdfRef = import("pdf-lib").PDFRef;

export interface PdfNativeFontReadingRepairPlan {
  readonly candidateBytes: Uint8Array;
  readonly candidateSha256: string;
  /** The fonts written again, each with the codes it now names. */
  readonly fonts: readonly Readonly<{
    baseFont: string;
    toUnicode: Readonly<{ objectNumber: number; generation: number }>;
    codes: readonly Readonly<{ code: number; text: string }>[];
    /** Codes that still draw as blanks (the program names them nothing). */
    left: number;
  }>[];
}

const refKey = (ref: PdfRef) => `${ref.objectNumber} ${ref.generationNumber}`;

/** The Type0 fonts a page draws with, its own and those of the forms it
 * draws, each once. `seen` keeps what a resource dictionary led to, for the
 * pages of one document that share it. */
function pageType0Fonts(lib: PdfLib, doc: import("pdf-lib").PDFDocument, pageNumber: number,
  seen: Map<unknown, readonly PdfDict[]> = new Map()): readonly PdfDict[] {
  const page = doc.getPages()[pageNumber - 1];
  if (!page) return [];
  const resolve = (value: unknown) => (value instanceof lib.PDFRef ? doc.context.lookup(value) : value);
  // (A dictionary reached again while it is still being read is read short
  // there: what was read through such a loop is not kept for the next page.)
  let looped = false;
  const walk = (resources: unknown, depth: number, open: Set<unknown>): readonly PdfDict[] => {
    const dict = resolve(resources);
    if (!(dict instanceof lib.PDFDict)) return [];
    if (depth > 8 || open.has(dict)) { looped = true; return []; }
    const known = seen.get(dict);
    if (known) return known;
    open.add(dict);
    const loopedBefore = looped;
    looped = false;
    const fonts = new Set<PdfDict>();
    const fontDict = resolve(dict.get(lib.PDFName.of("Font")));
    if (fontDict instanceof lib.PDFDict) {
      for (const [, value] of fontDict.entries()) {
        const font = resolve(value);
        if (!(font instanceof lib.PDFDict)) continue;
        const subtype = resolve(font.get(lib.PDFName.of("Subtype")));
        if (subtype instanceof lib.PDFName && subtype.decodeText() === "Type0") fonts.add(font);
      }
    }
    const xobjects = resolve(dict.get(lib.PDFName.of("XObject")));
    if (xobjects instanceof lib.PDFDict) {
      for (const [, value] of xobjects.entries()) {
        const form = resolve(value);
        if (!(form instanceof lib.PDFRawStream)) continue;
        const subtype = resolve(form.dict.get(lib.PDFName.of("Subtype")));
        if (subtype instanceof lib.PDFName && subtype.decodeText() === "Form") {
          for (const font of walk(form.dict.get(lib.PDFName.of("Resources")), depth + 1, open)) fonts.add(font);
        }
      }
    }
    open.delete(dict);
    const found = Object.freeze([...fonts]);
    if (!looped) seen.set(dict, found);
    looped ||= loopedBefore;
    return found;
  };
  return walk(page.node.Resources(), 0, new Set());
}

/** The program an Identity-encoded CIDFontType2 with an Identity CIDToGIDMap
 * draws with, and the ToUnicode stream it is read by: the fonts whose code
 * is their glyph id. Null for any other font. */
function identityTrueTypeFont(lib: PdfLib, doc: import("pdf-lib").PDFDocument, font: PdfDict): Readonly<{
  toUnicode: PdfRef; program: PdfRef;
}> | null {
  const resolve = (value: unknown) => (value instanceof lib.PDFRef ? doc.context.lookup(value) : value);
  const get = (dict: PdfDict, key: string) => resolve(dict.get(lib.PDFName.of(key)));
  const encoding = get(font, "Encoding");
  if (!(encoding instanceof lib.PDFName) || !/^Identity-[HV]$/u.test(encoding.decodeText())) return null;
  const descendants = get(font, "DescendantFonts");
  const descendant = descendants instanceof lib.PDFArray ? descendants.lookup(0) : null;
  if (!(descendant instanceof lib.PDFDict)) return null;
  const subtype = get(descendant, "Subtype");
  if (!(subtype instanceof lib.PDFName) || subtype.decodeText() !== "CIDFontType2") return null;
  const glyphMap = get(descendant, "CIDToGIDMap");
  if (glyphMap !== undefined && !(glyphMap instanceof lib.PDFName && glyphMap.decodeText() === "Identity")) return null;
  const descriptor = get(descendant, "FontDescriptor");
  const program = descriptor instanceof lib.PDFDict ? descriptor.get(lib.PDFName.of("FontFile2")) : null;
  const toUnicode = font.get(lib.PDFName.of("ToUnicode"));
  if (!(program instanceof lib.PDFRef) || !(toUnicode instanceof lib.PDFRef)) return null;
  if (!(doc.context.lookup(program) instanceof lib.PDFRawStream) ||
    !(doc.context.lookup(toUnicode) instanceof lib.PDFRawStream)) return null;
  return { toUnicode, program };
}

/** Every dictionary of the file that names `/ToUnicode`, by the stream it
 * names: the programs of the fonts that are read by each stream (`?` for a
 * font this does not read as an Identity TrueType one). */
function toUnicodeUsers(lib: PdfLib, doc: import("pdf-lib").PDFDocument): Map<string, Set<string>> {
  const users = new Map<string, Set<string>>();
  const key = lib.PDFName.of("ToUnicode");
  const visit = (value: unknown, depth: number) => {
    if (depth > 6) return;
    if (value instanceof lib.PDFRawStream) { visit(value.dict, depth); return; }
    if (value instanceof lib.PDFArray) {
      for (const entry of value.asArray()) if (!(entry instanceof lib.PDFRef)) visit(entry, depth + 1);
      return;
    }
    if (!(value instanceof lib.PDFDict)) return;
    const named = value.get(key);
    if (named instanceof lib.PDFRef) {
      const font = identityTrueTypeFont(lib, doc, value);
      let programs = users.get(refKey(named));
      if (!programs) users.set(refKey(named), programs = new Set());
      programs.add(font ? refKey(font.program) : "?");
    }
    for (const [, entry] of value.entries()) if (!(entry instanceof lib.PDFRef)) visit(entry, depth + 1);
  };
  for (const [, object] of doc.context.enumerateIndirectObjects()) visit(object, 0);
  return users;
}

/** What the stream's dictionary says besides how its bytes are stored (a
 * `/Type /CMap`, say): kept on the stream written again. */
function keptStreamEntries(lib: PdfLib, dict: PdfDict): string {
  const dropped = new Set(["Length", "Filter", "DecodeParms", "DL", "F", "FFilter", "FDecodeParms"]);
  let out = "";
  for (const [name, value] of dict.entries()) {
    if (dropped.has(name.decodeText())) continue;
    out += ` ${name.toString()} ${value.toString()}`;
  }
  return out;
}

type PlannedRepair = Readonly<{ ref: PdfRef; baseFont: string; repair: NonNullable<ReturnType<typeof planPdfToUnicodeRepair>>;
  dict: PdfDict }>;

/** One document's fonts as this reads them, each ToUnicode stream answered
 * once: its repair, or null where it has none to write. */
function repairReader(lib: PdfLib, doc: import("pdf-lib").PDFDocument) {
  const answered = new Map<string, PlannedRepair | null>();
  const resources = new Map<unknown, readonly PdfDict[]>();
  let users: Map<string, Set<string>> | null = null;
  const repairOf = (font: PdfDict): PlannedRepair | null => {
    const identity = identityTrueTypeFont(lib, doc, font);
    if (!identity) return null;
    const key = refKey(identity.toUnicode);
    const known = answered.get(key);
    if (known !== undefined) return known;
    let planned: PlannedRepair | null = null;
    try {
      const toUnicode = doc.context.lookup(identity.toUnicode) as import("pdf-lib").PDFRawStream;
      const program = doc.context.lookup(identity.program) as import("pdf-lib").PDFRawStream;
      const repair = planPdfToUnicodeRepair(lib.decodePDFRawStream(toUnicode).decode(), lib.decodePDFRawStream(program).decode());
      if (repair) {
        // Only a stream every font it reads draws with the one program.
        users ??= toUnicodeUsers(lib, doc);
        const programs = users.get(key);
        if (programs && programs.size === 1 && programs.has(refKey(identity.program))) {
          const baseFont = font.get(lib.PDFName.of("BaseFont"));
          planned = { ref: identity.toUnicode, repair, dict: toUnicode.dict,
            baseFont: baseFont instanceof lib.PDFName ? baseFont.decodeText() : "" };
        }
      }
    } catch { planned = null; }
    answered.set(key, planned);
    return planned;
  };
  return {
    /** The repairs of the fonts page `pageNumber` draws with, by object number. */
    ofPage(pageNumber: number): PlannedRepair[] {
      const found = new Map<string, PlannedRepair>();
      for (const font of pageType0Fonts(lib, doc, pageNumber, resources)) {
        const planned = repairOf(font);
        if (planned) found.set(refKey(planned.ref), planned);
      }
      return [...found.values()].sort((a, b) => a.ref.objectNumber - b.ref.objectNumber);
    },
  };
}

/** The largest file whose fonts are read here: the whole file is parsed for
 * it, in the reader's browser. A larger one keeps its fonts as they read
 * (its misnamed lines are refused, as before). */
const MAX_REPAIR_SOURCE_BYTES = 32 * 1024 * 1024;

async function loadForRepair(sourceBytes: Uint8Array) {
  if (sourceBytes.length > MAX_REPAIR_SOURCE_BYTES) return null;
  const lib = await import("pdf-lib");
  const doc = await lib.PDFDocument.load(Uint8Array.from(sourceBytes),
    { parseSpeed: pdfLibParseSpeed(), ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
  return doc.isEncrypted ? null : { lib, doc };
}

/**
 * The pages of the file that draw with a font this can read right (their
 * numbers, from 1): asked once for a document before its text is edited, so
 * a page that needs nothing costs nothing more. Empty where there is none.
 */
export async function inspectPdfNativeFontReadingRepair(input: {
  readonly sourceBytes: Uint8Array;
}): Promise<Readonly<{ pages: readonly number[] }>> {
  const loaded = await loadForRepair(input.sourceBytes);
  if (!loaded) return Object.freeze({ pages: Object.freeze([]) });
  const reader = repairReader(loaded.lib, loaded.doc);
  const pages: number[] = [];
  const count = loaded.doc.getPageCount();
  for (let pageNumber = 1; pageNumber <= count; pageNumber += 1) {
    if (reader.ofPage(pageNumber).length) pages.push(pageNumber);
  }
  return Object.freeze({ pages: Object.freeze(pages) });
}

/**
 * The file with the fonts of page `pageNumber` read right: one revision
 * appended, each repaired font's ToUnicode stream written again. Null where
 * the page draws with no font this can name the glyphs of (nothing is
 * written then).
 */
export async function planPdfNativeFontReadingRepair(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
}): Promise<PdfNativeFontReadingRepairPlan | null> {
  if (!Number.isSafeInteger(input.pageNumber) || input.pageNumber < 1) return null;
  const loaded = await loadForRepair(input.sourceBytes);
  if (!loaded) return null;
  const { lib } = loaded;
  const planned = repairReader(lib, loaded.doc).ofPage(input.pageNumber);
  if (planned.length === 0) return null;
  const revisions: { objectNumber: number; generation: number; bytes: Uint8Array }[] = [];
  const fonts: PdfNativeFontReadingRepairPlan["fonts"][number][] = [];
  for (const plan of planned) {
    const head = new TextEncoder().encode(`${plan.ref.objectNumber} ${plan.ref.generationNumber} obj\n` +
      `<<${keptStreamEntries(lib, plan.dict)} /Length ${plan.repair.stream.length} >>\nstream\n`);
    const tail = new TextEncoder().encode("\nendstream\nendobj");
    const bytes = new Uint8Array(head.length + plan.repair.stream.length + tail.length);
    bytes.set(head, 0);
    bytes.set(plan.repair.stream, head.length);
    bytes.set(tail, head.length + plan.repair.stream.length);
    revisions.push({ objectNumber: plan.ref.objectNumber, generation: plan.ref.generationNumber, bytes });
    fonts.push(Object.freeze({
      baseFont: plan.baseFont,
      toUnicode: Object.freeze({ objectNumber: plan.ref.objectNumber, generation: plan.ref.generationNumber }),
      codes: Object.freeze([...plan.repair.repaired].sort((a, b) => a[0] - b[0])
        .map(([code, text]) => Object.freeze({ code, text }))),
      left: plan.repair.left.length,
    }));
  }
  const candidateBytes = appendMraiNativeObjectRevisions(input.sourceBytes, revisions);
  return Object.freeze({
    candidateBytes,
    candidateSha256: await sha256PdfBytes(candidateBytes),
    fonts: Object.freeze(fonts),
  });
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Install the faces MRAI serves into a PDF for one write (WPS: a part of a
 * line set in a font the PDF does not hold is written in that font, the font
 * embedded as a subset).
 *
 * The planner asks for each part it sets in a family the PDF does not hold
 * (native-embedded-font-key.ts) and gets the part back as stretches, each in
 * one installed font: one font per slice of the face (a CJK face is ~100
 * slices; a line uses a few), so only the slices a text needs are fetched.
 * After planning, `packages()` builds every font's package: a subset program
 * of its slice (sfnt-subset.ts, at the face's weight when the slice is a
 * variable font) embedded as a Type0/CIDFontType2 under the face's own name.
 * The glyph numbers and advances the planner wrote with are the ones the
 * package carries, checked here; the server proves every glyph against the
 * pinned slice (native_font_extension.py).
 */

import { LATIN1_EXACT } from "./engine/native/latin1";
import {
  embeddableFontFamily,
  embeddableSlicesFor,
  selectEmbeddableFontFace,
  type EmbeddableFontFace,
  type EmbeddableFontManifest,
  type EmbeddableFontSlice,
} from "@mrnote/core/web/lib/office-fonts/embeddable-faces.ts";
import {
  pdfNativeEmbeddedFontRefusalTag,
  readPdfNativeEmbeddedFontKey,
  type PdfNativeEmbeddedFontRefusal,
} from "./native-embedded-font-key";
import { choosePdfPageFont, type PdfPageFont } from "./page-fonts";
import {
  buildPdfNativeEmbeddedFontPackage,
  pdfEmbeddedFontBaseName,
  type PdfNativeEmbeddedFontFace,
  type PdfNativeFontExtensionPackage,
} from "./engine/native/font-extension-package";
import { readSfntProgram, type SfntProgram } from "@mrnote/core/web/lib/pdf/engine/native/sfnt/sfnt-reader.ts";
import { buildPdfNativeSfntSubset, otRound } from "./engine/native/sfnt/sfnt-subset";
import { createSfntInstancer, normalizeSfntLocation, sfntVariationAxes } from "@mrnote/core/web/lib/pdf/engine/native/sfnt/sfnt-variation.ts";
import {
  pdfNativeTextFontIdentity,
  type PdfNativeEmbeddedFontProvider,
  type PdfNativeEmbeddedFontSession,
  type PdfNativeEmbeddedFontStretch,
  type PdfNativeIndirectObjectRef,
  type PdfNativeTextFontOption,
} from "./engine/native/list-text-fidelity";

/** Where an installation gets the faces: the lock's manifest and the
 * verified bytes of a slice (the browser fetches them; the engine host is
 * handed the ones the browser fetched, verified again by digest). */
export interface PdfNativeEmbeddedFontLibrary {
  readonly manifest: EmbeddableFontManifest;
  readonly loadSlice: (slice: EmbeddableFontSlice) => Promise<Uint8Array>;
  /** Fonts from readers' computers the page keeps (page-fonts.ts): a family
   * MRAI does not serve is set in the page's face of it, its program a
   * slice of its own (`loadSlice` hands its bytes over like a served one). */
  readonly pageFonts?: readonly PdfPageFont[];
  /** What the library holds, as one string (the programs by digest, the
   * served list, the page fonts): two libraries with the same fingerprint
   * install the same faces. Absent: a library a plan is not shared over. */
  readonly fingerprint?: string;
}

/** The slice a page font is installed from: the whole prepared face. */
export function pdfPageFontSlice(font: PdfPageFont): EmbeddableFontSlice {
  return Object.freeze({ unicodeRange: null, out: `page:${font.sha256}`, sha256: font.sha256, size: font.size });
}

export type PdfNativeEmbeddedFontErrorCode = PdfNativeEmbeddedFontRefusal["code"];

/** Refusals the editor says in the reader's language: the family, and for a
 * missing glyph the character. */
export class PdfNativeEmbeddedFontError extends Error {
  constructor(
    readonly code: PdfNativeEmbeddedFontErrorCode,
    readonly family: string,
    readonly character?: string,
  ) {
    // The refusal crosses the engine host as its message: it opens with the
    // tag the editor reads it back from (readPdfNativeEmbeddedFontRefusal).
    super(
      pdfNativeEmbeddedFontRefusalTag({ code, family, character }) + " " +
      (code === "missing-glyph"
        ? `The font ${family} has no glyph for “${character ?? ""}”.`
        : code === "unknown-family"
          ? `The font ${family} is not one MRAI can embed.`
          : code === "unsupported-face"
            ? `The font ${family} cannot be embedded in this style.`
            : code === "unsupported-place"
              ? "Fonts are installed into a page's own text only."
              : `The font ${family} could not be loaded.`),
    );
    this.name = "PdfNativeEmbeddedFontError";
  }
}

export function isPdfNativeEmbeddedFontError(error: unknown): error is PdfNativeEmbeddedFontError {
  return error instanceof PdfNativeEmbeddedFontError;
}

/** pdf-lib writes a Type0 font as five objects: the font at the first
 * number, then its program, descriptor, CIDFont and ToUnicode. */
export const PDF_NATIVE_EMBEDDED_FONT_PACKAGE_OBJECTS = 5;

export interface PdfNativeEmbeddedFontInstallation extends PdfNativeEmbeddedFontProvider {
  /** Every font the session installed, as its package, in object order.
   * Empty when nothing was set in an embedded face. */
  packages(): Promise<readonly PdfNativeFontExtensionPackage[]>;
}

interface InstalledFont {
  readonly slice: EmbeddableFontSlice;
  readonly face: EmbeddableFontFace;
  /** A face the page keeps (its package is marked `page`). */
  readonly page: boolean;
  readonly program: SfntProgram;
  readonly location: Readonly<Record<string, number>> | null;
  readonly resourceName: string;
  readonly firstObjectNumber: number;
  readonly option: PdfNativeTextFontOption;
  /** Glyph ids in the order the text first names them: their CIDs. */
  readonly glyphOrder: number[];
  readonly advances: Map<number, number>;
  /** Every character the write draws in it, in the order it draws them. */
  text: string;
}

const programCache = new Map<string, Promise<SfntProgram>>();
const PROGRAM_CACHE_LIMIT = 24;

function decodedProgram(slice: EmbeddableFontSlice, load: () => Promise<Uint8Array>): Promise<SfntProgram> {
  const cached = programCache.get(slice.sha256);
  if (cached) {
    programCache.delete(slice.sha256);
    programCache.set(slice.sha256, cached);
    return cached;
  }
  const decoding = load().then((bytes) => readSfntProgram(bytes));
  programCache.set(slice.sha256, decoding);
  decoding.catch(() => programCache.delete(slice.sha256));
  while (programCache.size > PROGRAM_CACHE_LIMIT) programCache.delete(programCache.keys().next().value!);
  return decoding;
}

async function sha256Hex(text: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

/** The location a face's slice is drawn at: its weight on a variable
 * slice whose default is another weight, else the slice's own default. */
function faceLocation(program: SfntProgram, face: EmbeddableFontFace): Readonly<Record<string, number>> | null {
  const weightAxis = sfntVariationAxes(program).find((axis) => axis.tag === "wght");
  if (!weightAxis || weightAxis.defaultValue === face.weight) return null;
  if (face.weight < weightAxis.min || face.weight > weightAxis.max) return null;
  return Object.freeze({ wght: face.weight });
}

export function createPdfNativeEmbeddedFontInstallation(input: {
  readonly library: PdfNativeEmbeddedFontLibrary;
  readonly sourceSha256: string;
  readonly pageNumber: number;
}): PdfNativeEmbeddedFontInstallation {
  let opened: Readonly<{ firstObjectNumber: number; page: PdfNativeIndirectObjectRef }> | null = null;
  const fonts = new Map<string, InstalledFont>();
  const instancers = new Map<string, ReturnType<typeof createSfntInstancer>>();

  const advanceOf = (font: InstalledFont, glyphId: number): number => {
    if (!font.location) return font.program.advances[glyphId];
    const key = `${font.slice.sha256}:${font.face.weight}`;
    let instancer = instancers.get(key);
    if (!instancer) {
      instancer = createSfntInstancer(font.program, normalizeSfntLocation(font.program, font.location));
      instancers.set(key, instancer);
    }
    return otRound(instancer.instance(glyphId).advance);
  };

  const installed = async (
    family: string,
    slice: EmbeddableFontSlice,
    face: EmbeddableFontFace,
    program: SfntProgram,
    page: boolean,
  ): Promise<InstalledFont> => {
    if (!opened) throw new PdfNativeEmbeddedFontError("unavailable", family);
    const key = `${slice.sha256}:${face.weight}:${face.italic ? 1 : 0}`;
    const existing = fonts.get(key);
    if (existing) return existing;
    const index = fonts.size;
    const firstObjectNumber = opened.firstObjectNumber + index * PDF_NATIVE_EMBEDDED_FONT_PACKAGE_OBJECTS;
    const digest = await sha256Hex([
      "mrai.native-embedded-font.v1", input.sourceSha256, input.pageNumber, slice.sha256, face.weight, face.italic ? 1 : 0, index,
    ].join("\u0000"));
    const resourceName = `MRAI_EF_${digest.slice(0, 24)}`;
    const baseFont = pdfEmbeddedFontBaseName({ family: face.substitute, weight: face.weight, italic: face.italic });
    const ref = (offset: number) => Object.freeze({ objectNumber: firstObjectNumber + offset, generation: 0 });
    const option: PdfNativeTextFontOption = Object.freeze({
      resourceName,
      baseFont,
      ...pdfNativeTextFontIdentity(baseFont),
      encoding: Object.freeze({ kind: "identity-h-cidfonttype2" as const, codeUnitBytes: 2 as const }),
      resources: Object.freeze({
        page: opened.page,
        font: ref(0),
        fontFile: ref(1),
        descriptor: ref(2),
        descendantFont: ref(3),
        toUnicode: ref(4),
      }),
    }) as PdfNativeTextFontOption;
    const font: InstalledFont = {
      slice, face, program, page,
      location: faceLocation(program, face),
      resourceName, firstObjectNumber, option,
      glyphOrder: [], advances: new Map(), text: "",
    };
    if (sfntVariationAxes(program).some((axis) => axis.tag === "wght") && font.location === null &&
      sfntVariationAxes(program).find((axis) => axis.tag === "wght")!.defaultValue !== face.weight) {
      throw new PdfNativeEmbeddedFontError("unsupported-face", family);
    }
    fonts.set(key, font);
    return font;
  };

  const session: PdfNativeEmbeddedFontSession = Object.freeze({
    async stretches(key: string, text: string): Promise<readonly PdfNativeEmbeddedFontStretch[]> {
      const request = readPdfNativeEmbeddedFontKey(key);
      if (!request) throw new PdfNativeEmbeddedFontError("unknown-family", key);
      const family = embeddableFontFamily(request.family);
      const manifest = input.library.manifest;
      // The faces a character may be set in, in stack order: each served
      // substitute's face for the style, else the page's own face of a
      // family from the reader's computer.
      const pageChoice = family ? null : choosePdfPageFont(input.library.pageFonts ?? [], request.family, request);
      if (!family && !pageChoice) throw new PdfNativeEmbeddedFontError("unknown-family", request.family);
      const choices = family
        ? family.substitutes.flatMap((substitute) => {
          const choice = selectEmbeddableFontFace(manifest, family, substitute, request);
          return choice ? [{ ...choice, page: false }] : [];
        })
        : [{
          face: Object.freeze({
            substitute: pageChoice!.font.family,
            weight: pageChoice!.font.weight,
            italic: pageChoice!.font.italic,
            slices: Object.freeze([pdfPageFontSlice(pageChoice!.font)]),
          }) as EmbeddableFontFace,
          syntheticBold: pageChoice!.syntheticBold,
          syntheticItalic: pageChoice!.syntheticItalic,
          page: true,
        }];
      type Placed = { font: InstalledFont; character: string; synthetic: { bold: boolean; italic: boolean } };
      const placed: Placed[] = [];
      for (const character of text) {
        const codePoint = character.codePointAt(0)!;
        let found: Placed | null = null;
        for (const choice of choices) {
          for (const slice of embeddableSlicesFor(choice.face, codePoint)) {
            let program: SfntProgram;
            try {
              program = await decodedProgram(slice, () => input.library.loadSlice(slice));
            } catch (error) {
              if (error instanceof Error && error.name === "PdfNativeSfntError" &&
                (error as Error & { code?: string }).code === "unsupported-outlines") {
                throw new PdfNativeEmbeddedFontError("unsupported-face", request.family);
              }
              throw new PdfNativeEmbeddedFontError("unavailable", request.family);
            }
            const glyphId = program.cmap.get(codePoint);
            if (glyphId === undefined || glyphId === 0) continue;
            found = {
              font: await installed(request.family, slice, choice.face, program, choice.page),
              character,
              synthetic: { bold: choice.syntheticBold, italic: choice.syntheticItalic },
            };
            break;
          }
          if (found) break;
        }
        if (!found) throw new PdfNativeEmbeddedFontError("missing-glyph", request.family, character);
        placed.push(found);
      }
      const stretches: PdfNativeEmbeddedFontStretch[] = [];
      let index = 0;
      while (index < placed.length) {
        const { font, synthetic } = placed[index];
        let stretchText = "";
        const glyphs = new Map<string, { cid: number; text: string; advanceAt1000: number }>();
        while (index < placed.length && placed[index].font === font &&
          placed[index].synthetic.bold === synthetic.bold && placed[index].synthetic.italic === synthetic.italic) {
          const character = placed[index].character;
          const glyphId = font.program.cmap.get(character.codePointAt(0)!)!;
          if (!font.glyphOrder.includes(glyphId)) font.glyphOrder.push(glyphId);
          let advance = font.advances.get(glyphId);
          if (advance === undefined) {
            advance = advanceOf(font, glyphId);
            font.advances.set(glyphId, advance);
          }
          glyphs.set(character, {
            cid: font.glyphOrder.indexOf(glyphId) + 1,
            text: character,
            advanceAt1000: (advance * 1_000) / font.program.unitsPerEm,
          });
          stretchText += character;
          index += 1;
        }
        font.text += stretchText;
        stretches.push(Object.freeze({
          text: stretchText,
          font: font.option,
          glyphs: Object.freeze([...glyphs.values()].map((glyph) => Object.freeze(glyph))),
          syntheticBold: synthetic.bold,
          slanted: synthetic.italic,
        }));
      }
      return Object.freeze(stretches);
    },
  });

  return Object.freeze({
    open(value: Readonly<{ firstObjectNumber: number; page: PdfNativeIndirectObjectRef }>) {
      if (opened && (opened.firstObjectNumber !== value.firstObjectNumber ||
        opened.page.objectNumber !== value.page.objectNumber || opened.page.generation !== value.page.generation)) {
        throw new PdfNativeEmbeddedFontError("unavailable", "");
      }
      opened = Object.freeze({ firstObjectNumber: value.firstObjectNumber, page: value.page });
      return session;
    },
    restart() {
      opened = null;
      fonts.clear();
    },
    async packages(): Promise<readonly PdfNativeFontExtensionPackage[]> {
      const packages: PdfNativeFontExtensionPackage[] = [];
      for (const font of fonts.values()) {
        const subset = buildPdfNativeSfntSubset({ program: font.program, text: font.text, location: font.location });
        // The subset numbers glyphs as the session handed them out.
        for (const glyph of subset.glyphs) {
          const glyphId = font.program.cmap.get(glyph.text.codePointAt(0)!)!;
          if (glyph.cid !== font.glyphOrder.indexOf(glyphId) + 1 || glyph.advance !== font.advances.get(glyphId)) {
            throw new PdfNativeEmbeddedFontError("unavailable", font.face.substitute);
          }
        }
        const face: PdfNativeEmbeddedFontFace = Object.freeze({
          family: font.face.substitute,
          weight: font.face.weight,
          italic: font.face.italic,
          ...(font.page ? { page: true as const } : {}),
        });
        const fontPackage = await buildPdfNativeEmbeddedFontPackage({
          subset,
          programSha256: font.slice.sha256,
          face,
          text: font.text,
          resourceName: font.resourceName,
          firstObjectNumber: font.firstObjectNumber,
        });
        const objectText = (offset: number) =>
          LATIN1_EXACT.decode(fontPackage.objects[offset]?.bytes ?? new Uint8Array());
        const at = (offset: number) => `${font.firstObjectNumber + offset} 0 R`;
        if (
          fontPackage.objects.length !== PDF_NATIVE_EMBEDDED_FONT_PACKAGE_OBJECTS ||
          fontPackage.objects.some((object, offset) => object.objectNumber !== font.firstObjectNumber + offset) ||
          !objectText(0).includes(`/DescendantFonts [ ${at(3)} ]`) ||
          !objectText(0).includes(`/ToUnicode ${at(4)}`) ||
          !objectText(2).includes(`/FontFile2 ${at(1)}`) ||
          !objectText(3).includes(`/FontDescriptor ${at(2)}`)
        ) {
          throw new PdfNativeEmbeddedFontError("unavailable", font.face.substitute);
        }
        packages.push(fontPackage);
      }
      return Object.freeze(packages);
    },
  });
}

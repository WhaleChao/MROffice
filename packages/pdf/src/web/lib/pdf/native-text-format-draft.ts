/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The format commands of a PDF's own text run, as drafts of its properties:
 * the font family, the size, Bold and Italic, the colour, Increase / Decrease
 * Indent, Clear Formatting and the Format Painter.
 *
 * The page window's text editor (PdfNativeListTextEditDialog) presses them on
 * the run it holds open; the AI's pdf.text.format_original presses the same
 * ones where no editor runs (engine/native/text-format-plan.ts). Both start
 * from the run's property proof and its current draft and get the next
 * draft, which the native planner writes into the page
 * (planMraiNativeListTextPropertyReplacement).
 */

import { pdfNativeFontFamilyKey, pdfNativeFontNameKey, pdfNativeFontNameReadsStyle } from "./native-font-identity";
import type {
  PdfNativeDeviceColor,
  PdfNativeTextFontOption,
  PdfNativeTextPropertyDraft,
  PdfNativeTextPropertyProof,
} from "./engine/native/list-text-fidelity";
import { pdfTextVerticalScale } from "./engine/native/text-point-size";
import {
  pdfNativeDrawnRunIsSlanted,
  pdfNativeSlantedTextMatrixLinear,
  pdfNativeSyntheticBoldDrawable,
  pdfNativeTextMatrixIsSlanted,
} from "./native-synthetic-style";
import { isPdfNativeEmbeddedFontKey, pdfNativeEmbeddedFontKey, readPdfNativeEmbeddedFontKey } from "./native-embedded-font-key";
import { classifyFontName } from "./font-resolve";
import { embeddableFontFamily } from "@mrnote/core/web/lib/office-fonts/embeddable-faces.ts";
import type { PdfNativeTextPaintFormat } from "./native-text-format-painter";

/** WPS's default text size, which 清除格式 sets (小四). */
export const NATIVE_TEXT_DEFAULT_POINTS = 12;

/** The Font Size box's steps (Grow Font / Shrink Font move between them). */
export const NATIVE_TEXT_FONT_SIZE_STEPS: readonly number[] = Object.freeze([
  8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72,
]);

/** Word's default tab stops: Increase/Decrease Indent move to the next one. */
const NATIVE_INDENT_TAB_STOP_POINTS = 36;

/** The face a run's font option stands for (one of the PDF's own, or a
 * family's served face named by an embedded-font key). */
export type NativeTextFace = Pick<PdfNativeTextFontOption, "resourceName" | "baseFont" | "family" | "bold" | "italic">;









export function nativeTextServedRegularFamily(face: Pick<NativeTextFace, "family" | "baseFont">): string {
  return embeddableFontFamily(face.family)?.office ?? embeddableFontFamily(face.baseFont)?.office ??
    nativeTextServedClassFamily(face);
}

/** The served family nearest a face's class: Times New Roman for a serif,
 * Courier New for a fixed pitch, else Arial. */
function nativeTextServedClassFamily(face: Pick<NativeTextFace, "family" | "baseFont">): string {
  const fontClass = classifyFontName(face.baseFont || face.family);
  return fontClass === "serif" ? "Times New Roman" : fontClass === "mono" ? "Courier New" : "Arial";
}

/** Whether the served faces a style is set in (Latin, Greek and Cyrillic
 * ones: Arimo, Tinos, Cousine...) draw any of the text. Chinese, Japanese or
 * Korean text alone they do not: such a part is leaned on its own face, as
 * WPS leans it, not set in a served italic face that draws none of it. */
export function nativeTextServedFaceDrawsAny(text: string): boolean {
  return /[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}0-9]/u.test(text);
}

/** Whether two faces are of one family, as the editor chooses a face: by
 * the family key (pdfNativeFontFamilyKey), so a black condensed face finds
 * the page's own regular condensed one, however each names its width. */
export function nativeTextSameFamily(
  left: Pick<NativeTextFace, "family"> & Partial<Pick<NativeTextFace, "baseFont">>,
  right: Pick<NativeTextFace, "family"> & Partial<Pick<NativeTextFace, "baseFont">>,
): boolean {
  return left.family === right.family ||
    pdfNativeFontFamilyKey(left.baseFont || left.family) === pdfNativeFontFamilyKey(right.baseFont || right.family);
}

/** A face as its own weight and slant are read: one of the PDF's fonts, or a
 * font option's fields. */
type NativeTextStyledFace = Pick<NativeTextFace, "family" | "baseFont" | "bold"> & Partial<Pick<NativeTextFace, "italic">>;

/**
 * A face's own weight and slant, as the editor reads them: what its font
 * option says (a style named after a separator, the replay's reading) and
 * what its name says run on with the family or in a TeX family's code
 * ("Univers-CondensedBold", CMTI12). The buttons and the drafts read a face
 * here alike. Read apart -- the buttons by the name, the drafts by the
 * option alone -- 倾斜 stood pressed on a CMTI12 line, a press asked the
 * draft for an upright run of a face it took for upright already, and the
 * line was left with nothing written and nothing said (pdf-walk-1000 0686).
 */
export function nativeTextFaceStyle(
  face: Readonly<{ bold?: boolean; italic?: boolean; baseFont?: string }> | null | undefined,
): Readonly<{ bold: boolean; italic: boolean }> {
  const named = face?.baseFont ? pdfNativeFontNameReadsStyle(face.baseFont) : null;
  return { bold: Boolean(face?.bold) || Boolean(named?.bold), italic: Boolean(face?.italic) || Boolean(named?.italic) };
}

/** A run of a box as the page draws it, as far as its weight and slant go
 * (PdfNativeSourceTextAppearanceRun). */
type NativeTextDrawnRun = Readonly<{
  text: string;
  font: Readonly<{ bold?: boolean; italic?: boolean; baseFont?: string }>;
  syntheticBold?: true;
  textMatrix: readonly number[];
  graphicsMatrix: readonly number[];
}>;

/**
 * A run's weight and slant as the page draws it: its face's own
 * (nativeTextFaceStyle), or drawn in place -- a bold by stroking the glyphs
 * in their fill colour, an italic by leaning the text matrix
 * (native-synthetic-style), as this writer sets a style the PDF holds no
 * face for.
 */
export function nativeTextDrawnRunStyle(run: NativeTextDrawnRun): Readonly<{ bold: boolean; italic: boolean }> {
  const face = nativeTextFaceStyle(run.font);
  return { bold: face.bold || run.syntheticBold === true, italic: face.italic || pdfNativeDrawnRunIsSlanted(run) };
}

/**
 * What a box no one run states (several shows, several lines) reads of its
 * weight and slant, off the runs the page draws it in: bold where every
 * inked run is, italic where every inked run is. The editor's buttons read
 * it so, and a press takes off what it reads: a box this writer leaned or
 * stroked where it stands opens again as leaned or bold, and the same press
 * again sets it back as it was. Neither with no inked run.
 */
export function nativeTextDrawnRunsStyle(runs: readonly NativeTextDrawnRun[] | null | undefined):
  Readonly<{ bold: boolean; italic: boolean }> {
  const inked = (runs ?? []).filter((run) => run.text.trim()).map(nativeTextDrawnRunStyle);
  return { bold: inked.length > 0 && inked.every((style) => style.bold),
    italic: inked.length > 0 && inked.every((style) => style.italic) };
}

/** TeX's italic and slanted text faces by their family's code, each with the
 * upright face of its own weight (the size follows the code: CMTI12, CMR12). */
const TEX_UPRIGHT_OF: Readonly<Record<string, string>> = Object.freeze({
  cmti: "CMR", cmsl: "CMR", cmsltt: "CMTT", cmbxti: "CMBX", cmmi: "CMR", cmmib: "CMBX",
});
/** The design sizes the TeX faces are cut in. */
const TEX_DESIGN_SIZES: readonly number[] = Object.freeze([5, 6, 7, 8, 9, 10, 12, 17]);

/**
 * The upright face of a TeX italic's own family that MRAI serves, at its
 * design size or the one nearest it: CMTI12's is CMR12, the roman the same
 * page sets its upright lines in. Null for a face that is no TeX italic, or
 * whose upright sibling is not served.
 */
function nativeTextTexUprightFamily(face: Pick<NativeTextFace, "family" | "baseFont">): string | null {
  for (const name of [face.baseFont, face.family]) {
    const plain = (name ?? "").replace(/^[A-Z]{6}\+/u, "").replace(/-Identity-[HV]$/u, "").replace(/\s+/gu, "");
    for (const code of [embeddableFontFamily(plain)?.office, plain]) {
      const tex = /^(cm[a-z]+)(\d+)$/iu.exec(code ?? "");
      const upright = tex ? TEX_UPRIGHT_OF[tex[1].toLowerCase()] : undefined;
      if (!tex || !upright) continue;
      const size = Number(tex[2]);
      for (const candidate of [...TEX_DESIGN_SIZES].sort((left, right) => Math.abs(left - size) - Math.abs(right - size))) {
        const served = embeddableFontFamily(`${upright}${candidate}`);
        if (served) return served.office;
      }
    }
  }
  return null;
}

/**
 * A style of a face's own taken off -- its bold, its italic or both -- when
 * the PDF holds no face of its family without it: the key of the served
 * face the run is set in, installed (a stroke cannot un-bold a face, a
 * slant cannot stand an italic up). An italic of a TeX family comes off
 * into that family's own upright face (CMTI12 into CMR12); any other style
 * into the face's own family where MRAI serves it, else into the served
 * family nearest its class (nativeTextServedRegularFamily), in the weight
 * and slant asked. Undefined when the face has no style to take off, or
 * the family has a face of its own without it, which the caller sets
 * instead. One rule for the box's format and a part's, for a bold and for
 * an italic.
 */
export function nativeTextStyleOffServedFaceKey(
  face: NativeTextStyledFace,
  fonts: Iterable<NativeTextStyledFace>,
  want: Readonly<{ bold: boolean; italic: boolean }>,
): string | undefined {
  const own = nativeTextFaceStyle(face);
  const boldOff = own.bold && !want.bold;
  const italicOff = own.italic && !want.italic;
  if (!boldOff && !italicOff) return undefined;
  for (const font of fonts) {
    if (!nativeTextSameFamily(font, face)) continue;
    const style = nativeTextFaceStyle(font);
    if ((want.bold || !style.bold) && (want.italic || !style.italic)) return undefined;
  }
  const served = (italicOff ? nativeTextTexUprightFamily(face) : null) ?? nativeTextServedRegularFamily(face);
  // A served family that is itself a face cut in the style to take off (a
  // TeX italic MRAI serves under its own name, with no upright of its code
  // served) would write that style again: the family of the face's class.
  const still = pdfNativeFontNameReadsStyle(served);
  const family = (boldOff && still.bold) || (italicOff && still.italic) ? nativeTextServedClassFamily(face) : served;
  // A family that is a styled face itself (CMBXTI10's upright is CMBX10,
  // bold as cut) carries that style: it is not asked of it again.
  const carried = pdfNativeFontNameReadsStyle(family);
  return pdfNativeEmbeddedFontKey({ family, bold: want.bold && !carried.bold, italic: want.italic && !carried.italic });
}










export function nativeTextUnboldServedFaceKey(
  face: NativeTextStyledFace,
  fonts: Iterable<NativeTextStyledFace>,
  italic: boolean,
): string | undefined {
  if (!face.bold) return undefined;
  return nativeTextStyleOffServedFaceKey(face, fonts, { bold: false, italic });
}

/**
 * The face of a run's family a weight and slant asked of it are set in, and
 * what that face lacks of them. The PDF's own faces first, and of those
 * only the ones that carry no style the run is not to have (a style of a
 * face's own is not drawn off it): the one with both as asked, else the
 * weight as asked in the run's own slant, else the slant as asked in the
 * run's own weight, else the run's own face; what the face chosen lacks is
 * drawn in place, a bold as a stroke, an italic as a slant. With no such
 * face in the PDF, the served face a style comes off into (`served`,
 * nativeTextStyleOffServedFaceKey).
 */
export function nativeTextStyledFace(
  fonts: readonly NativeTextFace[],
  face: NativeTextFace,
  want: Readonly<{ bold: boolean; italic: boolean }>,
): Readonly<{ face: NativeTextFace; syntheticBold: boolean; slanted: boolean }> | Readonly<{ served: string }> {
  const own = nativeTextFaceStyle(face);
  const fits = (font: NativeTextFace) => {
    const style = nativeTextFaceStyle(font);
    return (want.bold || !style.bold) && (want.italic || !style.italic);
  };
  // Its own face first (a family the page holds in two bold faces).
  const family = fonts.filter((font) => nativeTextSameFamily(font, face) && fits(font))
    .sort((left, right) => Number(right.resourceName === face.resourceName) - Number(left.resourceName === face.resourceName));
  const styled = (bold: boolean, italic: boolean) => family.find((font) => {
    const style = nativeTextFaceStyle(font);
    return style.bold === bold && style.italic === italic;
  });
  const match = styled(want.bold, want.italic) ?? styled(want.bold, own.italic) ?? styled(own.bold, want.italic) ??
    (fits(face) ? face : family[0]);
  if (!match) {
    const served = nativeTextStyleOffServedFaceKey(face, fonts, want);
    // (Always a key here: the face has a style to take off, and no face of
    // its family is without it.)
    if (served) return { served };
    return { face, syntheticBold: false, slanted: false };
  }
  const style = nativeTextFaceStyle(match);
  return { face: match, syntheticBold: want.bold && !style.bold, slanted: want.italic && !style.italic };
}














export function nativeTextNamedStyleOffFace(
  face: NativeTextFace,
  fonts: readonly NativeTextFace[],
  asked: Readonly<{ bold?: boolean; italic?: boolean }>,
): Readonly<{ fontResourceName: string; syntheticBold?: true; slanted?: true }> | undefined {
  const own = nativeTextFaceStyle(face);
  const italicOff = asked.italic === false && own.italic && !face.italic;
  const boldOff = asked.bold === false && own.bold && !face.bold;
  if (!italicOff && !boldOff) return undefined;
  const styled = nativeTextStyledFace(fonts, face, { bold: asked.bold ?? own.bold, italic: asked.italic ?? own.italic });
  if ("served" in styled) return Object.freeze({ fontResourceName: styled.served });
  if (styled.face.resourceName === face.resourceName) return undefined;
  return Object.freeze({ fontResourceName: styled.face.resourceName,
    ...(styled.syntheticBold ? { syntheticBold: true as const } : {}),
    ...(styled.slanted ? { slanted: true as const } : {}) });
}



















export function nativeTextBoxFormatDrawnStyle(input: Readonly<{
  asked: Readonly<{ bold?: boolean; italic?: boolean }>;
  /** A face is named for the format: one of the PDF's own, or a served one. */
  faceNamed: boolean;
  italicFaceStrokesBold: boolean;
  namedOff?: Readonly<{ syntheticBold?: true; slanted?: true }>;
}>): Readonly<{ syntheticBold?: boolean; slanted?: boolean }> {
  const { asked, faceNamed, italicFaceStrokesBold, namedOff } = input;
  return Object.freeze({
    ...(italicFaceStrokesBold ? { syntheticBold: true }
      : asked.bold === false || (asked.bold !== undefined && !faceNamed) ? { syntheticBold: asked.bold }
        : namedOff?.syntheticBold ? { syntheticBold: true } : {}),
    ...(asked.italic !== undefined && !faceNamed ? { slanted: asked.italic } : namedOff?.slanted ? { slanted: true } : {}),
  });
}

/** Whether a PDF font is the editor's default text face, the bundled Noto
 * Sans SC (a subset tag before its name or not). */
export function isNativeTextDefaultFace(baseFont: string | undefined): boolean {
  const name = pdfNativeFontNameKey(baseFont ?? "");
  return name.startsWith("notosanssc") || name.startsWith("notosanscjksc");
}

export const isBlackDeviceColor = (color: PdfNativeDeviceColor): boolean =>
  color.kind === "gray" ? color.components[0] === 0
    : color.kind === "rgb" ? color.components.every((component) => component === 0)
      : color.components[0] === 0 && color.components[1] === 0 && color.components[2] === 0 &&
        color.components[3] === 1;

export function nextNativeIndentStop(
  current: number,
  direction: "increase" | "decrease",
): number {
  const stop = NATIVE_INDENT_TAB_STOP_POINTS;
  const value = Number.isFinite(current) ? Math.max(0, current) : 0;
  const next =
    direction === "increase"
      ? (Math.floor(value / stop + 1e-6) + 1) * stop
      : (Math.ceil(value / stop - 1e-6) - 1) * stop;
  return Math.max(0, Math.round(next * 1000) / 1000);
}

/** Grow Font / Shrink Font from a size on paper: the next step, else a point. */
export function nativeTextFontSizeStep(points: number, delta: number): number {
  return delta > 0
    ? NATIVE_TEXT_FONT_SIZE_STEPS.find((value) => value > points) ?? Math.min(1000, points + 1)
    : [...NATIVE_TEXT_FONT_SIZE_STEPS].reverse().find((value) => value < points) ?? Math.max(1, points - 1);
}

export function nativeTextColorToHex(color: PdfNativeDeviceColor): string {
  const rgb =
    color.kind === "gray"
      ? [color.components[0], color.components[0], color.components[0]]
      : color.kind === "rgb"
        ? [...color.components]
        : [
            1 - Math.min(1, color.components[0] + color.components[3]),
            1 - Math.min(1, color.components[1] + color.components[3]),
            1 - Math.min(1, color.components[2] + color.components[3]),
          ];
  return `#${rgb
    .map((component) =>
      Math.round(Math.max(0, Math.min(1, component)) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;
}

export function nativeTextColorFromHex(value: string): PdfNativeDeviceColor {
  return {
    kind: "rgb",
    components: [
      Number.parseInt(value.slice(1, 3), 16) / 255,
      Number.parseInt(value.slice(3, 5), 16) / 255,
      Number.parseInt(value.slice(5, 7), 16) / 255,
    ],
  };
}

/**
 * Paper points per Tf unit for the run: the y scale of its text matrix and
 * of the page's transformation matrix around it (a "1 Tf" run under a 150
 * matrix inside "0.24 0 0 0.24 cm" draws 36 pt glyphs: Word for Mac). The
 * size of a slanted run is its height across the baseline: the slant leans
 * the glyphs, it does not make them taller (12 pt read 12.27 under a
 * synthetic italic).
 */
export function nativeTextMatrixScale(
  proof: PdfNativeTextPropertyProof | null | undefined,
  draft: PdfNativeTextPropertyDraft | null | undefined,
): number {
  const linear = draft?.textMatrixLinear ?? proof?.original.textMatrixLinear;
  if (!linear) return 1;
  return pdfTextVerticalScale(pdfNativeSlantedTextMatrixLinear(linear, false), proof?.graphicsMatrix);
}

/** Whether the run as the source draws it is a bold drawn in place (filled
 * and stroked, `2 Tr`) in a face that is not bold itself. */
export function nativeTextOriginalSyntheticBold(
  proof: PdfNativeTextPropertyProof,
  selectedFont: NativeTextFace | null | undefined,
): boolean {
  return proof.original.renderingMode === 2 && Boolean(selectedFont) && !selectedFont!.bold;
}

/** The run set whole in a served face named by its key: the stroke and
 * slant the run had in place of a style give way to the face's own (or to
 * the planner's, where it has none). */
function nativeTextWithServedRunFace(
  proof: PdfNativeTextPropertyProof | null | undefined,
  draft: PdfNativeTextPropertyDraft,
  key: string,
  originalSyntheticBold: boolean,
): PdfNativeTextPropertyDraft {
  const linear = draft.textMatrixLinear ?? proof?.original.textMatrixLinear;
  return {
    ...draft,
    fontResourceName: key,
    syntheticBold: originalSyntheticBold ? false : undefined,
    ...(linear && pdfNativeTextMatrixIsSlanted(linear)
      ? { textMatrixLinear: pdfNativeSlantedTextMatrixLinear(linear, false) }
      : {}),
  };
}

/** The run set whole in a family the PDF holds no face of, in a style: the
 * stroke and slant the run had in place of a style give way to the family's
 * own (or to the planner's, where it has none). */
export function nativeTextWithEmbeddedRunFont(
  proof: PdfNativeTextPropertyProof | null | undefined,
  draft: PdfNativeTextPropertyDraft,
  family: string,
  bold: boolean,
  italic: boolean,
  originalSyntheticBold: boolean,
): PdfNativeTextPropertyDraft {
  return nativeTextWithServedRunFace(proof, draft, pdfNativeEmbeddedFontKey({ family, bold, italic }), originalSyntheticBold);
}

/**
 * The run's own face again, as the PDF draws it: when the weight and slant
 * asked are the ones it has there, and the served face the draft stands in
 * is the upright face of its own TeX family an italic taken off set it in.
 * Italic off and on again on a CMTI12 line is that line as it was, nothing
 * to write -- not the served roman leaned. Null otherwise: a served family
 * of the face's class (Arial for a black title) may be the reader's own
 * choice from the font list, and keeps its own bold and italic.
 */
function nativeTextOwnFaceAgain(
  proof: PdfNativeTextPropertyProof,
  draft: PdfNativeTextPropertyDraft,
  editableFont: NativeTextFace,
  want: Readonly<{ bold: boolean; italic: boolean }>,
): PdfNativeTextPropertyDraft | null {
  const original = proof.original.font;
  if (isPdfNativeEmbeddedFontKey(original.resourceName)) return null;
  const own = nativeTextFaceStyle(original);
  const linear = proof.original.textMatrixLinear;
  if (want.bold !== (own.bold || nativeTextOriginalSyntheticBold(proof, original)) ||
    want.italic !== (own.italic || pdfNativeTextMatrixIsSlanted(linear))) return null;
  const upright = nativeTextTexUprightFamily(original);
  if (!upright || readPdfNativeEmbeddedFontKey(editableFont.resourceName)?.family !== upright ||
    nativeTextStyleOffServedFaceKey(original, proof.fonts, nativeTextFaceStyle(editableFont)) !== editableFont.resourceName) return null;
  return { ...draft, fontResourceName: original.resourceName, syntheticBold: undefined, textMatrixLinear: linear };
}

/**
 * Bold and Italic on the whole run: the face of its family that draws as
 * much of the style as the PDF has faces for; what it lacks is drawn in
 * place, a bold as a stroke, an italic as a slant
 * (lib/pdf/native-synthetic-style.ts). A style of the face's own that no
 * face of its family in the PDF is without comes off into a served face
 * (nativeTextStyledFace). Null when the bold would need a stroke and the
 * stroke width in force is unknown.
 */
export function nativeTextStyledDraft(
  input: Readonly<{
    proof: PdfNativeTextPropertyProof;
    draft: PdfNativeTextPropertyDraft;
    editableFont: NativeTextFace;
    originalSyntheticBold: boolean;
    bold: boolean;
    italic: boolean;
  }>,
): PdfNativeTextPropertyDraft | null {
  const { proof, draft, editableFont, originalSyntheticBold, bold, italic } = input;
  if (isPdfNativeEmbeddedFontKey(editableFont.resourceName)) {
    return nativeTextOwnFaceAgain(proof, draft, editableFont, { bold, italic }) ??
      nativeTextWithEmbeddedRunFont(proof, draft, editableFont.family, bold, italic, originalSyntheticBold);
  }
  const styled = nativeTextStyledFace(proof.fonts, editableFont, { bold, italic });
  if ("served" in styled) return nativeTextWithServedRunFace(proof, draft, styled.served, originalSyntheticBold);
  if (styled.syntheticBold && !pdfNativeSyntheticBoldDrawable(proof.original)) return null;
  const linear = draft.textMatrixLinear ?? proof.original.textMatrixLinear;
  return {
    ...draft,
    fontResourceName: styled.face.resourceName,
    syntheticBold: styled.syntheticBold === originalSyntheticBold ? undefined : styled.syntheticBold,
    ...(styled.slanted !== pdfNativeTextMatrixIsSlanted(linear)
      ? { textMatrixLinear: pdfNativeSlantedTextMatrixLinear(linear, styled.slanted) }
      : {}),
  };
}

/**
 * 清除格式 (WPS) on the whole run: back to the default text style -- the
 * default face, 12 pt, black, and no bold, italic, superscript or subscript.
 * The default face is the bundled Noto Sans SC; a run is set in it in place
 * when the page embeds that face, and otherwise keeps its own face at its
 * regular weight (`keptFace`). The rules beside the run are the caller's.
 */
export function nativeTextClearedDraft(
  input: Readonly<{
    proof: PdfNativeTextPropertyProof;
    draft: PdfNativeTextPropertyDraft;
    editableFont: NativeTextFace | null | undefined;
    originalSyntheticBold: boolean;
  }>,
): Readonly<{ draft: PdfNativeTextPropertyDraft; keptFace: boolean }> {
  const { proof, draft, editableFont, originalSyntheticBold } = input;
  const defaultFace = proof.fonts.find((font) =>
    !font.bold && !font.italic && isNativeTextDefaultFace(font.baseFont));
  const regularFace = editableFont && (editableFont.bold || editableFont.italic)
    ? proof.fonts.find((font) =>
      nativeTextSameFamily(font, editableFont) && !font.bold && !font.italic)
    : undefined;
  const face = defaultFace ?? regularFace ?? editableFont;
  const linear = draft.textMatrixLinear ?? proof.original.textMatrixLinear;
  // The draft without the style it asked for (a bold stroke, rules to take).
  const rest = Object.fromEntries(Object.entries(draft).filter(([key]) =>
    key !== "syntheticBold" && key !== "dropTextRules" && key !== "textRuleParts" && key !== "textRuleRanges")) as unknown as PdfNativeTextPropertyDraft;
  const next: PdfNativeTextPropertyDraft = {
    ...rest,
    ...(face ? { fontResourceName: face.resourceName } : {}),
    // The size on paper: the run stores it as Tf under its matrix.
    fontSize: Math.round(NATIVE_TEXT_DEFAULT_POINTS / nativeTextMatrixScale(proof, draft) * 1e6) / 1e6,
    // Black as the run already spells it, else black in RGB.
    color: isBlackDeviceColor(rest.color) ? rest.color : { kind: "rgb", components: [0, 0, 0] },
    textRise: 0,
    ...(originalSyntheticBold ? { syntheticBold: false } : {}),
    ...(pdfNativeTextMatrixIsSlanted(linear)
      ? { textMatrixLinear: pdfNativeSlantedTextMatrixLinear(linear, false) }
      : {}),
  };
  return {
    draft: next,
    keptFace: !defaultFace && Boolean(editableFont) && !isNativeTextDefaultFace(editableFont!.baseFont),
  };
}

/** The format the painter copies from a run (its whole format, on paper). */
export function nativeTextPaintFormatOf(
  input: Readonly<{
    proof: PdfNativeTextPropertyProof;
    draft: PdfNativeTextPropertyDraft;
    editableFont: NativeTextFace | null | undefined;
    bold: boolean;
    italic: boolean;
    color: string | null;
    underline: boolean;
    strikethrough: boolean;
  }>,
): PdfNativeTextPaintFormat {
  const scale = nativeTextMatrixScale(input.proof, input.draft);
  const rise = input.draft.textRise ?? input.proof.original.textRise ?? 0;
  return {
    fontFamily: input.editableFont?.family ?? null,
    bold: input.bold, italic: input.italic,
    fontSizePoints: Math.round(input.draft.fontSize * scale * 100) / 100,
    color: input.color,
    underline: input.underline, strikethrough: input.strikethrough,
    textRisePoints: Math.round(rise * scale * 100) / 100,
  };
}

/**
 * 格式刷 onto the whole run: the face of the painted family the run's PDF
 * has (else the run's own family), the style drawn in place where that face
 * lacks it, the size on paper, the colour and the rise. A value the run
 * already has on paper keeps the run's own spelling (the copied format
 * carries sizes to 0.01 pt and colours as #rrggbb), so painting a run with
 * its own format changes nothing. The rules the format adds or takes away
 * are returned beside the draft; the caller writes them.
 */
export function nativeTextPaintedDraft(
  input: Readonly<{
    proof: PdfNativeTextPropertyProof;
    draft: PdfNativeTextPropertyDraft;
    editableFont: NativeTextFace;
    originalSyntheticBold: boolean;
    format: PdfNativeTextPaintFormat;
    /** The rules standing beside the run, under all of it and under part. */
    rules: Readonly<{ underline: boolean; strikethrough: boolean }>;
    rulesStanding: Readonly<{ underline: boolean; strikethrough: boolean }>;
  }>,
): Readonly<{
  draft: PdfNativeTextPropertyDraft;
  add: readonly ("underline" | "strikethrough")[];
  drop: readonly ("underline" | "strikethrough")[];
}> {
  const { proof, draft: properties, editableFont, originalSyntheticBold, format } = input;
  const name = (value: string) => value.replace(/^[A-Z]{6}\+/, "").replace(/[\s_-]/g, "").toLowerCase();
  const named = format.fontFamily
    ? proof.fonts.filter((font) => name(font.family) === name(format.fontFamily!))
    : [];
  const pool: readonly NativeTextFace[] = named.length
    ? named
    : proof.fonts.filter((font) => nativeTextSameFamily(font, editableFont));
  const face =
    pool.find((font) => font.bold === format.bold && font.italic === format.italic) ??
    pool.find((font) => font.bold === format.bold && !font.italic) ??
    pool.find((font) => !font.bold && !font.italic) ??
    editableFont;
  const linear = properties.textMatrixLinear ?? proof.original.textMatrixLinear;
  const wantSyntheticBold = format.bold && !face.bold && pdfNativeSyntheticBoldDrawable(proof.original);
  const wantSlant = format.italic && !face.italic;
  const rest = Object.fromEntries(Object.entries(properties).filter(([key]) =>
    key !== "syntheticBold" && key !== "dropTextRules" && key !== "addTextRules" &&
    key !== "textRuleParts" && key !== "textRuleRanges")) as unknown as PdfNativeTextPropertyDraft;
  // A rule under part of the run is widened to all of it by an add (the
  // painted format has the rule) and taken away by a drop.
  const add = (["underline", "strikethrough"] as const).filter((style) => format[style] && !input.rules[style]);
  const drop = (["underline", "strikethrough"] as const).filter((style) => !format[style] && input.rulesStanding[style]);
  const round = (value: number) => Math.round(value * 1e6) / 1e6;
  const scale = nativeTextMatrixScale(proof, properties);
  const rise = properties.textRise ?? proof.original.textRise ?? 0;
  const samePoints = (points: number, current: number) => Math.abs(points - current) < 0.005;
  const next: PdfNativeTextPropertyDraft = {
    ...rest,
    fontResourceName: face.resourceName,
    ...(format.fontSizePoints && !samePoints(format.fontSizePoints, properties.fontSize * scale)
      ? { fontSize: round(format.fontSizePoints / scale) } : {}),
    ...(format.color && format.color.toLowerCase() !== nativeTextColorToHex(properties.color).toLowerCase()
      ? { color: nativeTextColorFromHex(format.color) } : {}),
    textRise: samePoints(format.textRisePoints, rise * scale)
      ? rise
      : round(format.textRisePoints / scale),
    ...(wantSyntheticBold !== originalSyntheticBold ? { syntheticBold: wantSyntheticBold } : {}),
    ...(wantSlant !== pdfNativeTextMatrixIsSlanted(linear)
      ? { textMatrixLinear: pdfNativeSlantedTextMatrixLinear(linear, wantSlant) }
      : {}),
  };
  return { draft: next, add, drop };
}











export function nativeTextDraftMayWiden(
  baseline: PdfNativeTextPropertyDraft | null | undefined,
  draft: PdfNativeTextPropertyDraft | null | undefined,
): boolean {
  if (!draft) return false;
  if (!baseline) return true;
  const wider = (next: number | undefined, base: number | undefined) =>
    next !== undefined && next > (base ?? 0) + 1e-9;
  if (wider(draft.fontSize, baseline.fontSize) ||
    wider(draft.characterSpacing, baseline.characterSpacing) ||
    wider(draft.wordSpacing, baseline.wordSpacing) ||
    wider(draft.horizontalScale, baseline.horizontalScale) ||
    draft.fontResourceName !== baseline.fontResourceName ||
    (draft.syntheticBold === true && baseline.syntheticBold !== true)) return true;
  const linear = draft.textMatrixLinear, base = baseline.textMatrixLinear;
  if (linear && base && (Math.hypot(linear[0], linear[1]) > Math.hypot(base[0], base[1]) + 1e-9 ||
    (pdfNativeTextMatrixIsSlanted(linear) && !pdfNativeTextMatrixIsSlanted(base)))) return true;
  const baseSegments = baseline.segments ?? [];
  return (draft.segments ?? []).some((segment) => {
    const same = baseSegments.find((entry) => entry.start === segment.start && entry.end === segment.end);
    if (same && JSON.stringify(same) === JSON.stringify(segment)) return false;
    return wider(segment.fontSize, draft.fontSize) ||
      wider(segment.characterSpacing, draft.characterSpacing ?? baseline.characterSpacing) ||
      (segment.fontResourceName !== undefined && segment.fontResourceName !== draft.fontResourceName) ||
      segment.syntheticBold === true || segment.slanted === true;
  });
}

/** The fields of a property draft that draw the run's glyphs otherwise and
 * move none of them. */
const LOOK_ONLY_FIELDS: ReadonlySet<string> = new Set([
  "color", "syntheticBold", "dropTextRules", "addTextRules", "textRuleParts", "textRuleRanges", "segments",
]);


























export function nativeTextDraftKeepsAdvances(
  baseline: PdfNativeTextPropertyDraft | null | undefined,
  draft: PdfNativeTextPropertyDraft | null | undefined,
  runSlanted: boolean,
  options?: Readonly<{ partFaces?: boolean }>,
): boolean {
  if (!baseline || !draft) return false;
  const placing = (value: PdfNativeTextPropertyDraft) => JSON.stringify(Object.keys(value)
    .filter((key) => !LOOK_ONLY_FIELDS.has(key) && (value as unknown as Record<string, unknown>)[key] !== undefined)
    .sort().map((key) => [key, (value as unknown as Record<string, unknown>)[key]]));
  if (placing(baseline) !== placing(draft)) return false;
  // Part by part: what places a part's glyphs (its face, size, rise, letter
  // spacing, slant), a field left out or stated as the run's own being the
  // run's own.
  const differs = (value: number | undefined, own: number) => value !== undefined && Math.abs(value - own) > 1e-9;
  const placingAt = (owner: PdfNativeTextPropertyDraft, at: number) => {
    const part = (owner.segments ?? []).find((segment) => segment.start <= at && at < segment.end);
    return JSON.stringify([
      options?.partFaces ? null
        : part?.fontResourceName !== undefined && part.fontResourceName !== owner.fontResourceName ? part.fontResourceName : null,
      differs(part?.fontSize, owner.fontSize) ? part!.fontSize : null,
      differs(part?.textRise, owner.textRise ?? 0) ? part!.textRise : null,
      differs(part?.characterSpacing, owner.characterSpacing) ? part!.characterSpacing : null,
      options?.partFaces ? null : part?.slanted ?? runSlanted,
    ]);
  };
  const bounds = new Set<number>();
  for (const segment of [...(baseline.segments ?? []), ...(draft.segments ?? [])]) {
    bounds.add(segment.start);
    bounds.add(segment.end);
  }
  for (const at of bounds) {
    if (placingAt(baseline, at) !== placingAt(draft, at)) return false;
  }
  return true;
}


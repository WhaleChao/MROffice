/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfNativeEmbeddedFontKey } from "./native-embedded-font-key";
import { pdfNativeFontNameReadsStyle } from "./native-font-identity";
import { officeFontSubstitution } from "@mrnote/core/web/lib/office-fonts/substitutes.ts";
import type { PdfNativeParagraphFormat, PdfNativeSourceTextAppearanceRun, PdfNativeTextFontOption } from "./engine/native/list-text-fidelity";

const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

/** The family of a TeX font MRAI serves its very design of (lib/office-fonts/
 * tex-pack): every optical size of Computer Modern, the AMS fonts and Latin
 * Modern, by the name a PDF gives them. The family's stack holds the run's own
 * face first, then its siblings of that size: a digit typed into a math italic
 * is the roman's, an element sign the symbols'. */
function texFamily(name: string | null | undefined): string | null {
  if (!name) return null;
  const face = name.replace(/^[A-Z]{6}\+/u, "").replace(/-Identity-[HV]$/u, "").replace(/\s+/gu, "");
  for (const candidate of [face, face.replace(/-[A-Za-z]+$/u, "")]) {
    const served = officeFontSubstitution(candidate);
    if (served?.substitutes[0]?.startsWith("MRTeX ")) return served.office;
  }
  return null;
}

/** TeX's bolds say so only in their family's code (CMBX10, CMMIB10). */
const TEX_BOLD = /^(?:[A-Z]{6}\+)?(?:cmb\d|cmbx|cmmib|cmbsy|cmssbx)/iu;

const runSize = (run: PdfNativeSourceTextAppearanceRun) => run.fontSize *
  Math.hypot(run.textMatrix[2] ?? 0, run.textMatrix[3] ?? 1) * Math.hypot(run.graphicsMatrix[2] ?? 0, run.graphicsMatrix[3] ?? 1);
/** A run the line's body stands larger than: a TeX script is positioned by a
 * chain of Td, so its size is what says so here (paragraph-rewrite's isScript
 * has the baselines). */
function runIsScript(run: PdfNativeSourceTextAppearanceRun, runs: readonly PdfNativeSourceTextAppearanceRun[]): boolean {
  return runs.some((other) => runSize(run) < runSize(other) * 0.9);
}

/** The served family closest to a run's own, for the characters its subset lacks. */
function similarFamily(name: string, cjk: boolean): string | null {
  const face = name.replace(/^[A-Z]{6}\+/u, "").toLowerCase();
  // The face itself, which a XeLaTeX line's Chinese is set in: what the run's
  // subset lacks is the same outlines, and its Latin stays the run's too.
  const fandol = /^fandol(song|hei|kai|fang)/u.exec(face);
  if (fandol) return `Fandol${fandol[1][0].toUpperCase()}${fandol[1].slice(1)}`;
  // Latin typed into a serif or a TeX face the pack does not hold; into a
  // sans, the served sans of its build (a condensed one's narrow, a fixed
  // pitch one's Courier New), as the draft shows it. Set in the one
  // substitute instead, the letters an exam cover's sans subset lacked came
  
  if (!cjk) return /song|sung|ming|simsun|serif|times|georgia|garamond|^lm|^cm|^mrtex|宋/u.test(face) ? "Times New Roman"
    : pdfNativeDraftLatinFamily(name);
  if (/kai|楷/u.test(face)) return "KaiTi";
  if (/fang|仿宋/u.test(face)) return "FangSong";
  if (/hei|sans|gothic|yahei|dengxian|arial|helvetica|黑/u.test(face)) return null;
  // Song and Ming, and a serif Latin face beside them: a TeX line's Latin Modern
  // and Computer Modern, as Word pairs Times New Roman with 宋体.
  if (/song|sung|ming|simsun|serif|times|georgia|garamond|^lm|^cm|^mrtex|宋/u.test(face)) return "SimSun";
  return null;
}

/**
 * The served families a draft sets a character the run's subset lacks in, as the
 * writer will: its Latin's, then its Chinese's. In the page's own face first, the
 * draft stood a typed 倘 in a system face until the exact page came.
 */








export function pdfNativeEastAsianFaceName(fontName: string | null | undefined): boolean {
  if (!fontName) return false;
  const face = fontName.replace(/^[A-Z]{6}\+/u, "").toLowerCase();
  const compact = face.replace(/[-_ ,.]+/gu, "");
  const family = face.split(/[-,]/u)[0].replace(/[_ .]+/gu, "");
  return /[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af\uf900-\ufaff]/u.test(face) ||
    /notosans(?:mono)?cjk|notoserifcjk|sourcehan|pingfang|hiragino|yahei|jhenghei|dengxian|simhei|simsun|nsimsun|heiti|songti|kaiti|fangsong|stheiti|stsong|stkaiti|stfangsong|stxihei|mingliu|meiryo|yugothic|yumincho|msgothic|msmincho|malgun|nanum|batang|gulim|dotum|wenquanyi|wqy|droidsansfallback|arialunicode|fzhei|fzshusong|fzkai|fzfangsong|hyshu|founder/u.test(compact) ||
    /(?:sans|serif|gothic|mincho|hei|song|kai|ming)(?:cjk)?(?:sc|tc|hk|jp|kr|cn|tw)$/u.test(family);
}

/**
 * The served Latin family a line's draft is shown in where its own face cannot
 * be (PDF.js maps an embedded subset's glyphs away from their characters, and
 * the browser drew the typed words in its default sans): a fixed-pitch face in
 * Courier New's served substitute, a condensed sans in Arial Narrow's, any
 * other sans in Arial's. The writer sets a sans run's missing Latin in the
 * same family (similarFamily), so the draft is what the page draws. Null for a serif, a TeX
 * or a CJK face (those have their own rules) and for an unnamed one.
 */
export function pdfNativeDraftLatinFamily(fontName: string | null | undefined): string | null {
  if (!fontName) return null;
  const face = fontName.replace(/^[A-Z]{6}\+/u, "").toLowerCase();
  if (texFamily(fontName) || pdfNativeEastAsianFaceName(fontName) ||
    /song|sung|ming|simsun|serif|times|georgia|garamond|minion|caslon|bodoni|^lm|^cm|^mrtex|宋|kai|楷|fang|hei|黑|gothic|mincho/u.test(face)) {
    return null;
  }
  // A fixed-pitch face in Courier New's served twin: drafted in Arial, a
  // code line's columns drifted until the page drew it.
  if (/courier|mono(?!type)|consol|menlo|monaco|typewriter|andale|fixedsys|letter ?gothic/u.test(face)) return "Courier New";
  if (/(?:^|[-_ ,])(?:cn|cond|condensed|narrow|compressed|comp)\b|[a-z](?:cn|condensed|narrow|compressed)(?:$|[-_ ,])|boldcn|bdcn|lightcn|blackcn|mediumcn/u.test(face)) {
    return "Arial Narrow";
  }
  return "Arial";
}

export function pdfNativeSimilarFamilies(fontName: string | null | undefined): readonly string[] {
  if (!fontName) return [];
  const families = [texFamily(fontName), similarFamily(fontName, false), similarFamily(fontName, true)].filter((family): family is string => family !== null);
  return [...new Set(families)];
}

/**
 * Characters typed that the run's own subset lacks, set in a served face like the
 * run's instead of the one sans substitute (a 倘 typed into a Kai line stood out
 * as bold beside its thin strokes). The words typed are one part of a paragraph
 * format; null where no served family is closer than the substitute.
 */
export function pdfNativeTypedWordsInSimilarFace(
  original: string,
  text: string,
  runs: readonly PdfNativeSourceTextAppearanceRun[] | null | undefined,
): PdfNativeParagraphFormat | null {
  if (!runs?.length) return null;
  const limit = Math.min(original.length, text.length);
  let prefix = 0;
  while (prefix < limit && original[prefix] === text[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < limit - prefix && original[original.length - 1 - suffix] === text[text.length - 1 - suffix]) suffix += 1;
  const end = text.length - suffix;
  if (!(end > prefix) || !text.slice(prefix, end).trim()) return null;
  const cjk = CJK.test(text.slice(prefix, end));
  // The run the words take their look from, as the paragraph rewrite takes it:
  // the word they are typed against after a blank, else the character before.
  const blankBefore = prefix > 0 && /\s/u.test(original[prefix - 1]);
  const at = (blankBefore && prefix < original.length) || prefix === 0
    ? prefix : prefix - 1;
  // Each character's run, as far as the look depends on them.
  let cursor = 0;
  const runAt: (PdfNativeSourceTextAppearanceRun | undefined)[] = [];
  for (const run of runs) {
    for (const character of run.text.replace(/\s+/gu, "")) {
      while (cursor < original.length && /\s/u.test(original[cursor])) cursor += 1;
      if (!original.startsWith(character, cursor)) { if (runAt[at]) break; return null; }
      runAt[cursor] = run;
      cursor += character.length;
    }
    if (cursor > at + 1) break;
  }
  let ownerAt = at;
  while (ownerAt < original.length && !runAt[ownerAt]) ownerAt += 1;
  let owner = runAt[ownerAt] ?? null;
  if (!owner) return null;
  // Typed after a script, the paragraph rewrite sets the words in the line's
  // body (typedAfter there), unless they stand inside the script.
  if (ownerAt < prefix && runIsScript(owner, runs) && !(runAt[prefix] && runIsScript(runAt[prefix]!, runs))) {
    for (let before = ownerAt - 1; before >= 0; before -= 1) {
      const candidate = runAt[before];
      if (candidate && !runIsScript(candidate, runs)) { owner = candidate; break; }
    }
  }
  let tex = cjk ? null : texFamily(owner.font.baseFont) ?? texFamily(owner.programFace?.family);
  // TeX sets a formula's digits and operators in the roman: a letter typed against
  // that = is the formula's math italic, if the line has one. A roman run holding
  // letters is text, and keeps its own.
  if (tex && /^CMR\d/u.test(tex) && /\p{L}/u.test(text.slice(prefix, end)) && !/\p{L}/u.test(owner.text)) {
    const own = runs.indexOf(owner);
    const italics = runs.flatMap((run, index) => {
      const family = texFamily(run.font.baseFont);
      return family && /^CMMI\d/u.test(family) ? [{ family, distance: Math.abs(index - own) }] : [];
    });
    if (italics.length) tex = italics.reduce((nearest, entry) => (entry.distance < nearest.distance ? entry : nearest)).family;
  }
  const family = tex ?? similarFamily(owner.programFace?.family ?? owner.font.baseFont, cjk);
  if (!family) return null;
  const style = pdfNativeFontNameReadsStyle(owner.font.baseFont);
  return Object.freeze({ parts: Object.freeze([Object.freeze({ start: prefix, end,
    fontResourceName: pdfNativeEmbeddedFontKey({ family, bold: Boolean(owner.font.bold) || style.bold || (tex !== null && TEX_BOLD.test(owner.font.baseFont)),
      italic: !cjk && (Boolean(owner.font.italic) || style.italic) }) })]) });
}













export function pdfNativeParagraphFormatInServedSimilarFaces(
  format: PdfNativeParagraphFormat,
  fontOptions: Iterable<PdfNativeTextFontOption>,
  text: string,
  


  original: string,
): PdfNativeParagraphFormat | null {
  if (original !== text || CJK.test(text) || /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Devanagari}\p{Script=Thai}]/u.test(text)) return null;
  const options = new Map([...fontOptions].map((option) => [option.resourceName, option]));
  let changed = false;
  let refused = false;
  // A weight drawn as a stroke, a slant drawn by the matrix: the served
  // family's own bold and italic faces carry them instead.
  const served = (name: string | undefined, strokeBold?: boolean, slanted?: boolean): string | undefined => {
    if (name === undefined) return undefined;
    const option = options.get(name);
    if (!option) return name;
    const face = option.baseFont || option.family;
    const family = pdfNativeServedSimilarFamily(face);
    if (!family) { refused = true; return name; }
    const style = pdfNativeFontNameReadsStyle(face);
    changed = true;
    return pdfNativeEmbeddedFontKey({ family, bold: Boolean(option.bold) || style.bold || strokeBold === true,
      italic: Boolean(option.italic) || style.italic || slanted === true });
  };
  // (Said off, not left out: a stroke the run had would stay on the new face.)
  const without = <T extends { syntheticBold?: boolean; slanted?: boolean }>(fields: T): T =>
    ({ ...fields, syntheticBold: false, slanted: false });
  const whole = served(format.fontResourceName, format.syntheticBold, format.slanted);
  const parts = format.parts?.map((part) => {
    const name = served(part.fontResourceName, part.syntheticBold ?? format.syntheticBold, part.slanted ?? format.slanted);
    return name === part.fontResourceName ? part : Object.freeze({ ...without(part), fontResourceName: name });
  });
  if (refused || !changed) return null;
  const base = whole !== undefined && whole !== format.fontResourceName ? without(format) : format;
  return Object.freeze({ ...base, ...(whole !== undefined ? { fontResourceName: whole } : {}),
    ...(parts ? { parts: Object.freeze(parts) } : {}) });
}

/** The served family a face is set in when it cannot set a line's words
 * (pdfNativeParagraphFormatInServedSimilarFaces), and the draft shows it in:
 * a serif in Times New Roman's, a sans of its build. Null for a TeX face, a
 * Chinese one, or none named. */
export function pdfNativeServedSimilarFamily(fontName: string | null | undefined): string | null {
  if (!fontName || texFamily(fontName)) return null;
  const face = fontName.replace(/^[A-Z]{6}\+/u, "").toLowerCase();
  if (pdfNativeEastAsianFaceName(fontName)) return null;
  return /song|sung|ming|simsun|serif|times|georgia|garamond|^lm|^cm|^mrtex|宋/u.test(face) ? "Times New Roman"
    : pdfNativeDraftLatinFamily(fontName);
}

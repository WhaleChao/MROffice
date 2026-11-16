/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text measurement for the layout engine. The browser host passes a canvas
 * measurer; tests and Node hosts use the metric approximation below
 * (full-width CJK one em, ASCII by its Aptos advance), which is what decides whether a
 * label fits and so which font size a node set settles on.
 */

/** A line's height as a multiple of its font, by what the line holds. */
import { OFFICE_SCENE_LABEL_LINE_SPACING } from "../model";
import { splitBehind } from "../../regexp-behind";

// After a hyphen (`(?<=-)`, tested in code: Safari before 16.4 cannot
// compile a lookbehind).
const AFTER_HYPHEN = { char: /-/ } as const;
export type LineFactors = { cjk: number; cjkEastAsian: number; latin: number; latinBesideCjk: number };

/** Width in points of `text` at `fontPt`; with the line heights of the face it measures, when not Aptos's. */
export type MeasureText = ((text: string, fontPt: number) => number) & {
  lines?: LineFactors;
  /** The same words set bold, where the face's bold advances are known (`latin-face`). */
  bold?: (text: string, fontPt: number) => number;
  /** The measure of words set in another face, where it is recorded (`latin-face`). */
  face?: (family: string) => MeasureText | undefined;
};

const POINTS_PER_MM = 72 / 25.4;

/**
 * Printable ASCII (U+0020-U+007E) in thousandths of an em: the advances of
 * Aptos, the theme font of the decks MRAI writes and of new Office decks.
 * Word-by-word, not per class: a label of narrow letters ("Engineering",
 * "Right") fits where an average letter would not, and PowerPoint sizes it
 * by its real width.
 */
const ASCII_ADVANCES = [
  203, 293, 370, 537, 534, 826, 643, 210, 293, 293, 457, 534, 286, 340, 286, 339,
  534, 534, 534, 534, 534, 534, 534, 534, 534, 534, 286, 286, 534, 534, 534, 501,
  895, 589, 604, 692, 686, 556, 524, 708, 707, 260, 331, 568, 500, 790, 706, 732,
  577, 732, 606, 566, 479, 681, 585, 892, 553, 540, 516, 294, 339, 294, 534, 460,
  552, 531, 561, 525, 561, 527, 301, 484, 551, 239, 239, 487, 260, 853, 551, 552,
  561, 561, 334, 486, 323, 559, 452, 721, 442, 452, 438, 294, 270, 294, 534,
];

function charEm(code: number): number {
  if (code >= 0x20 && code <= 0x7e) return ASCII_ADVANCES[code - 0x20] / 1000;
  if (
    (code >= 0x2e80 && code <= 0x9fff) ||
    (code >= 0xac00 && code <= 0xd7af) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xff00 && code <= 0xffef) ||
    (code >= 0x3000 && code <= 0x303f)
  ) {
    return 1;
  }
  
  
  if ((code >= 0x600 && code <= 0x77f) || (code >= 0xfb50 && code <= 0xfeff)) return 0.3;
  return 0.55;
}

/** Width in points of `text` at `fontPt`, from per-class advances. */
export const approximateMeasure: MeasureText = (text, fontPt) => {
  let em = 0;
  for (const ch of text) em += charEm(ch.codePointAt(0) ?? 0x20);
  return em * fontPt;
};

/** Printable ASCII in thousandths of an em: the advances of DengXian (等线). */
const DENGXIAN_ASCII_ADVANCES = [
  274, 284, 363, 597, 527, 500, 749, 222, 292, 292, 411, 667, 219, 500, 219, 378,
  527, 527, 527, 527, 527, 527, 527, 527, 527, 527, 219, 219, 667, 667, 667, 439,
  949, 636, 559, 620, 688, 503, 477, 677, 693, 247, 340, 553, 464, 865, 729, 757,
  552, 757, 577, 514, 518, 667, 612, 915, 576, 546, 570, 292, 373, 292, 667, 415,
  258, 501, 574, 453, 574, 514, 294, 574, 550, 224, 224, 470, 224, 841, 550, 573,
  574, 574, 339, 406, 318, 551, 466, 700, 440, 468, 458, 292, 232, 292, 667,
];

/**
 * A deck whose theme sets its Latin face to DengXian (等线, the theme of
 * Chinese Office's new decks) lines its words by DengXian, not Aptos: an East
 * Asian line 1.22 of the font, a Latin one 1.28. Measured on PowerPoint's own
 * layouts of 44 zh-CN diagrams in that theme (the 14 pt "计划阶段 / Build /
 * Testing and review" sets one to three points under the Aptos theme's) and of
 * Basic Cycle's "第一步" (34 pt in a 91 pt circle's square, two lines).
 */
export const dengXianMeasure: MeasureText = Object.assign(
  (text: string, fontPt: number) => {
    let em = 0;
    for (const ch of text) {
      const code = ch.codePointAt(0) ?? 0x20;
      em += code >= 0x20 && code <= 0x7e ? DENGXIAN_ASCII_ADVANCES[code - 0x20] / 1000 : charEm(code);
    }
    return em * fontPt;
  },
  { lines: { cjk: 1.22, cjkEastAsian: 1.22, latin: 1.28, latinBesideCjk: 1.28 } },
);

/**
 * DengXian's measure in a deck whose theme states no East Asian face of its
 * own (`a:ea typeface=""`, Chinese Office's default theme, its `Hans` script
 * face 等线): Chinese words in a run marked as Chinese take 等线 and stand
 * 1.22 of the font, but Chinese words in a run marked as another language
 * (`lang="en-US" altLang="zh-CN"`, as the golden 39-layout deck types them)
 * take the theme's empty `ea` slot and fall back to the system's face, 1.6
 * as Excel's fallback lines stand. Measured on the golden deck's 54 labels
 * against PowerPoint: 33 sized more than a point too large at 1.22, 6 at 1.6
 * (mean error 3.4 -> 0.6 pt); the 43 zh-CN-marked labels of the DengXian
 * sets are unchanged.
 */
export const dengXianNoEastAsianFaceMeasure: MeasureText = Object.assign(
  (text: string, fontPt: number) => dengXianMeasure(text, fontPt),
  { lines: { cjk: 1.6, cjkEastAsian: 1.22, latin: 1.28, latinBesideCjk: 1.28 } },
);

// Where a line may break: white space but the no-break spaces (U+00A0,
// U+2007, U+202F), which PowerPoint keeps with the words either side

// word to PowerPoint, fitted at 13 pt where broken in two it took 25), and
// East Asian characters.
const BREAKABLE = /([^\S\u00A0\u2007\u202F]+)|([⺀-鿿가-힯豈-﫿＀-￯　-〿])/u;

/** Break one paragraph into the lines it occupies at `widthPt`, words kept whole. */
export function wrapLines(
  text: string,
  fontPt: number,
  widthPt: number,
  measure: MeasureText,
  /** Break a word longer than the whole line where the line ends (see below). */
  breakLongWords = false,
): string[] {
  if (!text) return [""];
  // A line break (`a:br`) ends the line whatever room is left.
  if (text.includes("\n")) return text.split("\n").flatMap((part) => wrapLines(part, fontPt, widthPt, measure, breakLongWords));
  const pieces: string[] = [];
  let rest = text;
  while (rest.length) {
    const match = BREAKABLE.exec(rest);
    if (!match) {
      pieces.push(...splitBehind(rest, /(?=[^-])/, AFTER_HYPHEN));
      break;
    }
    // A word breaks after a hyphen when it must, as PowerPoint wraps
    // "draft-ietf-quic-multipath" in process3's narrow step.
    if (match.index > 0) pieces.push(...splitBehind(rest.slice(0, match.index), /(?=[^-])/, AFTER_HYPHEN));
    pieces.push(match[0]);
    rest = rest.slice(match.index + match[0].length);
  }
  const lines: string[] = [];
  let line = "";
  for (const piece of pieces) {
    const candidate = line + piece;
    if (line && measure(candidate.trimEnd(), fontPt) > widthPt) {
      lines.push(line.trimEnd());
      line = piece.trimStart();
    } else {
      line = candidate;
    }
    // A word longer than the whole line is broken where the line ends, once
    // no font rule can make it fit (`breakLongWords`), as PowerPoint breaks
    
    // p4: "https://www.marylandattorneygeneral.gov/…" at 13 pt over three
    // lines). Kept whole, it could never fit, and every rule ran out trying:
    // the items shrank to slivers 72 m tall. While a font rule can still
    // step, the word stays whole and the font shrinks (vList2's titles).
    while (breakLongWords && line.length > 1 && measure(line.trimEnd(), fontPt) > widthPt) {
      const head = longestFittingHead(line, fontPt, widthPt, measure);
      if (head <= 0 || head >= line.length) break;
      lines.push(line.slice(0, head));
      line = line.slice(head);
    }
  }
  lines.push(line.trimEnd());
  return lines;
}

/** {@link BREAKABLE} as a `split` pattern that keeps each break as one piece of its own. */
const BREAKABLE_SPLIT = new RegExp(`(${BREAKABLE.source.replace(/\((?!\?)/g, "(?:")})`, "u");

/**
 * A paragraph of runs at different sizes broken into lines at `widthPt`, words kept whole:
 * each piece measured at its run's size, each line with the largest size it holds. `null`
 * when a word is wider than the line.
 */
export function wrapSizedRuns(
  runs: readonly { text: string; pt?: number }[],
  fontPt: number,
  widthPt: number,
  measure: MeasureText,
): { text: string; pt: number }[] | null {
  type Piece = { text: string; pt: number; space: boolean; br?: boolean };
  const pieces: Piece[] = [];
  for (const run of runs) {
    const pt = run.pt ?? fontPt;
    if (run.text === "\n") {
      pieces.push({ text: "", pt, space: false, br: true });
      continue;
    }
    for (const part of run.text.split(BREAKABLE_SPLIT)) {
      if (part) pieces.push({ text: part, pt, space: !part.trim() });
    }
  }
  const lines: { text: string; pt: number }[] = [];
  let line: Piece[] = [];
  const widthOf = (items: readonly Piece[]) => {
    let end = items.length;
    while (end > 0 && items[end - 1].space) end -= 1;
    return items.slice(0, end).reduce((total, item) => total + measure(item.text, item.pt), 0);
  };
  const close = () => {
    const worded = line.filter((item) => !item.space);
    lines.push({
      text: line.map((item) => item.text).join("").trim(),
      pt: worded.length ? Math.max(...worded.map((item) => item.pt)) : (line[0]?.pt ?? fontPt),
    });
    line = [];
  };
  for (const piece of pieces) {
    if (piece.br) {
      close();
      continue;
    }
    if (!piece.space && line.some((item) => !item.space) && widthOf([...line, piece]) > widthPt) close();
    if (!line.length && piece.space) continue;
    line.push(piece);
    if (!piece.space && measure(piece.text, piece.pt) > widthPt + 0.01) return null;
  }
  if (line.length || !lines.length) close();
  return lines;
}

/** How many leading characters of `text` fit `widthPt` (at least one). */
function longestFittingHead(text: string, fontPt: number, widthPt: number, measure: MeasureText): number {
  const chars = Array.from(text);
  let low = 1;
  let high = chars.length;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    if (measure(chars.slice(0, middle).join(""), fontPt) <= widthPt) low = middle;
    else high = middle - 1;
  }
  return chars.slice(0, low).join("").length;
}

/** Whether a theme face is DengXian (等线, 等线 Light). */
export function isDengXianFace(face: string | null | undefined): boolean {
  return !!face && /^(等线|dengxian)(\s|$)/i.test(face.trim());
}

/** Whether a face holds East Asian glyphs (its name in CJK, or one of the CJK families Office ships or falls back to). */
export function isEastAsianFace(face: string | null | undefined): boolean {
  if (!face || !face.trim()) return false;
  if (CJK.test(face)) return true;
  return /(dengxian|simsun|simhei|yahei|kaiti|fangsong|pingfang|hiragino|heiti|songti|^st(song|heiti|kaiti|fangsong|xihei|zhongsong)|ms p?gothic|yu gothic|apple sd gothic|ms p?mincho|yu mincho|meiryo|malgun|batang|gulim|dotum|gungsuh|noto (sans|serif) cjk|source han|mingliu|jhenghei|biz ud|nanum|lisu|youyuan)/i.test(face.trim());
}

/**
 * How far apart PowerPoint stands the lines of a SmartArt label's Chinese
 * words (its 90% spacing included), by the face that draws them: the face's
 * own line, not the 1.2 em every Latin line takes. The golden deck laid out
 * by PowerPoint from its data (drawings stripped): "项目 1" at 65 pt in Lined
 * List's narrow label, one character a line -- zh-CN words in the theme's
 * 等线 79 pt apart (1.215), the same words typed in en-US drawn by Microsoft
 * YaHei 101 pt apart (1.554). Drawn at 0.9 x 1.2 such labels stood a fifth
 * to a third shorter than PowerPoint's.
 */
export function smartArtEastAsianLinePitch(face: string | null | undefined): number {
  if (isDengXianFace(face)) return 1.22;
  if (face && /(microsoft yahei|微软雅黑)/i.test(face.trim())) return 1.55;
  return TEXT_LINE_FACTORS.cjkEastAsian;
}

/**
 * The lines of Chinese words a theme gives no East Asian face (its `Hans`
 * and `ea` faces Latin or empty): they fall back to the system's CJK face and
 * stand 1.6 of the font. Measured on Excel's own drawings of every SmartArt
 * layout (an Excel for Mac zh-CN workbook whose theme states Calibri for
 * `ea`): 122 of 154 labels sized exactly and 151 within 2 pt at 1.6, against
 * 9 and 69 at the East Asian face's 1.3.
 */
export const FALLBACK_CJK_LINES: LineFactors = { cjk: 1.6, cjkEastAsian: 1.6, latin: 1.1, latinBesideCjk: 1.1 };

const SMARTART_LINE_SPACING = OFFICE_SCENE_LABEL_LINE_SPACING;

export type FitParagraph = {
  text: string;
  level: number;
  eastAsian?: boolean;
  /** Set in a face of its own, unlike the diagram's (the point's first run's `a:latin`). */
  face?: string;
  /** Set bold by the author (every run of the point `b="1"`): fitted by the bold face's advances. */
  bold?: boolean;
  /** Set in capitals by the author (`cap="all"`): fitted as capitals. */
  capitals?: boolean;
  




  lineSpacing?: number;
  /** The size this paragraph's words are fixed at (`custT`), whatever the shape's size settles on. */
  fixedPt?: number;
  /** With a fixed size, the size of each line its breaks make, where they differ. */
  linePts?: readonly (number | undefined)[];
  /** With a fixed size, its runs at the sizes they state, where those differ within it. */
  runPts?: readonly { text: string; pt?: number }[];
  /** With a fixed size, the space after it its own `a:spcAft` states, in points. */
  spaceAfterPt?: number;
};

export const CJK = /[⺀-鿿가-힯豈-﫿＀-￯　-〿]/u;
/** Whether words hold East Asian characters. */
export const EAST_ASIAN = CJK;













export const TEXT_LINE_FACTORS = { cjk: 1.6, cjkEastAsian: 1.3, latin: 1.105 };










export function lineFactorOf(text: string, eastAsian = false, lines: Omit<LineFactors, "latinBesideCjk"> = TEXT_LINE_FACTORS): number {
  return CJK.test(text) ? (eastAsian ? lines.cjkEastAsian : lines.cjk) : lines.latin;
}

/**
 * A Latin-only line of a paragraph that also holds East Asian words, as a
 * multiple of its font: the Latin face of a deck set for East Asian text
 * stands no taller than a Latin one. Measured on the 51- and 8-diagram sets
 * (Chinese labels "项目 N" wrapping their number onto its own line), with
 * their Chinese lines at the fallback face's 1.6: 1.1 sizes 43 of 58 labels
 * exactly, 1.12 40 and 1.15 38 (Hierarchy List's "项目 1" at 32 pt in two
 * lines, 86.4 of its 86.7 pt circle, as PowerPoint draws it).
 */
export const LATIN_LINE_BESIDE_CJK = 1.1;

















export function bulletMarginPt(fontPt: number): number {
  if (fontPt < 12) return 4.5;
  if (fontPt < 16) return 9;
  if (fontPt < 20) return 13.5;
  if (fontPt < 28) return 18;
  return 22.5 + 4.5 * Math.floor(Math.max(0, fontPt - 36) / 8 + (fontPt >= 36 ? 1 : 0));
}

export function paragraphSpaceAfter(fontPt: number, level: number): number {
  return level > 0 ? fontPt * 0.2 : fontPt * 1.2 * 0.35;
}

/** Whether paragraphs fit a box (mm) at a size, and the widest single word fits a line. */
export function textFits(
  paragraphs: readonly FitParagraph[],
  fontPt: number,
  secondaryPt: number,
  boxWidthMm: number,
  boxHeightMm: number,
  measure: MeasureText,
  lineSpacing?: number,
  /** The first level the shape bullets (`stBulletLvl` from its lowest):
   * those paragraphs are fitted after their bullet, as they are drawn. */
  bulletFrom?: number,
  /** Words longer than a line break where it ends (`wrapLines`). */
  breakLongWords = false,
): boolean {
  const widthPt = boxWidthMm * POINTS_PER_MM;
  const heightPt = boxHeightMm * POINTS_PER_MM;
  if (widthPt <= 0 || heightPt <= 0) return false;
  let used = 0;
  const diagramMeasure = measure;
  for (const [at, paragraph] of paragraphs.entries()) {
    // (A point set in a face of its own fits by it: List 1's title in Aptos Display over
    
    const measure = (paragraph.face && diagramMeasure.face?.(paragraph.face)) || diagramMeasure;
    const size = paragraph.fixedPt ?? (paragraph.level > 0 ? secondaryPt : fontPt);
    // A bulleted first-level paragraph (Picture Accent List's "• 项目2", the
    // shape's only words) stands after the bullet the scene draws: fitted
    // without it, 34 pt filled the line and the drawn words broke onto two
    // (Excel: 27 pt, one line). Its margin is the table's, its sublevels a step
    
    const indent = bulletFrom === 0 ? bulletMarginPt(size) * (paragraph.level + 1) : paragraph.level > 0 ? bulletMarginPt(size) * paragraph.level : bulletFrom !== undefined && paragraph.level >= bulletFrom ? bulletMarginPt(size) : 0;
    
    // "Integrantes" fills the line at 51 pt in Arial, 47 in Arial Bold).
    const face: MeasureText = paragraph.bold && measure.bold ? Object.assign(measure.bold, { lines: measure.lines }) : measure;
    const eastAsian = CJK.test(paragraph.text);
    const words = paragraph.capitals ? paragraph.text.toUpperCase() : paragraph.text;
    // Runs of different sizes the author fixed (custT) wrap as they are set: each word at
    
    // headings before 20 pt asides, measured whole at 31 they took 1.2 times the height).
    if (paragraph.runPts && paragraph.fixedPt !== undefined && !paragraph.capitals && !paragraph.bold) {
      const lines = wrapSizedRuns(paragraph.runPts, size, widthPt - indent, face);
      if (!lines) return false;
      for (const line of lines) {
        const factor = eastAsian && !CJK.test(line.text)
          ? (measure.lines?.latinBesideCjk ?? LATIN_LINE_BESIDE_CJK)
          : lineFactorOf(line.text, paragraph.eastAsian, measure.lines);
        const own = paragraph.lineSpacing;
        used += line.pt * (lineSpacing ?? (own !== undefined ? (factor * own) / SMARTART_LINE_SPACING : factor));
      }
      if (at < paragraphs.length - 1) used += paragraph.spaceAfterPt ?? paragraphSpaceAfter(size, bulletFrom !== undefined && paragraph.level >= bulletFrom ? Math.max(1, paragraph.level) : paragraph.level);
      if (used > heightPt + 0.01) return false;
      continue;
    }
    const broken = paragraph.linePts && paragraph.fixedPt !== undefined ? words.split("\n") : [words];
    for (const [row, part] of broken.entries()) {
      const pt = (broken.length > 1 ? paragraph.linePts?.[row] : undefined) ?? size;
      const lines = wrapLines(part, pt, widthPt - indent, face, breakLongWords);
      if (lines.some((line) => face(line, pt) > widthPt - indent + 0.01)) return false;
      // Each line stands as tall as the faces it holds: a wrapped label's
      // Latin-only line ("1" under "项目") is a Latin line (AlternatingHexagons'
      // "项目 1" fits PowerPoint's 24 pt in two lines of 64 pt).
      for (const line of lines) {
        const factor = eastAsian && !CJK.test(line)
          ? (measure.lines?.latinBesideCjk ?? LATIN_LINE_BESIDE_CJK)
          : lineFactorOf(line, paragraph.eastAsian, measure.lines);
        const own = paragraph.lineSpacing;
        used += pt * (lineSpacing ?? (own !== undefined ? (factor * own) / SMARTART_LINE_SPACING : factor));
      }
    }
    // A first-level paragraph the shape bullets is a bullet for its spacing too
    
    if (at < paragraphs.length - 1) used += paragraphSpaceAfter(size, bulletFrom !== undefined && paragraph.level >= bulletFrom ? Math.max(1, paragraph.level) : paragraph.level);
    if (used > heightPt + 0.01) return false;
  }
  return true;
}

export { POINTS_PER_MM };

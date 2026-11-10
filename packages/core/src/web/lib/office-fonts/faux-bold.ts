/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Bold text in an Office family that ships no bold face.
//
// Office draws such a family's bold by emboldening its regular (宋体/SimSun,
// 黑体, MS Gothic, Batang, ...: Excel for Mac's bold 宋体 heading is a heavy
// SimSun, r9/P18_48539). The alias stylesheet declares no bold face for them
// (scripts/sync-office-fonts.mjs, officeHasBold), leaving the bold to the
// browser's synthesis -- which WebKit and Firefox draw, but Chrome on macOS

// font): the bold text drew regular. The family list below is that
// rule over local-names.json (a test keeps the two equal); a browser found
// not to synthesize gets a text stroke in its place, set once on the root as
// --office-faux-bold-stroke (0 until measured, and everywhere synthesis works).

const NO_BOLD_FAMILIES: ReadonlySet<string> = new Set([
  "Abadi MT Condensed Extra Bold", "Arial Rounded MT Bold", "Avenir Next LT Pro Demi", "Baskerville Old Face", "Batang",
  "BatangChe", "Bauhaus 93", "Bernard MT Condensed", "Bookshelf Symbol 7", "Braggadocio",
  "Britannic Bold", "Cambria Math", "Century", "Colonna MT", "Cooper Black",
  "Copperplate Gothic Bold", "Curlz MT", "Desdemona", "Dotum", "DotumChe", "Edwardian Script ITC",
  "FangSong", "Franklin Gothic Book", "Franklin Gothic Demi", "Franklin Gothic Demi Cond",
  "Franklin Gothic Heavy", "Gabriola", "Gill Sans MT Condensed", "Gill Sans MT Ext Condensed Bold",
  "Gill Sans Ultra Bold", "Gloucester MT Extra Condensed", "Gulim", "GulimChe", "Gungsuh",
  "GungsuhChe", "HGGothicE", "HGMaruGothicMPRO", "HGMinchoE", "HGPGothicE", "HGPMinchoE",
  "HGPSoeiKakugothicUB", "HGP創英角ｺﾞｼｯｸUB", "HGP明朝E", "HGPｺﾞｼｯｸE", "HGSGothicE", "HGSMinchoE",
  "HGSSoeiKakugothicUB", "HGSoeiKakugothicUB", "HGS創英角ｺﾞｼｯｸUB", "HGS明朝E", "HGSｺﾞｼｯｸE",
  "HG丸ｺﾞｼｯｸM-PRO", "HG創英角ｺﾞｼｯｸUB", "HG明朝E", "HGｺﾞｼｯｸE", "Haettenschweiler", "Harrington",
  "Helvetica Neue Condensed",
  "Imprint MT Shadow", "KaiTi", "Kino MT", "Lucida Blackletter", "Lucida Calligraphy",
  "Lucida Console", "Lucida Handwriting", "Lucida Sans Unicode", "MS Gothic", "MS Mincho",
  "MS PGothic", "MS PMincho", "MS Reference Sans Serif", "MS Reference Specialty", "MS UI Gothic",
  "MT Extra", "Marlett", "Matura MT Script Capitals", "Microsoft Himalaya", "Microsoft Yi Baiti",
  "MingLiU", "MingLiU-ExtB", "MingLiU_HKSCS", "MingLiU_HKSCS-ExtB", "Mistral", "Modern No. 20", "Monaco",
  "Mongolian Baiti", "Monotype Corsiva", "Monotype Sorts", "NSimSun", "Nyala", "Onyx", "PMingLiU",
  "PMingLiU-ExtB", "STHupo", "STLiti", "STXingkai", "STXinwei", "STZhongsong", "Segoe UI Historic",
  "Segoe UI Symbol", "SimHei", "SimSun", "SimSun-ExtB", "Stencil", "Symbol",
  "Tw Cen MT Condensed Extra Bold", "Webdings", "Wide Latin", "Wingdings", "Wingdings 2",
  "Wingdings 3", "仿宋", "华文中宋", "华文新魏", "华文琥珀", "华文行楷", "华文隶书", "宋体", "新宋体", "新細明體", "新細明體-ExtB",
  "楷体", "細明體", "細明體-ExtB", "細明體_HKSCS", "細明體_HKSCS-ExtB", "黑体", "굴림", "굴림체", "궁서", "궁서체", "돋움",
  "돋움체", "바탕", "바탕체", "ＭＳ ゴシック", "ＭＳ 明朝", "ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝",
].map((family) => family.toLowerCase()));

/** The families the list holds (for the test that derives it again). */
export const OFFICE_NO_BOLD_FAMILIES: readonly string[] = [...NO_BOLD_FAMILIES];

/** The stroke width a faux-bold run carries: the root's measured value, else none. */
export const OFFICE_FAUX_BOLD_STROKE_WIDTH = "var(--office-faux-bold-stroke, 0px)";

/** About one pixel at 11pt, as Office's emboldening thickens the regular strokes. */
const FAUX_BOLD_STROKE = "0.06em";

/** The first family of a CSS font-family list, unquoted. */
function primaryFamily(fontFamily: string): string {
  const first = fontFamily.split(",")[0] ?? "";
  return first.trim().replace(/^["']|["']$/g, "").trim();
}

/**
 * Whether bold text set in `fontFamily` (a CSS family list; its first family
 * is the one asked for) is an Office family with no bold face, emboldened
 * from its regular.
 */
export function officeFamilyHasNoBold(fontFamily: string | null | undefined): boolean {
  if (typeof fontFamily !== "string" || !fontFamily) return false;
  return NO_BOLD_FAMILIES.has(primaryFamily(fontFamily).toLowerCase());
}

let probed = false;

/**
 * Measure once whether this browser synthesizes bold for a face that has
 * none, and when it does not, give faux-bold runs their stroke. A web font
 * declared at 400 only is drawn at 400 and 700 on a canvas: the same ink
 * means no synthesis. Idempotent; does nothing outside a browser.
 */
export function ensureOfficeFauxBoldProbe(): void {
  if (probed || typeof document === "undefined" || typeof FontFace === "undefined") return;
  probed = true;
  const face = new FontFace("mrai-faux-bold-probe", "url(/fonts/office/archivo-black/400-latin.woff2) format(\"woff2\")", { weight: "400" });
  void face.load().then((loaded) => {
    document.fonts.add(loaded);
    try {
      if (!synthesizesBold()) document.documentElement.style.setProperty("--office-faux-bold-stroke", FAUX_BOLD_STROKE);
    } finally {
      document.fonts.delete(loaded);
    }
  }).catch(() => {});
}

function synthesizesBold(): boolean {
  const canvas = document.createElement("canvas");
  canvas.width = 160;
  canvas.height = 40;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return true;
  const ink = (weight: number): number => {
    context.clearRect(0, 0, canvas.width, canvas.height);
    context.font = `${weight} 28px "mrai-faux-bold-probe"`;
    context.fillStyle = "#000";
    context.textBaseline = "top";
    context.fillText("HIIH", 4, 4);
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    let sum = 0;
    for (let index = 3; index < pixels.length; index += 4) sum += pixels[index];
    return sum;
  };
  const regular = ink(400);
  // Unreadable canvas (fingerprinting protection): assume the browser draws it.
  return regular === 0 || ink(700) > regular * 1.05;
}

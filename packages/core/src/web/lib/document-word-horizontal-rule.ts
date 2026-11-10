/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's horizontal line (Insert > Horizontal Line, and Pandoc's thematic
 * break): a VML `v:rect o:hr="t"` that is the only content of its paragraph.
 *
 * Measured in Word for Mac: the host paragraph is laid out exactly as an
 * empty paragraph (one line of its paragraph mark's height, with its own
 * spacing). The rule's box bottom sits at the bottom of that single line, so
 * the extra leading of a multiple line spacing goes below it; a rule taller
 * than the line minus 5pt grows the line until 5pt stay above it; exact line
 * spacing puts the box bottom at 80% of the line and at-least spacing at the
 * line's bottom. The rule spans the paragraph's text column scaled by
 * `o:hrpct` (tenths of a percent). A shaded rule is a groove ring shaded from
 * #5b5b5b to #c5c5c5; `o:hrnoshade` draws a solid `fillcolor` bar 0.4pt
 * shorter than the box, white when no fill color is given (VML's default).
 *
 * The server projection (projection.py `UniversalHorizontalRule`) reads the
 * same object; both importers emit it as the `wordHorizontalRule` attr of the
 * blank paragraph that hosts it.
 */

export interface WordHorizontalRule {
  version: 1;
  heightPt: number;
  widthPercent: number;
  align: "left" | "center" | "right";
  shade: boolean;
  color: string | null;
}

const VML_NS = "urn:schemas-microsoft-com:vml";
const VML_OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const ALIGNS = new Set(["left", "center", "right"]);
const HEX_COLOR = /^#[0-9a-f]{6}$/;
const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: "#000000", silver: "#c0c0c0", gray: "#808080", grey: "#808080",
  white: "#ffffff", maroon: "#800000", red: "#ff0000", purple: "#800080",
  fuchsia: "#ff00ff", green: "#008000", lime: "#00ff00", olive: "#808000",
  yellow: "#ffff00", navy: "#000080", blue: "#0000ff", teal: "#008080",
  aqua: "#00ffff",
};
function finite(value: unknown, min: number, max: number, exclusiveMin = true): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  if (exclusiveMin ? value <= min : value < min) return null;
  return value <= max ? value : null;
}

/** The validated attr value, or null. Invalid provenance draws nothing. */
export function normalizeWordHorizontalRule(value: unknown): WordHorizontalRule | null {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (raw.version !== 1) return null;
  const heightPt = finite(raw.heightPt, 0, 1584);
  const widthPercent = finite(raw.widthPercent, 0, 100);
  const align = typeof raw.align === "string" && ALIGNS.has(raw.align) ? raw.align as WordHorizontalRule["align"] : null;
  const shade = typeof raw.shade === "boolean" ? raw.shade : null;
  const color = raw.color === null || raw.color === undefined
    ? null
    : typeof raw.color === "string" && HEX_COLOR.test(raw.color) ? raw.color : undefined;
  if (heightPt === null || widthPercent === null || align === null || shade === null || color === undefined) return null;
  return { version: 1, heightPt, widthPercent, align, shade, color: shade ? null : color };
}

function vmlColor(value: string | null): string | null {
  // `black [3213]`: a color followed by its scheme index.
  const text = (value ?? "").trim().toLowerCase().replace(/\s*\[[0-9]+\]$/, "");
  if (HEX_COLOR.test(text)) return text;
  if (/^#[0-9a-f]{3}$/.test(text)) return `#${[...text.slice(1)].map(char => char + char).join("")}`;
  return NAMED_COLORS[text] ?? null;
}

const EMU_PER_UNIT: Readonly<Record<string, number>> = { pt: 12700, px: 9525, in: 914400, cm: 360000, mm: 36000 };

/** The VML style height in points, read as the server does: whole EMUs,
 * then points to three places. */
function vmlStyleHeightPt(style: string | null): number | null {
  const match = /(?:^|;)\s*height\s*:\s*([0-9.]+)\s*(pt|px|in|cm|mm)\b/i.exec(style ?? "");
  if (!match) return null;
  const emu = Math.trunc(Number(match[1]) * EMU_PER_UNIT[match[2].toLowerCase()]);
  return Number.isFinite(emu) && emu > 0 ? Math.round((emu / 12700) * 1000) / 1000 : null;
}

function localName(node: Element): string {
  return node.localName || node.nodeName.replace(/^.*:/, "");
}

/**
 * The paragraph's Word horizontal line when it is the paragraph's only
 * graphic: exactly one `w:pict` in a plain run directly under the paragraph,
 * holding exactly one childless `v:rect` marked `o:hr="t"`.
 *
 * The line must also OPEN the paragraph. Word's line is an inline object in
 * the paragraph's own flow, so the paragraph may hold text beside it (typing
 * at the end of the line appends a run there), but a page states the line as
 * the paragraph's attr, which cannot say it comes after the text — so a line
 * that follows text is left to the source alone. Mirrors the server reader.
 */
export function wordParagraphVmlHorizontalRule(paragraph: Element): WordHorizontalRule | null {
  const picts: Element[] = [];
  for (const node of Array.from(paragraph.getElementsByTagName("*"))) {
    const name = localName(node);
    if (name === "drawing" || name === "object" || name === "txbxContent" || name === "tbl" || name === "AlternateContent") return null;
    if ((name === "t" || name === "delText") && node.textContent && picts.length === 0) return null;
    if (name === "pict") picts.push(node);
  }
  if (picts.length !== 1) return null;
  const pict = picts[0];
  const run = pict.parentElement;
  if (!run || localName(run) !== "r" || run.parentElement !== paragraph) return null;
  const shapes = Array.from(pict.children);
  if (shapes.length !== 1) return null;
  const rect = shapes[0];
  if (rect.namespaceURI !== VML_NS || localName(rect) !== "rect" || rect.children.length > 0) return null;
  const office = (name: string) => (rect.getAttributeNS(VML_OFFICE_NS, name) ?? "").trim().toLowerCase();
  if (office("hr") !== "t" && office("hr") !== "true") return null;
  const heightPt = vmlStyleHeightPt(rect.getAttribute("style"));
  if (heightPt === null) return null;
  const rawPercent = office("hrpct");
  let widthPercent: number;
  if (rawPercent === "" || rawPercent === "0") widthPercent = 100;
  else if (/^[0-9]{1,4}$/.test(rawPercent) && Number(rawPercent) > 0 && Number(rawPercent) <= 1000) widthPercent = Number(rawPercent) / 10;
  else return null;
  const align = office("hralign") || "left";
  if (!ALIGNS.has(align)) return null;
  const shade = office("hrnoshade") !== "t" && office("hrnoshade") !== "true";
  return normalizeWordHorizontalRule({
    version: 1,
    heightPt,
    widthPercent,
    align,
    shade,
    color: shade ? null : vmlColor(rect.getAttribute("fillcolor")),
  });
}

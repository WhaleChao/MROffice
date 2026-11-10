/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A house icon as an SVG that draws by itself.
 *
 * Inline, a house icon takes its colours from houseIcon.module.css through
 * `data-office-icon-role`, a class and a `style` colour; out of the page (a
 * picture inserted on a slide, a file) none of that applies, so the same
 * markup draws nothing — and `style` alone makes it fail the inline-SVG
 * safety check. This states each role's paint as plain presentation
 * attributes, with the values the stylesheet gives in the light theme, and
 * drops what only the page understood.
 */

export interface HouseIconStandalonePalette {
  ink: string;
  accent: string;
  warm: string;
}

/** The stylesheet's light-theme defaults (houseIcon.module.css). */
export const HOUSE_ICON_STANDALONE_PALETTE: HouseIconStandalonePalette = Object.freeze({
  ink: "#1e293b",
  accent: "#176f92",
  warm: "#b96920",
});

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const ROLE_ATTRIBUTE = "data-office-icon-role";
// color-mix percentages of the stylesheet, as opacities.
const SOFT_STROKE_OPACITY = "0.68";
const SOFT_FILL_OPACITY = "0.16";
const PAPER_FILL_OPACITY = "0.07";

function paint(
  role: string,
  palette: HouseIconStandalonePalette,
): Record<string, string> | null {
  switch (role) {
    case "ink":
      return { fill: "none", stroke: palette.ink };
    case "ink-soft":
      return { fill: "none", stroke: palette.ink, "stroke-opacity": SOFT_STROKE_OPACITY };
    case "paper":
      return { fill: palette.ink, "fill-opacity": PAPER_FILL_OPACITY, stroke: palette.ink };
    case "accent":
      return { fill: "none", stroke: palette.accent };
    case "accent-fill":
      return { fill: palette.accent, stroke: palette.accent };
    case "accent-soft-fill":
      return { fill: palette.accent, "fill-opacity": SOFT_FILL_OPACITY, stroke: "none" };
    case "warm":
      return { fill: "none", stroke: palette.warm };
    case "warm-fill":
      return { fill: palette.warm, "fill-opacity": SOFT_FILL_OPACITY, stroke: palette.warm };
    default:
      return null;
  }
}

/**
 * `markup` is the icon's own `<svg>` (its outerHTML). Returns null when it
 * is not one, or when a shape carries a role this module does not know —
 * a shape left unpainted would silently vanish from the picture.
 */
export function houseIconStandaloneSvg(
  markup: string,
  palette: HouseIconStandalonePalette = HOUSE_ICON_STANDALONE_PALETTE,
): string | null {
  if (typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") return null;
  const parsed = new DOMParser().parseFromString(markup, "image/svg+xml");
  const root = parsed.documentElement;
  if (!root || root.localName !== "svg" || parsed.getElementsByTagName("parsererror").length) {
    return null;
  }
  for (const element of [root, ...Array.from(root.getElementsByTagName("*"))]) {
    const role = element.getAttribute(ROLE_ATTRIBUTE);
    if (role !== null) {
      const attributes = paint(role, palette);
      if (!attributes) return null;
      for (const [name, value] of Object.entries(attributes)) element.setAttribute(name, value);
    }
    for (const name of element.getAttributeNames()) {
      if (
        name === "class" ||
        name === "style" ||
        name === "focusable" ||
        name.startsWith("data-") ||
        name.startsWith("aria-")
      ) {
        element.removeAttribute(name);
      }
    }
  }
  if (!root.getAttribute("xmlns")) root.setAttribute("xmlns", SVG_NAMESPACE);
  return new XMLSerializer().serializeToString(root);
}

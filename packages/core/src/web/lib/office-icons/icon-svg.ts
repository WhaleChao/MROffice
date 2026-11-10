/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An Insert > Icons icon as the standalone SVG the window inserts: the
 * lucide glyph at 96 px, stroke 1.6, in one colour, with nothing a page's
 * stylesheet would have painted (PresentationWindow insertPresentationIconGlyph
 * renders the same glyph with React and strips it with
 * houseIconStandaloneSvg). Built from the glyph's own drawing data, so it
 * needs no renderer: the command host prepares the AI's icons with it.
 *
 * It loads every glyph (`./glyphs`), so only the command host and tests
 * import it; a window loads the glyphs when its gallery opens.
 */

import { OFFICE_ICON_GLYPHS } from "./glyphs";
import type { OfficeIconEntry } from "./catalog";

/** The size and stroke the window draws an inserted icon at. */
export const OFFICE_ICON_SIZE = 96;
export const OFFICE_ICON_STROKE_WIDTH = 1.6;

type IconNode = readonly (readonly [string, Readonly<Record<string, string | number>>])[];

const SAFE_TAG = /^(?:path|circle|ellipse|rect|line|polyline|polygon)$/;
const SAFE_ATTRIBUTE = /^(?:d|cx|cy|r|rx|ry|x|y|width|height|x1|y1|x2|y2|points|fill|stroke|stroke-width)$/;
const HEX_COLOR = /^#[0-9a-f]{6}$/i;

const escape = (value: string) =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** The glyph's drawing: what its component hands lucide's Icon. */
function glyphNode(glyph: string): IconNode | null {
  const component = (OFFICE_ICON_GLYPHS as Readonly<Record<string, unknown>>)[glyph] as
    | { render?: (props: object, ref: null) => { props?: { iconNode?: unknown } } }
    | undefined;
  const node = component?.render?.({}, null)?.props?.iconNode;
  return Array.isArray(node) ? (node as IconNode) : null;
}

/** The icon's standalone SVG in `color` (`#rrggbb`), or null. */
export function officeIconSvg(entry: OfficeIconEntry, color: string): string | null {
  if (!HEX_COLOR.test(color)) return null;
  const node = glyphNode(entry.glyph);
  if (!node?.length) return null;
  const children: string[] = [];
  for (const [tag, attributes] of node) {
    if (!SAFE_TAG.test(tag)) return null;
    const written = Object.entries(attributes)
      .filter(([name]) => name !== "key")
      .map(([name, value]) => {
        if (!SAFE_ATTRIBUTE.test(name)) return null;
        return `${name}="${escape(String(value))}"`;
      });
    if (written.some((attribute) => attribute === null)) return null;
    children.push(`<${tag}${written.length ? ` ${written.join(" ")}` : ""}/>`);
  }
  // `color` too: a few glyphs fill a dot with currentColor, which a
  // standalone SVG would otherwise draw black.
  const hex = color.toLowerCase();
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${OFFICE_ICON_SIZE}" height="${OFFICE_ICON_SIZE}" ` +
    `viewBox="0 0 24 24" fill="none" stroke="${hex}" stroke-width="${OFFICE_ICON_STROKE_WIDTH}" ` +
    `stroke-linecap="round" stroke-linejoin="round" color="${hex}">${children.join("")}</svg>`
  );
}

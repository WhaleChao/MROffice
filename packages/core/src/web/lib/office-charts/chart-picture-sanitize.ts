/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import DOMPurify from "dompurify";

import { withOfficeFontSubstitutes } from "../office-fonts/substitutes";

/** The marks a picture keeps for the browser's hit-testing and part editing:
 * ECharts' series / point marks and the chart engine's part mark. Values
 * are numbers or one closed word. */
export const CHART_PICTURE_MARKS = ["ecmeta_series_index", "ecmeta_data_index", "ecmeta_part", "ecmeta_label"] as const;

const MAX_PICTURE_CHARS = 2_000_000;

/**
 * The one reference an `<image>` may keep: an inline PNG / JPEG / GIF / WebP
 * (a chart element's picture fill, read from the package's own chart part).
 * Never SVG (it could carry script or further loads), never a URL.
 */
const INLINE_RASTER = /^data:image\/(?:png|jpeg|gif|webp);base64,[A-Za-z0-9+/]+={0,2}$/;

/**
 * A server-drawn chart picture made safe to insert as inline SVG.
 *
 * The chart engine draws only ECharts-generated markup (its WordArt titles
 * already inlined as nested vector), never imported SVG; this is the second
 * fence. Inline SVG inherits the page's installed and web fonts (an <img>
 * would lose the document's @font-face rules); labels are escaped by ECharts.
 * The picture fills its box (`viewBox` = the size it was drawn at).
 */
export function sanitizeChartPicture(svg: string, width: number, height: number): string {
  if (svg.length > MAX_PICTURE_CHARS) throw new Error("Chart picture exceeds budget");
  // The same picture is made safe once: a chart's selection, its contextual
  // tab and its style gallery each insert the picture they were handed, and
  
  const key = width + "x" + height + "\u0000" + svg;
  const kept = sanitizedPictures.get(key);
  if (kept !== undefined) {
    sanitizedPictures.delete(key);
    sanitizedPictures.set(key, kept);
    return kept;
  }
  const made = sanitizeChartPictureOnce(svg, width, height);
  sanitizedPictures.set(key, made);
  sanitizedChars += key.length + made.length;
  for (const [oldest, value] of sanitizedPictures) {
    if (sanitizedPictures.size <= SANITIZED_PICTURES_MAX && sanitizedChars <= SANITIZED_PICTURES_MAX_CHARS) break;
    sanitizedPictures.delete(oldest);
    sanitizedChars -= oldest.length + value.length;
  }
  return made;
}

/** The pictures made safe lately, least recently used first (a bounded
 * memo: the result is a function of the picture and its size alone). */
const sanitizedPictures = new Map<string, string>();
let sanitizedChars = 0;
const SANITIZED_PICTURES_MAX = 96;
const SANITIZED_PICTURES_MAX_CHARS = 12_000_000;

function sanitizeChartPictureOnce(svg: string, width: number, height: number): string {
  if (!/^\s*<svg[\s>]/.test(svg)) throw new Error("No chart picture");
  const sanitized = DOMPurify.sanitize(svg, {
    USE_PROFILES: { svg: true, svgFilters: true },
    ADD_ATTR: [...CHART_PICTURE_MARKS],
    // A pattern fill drawn as its own marks tiles them with `<use>` of a
    // `<symbol>` (`svgPatternFillsAsShapes`); every `href` but one inside the
    // picture is removed below, so a `<use>` reaches nothing outside it.
    ADD_TAGS: ["use"],
    FORBID_TAGS: [
      "style",
      "script",
      "foreignObject",
      "animate",
      "animateMotion",
      "animateTransform",
      "set",
    ],
  });
  // The sanitizer serializes as HTML: a no-break space comes back as `&nbsp;`,
  // which XML does not know, and the whole chart was refused and drew nothing
  
  const doc = new DOMParser().parseFromString(sanitized.replace(/&nbsp;/g, "&#160;"), "image/svg+xml");
  if (doc.documentElement.localName !== "svg" || doc.querySelector("parsererror"))
    throw new Error("Invalid chart picture");
  for (const node of Array.from(doc.querySelectorAll("*"))) {
    // A picture fill's tile: an image whose only reference is an inline
    // raster picture, and nothing else of an image survives.
    let inlinePicture: Attr | null = null;
    if (node.localName === "image") {
      const references = Array.from(node.attributes).filter((attribute) => attribute.localName === "href");
      if (references.length !== 1 || !INLINE_RASTER.test(references[0].value)) {
        node.remove();
        continue;
      }
      inlinePicture = references[0];
    }
    servedFaces(node);
    for (const attribute of Array.from(node.attributes)) {
      if (attribute === inlinePicture) continue;
      if (
        (attribute.localName === "href" && !attribute.value.startsWith("#")) ||
        /(?:https?|data|javascript):|@import|\\/i.test(attribute.value) ||
        // Every part a click selects (chart-hit's clickable parts): the axis
        // labels and data labels too, or a click on them selects nothing.
        (attribute.localName === "ecmeta_part" &&
          !/^(?:title|axisX|axisY|legend|axisLabelX|axisLabelY|dataLabels)$/.test(attribute.value)) ||
        (attribute.localName === "ecmeta_label" && !/^\d{1,4}:\d{1,6}$/.test(attribute.value))
      )
        node.removeAttributeNode(attribute);
    }
  }
  doc.documentElement.setAttribute("viewBox", "0 0 " + width + " " + height);
  doc.documentElement.setAttribute("width", "100%");
  doc.documentElement.setAttribute("height", "100%");
  doc.documentElement.setAttribute("preserveAspectRatio", "none");
  return new XMLSerializer().serializeToString(doc.documentElement);
}









function servedFaces(node: Element): void {
  const own = node.getAttribute("font-family");
  if (own) node.setAttribute("font-family", servedStack(own));
  const style = node.getAttribute("style");
  if (style && style.includes("font-family"))
    node.setAttribute("style", style.replace(/(font-family\s*:\s*)([^;]+)/g, (_, name: string, stack: string) => name + servedStack(stack)));
}

const servedStacks = new Map<string, string>();

function servedStack(stack: string): string {
  let served = servedStacks.get(stack);
  if (served === undefined) {
    served = withOfficeFontSubstitutes(stack).replace(/"([^"]*)"/g, "'$1'");
    if (servedStacks.size > 256) servedStacks.clear();
    servedStacks.set(stack, served);
  }
  return served;
}

/** Each DOM copy owns its clip / gradient ids while sharing one drawing.
 * Rewrites are limited to tags, never the text between them. */
export function namespaceChartPictureIds(markup: string, prefix: string): string {
  return markup.replace(/<[^>]+>/g, (tag) =>
    tag
      .replace(/\bid="([^"]+)"/g, (_, id) => 'id="' + prefix + "-" + id + '"')
      .replace(/url\(#([^)]+)\)/g, (_, id) => "url(#" + prefix + "-" + id + ")")
      .replace(/((?:xlink:)?href=")#([^"]+)"/g, (_, attribute, id) => attribute + "#" + prefix + "-" + id + '"'),
  );
}

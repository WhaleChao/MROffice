/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import DOMPurify from "dompurify";
import {
  projectDocumentFontSlotsToDom,
  createEffectiveKerningThresholdResolver,
} from "@mrnote/core/web/components/console/editor/extensions/TextStyling.ts";
import {
  parseOfficeFontFace,
  officeFontFaceRule,
} from "@mrnote/core/web/components/office-ribbon/officeFontFace.ts";
import {
  DOMParser as ProseMirrorDOMParser,
  DOMSerializer,
  type Schema,
} from "@tiptap/pm/model";
import { EditorState } from "@tiptap/pm/state";

import { mergePagedRuntimeContinuationsJson } from "@mrnote/core/web/components/console/editor/native-document/paged/serialization.ts";
import { flattenPageBlocks } from "@mrnote/core/web/components/console/editor/native-document/paged/pageBlockExtension.ts";
import type { TipTapNode } from "@/lib/document-export/tiptap-json";
import {
  MRNOTE_HTML_DOCUMENT_SCRIPT_ID,
  MRNOTE_HTML_DOCUMENT_VERSION,
  serializeMrnoteHtmlDocumentEnvelope,
} from "@mrnote/core/web/lib/document-html-envelope.ts";

import {
  applyRevisionPolicyToDom,
  projectTrackedChanges,
  type RevisionExportPolicy,
} from "@mrnote/core/web/lib/document-export/revision-policy.ts";
import { OfflineHtmlExportError } from "./html-export-error";
import {
  STATIC_SEMANTIC_LIST_CLASS,
  projectStaticNativeNodes,
  isProjectedChartImage,
} from "./static-node-projection";

export { OfflineHtmlExportError } from "./html-export-error";
export { prefetchDocumentExportCharts } from "./html-export-charts";
import type { DocumentExportChartPictures } from "./html-export-charts";

const NETWORKED_OR_EXECUTABLE_TAGS = [
  "applet",
  "audio",
  "base",
  "bgsound",
  "embed",
  "fencedframe",
  "form",
  "frame",
  "iframe",
  "input",
  "link",
  "meta",
  "object",
  "script",
  "source",
  "style",
  "template",
  "textarea",
  "track",
  "video",
] as const;

const MAX_INLINE_STYLE_CHARS = 4_096;
const MAX_INLINE_STYLE_DECLARATIONS = 64;
const MAX_INLINE_STYLE_VALUE_CHARS = 512;
const UNSAFE_INLINE_STYLE_SYNTAX_RE =
  /[\\\u0000-\u001f\u007f{}@]|\/\*|\*\/|!\s*important/i;
const CSS_NUMBER_SOURCE = String.raw`-?(?:\d+(?:\.\d+)?|\.\d+)`;
const CSS_LENGTH_RE = new RegExp(
  `^${CSS_NUMBER_SOURCE}(?:px|pt|em|rem|%)?$`,
  "i",
);
const CSS_POLYGON_POINT_SOURCE = `${CSS_NUMBER_SOURCE}%\\s+${CSS_NUMBER_SOURCE}%`;
const CSS_POLYGON_RE = new RegExp(
  `^polygon\\(\\s*${CSS_POLYGON_POINT_SOURCE}(?:\\s*,\\s*${CSS_POLYGON_POINT_SOURCE}){2,}\\s*\\)$`,
  "i",
);
const SAFE_COLOR_RE = /^(?:#[0-9a-f]{3,8}|currentcolor|transparent|[a-z]+)$/i;
const SAFE_FONT_FAMILY_RE = /^[\p{L}\p{N}\s"',._-]+$/u;
const SAFE_BORDER_RE = new RegExp(
  `^(?:none|${CSS_NUMBER_SOURCE}(?:px|pt|em|rem)\\s+(?:solid|double|dotted|dashed)\\s+(?:#[0-9a-f]{3,8}|currentcolor|transparent|[a-z]+))$`,
  "i",
);
const SAFE_DATA_IMAGE_RE =
  /^data:image\/(?:png|jpe?g|gif|webp|bmp);base64,[a-z0-9+/]+={0,2}$/i;
const SAFE_LINK_SCHEMES = new Set([
  "ftp:",
  "ftps:",
  "http:",
  "https:",
  "mailto:",
  "sms:",
  "tel:",
]);

const OFFLINE_HTML_CSP = [
  "default-src 'none'",
  "base-uri 'none'",
  "connect-src 'none'",
  "font-src data:",
  "form-action 'none'",
  "frame-src 'none'",
  "img-src data:",
  "media-src 'none'",
  "object-src 'none'",
  "script-src 'none'",
  "style-src 'unsafe-inline'",
].join("; ");

const OFFLINE_HTML_STYLE =
  "body{font-family:system-ui,'Segoe UI',sans-serif;max-width:52rem;" +
  "margin:2rem auto;padding:0 1rem;line-height:1.6}" +
  "table{border-collapse:collapse;max-width:100%}" +
  "td,th{border:1px solid #ccc;padding:4px 8px}" +
  "img{max-width:100%}" +
  ".external-image-omitted{display:block;border:1px dashed #999;" +
  "padding:.5rem;color:#555}" +
  `.${STATIC_SEMANTIC_LIST_CLASS}{list-style:none;padding-left:0}`;

function escapedText(value: string): string {
  const node = document.createElement("span");
  node.textContent = value;
  return node.innerHTML;
}

function safeLinkHref(rawHref: string): boolean {
  if (rawHref.startsWith("#")) return true;
  try {
    const parsed = new URL(rawHref);
    return SAFE_LINK_SCHEMES.has(parsed.protocol);
  } catch {
    return false;
  }
}

function oneOf(value: string, allowed: readonly string[]): boolean {
  return allowed.includes(value.toLowerCase());
}

function safeLengthList(
  value: string,
  options: { allowAuto?: boolean; allowFitContent?: boolean } = {},
): boolean {
  const tokens = value.trim().split(/\s+/);
  if (tokens.length === 0 || tokens.length > 4) return false;
  return tokens.every(
    (token) =>
      CSS_LENGTH_RE.test(token) ||
      (options.allowAuto === true && token.toLowerCase() === "auto") ||
      (options.allowFitContent === true &&
        token.toLowerCase() === "fit-content"),
  );
}

function safeScaleX(value: string): boolean {
  const match = new RegExp(`^scaleX\\((${CSS_NUMBER_SOURCE})\\)$`, "i").exec(
    value,
  );
  if (!match) return false;
  const scale = Number(match[1]);
  return Number.isFinite(scale) && scale >= 0.01 && scale <= 6;
}

function safeAspectRatio(value: string): boolean {
  const match = /^(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)$/.exec(value);
  return Boolean(match && Number(match[1]) > 0 && Number(match[2]) > 0);
}

function safeInlineStyleValue(property: string, value: string): boolean {
  if (
    value.length === 0 ||
    value.length > MAX_INLINE_STYLE_VALUE_CHARS ||
    value.includes(":") ||
    UNSAFE_INLINE_STYLE_SYNTAX_RE.test(value)
  ) {
    return false;
  }

  switch (property) {
    case "font-variation-settings":
      return (
        value === "normal" ||
        /^(?:"[A-Za-z][A-Za-z0-9 ]{3}" -?\d+(?:\.\d+)?)(?:,\s*"[A-Za-z][A-Za-z0-9 ]{3}" -?\d+(?:\.\d+)?){0,15}$/u.test(
          value,
        )
      );
    case "font-kerning":
      return value === "normal" || value === "none" || value === "auto";
    case "font-synthesis":
      return value === "none";
    case "color":
    case "background-color":
    case "text-decoration-color":
      return SAFE_COLOR_RE.test(value);
    case "font-family":
      return SAFE_FONT_FAMILY_RE.test(value);
    case "font-size":
    case "line-height":
    case "letter-spacing":
    case "text-indent":
    case "tab-size":
      return CSS_LENGTH_RE.test(value);
    case "width":
    case "min-width":
    case "max-width":
    case "height":
    case "min-height":
    case "max-height":
      return safeLengthList(value, { allowFitContent: true });
    case "margin":
    case "margin-top":
    case "margin-right":
    case "margin-bottom":
    case "margin-left":
      return safeLengthList(value, { allowAuto: true });
    case "padding":
    case "padding-top":
    case "padding-right":
    case "padding-bottom":
    case "padding-left":
    case "border-radius":
      return safeLengthList(value);
    case "border":
    case "border-top":
    case "border-right":
    case "border-bottom":
    case "border-left":
      return SAFE_BORDER_RE.test(value);
    case "text-align":
      return oneOf(value, [
        "left",
        "right",
        "center",
        "justify",
        "start",
        "end",
      ]);
    case "vertical-align":
      return (
        CSS_LENGTH_RE.test(value) ||
        oneOf(value, [
          "baseline",
          "middle",
          "top",
          "bottom",
          "text-top",
          "text-bottom",
          "super",
          "sub",
        ])
      );
    case "font-weight":
      return /^(?:normal|bold|[1-9]00)$/i.test(value);
    case "font-style":
      return oneOf(value, ["normal", "italic"]);
    case "font-variant":
    case "font-variant-caps":
      return oneOf(value, ["normal", "small-caps"]);
    case "text-transform":
      return oneOf(value, ["none", "uppercase", "lowercase", "capitalize"]);
    case "display":
      return oneOf(value, ["inline", "inline-block", "inline-flex", "block"]);
    case "transform":
      return safeScaleX(value);
    case "transform-origin":
      return oneOf(value, ["left center"]);
    case "white-space":
      return oneOf(value, ["nowrap"]);
    case "text-combine-upright":
      return oneOf(value, ["all"]);
    case "-webkit-text-combine":
      return oneOf(value, ["horizontal"]);
    case "flex-direction":
      return oneOf(value, ["column"]);
    case "text-decoration-line":
      return oneOf(value, ["underline"]);
    case "text-decoration-style":
      return oneOf(value, ["solid", "double", "dotted", "dashed", "wavy"]);
    case "text-emphasis":
    case "-webkit-text-emphasis":
      return oneOf(value, ["filled sesame", "open circle", "filled dot"]);
    case "text-emphasis-position":
    case "-webkit-text-emphasis-position":
      return oneOf(value, ["under right"]);
    case "align-items":
    case "justify-content":
      return oneOf(value, ["center"]);
    case "box-sizing":
      return oneOf(value, ["content-box", "border-box"]);
    case "aspect-ratio":
      return safeAspectRatio(value);
    case "clip-path":
    case "-webkit-clip-path":
      return CSS_POLYGON_RE.test(value);
    case "page-break-before":
      return oneOf(value, ["auto", "always"]);
    case "break-before":
      return oneOf(value, ["auto", "page", "column"]);
    case "-webkit-column-break-before":
      return oneOf(value, ["auto", "always"]);
    case "float":
      return oneOf(value, ["left", "right"]);
    case "clear":
      return oneOf(value, ["both"]);
    case "position":
      return oneOf(value, ["relative"]);
    case "overflow":
      return oneOf(value, ["hidden"]);
    case "pointer-events":
      return oneOf(value, ["none"]);
    case "z-index":
      return /^-?\d{1,4}$/.test(value);
    case "counter-increment":
      return /^mrn-page [1-9]\d{0,3}$/.test(value);
    default:
      return false;
  }
}

function assertSafeInlineStyle(style: string): void {
  if (
    style.length > MAX_INLINE_STYLE_CHARS ||
    UNSAFE_INLINE_STYLE_SYNTAX_RE.test(style)
  ) {
    throw new OfflineHtmlExportError(
      "Document contains an unsafe inline style",
    );
  }
  const declarations = style
    .split(";")
    .map((declaration) => declaration.trim())
    .filter(Boolean);
  if (declarations.length > MAX_INLINE_STYLE_DECLARATIONS) {
    throw new OfflineHtmlExportError(
      "Document contains too many inline style declarations",
    );
  }
  const seenProperties = new Set<string>();
  for (const declaration of declarations) {
    const separator = declaration.indexOf(":");
    if (separator <= 0 || declaration.indexOf(":", separator + 1) !== -1) {
      throw new OfflineHtmlExportError(
        "Document contains a malformed inline style",
      );
    }
    const property = declaration.slice(0, separator).trim().toLowerCase();
    const value = declaration.slice(separator + 1).trim();
    if (
      !/^-?[a-z]+(?:-[a-z]+)*$/.test(property) ||
      seenProperties.has(property) ||
      !safeInlineStyleValue(property, value)
    ) {
      throw new OfflineHtmlExportError(
        "Document contains an unsupported or unsafe inline style",
      );
    }
    seenProperties.add(property);
  }
}

function replaceExternalImage(image: HTMLImageElement): void {
  const replacement = document.createElement("span");
  replacement.className = "external-image-omitted";
  replacement.textContent =
    image.getAttribute("alt") || "External image omitted";
  image.replaceWith(replacement);
}

/**
 * Build a standalone, no-network HTML export from schema-rendered editor HTML.
 *
 * The editor schema is the first validation boundary. This function is the
 * independent serialization boundary: DOMPurify removes active markup, all
 * automatic-fetch elements are denied, remote/blob images become an explicit
 * placeholder, active CSS is rejected, and a meta CSP keeps the downloaded
 * file inert even if a future editor extension accidentally emits a new sink.
 */
export function buildOfflineDocumentHtml(
  editorHtml: string,
  documentTitle: string,
  revisionPolicy: RevisionExportPolicy = "accept",
  editorJson?: TipTapNode,
  /** The document's chart pictures (`prefetchDocumentExportCharts`). */
  chartPictures?: DocumentExportChartPictures,
): string {
  if (typeof document === "undefined") {
    throw new OfflineHtmlExportError(
      "Offline HTML export requires a browser DOM",
    );
  }

  // Parsing untrusted markup into an ordinary detached element is not inert:
  // browsers may start fetching subresources before sanitization. Keep the
  // complete projection and sanitization pipeline inside template contents,
  // which are inert even if a future editor node emits a networked element.
  const outerTemplate = document.createElement("template");
  outerTemplate.innerHTML = "<div></div>";
  const root = outerTemplate.content.firstElementChild;
  if (!(root instanceof HTMLDivElement)) {
    throw new OfflineHtmlExportError("Unable to create an inert export root");
  }
  const sourceTemplate = document.createElement("template");
  sourceTemplate.innerHTML = editorHtml;
  root.append(sourceTemplate.content);

  applyRevisionPolicyToDom(root, revisionPolicy);
  const semanticSnapshot = editorJson
    ? projectTrackedChanges(editorJson, revisionPolicy)
    : undefined;
  projectStaticNativeNodes(root, semanticSnapshot, chartPictures);
  projectDocumentFontSlotsToDom(root);
  DOMPurify.sanitize(root, {
    IN_PLACE: true,
    FORBID_TAGS: [...NETWORKED_OR_EXECUTABLE_TAGS],
    FORBID_ATTR: [
      "action",
      "background",
      "formaction",
      "ping",
      "poster",
      "srcdoc",
      "srcset",
    ],
    USE_PROFILES: { html: true },
  });

  for (const element of root.querySelectorAll<HTMLElement>("[style]")) {
    const style = element.getAttribute("style") ?? "";
    assertSafeInlineStyle(style);
  }

  for (const image of root.querySelectorAll<HTMLImageElement>("img")) {
    const src = image.getAttribute("src") ?? "";
    if (!SAFE_DATA_IMAGE_RE.test(src) && !isProjectedChartImage(image)) replaceExternalImage(image);
  }

  for (const anchor of root.querySelectorAll<HTMLAnchorElement>("a")) {
    const href = anchor.getAttribute("href") ?? "";
    if (!safeLinkHref(href)) {
      anchor.removeAttribute("href");
      anchor.removeAttribute("target");
      anchor.removeAttribute("rel");
      continue;
    }
    if (!href.startsWith("#")) {
      anchor.rel = "nofollow noopener noreferrer";
      if (anchor.target && anchor.target !== "_blank") {
        anchor.removeAttribute("target");
      }
    }
  }

  const title = escapedText(documentTitle || "document");
  const fontRules = new Map<string, string>();
  for (const element of root.querySelectorAll(
    "[data-office-font-face],[data-office-cjk-font-face]",
  )) {
    for (const name of ["data-office-font-face", "data-office-cjk-font-face"]) {
      const raw = element.getAttribute(name);
      if (raw === null) continue;
      const face = parseOfficeFontFace(raw);
      if (!face) throw new OfflineHtmlExportError("Invalid font face identity");
      fontRules.set(face.postscriptName, officeFontFaceRule(face));
    }
  }
  const embeddedDocument = semanticSnapshot
    ? `<script id="${MRNOTE_HTML_DOCUMENT_SCRIPT_ID}" type="application/json">${serializeMrnoteHtmlDocumentEnvelope(
        documentTitle || "document",
        semanticSnapshot,
      )}</script>`
    : "";
  return (
    '<!doctype html><html><head><meta charset="utf-8">' +
    `<meta http-equiv="Content-Security-Policy" content="${OFFLINE_HTML_CSP}">` +
    `<meta name="mrnote-document-format" content="${MRNOTE_HTML_DOCUMENT_VERSION}">` +
    `<title>${title}</title>` +
    `<style>${OFFLINE_HTML_STYLE}${[...fontRules.values()].join("")}</style>` +
    `</head><body>${root.innerHTML}${embeddedDocument}</body></html>`
  );
}

/**
 * Build standalone HTML from one paged-editor JSON snapshot.
 *
 * Structural sheets and split code/list/note/table nodes are a live-editor
 * projection, never document semantics.  Deriving markup from `editor.getHTML()`
 * after separately merging `editor.getJSON()` mixes those two models: the DOM
 * still contains continuation tails while atom projection is bound to the
 * merged JSON.  Besides duplicate ordinary content, that can either leak
 * pagination attrs or fail the closed atom count check.
 *
 * This boundary owns the complete conversion instead:
 *   physical JSON -> merged semantic JSON -> flat flow -> revision projection
 *   -> schema HTML -> inert/sanitized offline document.
 *
 * Callers therefore cannot accidentally pair semantic JSON with physical DOM
 * HTML. The function is intentionally synchronous so the editor state cannot
 * change between any of these steps.
 */
export function buildOfflineDocumentHtmlFromPagedJson(
  editorJson: TipTapNode,
  documentTitle: string,
  schema: Schema,
  revisionPolicy: RevisionExportPolicy = "accept",
  /** The document's chart pictures, drawn by the server beforehand
   * (`prefetchDocumentExportCharts`); a chart without one refuses the export. */
  chartPictures?: DocumentExportChartPictures,
): string {
  const semanticJson = flattenPageBlocks(
    mergePagedRuntimeContinuationsJson(editorJson),
  ) as TipTapNode;
  const projectedJson = projectTrackedChanges(semanticJson, revisionPolicy);

  let editorHtml: string;
  try {
    if (projectedJson.type !== schema.topNodeType.name) {
      throw new Error("document root does not match the editor schema");
    }
    const projectedDocument = schema.nodeFromJSON(projectedJson);
    const serializer = DOMSerializer.fromSchema(schema);
    // The live font layout includes an effective-kerning decoration. Project
    // it from the SAME semantic state; don't mix live DOM with flat JSON or
    // silently let HTML's default "auto" change the font's measured advances.
    const template = document.createElement("template");
    template.content.append(
      serializer.serializeFragment(projectedDocument.content, {
        document: template.content.ownerDocument,
      }),
    );
    const positions: Array<{ node: Node; offset: number; pos?: number }> = [];
    const walker = template.content.ownerDocument.createTreeWalker(
      template.content,
      NodeFilter.SHOW_TEXT,
    );
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (
        node.parentElement?.closest(
          '[data-latex],.katex,[contenteditable="false"]',
        )
      )
        continue;
      positions.push({ node, offset: 0 });
    }
    const parsed = ProseMirrorDOMParser.fromSchema(schema).parse(
      template.content,
      {
        topNode: schema.topNodeType.create(projectedDocument.attrs),
        findPositions: positions,
      },
    );
    const state = EditorState.create({ schema, doc: parsed });
    const kerning = createEffectiveKerningThresholdResolver(state);
    for (const position of positions) {
      if (position.pos === undefined) continue;
      const node = parsed.nodeAt(position.pos);
      if (!node?.isText) continue;
      const value = kerning(position.pos, node);
      const enabled =
        value.valid &&
        value.halfPoints !== null &&
        value.fontSizePx !== null &&
        value.fontSizePx * 1.5 >= value.halfPoints;
      const span = template.content.ownerDocument.createElement("span");
      span.dataset.docsEffectiveKerning = enabled ? "on" : "off";
      span.style.fontKerning = enabled ? "normal" : "none";
      position.node.parentNode?.replaceChild(span, position.node);
      span.append(position.node);
    }
    editorHtml = template.innerHTML;
  } catch (error) {
    const detail = error instanceof Error ? error.message : "unknown error";
    throw new OfflineHtmlExportError(
      `Unable to serialize the semantic document as HTML: ${detail}`,
    );
  }

  // The revision policy has already been applied to projectedJson before DOM
  // serialization. "preserve" here prevents a second, DOM-based projection;
  // atom binding receives the exact same projected snapshot that made the DOM.
  return buildOfflineDocumentHtml(
    editorHtml,
    documentTitle,
    "preserve",
    projectedJson,
    chartPictures,
  );
}

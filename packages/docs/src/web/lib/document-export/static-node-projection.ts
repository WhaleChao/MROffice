/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TipTapNode } from "@/lib/document-export/tiptap-json";
import {
  captionDisplayLabel,
  captionListBuiltinKind,
  captionListLabelText,
  captionsForListLabel,
  collectCaptions,
  generatedCaptionListEntries,
  normalizeCaptionListLabel,
  normalizeGeneratedCaptionListEntries,
} from "@mrnote/core/web/components/console/editor/native-document/captions.ts";
import {
  buildBibliography,
  findSource,
  formatInText,
  MAX_SOURCES,
  normalizeSources,
  resolveCitationStyle,
  sanitizeCslField,
  sanitizeLocator,
  sanitizeSourceId,
} from "@mrnote/core/web/components/console/editor/native-document/citations.ts";
import {
  findFootnote,
  MAX_FOOTNOTES,
  normalizeFootnotes,
  noteStreamOf,
  numberNoteStreams,
  resolveFootnoteMode,
  sanitizeCustomMark,
  sanitizeFootnoteId,
  type FootnoteRichBody,
  type OrderedNoteRef,
} from "@mrnote/core/web/components/console/editor/native-document/footnotes.ts";
import { documentNoteSettings } from "@mrnote/core/web/components/console/editor/native-document/noteSectionSettings.ts";
import {
  collectHeadings,
  collectTocEntries,
  normalizeTocEntrySnapshot,
  sanitizeAnchorId,
} from "@mrnote/core/web/components/console/editor/native-document/tableOfContents.ts";
import {
  normalizeGeneratedListSettings,
  normalizeTocSettings,
  tocEntryMaximumLevel,
  tocSourceLevelResolver,
  type TocTabLeader,
} from "@mrnote/core/web/components/console/editor/native-document/referenceSettings.ts";
import {
  collectBookmarks,
  sanitizeBookmarkName,
} from "@mrnote/core/web/components/console/editor/native-document/bookmarks.ts";
import { OfflineHtmlExportError } from "./html-export-error";
import {
  documentCoverPageContext,
  documentCoverPageScene,
} from "@mrnote/core/web/lib/document-cover-pages/coverPageContext.ts";
import { coverSceneSvgMarkup } from "@mrnote/core/web/lib/document-cover-pages/coverSceneSvg.ts";
import { normalizeChartBlockAttrs } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { documentExportChartKey, documentExportChartRequest, type DocumentExportChartPictures } from "./html-export-charts";

export type HtmlExportAtomDisposition = "passthrough" | "project" | "reject";

/**
 * Every projected list carries its own visible semantic label (for example
 * `Figure 1`, `[1]`, or a Word footnote marker). Browsers must therefore not
 * add a second marker from the backing `<ol>` element.
 */
export const STATIC_SEMANTIC_LIST_CLASS = "mrnote-static-semantic-list";

/**
 * Closed policy for every extension `NodeType.isAtom` in the native editor
 * schema. ProseMirror's intrinsic `text` scalar also reports `isAtom` because
 * it is a leaf, but it is serialized by the ordinary text-node path and is
 * intentionally excluded from this NodeView policy.
 *
 * Keeping this declaration independent from the schema is deliberate: the
 * contract test instantiates the real registry and requires exact equality, so
 * adding an atom cannot silently inherit a lossy default.
 */
export const HTML_EXPORT_ATOM_POLICIES = {
  ai_output: "project",
  bibliography: "project",
  // Word's point bookmark is an empty marker span; it passes through and
  // takes its bookmark's anchor id (bindBookmarkIds) so links still land.
  bookmarkPoint: "passthrough",
  // An editing exception's invisible marker (w:permStart / w:permEnd).
  rangePermission: "passthrough",
  caption: "project",
  chart: "project",
  citation: "project",
  // Reproducing the code-chart faithfully offline would require re-running
  // the sandboxed Python runtime; offline HTML export fails closed exactly
  // like chart/smartArt/whiteboard (the server HTML export embeds the PNG).
  codeChart: "reject",
  columnBreak: "passthrough",
  crossRef: "project",
  // This text belongs to a source-package altChunk. Static HTML cannot carry
  // its custody contract, so it must never flatten the carrier into prose.
  documentAltChunkReadOnly: "reject",
  // A cover page is drawn as one static picture of its page (shapes and the
  // property values it prints).
  documentCoverPage: "project",
  documentField: "passthrough",
  // The native renderer already emits a bounded, inert link fallback rather
  // than an iframe/player. Preserve that deterministic fallback offline.
  documentMedia: "passthrough",
  // These NodeViews do not yet have deterministic static renderers. Their
  // renderHTML fallbacks omit authored geometry, signer details, or table
  // cells, so exporting them would silently discard visible document data.
  documentShape: "reject",
  documentSignatureLine: "reject",
  documentVirtualTable: "reject",
  file: "project",
  flashcard: "project",
  footnoteList: "project",
  footnoteRef: "project",
  frontmatter: "project",
  hardBreak: "passthrough",
  horizontalRule: "passthrough",
  image: "passthrough",
  // The In Line with Text spellings of a chart, picture and shape follow their
  // block nodes: the chart is drawn as the same static picture on its line,
  // the picture passes through, and the shape fails closed like the block.
  inlineChart: "project",
  inlineDocumentShape: "reject",
  inlineImage: "passthrough",
  inlineMath: "project",
  listOfFigures: "project",
  mathBlock: "project",
  mergeField: "passthrough",
  // Rendering mermaid deterministically requires the browser engine; offline
  // HTML export fails closed exactly like chart/smartArt/whiteboard.
  mermaid: "reject",
  pageBreak: "passthrough",
  reference: "project",
  ruby: "passthrough",
  sectionBreak: "passthrough",
  smartArt: "reject",
  sourcePageBreak: "passthrough",
  tableOfContents: "project",
  task: "project",
  whiteboard: "reject",
  // A whiteboard's scene objects exist only inside a whiteboard, which is
  // refused as a whole above.
  whiteboardObject: "passthrough",
} as const satisfies Record<string, HtmlExportAtomDisposition>;

export type HtmlExportAtomName = keyof typeof HTML_EXPORT_ATOM_POLICIES;
type ProjectedAtomName = {
  [Name in HtmlExportAtomName]: (typeof HTML_EXPORT_ATOM_POLICIES)[Name] extends
    | "project"
    | "reject"
    ? Name
    : never;
}[HtmlExportAtomName];

const PROJECTED_ATOM_SELECTORS = {
  ai_output: '[data-type="ai_output"]',
  bibliography: '[data-type="bibliography"]',
  caption: 'figcaption[data-type="caption"]',
  chart: '[data-type="chart"]',
  citation: 'span[data-type="citation"]',
  codeChart: '[data-type="codeChart"]',
  crossRef: '[data-type="cross-ref"]',
  documentAltChunkReadOnly: "[data-document-altchunk-read-only]",
  documentCoverPage: 'div[data-type="document-cover-page"]',
  documentShape: "figure[data-document-shape]",
  documentSignatureLine: 'div[data-type="document-signature-line"]',
  documentVirtualTable: '[data-type="documentVirtualTable"]',
  file: '[data-type="file"]',
  flashcard: '[data-type="flashcard"]',
  footnoteList: '[data-type="footnote-list"]',
  footnoteRef: 'sup[data-type="footnote-ref"]',
  frontmatter: '[data-type="frontmatter"]',
  inlineChart: 'span[data-type="inline-chart"]',
  inlineDocumentShape: "span[data-inline-document-shape]",
  inlineMath: 'span[data-type="inline-math"]',
  listOfFigures: '[data-type="list-of-figures"]',
  mathBlock: '[data-type="math-block"]',
  mermaid: '[data-type="mermaid"]',
  reference: '[data-type="reference"]',
  smartArt: '[data-type="smartArt"]',
  tableOfContents: '[data-type="table-of-contents"]',
  task: '[data-type="task"]',
  whiteboard: '[data-type="whiteboard"]',
} as const satisfies Record<ProjectedAtomName, string>;

const PROJECTED_ATOM_NAMES = Object.keys(
  PROJECTED_ATOM_SELECTORS,
) as ProjectedAtomName[];

const REJECTED_ATOM_NAMES = Object.entries(HTML_EXPORT_ATOM_POLICIES)
  .filter(([, disposition]) => disposition === "reject")
  .map(([name]) => name) as ProjectedAtomName[];

interface BoundNode {
  node: TipTapNode;
  element: HTMLElement;
}

interface ProjectionContext {
  doc: TipTapNode;
  attrs: Record<string, unknown>;
  bindings: Map<ProjectedAtomName, BoundNode[]>;
  headings: ReturnType<typeof collectHeadings>;
  captions: ReturnType<typeof collectCaptions>;
  bookmarks: ReturnType<typeof collectBookmarks>;
  footnotes?: ReturnType<typeof footnoteProjection>;
  chartPictures?: DocumentExportChartPictures;
}

function sortedUnique(values: readonly string[]): string[] {
  return Array.from(new Set(values)).sort();
}

/** Assert exact atom coverage; used by the real-schema CI contract. */
export function assertHtmlExportAtomCoverage(
  atomNames: readonly string[],
): void {
  const actual = sortedUnique(atomNames);
  const declared = Object.keys(HTML_EXPORT_ATOM_POLICIES).sort();
  const undeclared = actual.filter((name) => !declared.includes(name));
  const stale = declared.filter((name) => !actual.includes(name));
  if (undeclared.length === 0 && stale.length === 0) return;
  throw new OfflineHtmlExportError(
    [
      undeclared.length
        ? `undeclared native atoms: ${undeclared.join(", ")}`
        : "",
      stale.length ? `stale atom policies: ${stale.join(", ")}` : "",
    ]
      .filter(Boolean)
      .join("; "),
  );
}

function attrsOf(node: TipTapNode): Record<string, unknown> {
  return node.attrs && typeof node.attrs === "object"
    ? (node.attrs as Record<string, unknown>)
    : {};
}

function boundedText(value: unknown, maxChars = 8_000): string {
  if (value === undefined || value === null) return "";
  if (typeof value !== "string") {
    throw new OfflineHtmlExportError(
      "Offline HTML export found a non-text value in a text field",
    );
  }
  const cleaned = value.replace(
    /[\u0000-\u0008\u000b-\u001f\u007f-\u009f]+/g,
    " ",
  );
  if (cleaned.length > maxChars) {
    throw new OfflineHtmlExportError(
      `Offline HTML export text exceeds the ${maxChars}-character fidelity limit`,
    );
  }
  return cleaned;
}

function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const result = document.createElement(tag);
  if (className) result.className = className;
  return result;
}

function wrapped<K extends keyof HTMLElementTagNameMap>(tag: K, child: Node): HTMLElementTagNameMap[K] {
  const result = document.createElement(tag);
  result.append(child);
  return result;
}

function textElement<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  text: string,
  className?: string,
): HTMLElementTagNameMap[K] {
  const result = element(tag, className);
  result.textContent = text;
  return result;
}

/** Preserve ordered-list semantics while delegating visible markers to data. */
function semanticList(): HTMLOListElement {
  return element("ol", STATIC_SEMANTIC_LIST_CLASS);
}

function addLabelledText(
  parent: HTMLElement,
  label: string,
  value: unknown,
  maxChars = 8_000,
): void {
  const text = boundedText(value, maxChars);
  if (!text.trim()) return;
  const row = element("p");
  row.append(
    textElement("strong", `${label}: `),
    document.createTextNode(text),
  );
  parent.append(row);
}

function collectNodes(doc: TipTapNode): TipTapNode[] {
  const ordered: TipTapNode[] = [];
  const stack: TipTapNode[] = [doc];
  while (stack.length > 0) {
    const node = stack.pop();
    if (!node) continue;
    ordered.push(node);
    if (!Array.isArray(node.content)) continue;
    for (let index = node.content.length - 1; index >= 0; index -= 1) {
      const child = node.content[index];
      if (child && typeof child === "object") stack.push(child);
    }
  }
  return ordered;
}

function selectorName(element: Element): ProjectedAtomName | null {
  for (const name of PROJECTED_ATOM_NAMES) {
    if (element.matches(PROJECTED_ATOM_SELECTORS[name])) return name;
  }
  return null;
}

function bindProjectedAtoms(
  root: HTMLElement,
  doc: TipTapNode,
): Map<ProjectedAtomName, BoundNode[]> {
  const jsonAtoms = collectNodes(doc).filter(
    (node): node is TipTapNode & { type: ProjectedAtomName } =>
      typeof node.type === "string" &&
      Object.hasOwn(PROJECTED_ATOM_SELECTORS, node.type),
  );
  const selector = PROJECTED_ATOM_NAMES.map(
    (name) => PROJECTED_ATOM_SELECTORS[name],
  ).join(",");
  const domAtoms = Array.from(root.querySelectorAll<HTMLElement>(selector));

  if (jsonAtoms.length !== domAtoms.length) {
    const jsonCounts = new Map<string, number>();
    const domCounts = new Map<string, number>();
    for (const node of jsonAtoms) {
      const name = node.type ?? "unknown";
      jsonCounts.set(name, (jsonCounts.get(name) ?? 0) + 1);
    }
    for (const domAtom of domAtoms) {
      const name = selectorName(domAtom) ?? "unknown";
      domCounts.set(name, (domCounts.get(name) ?? 0) + 1);
    }
    const changed = PROJECTED_ATOM_NAMES.filter(
      (name) => (jsonCounts.get(name) ?? 0) !== (domCounts.get(name) ?? 0),
    );
    throw new OfflineHtmlExportError(
      `Offline HTML atom count mismatch: ${changed.join(", ") || "unknown"}`,
    );
  }

  const bindings = new Map<ProjectedAtomName, BoundNode[]>();
  for (let index = 0; index < jsonAtoms.length; index += 1) {
    const jsonAtom = jsonAtoms[index];
    const domAtom = domAtoms[index];
    const jsonName = jsonAtom.type as ProjectedAtomName;
    const domName = selectorName(domAtom);
    if (domName !== jsonName) {
      throw new OfflineHtmlExportError(
        `Offline HTML atom order mismatch at ${index}: expected ${jsonName}, found ${domName ?? "unknown"}`,
      );
    }
    const bucket = bindings.get(jsonName) ?? [];
    bucket.push({ node: jsonAtom, element: domAtom });
    bindings.set(jsonName, bucket);
  }
  return bindings;
}

function assertJsonRequiredForLossyAtoms(root: HTMLElement): void {
  const selector = PROJECTED_ATOM_NAMES.map(
    (name) => PROJECTED_ATOM_SELECTORS[name],
  ).join(",");
  const first = root.querySelector<HTMLElement>(selector);
  if (!first) return;
  throw new OfflineHtmlExportError(
    `Offline HTML export requires the matching editor JSON snapshot for ${selectorName(first) ?? "a native atom"}`,
  );
}

function bindHeadingIds(root: HTMLElement, doc: TipTapNode): void {
  const headings = collectHeadings(doc);
  const elements = Array.from(
    root.querySelectorAll<HTMLHeadingElement>("h1,h2,h3"),
  );
  if (elements.length !== headings.length) {
    throw new OfflineHtmlExportError(
      "Offline HTML heading count does not match the editor JSON snapshot",
    );
  }
  for (let index = 0; index < headings.length; index += 1) {
    const heading = headings[index];
    const target = elements[index];
    if (Number(target.tagName.slice(1)) !== heading.level) {
      throw new OfflineHtmlExportError(
        `Offline HTML heading order mismatch at ${index}`,
      );
    }
    target.id = heading.id;
  }
}

function bindBookmarkIds(root: HTMLElement, doc: TipTapNode): void {
  const expected = new Set(collectBookmarks(doc).map(({ name }) => name));
  for (const target of root.querySelectorAll<HTMLElement>(
    "[data-bookmark-name], [data-bookmark-point]",
  )) {
    const name = sanitizeBookmarkName(
      target.dataset.bookmarkName ?? target.dataset.bookmarkPoint,
    );
    if (!name || !expected.has(name)) continue;
    target.id = `bookmark-${sanitizeAnchorId(name)}`;
    expected.delete(name);
  }
  if (expected.size > 0) {
    throw new OfflineHtmlExportError(
      "Offline HTML bookmark count does not match the editor JSON snapshot",
    );
  }
}

function projectMath(node: TipTapNode, inline: boolean): HTMLElement {
  const latex = boundedText(attrsOf(node).latex, 8_000);
  if (inline) return textElement("code", latex, "mrnote-static-inline-math");
  const pre = element("pre", "mrnote-static-math-block");
  pre.append(textElement("code", latex));
  return pre;
}

function projectFile(node: TipTapNode): HTMLElement {
  const attrs = attrsOf(node);
  const section = element("section", "mrnote-static-file");
  const filename = boundedText(attrs.filename, 1_000);
  section.append(textElement("h2", filename.trim() ? filename : "Attachment"));
  addLabelledText(section, "Media type", attrs.mime_type, 255);
  const size = attrs.size_bytes;
  if (typeof size === "number" && Number.isSafeInteger(size) && size >= 0) {
    addLabelledText(section, "Size", `${size} bytes`, 64);
  } else if (size !== undefined && size !== null) {
    throw new OfflineHtmlExportError(
      "Offline HTML export found an invalid attachment size",
    );
  }
  section.append(
    textElement(
      "p",
      "Attachment content is not embedded in this standalone HTML file.",
      "mrnote-static-file-omitted",
    ),
  );
  return section;
}

function projectAiOutput(node: TipTapNode): HTMLElement {
  const attrs = attrsOf(node);
  const section = element("section", "mrnote-static-ai-output");
  section.append(textElement("h2", "AI output"));
  addLabelledText(section, "Action", attrs.action_type, 255);
  addLabelledText(section, "Model", attrs.model_id, 255);
  addLabelledText(section, "Trace id", attrs.action_log_id, 255);
  const pre = element("pre");
  pre.append(textElement("code", boundedText(attrs.content_markdown, 100_000)));
  section.append(pre);
  if (attrs.sources !== undefined && !Array.isArray(attrs.sources)) {
    throw new OfflineHtmlExportError(
      "Offline HTML export found a malformed AI source list",
    );
  }
  if (Array.isArray(attrs.sources)) {
    if (attrs.sources.length > 1_000) {
      throw new OfflineHtmlExportError(
        "Offline HTML export AI source list exceeds the 1000-entry fidelity limit",
      );
    }
    const list = element("ul");
    for (const candidate of attrs.sources) {
      if (!candidate || typeof candidate !== "object") {
        throw new OfflineHtmlExportError(
          "Offline HTML export found a malformed AI source entry",
        );
      }
      const source = candidate as Record<string, unknown>;
      const title = boundedText(source.title, 1_000);
      const type = boundedText(source.type, 100);
      const id = boundedText(source.id, 255);
      const label = [type, title, id]
        .filter((part) => part.trim().length > 0)
        .join(" · ");
      if (!label.trim()) {
        throw new OfflineHtmlExportError(
          "Offline HTML export found an AI source without display metadata",
        );
      }
      list.append(textElement("li", label));
    }
    if (list.childElementCount > 0) section.append(list);
  }
  return section;
}

function projectReference(node: TipTapNode): HTMLElement {
  const attrs = attrsOf(node);
  const section = element("section", "mrnote-static-reference");
  const title = boundedText(attrs.title, 1_000);
  section.append(textElement("h2", title.trim() ? title : "Reference"));
  addLabelledText(section, "Excerpt", attrs.snippet, 8_000);
  addLabelledText(section, "Target type", attrs.target_type, 100);
  addLabelledText(section, "Target id", attrs.target_id, 255);
  return section;
}

function projectTask(node: TipTapNode): HTMLElement {
  const attrs = attrsOf(node);
  const section = element("section", "mrnote-static-task");
  const completed = attrs.completed === true;
  const title = boundedText(attrs.title, 2_000);
  section.append(
    textElement(
      "h2",
      `${completed ? "☑" : "☐"} ${title.trim() ? title : "Task"}`,
    ),
  );
  addLabelledText(section, "Description", attrs.description, 8_000);
  addLabelledText(section, "Due", attrs.due_date, 100);
  return section;
}

function projectFlashcard(node: TipTapNode): HTMLElement {
  const attrs = attrsOf(node);
  const section = element("section", "mrnote-static-flashcard");
  section.append(textElement("h2", "Flashcard"));
  addLabelledText(section, "Front", attrs.front, 8_000);
  addLabelledText(section, "Back", attrs.back, 8_000);
  return section;
}

function projectCoverPage(
  node: TipTapNode,
  context: ProjectionContext,
): HTMLElement {
  const figure = element("figure", "mrnote-static-cover-page");
  figure.setAttribute("data-type", "document-cover-page");
  const coverContext = documentCoverPageContext(context.attrs);
  const scene = documentCoverPageScene(attrsOf(node), coverContext);
  if (!scene) {
    throw new OfflineHtmlExportError(
      "Offline HTML export cannot draw a cover page that names no design",
    );
  }
  const svg = coverSceneSvgMarkup(
    scene,
    coverContext.values,
    coverContext.fontFamily ?? "sans-serif",
  );
  const image = element("img");
  const bytes = new TextEncoder().encode(svg);
  image.src = `data:image/svg+xml;base64,${btoa(Array.from(bytes, (byte) => String.fromCharCode(byte)).join(""))}`;
  image.alt = coverContext.values.title ?? "";
  image.style.cssText = "display:block;width:100%;height:auto";
  figure.style.cssText = "margin:0";
  figure.append(image);
  return figure;
}

function projectFrontmatter(node: TipTapNode): HTMLElement {
  const section = element("section", "mrnote-static-frontmatter");
  section.append(textElement("h2", "Front matter"));
  const pre = element("pre");
  pre.append(textElement("code", boundedText(attrsOf(node).yaml, 100_000)));
  section.append(pre);
  return section;
}

function staticGeneratedListLeader(
  rightAlignPageNumbers: boolean,
  tabLeader: TocTabLeader,
): string {
  if (!rightAlignPageNumbers) return " ";
  return tabLeader === "dot"
    ? " ........ "
    : tabLeader === "hyphen"
      ? " -------- "
      : tabLeader === "underscore"
        ? " ________ "
        : tabLeader === "middleDot"
          ? " \u00b7\u00b7\u00b7\u00b7\u00b7\u00b7\u00b7\u00b7 "
          : " ";
}

function projectToc(node: TipTapNode, context: ProjectionContext): HTMLElement {
  const attrs = attrsOf(node);
  const settings = normalizeTocSettings(attrs);
  const maximumLevel = tocEntryMaximumLevel(settings);
  const headings =
    normalizeTocEntrySnapshot(attrs.cachedEntries) ??
    collectTocEntries(context.doc, tocSourceLevelResolver(settings));
  const nav = element("nav", "mrnote-static-toc");
  nav.setAttribute("aria-label", "Table of contents");
  nav.setAttribute(
    "data-show-page-numbers",
    settings.showPageNumbers ? "true" : "false",
  );
  nav.setAttribute(
    "data-right-align-page-numbers",
    settings.rightAlignPageNumbers ? "true" : "false",
  );
  nav.setAttribute("data-tab-leader", settings.tabLeader);
  // The table's own title (Word's "TOC Heading" paragraph), as the editor
  // shows it; a table without one has none.
  if (settings.title !== null) nav.append(textElement("h2", settings.title));
  const list = semanticList();
  for (const heading of headings) {
    if (heading.level > maximumLevel) continue;
    const item = element("li");
    item.setAttribute("data-level", String(heading.level));
    const label = settings.hyperlinks ? element("a") : element("span");
    if (settings.hyperlinks) {
      (label as HTMLAnchorElement).href = `#${heading.id}`;
    }
    // A numbered heading's row: its number, a tab, its words (Word's).
    if (heading.number) {
      label.append(
        textElement("span", heading.number, "mrnote-static-toc-number"),
        "\t",
      );
    }
    if (heading.runs && heading.text) {
      // The emphasis on part of the heading the row keeps.
      for (const run of heading.runs) {
        let piece: Node = document.createTextNode(run.text);
        if (run.underline) piece = wrapped("u", piece);
        if (run.bold) piece = wrapped("strong", piece);
        if (run.italic) piece = wrapped("em", piece);
        label.append(piece);
      }
    } else {
      label.append(heading.text || heading.id);
    }
    item.append(label);
    if (settings.showPageNumbers) {
      const leader = staticGeneratedListLeader(
        settings.rightAlignPageNumbers,
        settings.tabLeader,
      );
      item.append(textElement("span", leader, "mrnote-static-toc-leader"));
      item.append(
        textElement(
          "span",
          // The page Word printed in an imported row ("iv") until updated.
          heading.pageText ?? String(heading.pageNumber ?? 1),
          "mrnote-static-toc-page-number",
        ),
      );
    }
    list.append(item);
  }
  nav.append(list);
  return nav;
}

function projectCaption(
  caption: ProjectionContext["captions"][number],
): HTMLElement {
  const result = textElement(
    "figcaption",
    captionDisplayLabel(caption),
    "mrnote-static-caption",
  );
  result.id = caption.id;
  result.setAttribute("data-caption-kind", caption.kind);
  return result;
}

function projectListOfFigures(
  node: TipTapNode,
  context: ProjectionContext,
): HTMLElement {
  const attrs = attrsOf(node);
  const kind = normalizeCaptionListLabel(attrs.kind);
  const settings = normalizeGeneratedListSettings(attrs);
  const entries =
    normalizeGeneratedCaptionListEntries(attrs.cachedEntries) ??
    generatedCaptionListEntries(
      captionsForListLabel(collectCaptions(context.doc), kind),
      { includeLabel: settings.includeLabel },
    );
  const section = element("section", "mrnote-static-list-of-figures");
  section.setAttribute(
    "data-show-page-numbers",
    settings.showPageNumbers ? "true" : "false",
  );
  section.setAttribute(
    "data-right-align-page-numbers",
    settings.rightAlignPageNumbers ? "true" : "false",
  );
  section.setAttribute("data-tab-leader", settings.tabLeader);
  // The heading names the caption label the list collects (a Chinese
  // session's 表格 is the built-in Table), as the editor's header does.
  const builtinKind = captionListBuiltinKind(kind);
  section.append(
    textElement(
      "h2",
      kind === "none"
        ? "List of captions"
        : builtinKind === null
          ? `${captionListLabelText(kind)} list`
          : `List of ${builtinKind === "table" ? "tables" : builtinKind === "equation" ? "equations" : "figures"}`,
    ),
  );
  const list = semanticList();
  for (const entry of entries) {
    const item = element("li");
    const content = element(
      settings.hyperlinks ? "a" : "span",
      "mrnote-static-list-of-figures-entry",
    );
    if (settings.hyperlinks) {
      (content as HTMLAnchorElement).href = `#${entry.id}`;
    }
    content.append(
      textElement("span", entry.text, "mrnote-static-list-of-figures-label"),
    );
    if (settings.showPageNumbers) {
      content.append(
        textElement(
          "span",
          staticGeneratedListLeader(
            settings.rightAlignPageNumbers,
            settings.tabLeader,
          ),
          "mrnote-static-list-of-figures-leader",
        ),
        textElement(
          "span",
          String(entry.pageNumber),
          "mrnote-static-list-of-figures-page-number",
        ),
      );
    }
    item.append(content);
    list.append(item);
  }
  section.append(list);
  return section;
}

function crossRefTarget(
  node: TipTapNode,
  context: ProjectionContext,
): { id: string; label: string } {
  const attrs = attrsOf(node);
  const targetType = attrs.targetType ?? "heading";
  const bookmarkName =
    targetType === "bookmark" ? sanitizeBookmarkName(attrs.targetAnchorId) : "";
  const noteId =
    targetType === "footnote" || targetType === "endnote"
      ? sanitizeFootnoteId(attrs.targetAnchorId)
      : null;
  const id = bookmarkName
    ? `bookmark-${sanitizeAnchorId(bookmarkName)}`
    : noteId
      ? `note-ref-${noteId}`
      : sanitizeAnchorId(attrs.targetAnchorId);
  if (
    attrs.targetType !== undefined &&
    attrs.targetType !== null &&
    attrs.targetType !== "heading" &&
    attrs.targetType !== "figure" &&
    attrs.targetType !== "table" &&
    attrs.targetType !== "bookmark" &&
    attrs.targetType !== "footnote" &&
    attrs.targetType !== "endnote"
  ) {
    throw new OfflineHtmlExportError(
      "Offline HTML export found an invalid cross-reference target type",
    );
  }
  const type =
    attrs.targetType === "figure" || attrs.targetType === "table"
      ? attrs.targetType
      : attrs.targetType === "bookmark"
        ? "bookmark"
        : attrs.targetType === "footnote" || attrs.targetType === "endnote"
          ? attrs.targetType
          : "heading";
  if (!id) {
    throw new OfflineHtmlExportError(
      "Offline HTML export found a cross-reference without a safe target id",
    );
  }
  if (type === "bookmark") {
    if (!context.bookmarks.some((bookmark) => bookmark.name === bookmarkName)) {
      throw new OfflineHtmlExportError(
        `Offline HTML export cannot resolve bookmark cross-reference target ${bookmarkName}`,
      );
    }
    return { id, label: boundedText(attrs.label || bookmarkName, 300) };
  }
  if (type === "footnote" || type === "endnote") {
    const mode = resolveFootnoteMode(context.attrs.footnoteMode);
    const footnotes = context.footnotes;
    if (
      !noteId ||
      !footnotes ||
      noteStreamOf(noteId, mode) !== type ||
      !footnotes.refs.some((ref) => ref.id === noteId)
    ) {
      throw new OfflineHtmlExportError(
        `Offline HTML export cannot resolve ${type} cross-reference target ${noteId ?? ""}`,
      );
    }
    return {
      id,
      label:
        footnotes.labels.get(noteId) || boundedText(attrs.label || noteId, 300),
    };
  }
  if (type === "heading") {
    const target = context.headings.find((heading) => heading.id === id);
    if (target) return { id, label: target.text || target.id };
  } else {
    const target = context.captions.find((caption) => caption.id === id);
    if (target) return { id, label: captionDisplayLabel(target) };
  }
  throw new OfflineHtmlExportError(
    `Offline HTML export cannot resolve cross-reference target ${id}`,
  );
}

function projectCrossRef(
  node: TipTapNode,
  context: ProjectionContext,
): HTMLElement {
  const target = crossRefTarget(node, context);
  const hyperlink = attrsOf(node).hyperlink;
  if (hyperlink !== undefined && typeof hyperlink !== "boolean") {
    throw new OfflineHtmlExportError(
      "Offline HTML export found an invalid cross-reference hyperlink option",
    );
  }
  if (hyperlink === false) {
    const text = textElement("span", target.label, "mrnote-static-cross-ref");
    text.dataset.targetAnchorId = target.id;
    return text;
  }
  const link = textElement("a", target.label, "mrnote-static-cross-ref");
  link.href = `#${target.id}`;
  return link;
}

/**
 * The source records of the document's citation store. The editor keeps the
 * store as a versioned envelope (`{ version: 1, items, nativeAuthority? }`,
 * the doc attribute's default); older snapshots carry the bare array. Any
 * other shape is refused rather than exported with its citations unresolved.
 * `nativeAuthority` is the DOCX save path's receipt and has no HTML meaning.
 */
function citationStoreItems(value: unknown): unknown[] | undefined {
  if (value === undefined) return undefined;
  if (Array.isArray(value)) return value;
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    if (record.version === 1 && Array.isArray(record.items)) return record.items;
  }
  throw new OfflineHtmlExportError(
    "Offline HTML export found a malformed citation store",
  );
}

function citationProjection(context: ProjectionContext): {
  nodes: TipTapNode[];
  labels: string[];
  bibliography: ReturnType<typeof buildBibliography>;
} {
  const nodes = collectNodes(context.doc).filter(
    (node) => node.type === "citation",
  );
  const items = citationStoreItems(context.attrs.sources);
  if (items !== undefined && items.length > MAX_SOURCES) {
    throw new OfflineHtmlExportError(
      `Offline HTML export citation store exceeds the ${MAX_SOURCES}-source fidelity limit`,
    );
  }
  const sources = normalizeSources(items);
  const style = resolveCitationStyle(context.attrs.citationStyle);
  const orderedIds: string[] = [];
  const labels = nodes.map((node) => {
    const attrs = attrsOf(node);
    const id = sanitizeSourceId(attrs.sourceId);
    if (!id) {
      throw new OfflineHtmlExportError(
        "Offline HTML export found a citation without a safe source id",
      );
    }
    const source = findSource(sources, id);
    if (!source) {
      throw new OfflineHtmlExportError(
        `Offline HTML export is missing citation source ${id}`,
      );
    }
    orderedIds.push(id);
    return formatInText(style, source, orderedIds, {
      locator: sanitizeLocator(attrs.locator),
      prefix: sanitizeCslField(attrs.prefix),
    });
  });
  // IEEE numbering depends on the complete first-citation order, not the
  // prefix seen while mapping. Recompute after collecting all ids.
  for (let index = 0; index < nodes.length; index += 1) {
    const attrs = attrsOf(nodes[index]);
    const id = sanitizeSourceId(attrs.sourceId);
    const source = id ? findSource(sources, id) : null;
    if (!source) continue;
    labels[index] = formatInText(
      style,
      source,
      orderedIds,
      {
        locator: sanitizeLocator(attrs.locator),
        prefix: sanitizeCslField(attrs.prefix),
      },
      sources,
    );
  }
  return {
    nodes,
    labels,
    bibliography: buildBibliography(style, sources, orderedIds),
  };
}

function renderBibliography(
  entries: ReturnType<typeof buildBibliography>,
): HTMLElement {
  const section = element("section", "mrnote-static-bibliography");
  section.append(textElement("h2", "Bibliography"));
  const list = semanticList();
  for (const entry of entries) {
    list.append(
      textElement(
        "li",
        entry.label === null ? entry.text : `${entry.label} ${entry.text}`,
      ),
    );
  }
  section.append(list);
  return section;
}

function appendRichFootnote(parent: HTMLElement, rich: FootnoteRichBody): void {
  for (const paragraph of rich) {
    const p = element("p");
    for (const run of paragraph) {
      if (typeof run.math === "string" && run.math) {
        // A note formula: the same static in-line equation as the body's.
        p.append(textElement("code", boundedText(run.math, 8_000), "mrnote-static-inline-math"));
        continue;
      }
      let child: HTMLElement = textElement("span", run.text);
      for (const [flag, tag] of [
        ["bold", "strong"],
        ["italic", "em"],
        ["underline", "u"],
        ["strike", "s"],
        ["superscript", "sup"],
        ["subscript", "sub"],
      ] as const) {
        if (run[flag] !== true) continue;
        const wrapper = element(tag);
        wrapper.append(child);
        child = wrapper;
      }
      if (typeof run.highlight === "string" && /^#[0-9a-f]{6}$/i.test(run.highlight)) {
        const wrapper = element("mark");
        wrapper.setAttribute("data-highlight-color", run.highlight);
        wrapper.setAttribute("style", `background-color: ${run.highlight}; color: inherit`);
        wrapper.append(child);
        child = wrapper;
      }
      if (run.hidden === true) {
        const wrapper = element("span");
        wrapper.setAttribute("data-hidden-text", "");
        wrapper.setAttribute("style", "display: none");
        wrapper.append(child);
        child = wrapper;
      }
      p.append(child);
    }
    parent.append(p);
  }
}

function footnoteProjection(context: ProjectionContext): {
  refs: OrderedNoteRef[];
  labels: Map<string, string>;
  renderList: () => HTMLElement;
} {
  // Each marker numbers with its section's settings (Word's per-section
  // w:footnotePr / w:endnotePr), so remember which section holds it.
  const sectionOf = new Map<TipTapNode, number>();
  let sectionIndex = 0;
  const nodes = collectNodes(context.doc).filter((node) => {
    if (node.type === "sectionBreak") sectionIndex += 1;
    if (node.type !== "footnoteRef") return false;
    sectionOf.set(node, sectionIndex);
    return true;
  });
  if (
    context.attrs.footnotes !== undefined &&
    !Array.isArray(context.attrs.footnotes)
  ) {
    throw new OfflineHtmlExportError(
      "Offline HTML export found a malformed footnote store",
    );
  }
  if (
    Array.isArray(context.attrs.footnotes) &&
    context.attrs.footnotes.length > MAX_FOOTNOTES
  ) {
    throw new OfflineHtmlExportError(
      `Offline HTML export footnote store exceeds the ${MAX_FOOTNOTES}-note fidelity limit`,
    );
  }
  const notes = normalizeFootnotes(context.attrs.footnotes);
  const refs: OrderedNoteRef[] = nodes.map((node) => {
    const attrs = attrsOf(node);
    const id = sanitizeFootnoteId(attrs.noteId);
    if (!id) {
      throw new OfflineHtmlExportError(
        "Offline HTML export found a footnote without a safe note id",
      );
    }
    if (!findFootnote(notes, id)) {
      throw new OfflineHtmlExportError(
        `Offline HTML export is missing footnote record ${id}`,
      );
    }
    return {
      id,
      customMark: sanitizeCustomMark(attrs.customMark),
      sectionIndex: sectionOf.get(node) ?? 0,
    };
  });
  const streams = numberNoteStreams(
    notes,
    refs,
    resolveFootnoteMode(context.attrs.footnoteMode),
    documentNoteSettings(context.attrs),
  );
  const labels = new Map(
    [...streams.footnotes, ...streams.endnotes].map((note) => [
      note.id,
      note.label,
    ]),
  );
  const renderList = (): HTMLElement => {
    const section = element("section", "mrnote-static-footnotes");
    for (const [heading, notesInStream] of [
      ["Footnotes", streams.footnotes],
      ["Endnotes", streams.endnotes],
    ] as const) {
      if (notesInStream.length === 0) continue;
      section.append(textElement("h2", heading));
      const list = semanticList();
      for (const note of notesInStream) {
        const item = element("li");
        item.id = `note-${note.id}`;
        item.append(
          textElement("span", `${note.label} `, "mrnote-static-note-label"),
        );
        if (note.rich) appendRichFootnote(item, note.rich);
        else item.append(document.createTextNode(note.body));
        list.append(item);
      }
      section.append(list);
    }
    return section;
  };
  return { refs, labels, renderList };
}

const projectedChartImages = new WeakSet<HTMLImageElement>();
export function isProjectedChartImage(image: HTMLImageElement): boolean {
  return projectedChartImages.has(image);
}

/** Only the bounded local drop-shadow form emitted by the chart renderer. */
export function isSafeRenderedChartShadowFilters(root: Element): boolean {
  const filters = Array.from(root.querySelectorAll("filter"));
  if (filters.length > 128) return false;
  const ids = new Set<string>();
  const number = (value: string | null, min: number, max: number) => value !== null && value.trim() !== "" && Number.isFinite(Number(value)) && Number(value) >= min && Number(value) <= max;
  const color = (value: string) => {
    if (/^#(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value)) return true;
    const match = /^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/.exec(value);
    return Boolean(match && match.slice(1).every(part => number(part, 0, 255)));
  };
  for (const filter of filters) {
    const id = filter.getAttribute("id");
    if (!id || !/^zr\d+-s\d+$/.test(id) || ids.has(id) || filter.namespaceURI !== "http://www.w3.org/2000/svg" ||
        filter.attributes.length !== 5 || filter.getAttribute("x") !== "-100%" || filter.getAttribute("y") !== "-100%" ||
        filter.getAttribute("width") !== "300%" || filter.getAttribute("height") !== "300%" || filter.children.length !== 1) return false;
    const shadow = filter.children[0];
    if (shadow.localName !== "feDropShadow" || shadow.namespaceURI !== "http://www.w3.org/2000/svg" || shadow.textContent?.trim() || shadow.children.length || shadow.attributes.length !== 5 ||
        !number(shadow.getAttribute("dx"), -64, 64) || !number(shadow.getAttribute("dy"), -64, 64) ||
        !number(shadow.getAttribute("flood-opacity"), 0, 1) || !color(shadow.getAttribute("flood-color") ?? "")) return false;
    const deviations = (shadow.getAttribute("stdDeviation") ?? "").trim().split(/\s+/);
    if (deviations.length !== 2 || !deviations.every(value => number(value, 0, 32))) return false;
    ids.add(id);
  }
  if (root.querySelectorAll("feDropShadow").length !== filters.length) return false;
  for (const node of Array.from(root.querySelectorAll("*"))) {
    const filter = node.getAttribute("filter");
    if (filter !== null) {
      const id = /^url\(#(zr\d+-s\d+)\)$/.exec(filter)?.[1];
      if (!id || !ids.has(id)) return false;
    }
    if (/filter\s*:/i.test(node.getAttribute("style") ?? "")) return false;
  }
  return true;
}

function projectChart(
  node: TipTapNode,
  documentAttrs: Record<string, unknown>,
  pictures: DocumentExportChartPictures | undefined,
): HTMLElement {
  const attrs=attrsOf(node), normalized=normalizeChartBlockAttrs(attrs);
  if (!normalized.validation.valid) throw new OfflineHtmlExportError("Invalid chart semantic data");
  // Drawn by the server's chart engine with what the editor's canvas draws it
  // with (the imported look, the document theme a Change Colors pick resolves
  // against, the sheet's number format), fetched before this synchronous
  // projection (`prefetchDocumentExportCharts`).
  const picture=pictures?.get(documentExportChartKey(attrs,documentAttrs));
  if (!picture) throw new OfflineHtmlExportError("This chart needs a connection to be drawn");
  if (picture.flattened || picture.renderMode!=="native" || normalized.attrs.chartType==="map")
    throw new OfflineHtmlExportError("This chart requires an unavailable offline renderer");
  const { width, height } = documentExportChartRequest(attrs,documentAttrs);
  if (width*height>16_000_000) throw new OfflineHtmlExportError("Chart dimensions exceed the offline budget");
  // Only renderer-generated SVG enters this branch. Independently verify its
  // inert element/attribute grammar; labels were escaped by the SVG renderer.
  const parsed=new DOMParser().parseFromString(picture.svg,"image/svg+xml");
  if (parsed.querySelector("parsererror")) throw new OfflineHtmlExportError("Chart renderer produced non-inert SVG");
  parsed.documentElement.removeAttribute("xmlns:xlink");
  parsed.querySelectorAll("style").forEach(style=>style.remove());
  // A WordArt title is the engine's inlined vector paint (the part marked
  // "title", the mark on the group that places it): set aside and checked on
  // its own bounded grammar, then put back, as the title vector was inserted
  // after this check before.
  const unsafe=(el: Element, allowed: ReadonlySet<string>)=>!allowed.has(el.localName) || [...el.attributes].some(a=>
      a.name==="xmlns" ? a.value!=="http://www.w3.org/2000/svg" :
      /^on/i.test(a.name) || /(?:href|src)$/i.test(a.name) || /(?:javascript:|data:|https?:|@import)/i.test(a.value));
  const titleAllowed=new Set(["svg","defs","linearGradient","stop","g","text","tspan","rect","mask","filter","feDropShadow"]);
  const titles=[...parsed.documentElement.querySelectorAll('svg[ecmeta_part="title"], g[ecmeta_part="title"] > svg')].map(vector=>{
    if (vector.querySelectorAll("*").length>80 || [vector, ...vector.querySelectorAll("*")].some(el=>unsafe(el,titleAllowed)))
      throw new OfflineHtmlExportError("Chart title exceeds the bounded vector surface");
    const placeholder=parsed.createComment("generated-title");
    vector.replaceWith(placeholder);
    return { vector, placeholder };
  });
  // The engine's picture keeps the editor's hit marks; the file does not.
  for (const el of [...parsed.querySelectorAll("*"), ...titles.flatMap(({vector})=>[vector, ...vector.querySelectorAll("*")])])
    for (const a of [...el.attributes]) if (a.name.startsWith("ecmeta_")) el.removeAttribute(a.name);
  const allowed=new Set(["svg","g","path","rect","circle","ellipse","line","polyline","polygon","text","tspan","defs","clipPath","linearGradient","radialGradient","stop","filter","feDropShadow"]);
  if (!isSafeRenderedChartShadowFilters(parsed.documentElement) || [...parsed.querySelectorAll("*")].some(el=>unsafe(el,allowed)))
    throw new OfflineHtmlExportError("Chart renderer produced non-inert SVG");
  for (const { vector, placeholder } of titles) placeholder.replaceWith(vector);
  const svg=new XMLSerializer().serializeToString(parsed.documentElement);
  const figure=element("figure");
  figure.setAttribute("data-type","chart");
  figure.style.cssText=`margin:0;width:${width}px;height:${height}px`;
  const image=element("img");
  const bytes=new TextEncoder().encode(svg);
  image.src=`data:image/svg+xml;base64,${btoa(Array.from(bytes,byte=>String.fromCharCode(byte)).join(""))}`;
  image.alt=normalized.attrs.title;
  image.style.cssText="display:block;width:100%;height:100%";
  projectedChartImages.add(image);
  figure.append(image);
  return figure;
}

function projectSimpleAtoms(context: ProjectionContext): void {
  const replacements: Partial<
    Record<ProjectedAtomName, (node: TipTapNode) => HTMLElement>
  > = {
    chart: (node) => projectChart(node, context.attrs, context.chartPictures),
    inlineChart: (node) => {
      // The same static chart picture, as a box on its text line whose
      // bottom stands on the baseline (Word's `wp:inline` chart).
      const figure = projectChart(node, context.attrs, context.chartPictures);
      const run = element("span");
      run.setAttribute("data-type", "inline-chart");
      run.style.cssText = `${figure.style.cssText};display:inline-block;vertical-align:baseline`;
      run.append(...Array.from(figure.childNodes));
      return run;
    },
    mathBlock: (node) => projectMath(node, false),
    inlineMath: (node) => projectMath(node, true),
    file: projectFile,
    ai_output: projectAiOutput,
    reference: projectReference,
    task: projectTask,
    flashcard: projectFlashcard,
    frontmatter: projectFrontmatter,
  };
  for (const [name, render] of Object.entries(replacements) as Array<
    [ProjectedAtomName, (node: TipTapNode) => HTMLElement]
  >) {
    for (const binding of context.bindings.get(name) ?? []) {
      binding.element.replaceWith(render(binding.node));
    }
  }
}

/**
 * Replace every NodeView-only/lossy native atom with deterministic static DOM.
 * The source HTML and JSON must contain the exact same projected atom sequence;
 * user data is assigned through textContent/attributes only, never parsed as
 * markup. The caller keeps this root inert and sanitizes it afterwards.
 */
export function projectStaticNativeNodes(
  root: HTMLElement,
  editorJson?: TipTapNode,
  /** The document's chart pictures, drawn by the server beforehand. */
  chartPictures?: DocumentExportChartPictures,
): void {
  if (!editorJson) {
    assertJsonRequiredForLossyAtoms(root);
    return;
  }
  const bindings = bindProjectedAtoms(root, editorJson);
  for (const rejected of REJECTED_ATOM_NAMES) {
    if ((bindings.get(rejected)?.length ?? 0) > 0) {
      throw new OfflineHtmlExportError(
        `Offline HTML export cannot safely render a ${rejected} without a deterministic static renderer`,
      );
    }
  }

  bindHeadingIds(root, editorJson);
  bindBookmarkIds(root, editorJson);
  const context: ProjectionContext = {
    doc: editorJson,
    attrs: attrsOf(editorJson),
    bindings,
    headings: collectHeadings(editorJson),
    captions: collectCaptions(editorJson),
    bookmarks: collectBookmarks(editorJson),
    chartPictures,
  };

  projectSimpleAtoms(context);

  const footnotes = footnoteProjection(context);
  context.footnotes = footnotes;

  for (const binding of bindings.get("tableOfContents") ?? []) {
    binding.element.replaceWith(projectToc(binding.node, context));
  }
  for (const binding of bindings.get("documentCoverPage") ?? []) {
    binding.element.replaceWith(projectCoverPage(binding.node, context));
  }

  const captionBindings = bindings.get("caption") ?? [];
  if (captionBindings.length !== context.captions.length) {
    throw new OfflineHtmlExportError(
      "Offline HTML caption count does not match the editor JSON snapshot",
    );
  }
  for (let index = 0; index < captionBindings.length; index += 1) {
    captionBindings[index].element.replaceWith(
      projectCaption(context.captions[index]),
    );
  }
  for (const binding of bindings.get("listOfFigures") ?? []) {
    binding.element.replaceWith(projectListOfFigures(binding.node, context));
  }
  for (const binding of bindings.get("crossRef") ?? []) {
    binding.element.replaceWith(projectCrossRef(binding.node, context));
  }

  const citation = citationProjection(context);
  const citationBindings = bindings.get("citation") ?? [];
  for (let index = 0; index < citationBindings.length; index += 1) {
    citationBindings[index].element.replaceWith(
      textElement("span", citation.labels[index], "mrnote-static-citation"),
    );
  }
  for (const binding of bindings.get("bibliography") ?? []) {
    binding.element.replaceWith(renderBibliography(citation.bibliography));
  }

  const footnoteBindings = bindings.get("footnoteRef") ?? [];
  for (let index = 0; index < footnoteBindings.length; index += 1) {
    const ref = footnotes.refs[index];
    const label = footnotes.labels.get(ref.id);
    if (!label) {
      throw new OfflineHtmlExportError(
        `Offline HTML export cannot number footnote ${ref.id}`,
      );
    }
    const sup = element("sup", "mrnote-static-footnote-ref");
    sup.id = `note-ref-${ref.id}`;
    const link = textElement("a", label);
    link.href = `#note-${ref.id}`;
    sup.append(link);
    footnoteBindings[index].element.replaceWith(sup);
  }
  const footnoteLists = bindings.get("footnoteList") ?? [];
  if (footnotes.refs.length > 0 && footnoteLists.length === 0) {
    throw new OfflineHtmlExportError(
      "Offline HTML export would omit footnote bodies because the document has no footnote list",
    );
  }
  if (footnoteLists.length > 1) {
    throw new OfflineHtmlExportError(
      "Offline HTML export found multiple footnote lists",
    );
  }
  for (const binding of footnoteLists) {
    binding.element.replaceWith(footnotes.renderList());
  }
}

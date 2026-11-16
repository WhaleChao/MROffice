/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Typed truth table for native-document feature support.
 *
 * A registered editor atom must have an explicit decision at every binary
 * boundary. `unverified` is intentionally different from `unsupported`: it is
 * a visible audit debt, never permission for an exporter to guess.
 */

export type DocumentFeatureCapabilityState =
  | "native"
  | "projected"
  | "metadata-only"
  | "unsupported"
  | "unverified";

export interface DocumentAtomCapability {
  kind: "atom" | "document-attribute";
  editor: DocumentFeatureCapabilityState;
  docxImport: DocumentFeatureCapabilityState;
  docxOfflineExport: DocumentFeatureCapabilityState;
  docxServerExport: DocumentFeatureCapabilityState;
}

function atom(
  editor: DocumentFeatureCapabilityState,
  docxImport: DocumentFeatureCapabilityState,
  docxOfflineExport: "native" | "unsupported",
  docxServerExport: "native" | "unsupported",
): DocumentAtomCapability & { kind: "atom" } {
  return {
    kind: "atom",
    editor,
    docxImport,
    docxOfflineExport,
    docxServerExport,
  };
}

/**
 * Closed atom universe shared by the schema contract and both DOCX gates.
 * Keep the server column aligned with `document_export_capabilities.py`; the
 * parity test fails on either a missing key or a changed disposition.
 */
export const DOCUMENT_ATOM_CAPABILITIES = {
  ai_output: atom("native", "unverified", "unsupported", "unsupported"),
  bibliography: atom("native", "native", "native", "native"),
  // Word's point bookmark (Insert > Bookmark on a collapsed caret): an
  // adjacent bookmarkStart/bookmarkEnd pair.
  bookmarkPoint: atom("native", "native", "native", "native"),
  // Word's editing exception markers (w:permStart / w:permEnd).
  rangePermission: atom("native", "native", "native", "native"),
  caption: atom("native", "native", "native", "native"),
  chart: atom("native", "native", "native", "native"),
  citation: atom("native", "native", "native", "native"),
  codeChart: atom("native", "unverified", "unsupported", "unsupported"),
  columnBreak: atom("native", "native", "native", "native"),
  crossRef: atom("native", "native", "native", "native"),
  // The server writes an imported HTML chunk back as the w:altChunk and part
  // it was, from the bytes the carrier holds (refused when it holds none).
  documentAltChunkReadOnly: atom(
    "native",
    "projected",
    "unsupported",
    "native",
  ),
  documentField: atom("native", "native", "native", "native"),
  // One Word cover page (Insert > Cover Page): an MRAI design or a cover
  // imported from Word, written back as the Cover Pages gallery content control.
  documentCoverPage: atom("native", "native", "native", "native"),
  documentMedia: atom("native", "native", "unsupported", "native"),
  documentShape: atom("native", "native", "native", "native"),
  // Editor support includes the closed unsigned authoring dialog and preserving
  // placeholder view. Digital signing itself remains outside this atom.
  documentSignatureLine: atom(
    "native",
    "native",
    "native",
    "native",
  ),
  documentVirtualTable: atom(
    "native",
    "native",
    "native",
    "native",
  ),
  // The Word export embeds an attached file as Word's own embedded object
  // (an Office file as its package, any other in an OLE Package).
  file: atom("native", "unverified", "unsupported", "native"),
  flashcard: atom("native", "unverified", "unsupported", "unsupported"),
  footnoteList: atom("native", "native", "native", "native"),
  footnoteRef: atom("native", "native", "native", "native"),
  frontmatter: atom("native", "unverified", "unsupported", "unsupported"),
  hardBreak: atom("native", "native", "native", "native"),
  horizontalRule: atom("native", "native", "native", "native"),
  image: atom("native", "native", "native", "native"),
  // Word "In Line with Text" chart / picture / shape (`wp:inline` run among a
  // paragraph's text); both exporters write the run.
  inlineChart: atom("native", "projected", "native", "native"),
  inlineDocumentShape: atom("native", "projected", "native", "native"),
  inlineImage: atom("native", "projected", "native", "native"),
  inlineMath: atom("native", "native", "native", "native"),
  listOfFigures: atom("native", "native", "native", "native"),
  mathBlock: atom("native", "native", "native", "native"),
  mergeField: atom("native", "native", "native", "native"),
  mermaid: atom("native", "unverified", "unsupported", "unsupported"),
  pageBreak: atom("native", "native", "native", "native"),
  reference: atom("native", "unverified", "unsupported", "unsupported"),
  ruby: atom("native", "native", "native", "native"),
  sectionBreak: atom("native", "native", "native", "native"),
  smartArt: atom("native", "unverified", "unsupported", "unsupported"),
  sourcePageBreak: atom("native", "native", "native", "native"),
  tableOfContents: atom("native", "native", "native", "native"),
  task: atom("native", "unverified", "unsupported", "unsupported"),
  whiteboard: atom("native", "unverified", "unsupported", "unsupported"),
  // A board's own objects: exported (or refused) only as part of their board.
  whiteboardObject: atom("native", "unverified", "unsupported", "unsupported"),
} as const satisfies Record<string, DocumentAtomCapability>;

export type DocumentAtomName = keyof typeof DOCUMENT_ATOM_CAPABILITIES;
export type DocxExportAtomDisposition = "native" | "reject";
export type DocxExportAtomPolicyMap = {
  readonly [Name in DocumentAtomName]: DocxExportAtomDisposition;
};

function deriveDocxPolicy(
  surface: "docxOfflineExport" | "docxServerExport",
): DocxExportAtomPolicyMap {
  return Object.freeze(
    Object.fromEntries(
      Object.entries(DOCUMENT_ATOM_CAPABILITIES).map(([name, capability]) => [
        name,
        capability[surface] === "native" ? "native" : "reject",
      ]),
    ),
  ) as DocxExportAtomPolicyMap;
}

/** Browser/offline serializer policy consumed by `docx-export.ts`. */
export const DOCX_EXPORT_ATOM_POLICIES = deriveDocxPolicy("docxOfflineExport");

/** Canonical server declaration expected from the independent Python gate. */
export const SERVER_DOCX_ATOM_POLICIES = deriveDocxPolicy("docxServerExport");

/**
 * Closed document-attribute universe shared with the canonical Python DOCX
 * export gate.  The server disposition is proven independently below; import
 * and browser-export stay `unverified` until their own per-attribute evidence
 * exists.  This prevents a native server renderer from being mistaken for a
 * verified end-to-end round trip.
 */
const SERVER_NATIVE_DOCUMENT_ATTRIBUTE_NAMES = [
  "headerText",
  "footerText",
  "showPageNumbers",
  "pageNumberFormat",
  "pageNumberStart",
  "pageNumberStartSection",
  "documentMetadata",
  "differentFirstPage",
  "firstPageHeader",
  "firstPageFooter",
  "differentOddEven",
  "evenPageHeader",
  "evenPageFooter",
  "headerRich",
  "footerRich",
  "firstPageHeaderRich",
  "firstPageFooterRich",
  "evenPageHeaderRich",
  "evenPageFooterRich",
  "sectionBands",
  "sectionLayouts",
  "sectionStoryGraph",
  "autoHyphenation",
  "wordSettings",
  "pageSize",
  "pageMargins",
  "pageOrientation",
  "documentProtection",
  // Track Changes > For Everyone (`<w:trackRevisions/>` in settings.xml).
  "trackRevisions",
  "customPageSize",
  "customPageMargins",
  "columns",
  "columnSeparator",
  "writingMode",
  "manuscriptGrid",
  "docGridLinePitch",
  "defaultFontFamily",
  "defaultFontSize",
  "defaultTextColor",
  "defaultLineHeight",
  "defaultTextAlign",
  "characterSpacingControl",
  "doNotUseEastAsianBreakRules",
  "defaultKinsoku",
  "defaultParagraphSpaceBefore",
  "defaultParagraphSpaceAfter",
  // Normal's indents (Paragraph > Set As Default), written as its w:ind.
  "defaultIndentLeftPt",
  "defaultIndentRightPt",
  "defaultFirstLineIndentPt",
  "namedStyles",
  "headingStyles",
  // Modified TOC 1..9 / Table of Figures styles (Modify Style).
  "referenceStyles",
  "styleDefinitions",
  "watermark",
  "pageBorder",
  "lineNumbers",
  "pageBackgroundColor",
  // The page's Fill Effects gradient (w:background/v:background/v:fill).
  "pageBackgroundFill",
  // File > Properties > Custom (docProps/custom.xml).
  "customProperties",
  // Layout > Hyphenation > Hyphenation Options (zone, limit, CAPS).
  "hyphenationOptions",
  // A handwriting document's cover and paper (lib/handwriting-paper): the
  // paper is written as a picture behind the text in each header, the way
  // Word keeps a picture watermark.
  "handwritingDocument",
  // The Design tab's colour palette, carried by the package's theme part.
  "designColorSet",
  // The Design tab's theme schemes, style set and Paragraph Spacing
  // (lib/document-design.ts): the package's theme part and styles.xml.
  "documentDesign",
  "footnotes",
  "footnoteMode",
  "footnoteNumFmt",
  "endnoteNumFmt",
  "footnoteNumStart",
  "endnoteNumStart",
  "footnotePosition",
  "endnotePosition",
  "footnoteNumRestart",
  "endnoteNumRestart",
  "sources",
  "citationStyle",
  "headingNumbering",
  // Word list styles (Define New List Style): both writers emit them as
  // numbering styles (abstractNum w:styleLink + w:numStyleLink users) and
  // the importer reads them back.
  "listStyles",
  "commentThreads",
] as const;

const SERVER_METADATA_DOCUMENT_ATTRIBUTE_NAMES = [
  // Read-only evidence copied out of the source font table so the editor can
  // reproduce Word's East Asian single-line pitch. Word rebuilds its own font
  // table from the fonts a document actually uses, so this must never block or
  // alter DOCX serialization.
  "eastAsianFontFamilies",
  // The source's Normal/docDefaults eastAsia face where their ascii face is an
  // East Asian one too: the face the defaults' East Asian text is set in.
  "defaultEastAsianFontFamily",
  // The source's Word compatibility mode, read for layout rules that differ
  // by mode; the package keeps its own compat settings.
  "wordCompatibilityMode",
  "wordTableLinesOnGrid",
  // Word's "Don't use HTML paragraph auto spacing": paragraph spacing adds
  // up instead of overlapping. Read-only layout evidence.
  "wordSpacingAddsUp",
  "eastAsianLanguage",
  "wordStrictKinsoku",
  "wordBalanceSbcsDbcs",
  "wordBalancedSpaces",
  "wordDoNotBreakWrappedTables",
  "mailMergeData",
  "pageBlockSizes",
  "pageBlockLayoutKey",
  "pageRegionLayoutReceipt",
  "tableStyleVisualCatalog",
  "documentThemeColors",
  // The "footnote text" style the browser paints notes in; the server
  // renderer keeps the document's own styles part.
  "footnoteTextStyle",
  "endnoteTextStyle",
  // Word's building blocks saved in the document (Save in: this document).
  "quickParts",
  // The pen's ruled, squared or dotted paper (Draw > Background). Word has no
  // page pattern to carry it, so a saved file is plain paper.
  "inkPaper",
  // Encrypt with Password: the browser encrypts the exported .docx with the
  // session's password (lib/office/office-encryption); the server writes the
  // package as ever and never sees a password with it.
  "encryptWithPassword",
] as const;

const SERVER_REJECTED_DOCUMENT_ATTRIBUTE_NAMES = [] as const;

export type DocumentAttributeName =
  | (typeof SERVER_NATIVE_DOCUMENT_ATTRIBUTE_NAMES)[number]
  | (typeof SERVER_METADATA_DOCUMENT_ATTRIBUTE_NAMES)[number]
  | (typeof SERVER_REJECTED_DOCUMENT_ATTRIBUTE_NAMES)[number];

export type DocxServerDocumentAttributeDisposition =
  | "native"
  | "editor_metadata"
  | "reject";

const BROWSER_NATIVE_DOCUMENT_ATTRIBUTE_NAMES = new Set<DocumentAttributeName>([
  "sectionLayouts",
  "documentMetadata",
  "headerText",
  "footerText",
  "showPageNumbers",
  "pageNumberFormat",
  "pageNumberStart",
  "pageNumberStartSection",
  "differentFirstPage",
  "firstPageHeader",
  "firstPageFooter",
  "autoHyphenation",
  "wordSettings",
  "pageSize",
  "pageMargins",
  "pageOrientation",
  "customPageSize",
  "customPageMargins",
  "columns",
  "columnSeparator",
  "writingMode",
  "manuscriptGrid",
  "docGridLinePitch",
  "defaultFontFamily",
  "defaultFontSize",
  "defaultParagraphSpaceBefore",
  "defaultParagraphSpaceAfter",
  "namedStyles",
  "referenceStyles",
  // `headingStyles` is deliberately absent: the server renderer and the
  // offline serializer both emit the Heading 1..6 definitions, but the browser
  // importer does not yet read them back, so there is no independent browser
  // round-trip evidence to claim. Import/offline stay `unverified` until it
  // does, exactly as this list's contract requires.
  "pageBackgroundColor",
  "pageBorder",
  "lineNumbers",
  "footnotes",
  "footnoteMode",
  "footnoteNumFmt",
  "endnoteNumFmt",
  "footnoteNumStart",
  "endnoteNumStart",
  "footnotePosition",
  "endnotePosition",
  "footnoteNumRestart",
  "endnoteNumRestart",
  "sources",
  "citationStyle",
  "commentThreads",
  "watermark",
  "documentProtection",
  "defaultTextColor",
  "defaultLineHeight",
  "defaultTextAlign",
  "characterSpacingControl",
  "doNotUseEastAsianBreakRules",
  "defaultKinsoku",
  "headingNumbering",
  "listStyles",
]);

function documentAttribute(
  name: DocumentAttributeName,
  docxServerExport: "native" | "metadata-only" | "unsupported",
): DocumentAtomCapability & { kind: "document-attribute" } {
  if (name === "sectionStoryGraph") {
    return Object.freeze({
      kind: "document-attribute",
      editor: "metadata-only",
      docxImport: "projected",
      docxOfflineExport: "unverified",
      docxServerExport: "unverified",
    });
  }
  if (name === "styleDefinitions") {
    return Object.freeze({
      kind: "document-attribute",
      // The browser persists references and definitions but does not yet offer
      // a complete native editor for every OOXML style property.
      editor: "metadata-only",
      docxImport: "native",
      docxOfflineExport: "unsupported",
      docxServerExport: "native",
    });
  }
  const hasIndependentRoundTripEvidence =
    BROWSER_NATIVE_DOCUMENT_ATTRIBUTE_NAMES.has(name);
  return Object.freeze({
    kind: "document-attribute",
    editor: "native",
    docxImport: hasIndependentRoundTripEvidence ? "native" : "unverified",
    docxOfflineExport: hasIndependentRoundTripEvidence
      ? "native"
      : "unverified",
    docxServerExport,
  });
}

export const DOCUMENT_ATTRIBUTE_CAPABILITIES = Object.freeze(
  Object.fromEntries([
    ...SERVER_NATIVE_DOCUMENT_ATTRIBUTE_NAMES.map((name) => [
      name,
      documentAttribute(name, "native"),
    ]),
    ...SERVER_METADATA_DOCUMENT_ATTRIBUTE_NAMES.map((name) => [
      name,
      documentAttribute(name, "metadata-only"),
    ]),
    ...SERVER_REJECTED_DOCUMENT_ATTRIBUTE_NAMES.map((name) => [
      name,
      documentAttribute(name, "unsupported"),
    ]),
  ]),
) as Readonly<
  Record<
    DocumentAttributeName,
    DocumentAtomCapability & { kind: "document-attribute" }
  >
>;

/** Canonical server declaration expected from the independent Python gate. */
export const SERVER_DOCX_DOCUMENT_ATTRIBUTE_POLICIES = Object.freeze(
  Object.fromEntries(
    Object.entries(DOCUMENT_ATTRIBUTE_CAPABILITIES).map(
      ([name, capability]) => [
        name,
        // Python's `native` policy authenticates its strict writer subset;
        // valid opaque graphs keep the public whole-domain claim unverified.
        name === "sectionStoryGraph"
          ? "native"
          : capability.docxServerExport === "native"
            ? "native"
            : capability.docxServerExport === "metadata-only"
              ? "editor_metadata"
              : "reject",
      ],
    ),
  ),
) as Readonly<
  Record<DocumentAttributeName, DocxServerDocumentAttributeDisposition>
>;

/** The visible Writer ribbon tabs. macOS menu placement is expressed by
 * `entry`, not by inventing a pseudo-tab. */
export type RibbonTabId =
  | "home"
  | "insert"
  | "draw"
  | "design"
  | "layout"
  | "references"
  | "mailings"
  | "review"
  | "view";

export type DocumentWordWebSupport =
  | "confirmed"
  | "candidate"
  | "providerBound"
  | "unsupported";

export type DocumentMaturePreserveDisposition =
  | "none"
  | "visible"
  | "hidden";

export interface DocumentCapabilityContext {
  editable: boolean;
  hasSelection: boolean;
  inTable: boolean;
  hasExactSectionModel: boolean;
  online: boolean;
  /** The page writes back into an imported Word original (patchback). */
  sourcePackageBound?: boolean;
}

/** Unified contract consumed by insert surfaces and parity/QA reporting. */
export interface DocumentFeatureCapability {
  id: string;
  tabs: readonly RibbonTabId[];
  entry: "ribbon" | "menu" | "both";
  authoring: "native" | "projection" | "external" | "unsupported";
  docxImport: "native" | "preserved" | "collapsed" | "none";
  offlineDocxExport: "native" | "fallback" | "reject";
  canonicalDocxExport: "native" | "fallback" | "reject";
  word: "editable" | "visible" | "unsupported";
  wps: "editable" | "visible" | "unsupported";
  /** Current Word-for-the-web authoring evidence. Desktop OOXML support is
   * deliberately not sufficient for this field. */
  wordWeb: DocumentWordWebSupport;
  /** Separately records an already-supported MRAI enhancement. `hidden`
   * preserves its model/writer without placing it on the Web surface. */
  maturePreserve: DocumentMaturePreserveDisposition;
  /** Frozen guide/fixture key supporting the Web disposition. */
  webEvidence: string;
  /** Stable UI command/translation key. It is deliberately separate from
   * `entry`, whose only job is to describe ribbon/menu placement. */
  commandKey: string;
  /** Closed predicate id evaluated by `isDocumentFeatureAvailable`. */
  contextPredicate?:
    | "always"
    | "editable"
    | "selection"
    | "exactSection"
    | "online";
  /** Translation key explaining a disabled/contextual/unverified state. */
  reasonKey?: string;
  /**
   * What the save into an imported Word original (patchback) does with the
   * object this command makes: `native` (default) writes it; `reject` has no
   * writer, so the command is unavailable on such a page instead of being
   * declined after the save. Mirrors the API's
   * `_docx_patchback_unwritable_nodes` (every DOCX `reject` atom + `file`)
   * and its Online Video decline.
   */
  sourcePatchback?: "native" | "reject";
}

function leaf(
  capability: DocumentFeatureCapability,
): DocumentFeatureCapability {
  return Object.freeze(capability);
}

const nativeInsert = (
  id: string,
  commandKey: string,
  overrides: Partial<DocumentFeatureCapability> = {},
): DocumentFeatureCapability =>
  leaf({
    id,
    tabs: ["insert"],
    entry: "both",
    authoring: "native",
    docxImport: "native",
    offlineDocxExport: "native",
    canonicalDocxExport: "native",
    word: "editable",
    wps: "editable",
    wordWeb: "confirmed",
    maturePreserve: "none",
    webEvidence: "word-web-guide-2026-08-25:insert",
    commandKey,
    contextPredicate: "editable",
    ...overrides,
  });

const ribbonLeaf = (
  tab: RibbonTabId,
  id: string,
  commandKey: string,
  overrides: Partial<DocumentFeatureCapability> = {},
): DocumentFeatureCapability =>
  leaf({
    id,
    tabs: [tab],
    entry: "ribbon",
    authoring: "native",
    docxImport: "native",
    offlineDocxExport: "native",
    canonicalDocxExport: "native",
    word: "editable",
    wps: "editable",
    wordWeb: ["home", "insert", "layout", "references", "review", "view"].includes(
      tab,
    )
      ? "confirmed"
      : "unsupported",
    maturePreserve: ["home", "insert", "layout", "references", "review", "view"].includes(
      tab,
    )
      ? "none"
      : "hidden",
    webEvidence: `word-web-guide-2026-08-25:${tab}`,
    commandKey,
    contextPredicate: "editable",
    ...overrides,
  });

const pendingRibbonLeaf = (
  tab: RibbonTabId,
  id: string,
  commandKey: string,
  overrides: Partial<DocumentFeatureCapability> = {},
): DocumentFeatureCapability =>
  ribbonLeaf(tab, id, commandKey, {
    authoring: "unsupported",
    docxImport: "preserved",
    offlineDocxExport: "reject",
    canonicalDocxExport: "reject",
    word: "editable",
    wps: "editable",
    contextPredicate: "always",
    reasonKey: "document.capability.notImplemented",
    wordWeb: "unsupported",
    maturePreserve: "none",
    webEvidence: `word-web-guide-2026-08-25:${id}`,
    ...overrides,
  });

const hiddenMatureWebContract = (
  id: string,
): Pick<
  DocumentFeatureCapability,
  "wordWeb" | "maturePreserve" | "webEvidence"
> => ({
  wordWeb: "unsupported",
  maturePreserve: "hidden",
  webEvidence: `mr-docs-web-guide-2026-08-25:${id}`,
});

const visibleMatureWebContract = (
  id: string,
): Pick<
  DocumentFeatureCapability,
  "wordWeb" | "maturePreserve" | "webEvidence"
> => ({
  wordWeb: "unsupported",
  maturePreserve: "visible",
  webEvidence: `mr-docs-web-guide-2026-08-25:${id}`,
});

const providerBoundWebContract = (
  id: string,
): Pick<
  DocumentFeatureCapability,
  "wordWeb" | "maturePreserve" | "webEvidence"
> => ({
  wordWeb: "providerBound",
  maturePreserve: "none",
  webEvidence: `mr-docs-web-guide-2026-08-25:${id}`,
});

export function isDocumentFeatureAvailable(
  capability: DocumentFeatureCapability,
  context: DocumentCapabilityContext,
): boolean {
  // Context can further restrict a real command, but it must never make a
  // catalog-only leaf executable merely because its predicate is `always`.
  if (capability.authoring === "unsupported") return false;
  switch (capability.contextPredicate ?? "always") {
    case "always":
      return true;
    case "editable":
      return context.editable;
    case "selection":
      return context.editable && context.hasSelection;
    case "exactSection":
      return context.editable && context.hasExactSectionModel;
    case "online":
      return context.editable && context.online;
  }
}

/**
 * Visible leaf actions in the current Writer insert/reference surfaces. Online
 * catalog products stop at their catalog entry; individual assets are dynamic.
 */
export const DOCUMENT_FEATURE_CAPABILITY_REGISTRY: readonly DocumentFeatureCapability[] =
  Object.freeze([
    nativeInsert("insert.blankPage", "toolbar.blankPage"),
    nativeInsert("insert.coverPage", "toolbar.coverPage", {
      contextPredicate: "editable",
      ...visibleMatureWebContract("insert.coverPage"),
    }),
    nativeInsert("insert.pageBreak", "toolbar.pageBreak"),
    nativeInsert("insert.columnBreak", "toolbar.columnBreak"),
    nativeInsert(
      "insert.sectionBreak.nextPage",
      "toolbar.sectionBreak.nextPage",
      visibleMatureWebContract("insert.sectionBreak.nextPage"),
    ),
    nativeInsert(
      "insert.sectionBreak.continuous",
      "toolbar.sectionBreak.continuous",
      visibleMatureWebContract("insert.sectionBreak.continuous"),
    ),
    nativeInsert(
      "insert.sectionBreak.evenPage",
      "toolbar.sectionBreak.evenPage",
      visibleMatureWebContract("insert.sectionBreak.evenPage"),
    ),
    nativeInsert(
      "insert.sectionBreak.oddPage",
      "toolbar.sectionBreak.oddPage",
      visibleMatureWebContract("insert.sectionBreak.oddPage"),
    ),
    nativeInsert("insert.pageNumber", "toolbar.pageNumber"),
    nativeInsert("insert.header", "toolbar.header"),
    nativeInsert("insert.footer", "toolbar.footer"),
    nativeInsert("insert.table", "toolbar.table"),
    nativeInsert("insert.picture.local", "toolbar.picture.local"),
    // Insert > Pictures > AI Picture: a description drawn on the server and
    // placed as an ordinary picture, so the document round-trips it as one.
    // Word's own menu offers a stock library here; MRAI draws instead.
    nativeInsert("insert.picture.aiGenerate", "toolbar.aiImage", {
      word: "editable",
      wps: "editable",
      // Bound to the draw provider, authored natively: shown and live.
      ...providerBoundWebContract("insert.picture.aiGenerate"),
    }),
    nativeInsert("insert.picture.onlineProvider", "toolbar.onlinePictures", {
      authoring: "unsupported",
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "unsupported",
      wps: "unsupported",
      contextPredicate: "always",
      reasonKey: "document.capability.imageProviderUnavailable",
      ...providerBoundWebContract("insert.picture.onlineProvider"),
    }),
    nativeInsert("insert.picture.device", "toolbar.picture.device", {
      authoring: "external",
      reasonKey: "document.capability.deviceDependent",
    }),
    // Contextual Picture Format commands are catalogued against Home so this
    // base-tab registry remains the closed nine-tab Writer contract. Their
    // actual entry point is the selection-only Picture contextual tab.
    ribbonLeaf(
      "home",
      "picture.change.fromDevice",
      "toolbar.pictureFromDevice",
      { contextPredicate: "selection" },
    ),
    ribbonLeaf("home", "picture.reset", "presentation.image.reset", {
      contextPredicate: "selection",
    }),
    ribbonLeaf(
      "home",
      "picture.resetAndSize",
      "presentation.image.resetAndSize",
      { contextPredicate: "selection" },
    ),
    ribbonLeaf(
      "home",
      "picture.formatPane",
      "presentation.image.formatPicture",
      {
        contextPredicate: "selection",
        webEvidence:
          "picture-set-format-nonzero-crop-wordweb-4c35d9f:nonzero-crop",
      },
    ),
    ...(
      [
        ["picture.change.onlineProvider", "toolbar.onlinePictures"],
      ] as const
    ).map(([id, commandKey]) =>
      pendingRibbonLeaf("home", id, commandKey, {
        reasonKey: "document.capability.imageProviderUnavailable",
        ...providerBoundWebContract(id),
      }),
    ),
    ...(
      [
        ["picture.removeBackground", "pdfViewer.backgroundRemove"],
      ] as const
    ).map(([id, commandKey]) =>
      pendingRibbonLeaf("home", id, commandKey, {
        reasonKey: "document.capability.drawingMlPending",
      }),
    ),
    ribbonLeaf("home", "picture.transparency", "presentation.imageOpacity", {
      contextPredicate: "selection",
    }),
    ribbonLeaf("home", "picture.color", "toolbar.ribbonColor", {
      contextPredicate: "selection",
    }),
    ribbonLeaf("home", "picture.corrections", "toolbar.designEffects", {
      contextPredicate: "selection",
    }),
    nativeInsert("insert.screenshot", "toolbar.screenshot", {
      authoring: "external",
      reasonKey: "document.capability.deviceDependent",
      ...visibleMatureWebContract("insert.screenshot"),
    }),
    nativeInsert("insert.media.onlineVideo", "toolbar.onlineVideo", {
      // The ribbon writes a typed documentMedia atom. Playback stays
      // provider-owned and requires connectivity; browser DOCX cannot emit the
      // relationship, while canonical server export preserves it natively.
      entry: "ribbon",
      offlineDocxExport: "reject",
      contextPredicate: "online",
      reasonKey: "document.capability.onlineRequired",
      sourcePatchback: "reject",
    }),
    nativeInsert("insert.drawingCanvas", "toolbar.drawingCanvas", {
      authoring: "native",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      // A writer: disabled when the document is read-only. This was
      // `always`, carried over from the placeholder era when the
      // capability was `unsupported` and the predicate never mattered;
      // the Draw tab masked it by disabling its own writers until
      // `45f97ff311` moved them to Insert.
      contextPredicate: "editable",
    }),
    nativeInsert("insert.shape", "toolbar.shape", {
      authoring: "native",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      // A writer: disabled when the document is read-only. This was
      // `always`, carried over from the placeholder era when the
      // capability was `unsupported` and the predicate never mattered;
      // the Draw tab masked it by disabling its own writers until
      // `45f97ff311` moved them to Insert.
      contextPredicate: "editable",
    }),
    nativeInsert("insert.icon", "toolbar.icon", {
      entry: "ribbon",
      authoring: "projection",
      docxImport: "collapsed",
      offlineDocxExport: "fallback",
      canonicalDocxExport: "fallback",
      word: "visible",
      wps: "visible",
      reasonKey: "document.capability.iconProjection",
      ...visibleMatureWebContract("insert.icon"),
    }),
    nativeInsert("insert.textBox.horizontal", "toolbar.textBox.horizontal"),
    nativeInsert("insert.textBox.vertical", "toolbar.textBox.vertical"),
    // WordArt inserts the shape/text-box node with the Shape Format tab's
    // WordArt text paint (lib/document-shapes DOCUMENT_WORD_ART_STYLES), which
    // the DOCX writer emits as w14:textFill / w14:textOutline / text effects
    // and the importer reads back — the same native contract as insert.shape.
    nativeInsert("insert.wordArt", "toolbar.wordArt", {
      authoring: "native",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      ...visibleMatureWebContract("insert.wordArt"),
    }),
    nativeInsert("insert.chart", "toolbar.chart", {
      ...visibleMatureWebContract("insert.chart"),
    }),
    nativeInsert("insert.mermaid", "toolbar.insertMermaid", {
      tabs: ["insert", "draw"],
      entry: "ribbon",
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "unsupported",
      wps: "unsupported",
      sourcePatchback: "reject",
      
      // still draws and edits in place.
      ...hiddenMatureWebContract("insert.mermaid"),
      authoring: "unsupported",
      reasonKey: "document.capability.retired",
    }),
    nativeInsert("insert.codeChart", "toolbar.insertCodeChart", {
      tabs: ["insert", "draw"],
      entry: "ribbon",
      authoring: "projection",
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "unsupported",
      wps: "unsupported",
      reasonKey: "document.capability.losslessExportUnavailable",
      sourcePatchback: "reject",
      ...visibleMatureWebContract("insert.codeChart"),
    }),
    nativeInsert("insert.smartArt", "toolbar.smartArt", {
      // Insert lays the diagram out with the shared Office SmartArt engine and
      // materialises it as ONE native shape group, so it writes and re-imports
      // through the `wpg:wgp` writer like Flowchart and Mind Map do. `preserved`
      // still describes IMPORT, which is a different object: a native Word
      // `dgm` part keeps its own layout parts and stays text-only editable.
      docxImport: "preserved",
      ...visibleMatureWebContract("insert.smartArt"),
    }),
    nativeInsert("insert.flowchart", "toolbar.flowchart", {
      entry: "ribbon",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      contextPredicate: "editable",
      
      ...hiddenMatureWebContract("insert.flowchart"),
      authoring: "unsupported",
      reasonKey: "document.capability.retired",
    }),
    nativeInsert("insert.mindMap", "toolbar.mindMap", {
      entry: "ribbon",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      contextPredicate: "editable",
      
      ...hiddenMatureWebContract("insert.mindMap"),
      authoring: "unsupported",
      reasonKey: "document.capability.retired",
    }),
    nativeInsert("insert.symbol", "toolbar.symbol"),
    nativeInsert("insert.equation", "toolbar.equation"),
    nativeInsert("insert.latex", "toolbar.latex", {
      authoring: "projection",
      docxImport: "collapsed",
      offlineDocxExport: "fallback",
      canonicalDocxExport: "fallback",
      word: "editable",
      wps: "editable",
      reasonKey: "document.capability.ommlProjection",
    }),
    nativeInsert("insert.comment", "toolbar.comment", {
      tabs: ["insert", "review"],
    }),
    nativeInsert("insert.hyperlink", "toolbar.hyperlink"),
    nativeInsert("insert.bookmark", "toolbar.bookmark", {
      contextPredicate: "selection",
      reasonKey: "document.capability.selectionRequired",
    }),
    nativeInsert("insert.footnote", "toolbar.footnote", {
      tabs: ["insert", "references"],
    }),
    nativeInsert("insert.endnote", "toolbar.endnote", {
      tabs: ["insert", "references"],
    }),
    nativeInsert("insert.caption.figure", "toolbar.caption.figure", {
      tabs: ["insert", "references"],
      ...hiddenMatureWebContract("insert.caption.figure"),
    }),
    nativeInsert("insert.caption.table", "toolbar.caption.table", {
      tabs: ["insert", "references"],
      ...hiddenMatureWebContract("insert.caption.table"),
    }),
    nativeInsert("insert.crossReference", "toolbar.crossReference", {
      tabs: ["insert", "references"],
      ...visibleMatureWebContract("insert.crossReference"),
    }),
    nativeInsert("insert.tableOfContents", "toolbar.tableOfContents", {
      tabs: ["insert", "references"],
    }),
    nativeInsert("insert.listOfFigures", "toolbar.listOfFigures", {
      tabs: ["insert", "references"],
      ...hiddenMatureWebContract("insert.listOfFigures"),
    }),
    nativeInsert("insert.listOfTables", "toolbar.listOfTables", {
      tabs: ["insert", "references"],
      ...hiddenMatureWebContract("insert.listOfTables"),
    }),
    nativeInsert("insert.index", "toolbar.index", {
      tabs: ["insert", "references"],
      ...hiddenMatureWebContract("insert.index"),
    }),
    nativeInsert("insert.bibliography", "toolbar.bibliography", {
      tabs: ["insert", "references"],
    }),
    nativeInsert("insert.dropCap", "toolbar.dropCap", {
      ...visibleMatureWebContract("insert.dropCap"),
    }),
    nativeInsert("insert.dateTime", "toolbar.dateTime"),
    nativeInsert("insert.signatureLine", "toolbar.signatureLine", {
      authoring: "native",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      ...visibleMatureWebContract("insert.signatureLine"),
    }),
    nativeInsert("insert.file", "toolbar.file", {
      // MRAI-owned file attachment. The canonical Word export writes it as
      // the embedded object Word makes of a file, shown as an icon (an
      // Office file as its own package, any other in an OLE Package); one
      // past the page's object budget is written as its name and reported.
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      reasonKey: "document.capability.attachmentExportUnavailable",
      sourcePatchback: "reject",
      ...visibleMatureWebContract("insert.file"),
    }),
    nativeInsert("insert.object", "toolbar.object", {
      authoring: "external",
      docxImport: "preserved",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "editable",
      wps: "editable",
      contextPredicate: "always",
      reasonKey: "document.capability.oleUnsupported",
      ...hiddenMatureWebContract("insert.object"),
    }),
    nativeInsert("insert.documentParts", "toolbar.documentParts", {
      authoring: "projection",
      docxImport: "collapsed",
      offlineDocxExport: "fallback",
      canonicalDocxExport: "fallback",
      word: "visible",
      wps: "visible",
      reasonKey: "document.capability.quickPartsPartial",
      ...hiddenMatureWebContract("insert.documentParts"),
    }),
    nativeInsert("insert.watermark", "toolbar.watermark", {
      ...hiddenMatureWebContract("insert.watermark"),
    }),
    nativeInsert("layout.sectionLayout", "toolbar.sectionLayout", {
      tabs: ["layout"],
      entry: "ribbon",
      contextPredicate: "exactSection",
      reasonKey: "document.capability.exactSectionModelRequired",
    }),
    ...(
      [
        ["DATE", "native"],
        ["TIME", "native"],
        ["PAGE", "native"],
        ["NUMPAGES", "native"],
        ["AUTHOR", "native"],
        ["FILENAME", "native"],
        ["TITLE", "native"],
        ["REF", "native"],
        ["SEQ", "native"],
        ["TOC", "native"],
        ["CITATION", "native"],
        ["INDEX", "native"],
        ["TOA", "native"],
        ["MERGEFIELD", "native"],
        ["IF", "native"],
      ] as const
    ).map(([fieldType, importState]) =>
      nativeInsert(`insert.field.${fieldType}`, `toolbar.field_${fieldType}`, {
        docxImport: importState,
        ...(fieldType === "REF"
          ? hiddenMatureWebContract("insert.field.REF")
          : {}),
      }),
    ),
    ...(
      [
        ["home.clipboard.paste", "toolbar.paste"],
        ["home.clipboard.cut", "toolbar.cut"],
        ["home.clipboard.copy", "toolbar.copy"],
        ["home.clipboard.formatPainter", "toolbar.formatPainter"],
        ["home.font.family", "toolbar.fontFamily"],
        ["home.font.size", "toolbar.fontSize"],
        ["home.font.bold", "toolbar.bold"],
        ["home.font.italic", "toolbar.italic"],
        ["home.font.underline", "toolbar.underline"],
        ["home.font.underlineStyle", "toolbar.underlineStyleMenu"],
        ["home.font.underlineColor", "toolbar.underlineColorLabel"],
        ["home.font.strike", "toolbar.strikethrough"],
        ["home.font.superscript", "toolbar.superscript"],
        ["home.font.subscript", "toolbar.subscript"],
        ["home.font.textColor", "toolbar.textColor"],
        ["home.font.highlight", "toolbar.highlight"],
        ["home.font.characterSpacing", "toolbar.characterSpacing"],
        ["home.font.scale", "toolbar.characterScale"],
        ["home.font.emphasisDots", "toolbar.emphasisDots"],
        ["home.font.enclosedCharacter", "toolbar.enclosedCharacter"],
        ["home.font.eastAsianLayout", "toolbar.eastAsianLayout"],
        ["home.font.clearFormatting", "toolbar.clearFormatting"],
        ["home.paragraph.bullets", "toolbar.bulletList"],
        ["home.paragraph.numbering", "toolbar.orderedList"],
        ["home.paragraph.multilevelList", "toolbar.multilevelList"],
        ["home.paragraph.changeListLevel", "toolbar.changeListLevel"],
        ["home.paragraph.alignLeft", "toolbar.alignLeft"],
        ["home.paragraph.alignCenter", "toolbar.alignCenter"],
        ["home.paragraph.alignRight", "toolbar.alignRight"],
        ["home.paragraph.justify", "toolbar.justify"],
        ["home.paragraph.indentDecrease", "toolbar.indentDecrease"],
        ["home.paragraph.indentIncrease", "toolbar.indentIncrease"],
        ["home.paragraph.lineSpacing", "toolbar.lineSpacing"],
        [
          "home.paragraph.contextualSpacing",
          "toolbar.paragraphDontAddSameStyle",
        ],
        ["home.paragraph.borders", "toolbar.paragraphBorders"],
        ["home.paragraph.shading", "toolbar.paragraphShading"],
        ["home.paragraph.formattingMarks", "toolbar.formattingMarks"],
        ["home.styles.gallery", "toolbar.stylesMenu"],
        ["home.styles.manage", "toolbar.stylesPane"],
        ["home.edit.findReplace", "toolbar.findReplace"],
        ["home.edit.goTo", "toolbar.goTo"],
        ["home.edit.selectAll", "toolbar.selectAll"],
      ] as const
    ).map(([id, commandKey]) => ribbonLeaf("home", id, commandKey)),
    // The server's Word writer states it as Word's own w:dstrike.
    ribbonLeaf("home", "home.font.doubleStrike", "toolbar.doubleStrikethrough"),
    ribbonLeaf(
      "home",
      "home.font.hiddenText",
      "toolbar.formattingMarksHiddenText",
    ),
    ribbonLeaf(
      "home",
      "home.font.kerningThreshold",
      "toolbar.kerningThreshold",
    ),
    ...(
      [
        ["home.font.openType", "toolbar.openTypeFeatures"],
        ["home.font.textEffects", "toolbar.textEffects"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("home", id, commandKey, {
        // These marks are native to the browser editor and saved HTML, but the
        // DOCX importers/exporters do not model their Office run properties yet.
        // Export therefore keeps readable text while dropping the effect.
        docxImport: "none",
        offlineDocxExport: "fallback",
        canonicalDocxExport: "fallback",
        word: "visible",
        wps: "visible",
        reasonKey: "document.capability.losslessExportUnavailable",
      }),
    ),
    ...(
      [
        ["home.paragraph.defineNewBullet", "toolbar.defineNewBullet"],
        [
          "home.paragraph.defineNewNumberFormat",
          "toolbar.defineNewNumberFormat",
        ],
        [
          "home.paragraph.defineNewMultilevelList",
          "toolbar.defineNewMultilevelList",
        ],
        ["home.paragraph.defineNewListStyle", "toolbar.defineNewListStyle"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("home", id, commandKey, {
        // Authoring applies a real list immediately. DOCX keeps the editable
        // list and visible markers, but the browser-only definition identity
        // (including a locally named list style) is collapsed on round-trip.
        docxImport: "collapsed",
        offlineDocxExport: "fallback",
        canonicalDocxExport: "fallback",
        reasonKey: "document.capability.losslessExportUnavailable",
      }),
    ),
    ...(
      [
        ["tabs", "toolbar.formattingMarksTabs"],
        ["spaces", "toolbar.formattingMarksSpaces"],
        ["paragraphs", "toolbar.formattingMarksParagraphs"],
        ["optionalHyphens", "toolbar.formattingMarksOptionalHyphens"],
        ["objectAnchors", "toolbar.formattingMarksObjectAnchors"],
        ["optionalBreaks", "toolbar.formattingMarksOptionalBreaks"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("home", `home.paragraph.formattingMarks.${id}`, commandKey),
    ),
    ribbonLeaf(
      "home",
      "home.paragraph.formattingMarks.hiddenText",
      "toolbar.formattingMarksHiddenText",
    ),
    ribbonLeaf("home", "home.edit.selectObjects", "toolbar.selectObjects"),
    ribbonLeaf("home", "home.voice.dictation", "toolbar.dictation", {
      // The recognition provider is browser-owned and leaves no provenance in
      // the document. Its final transcript is ordinary editable text, so the
      // existing text/DOCX save path remains the sole persistence authority.
      authoring: "external",
      docxImport: "none",
      contextPredicate: "online",
      reasonKey: "toolbar.dictationUnavailable",
    }),
    ribbonLeaf("home", "home.voice.transcribe", "toolbar.transcribe", {
      // Word Web stores recordings and transcript state with an authenticated
      // cloud provider. MRAI has no workspace-scoped custody/provider contract
      // for that data, so this catalog leaf stays visible but deliberately
      // unbound instead of reusing ephemeral browser dictation.
      authoring: "external",
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      wps: "unsupported",
      contextPredicate: "online",
      reasonKey: "document.capability.transcriptionProviderUnavailable",
    }),
    pendingRibbonLeaf(
      "home",
      "home.edit.selectSimilarFormatting",
      "toolbar.selectSimilarFormatting",
      { reasonKey: "document.capability.multiRangeSelectionUnavailable" },
    ),
    ribbonLeaf(
      "home",
      "home.edit.selectionPane",
      "toolbar.selectionPane",
      {
        // MRNote extension: a session-only projection of the document's real
        // selectable object atoms. It selects by live ProseMirror identity and
        // never serializes pane state. Visibility, stacking, rename and
        // multi-select remain intentionally absent until their models exist.
        authoring: "projection",
        docxImport: "none",
        contextPredicate: "always",
      },
    ),
    ribbonLeaf(
      "home",
      "home.clipboard.pasteSpecial.formattedText",
      "toolbar.pasteSpecialFormattedText",
    ),
    ribbonLeaf(
      "home",
      "home.clipboard.pasteSpecial.unformattedText",
      "toolbar.pasteSpecialUnformattedText",
    ),
    ribbonLeaf(
      "home",
      "home.clipboard.pasteSpecial.picture",
      "toolbar.pasteSpecialPicture",
    ),
    pendingRibbonLeaf(
      "home",
      "home.clipboard.pasteSpecial.pasteLink",
      "toolbar.pasteSpecialLink",
      { reasonKey: "document.capability.pasteLinkUnavailable" },
    ),
    pendingRibbonLeaf(
      "home",
      "home.clipboard.officeClipboard",
      "toolbar.officeClipboard",
      { reasonKey: "document.capability.clipboardHistoryUnavailable" },
    ),
    ribbonLeaf("draw", "draw.canvas", "toolbar.insertWhiteboard", {
      authoring: "projection",
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "unsupported",
      wps: "unsupported",
      reasonKey: "document.capability.losslessExportUnavailable",
    }),
    ribbonLeaf("draw", "draw.shapes", "toolbar.insertShapes"),
    ribbonLeaf("draw", "draw.mermaid", "toolbar.insertMermaid", {
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "unsupported",
      wps: "unsupported",
      
      maturePreserve: "hidden",
      authoring: "unsupported",
      reasonKey: "document.capability.retired",
    }),
    ribbonLeaf("draw", "draw.chart", "toolbar.insertChart"),
    // Word's Draw tab ("Draw and write with ink in Office"): Draw, an Eraser
    // (Stroke, Small and Medium), pens the user adds and deletes (Pen,
    // Highlighter, Pencil), each pen's thickness, colour, effect and More
    // Colors, and Draw with Trackpad (Word for Mac). A stroke is a
    // `documentShape` with its own outline, so both DOCX writers already
    // export it as the freeform line Word draws (an effect as its solid
    // stand-in); Word's InkML content part is not read, so an inked Word
    // file's strokes are not editable ink here. Word for the web offers pens,
    // a highlighter, thickness, colour and the stroke eraser.
    ...(
      [
        ["draw.pen", "toolbar.drawPen"],
        ["draw.highlighter", "toolbar.drawHighlighter"],
        ["draw.eraser", "toolbar.drawEraser"],
        ["draw.inkColor", "toolbar.drawInkColor"],
        ["draw.inkWidth", "toolbar.drawInkWidth"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("draw", id, commandKey, {
        docxImport: "none",
        wordWeb: "confirmed",
        webEvidence: `ms-support-draw-and-write-with-ink-in-office-6d76c674:web:${id}`,
      }),
    ),
    // Desktop Word's pencil, effects, eraser kinds, Add Pen / Delete and the
    // trackpad pad: visible here, not named in the article's Web section.
    ...(
      [
        ["draw.pencil", "toolbar.drawPencil"],
        ["draw.inkEffect", "toolbar.drawInkEffect"],
        ["draw.eraserKind", "toolbar.drawEraserKinds"],
        ["draw.addPen", "toolbar.drawAddPen"],
        ["draw.trackpad", "toolbar.drawTrackpad"],
        ["draw.touch", "toolbar.drawTouch"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("draw", id, commandKey, {
        docxImport: "none",
        ...visibleMatureWebContract(id),
      }),
    ),
    ribbonLeaf("draw", "draw.connector", "toolbar.drawConnector", {
      authoring: "native",
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      word: "editable",
      wps: "editable",
      contextPredicate: "always",
    }),
    ribbonLeaf("draw", "draw.flowShape", "toolbar.insertFlowchart", {
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      
      maturePreserve: "hidden",
      authoring: "unsupported",
      reasonKey: "document.capability.retired",
    }),
    ribbonLeaf("draw", "draw.mindMap", "toolbar.insertMindmap", {
      docxImport: "native",
      offlineDocxExport: "native",
      canonicalDocxExport: "native",
      
      maturePreserve: "hidden",
      authoring: "unsupported",
      reasonKey: "document.capability.retired",
    }),
    ribbonLeaf("design", "design.watermark", "toolbar.watermarkLabel"),
    ribbonLeaf("design", "design.pageColor", "toolbar.pageColor"),
    ribbonLeaf("design", "design.pageBorder", "toolbar.pageBorder"),
    // Word desktop's Design > Document Formatting commands (Word for Mac
    // 16.86: Themes, the style-set gallery, Colors, Fonts, Paragraph
    // Spacing), on the Design tab the Docs ribbon carries in every profile.
    // The document's design (lib/document-design.ts) restyles the Quick
    // Styles and Normal; the server writes the theme part it names.
    ...(
      [
        ["design.theme", "toolbar.designTheme"],
        ["design.styleSet", "toolbar.designStyleSet"],
        ["design.colorSet", "toolbar.designColors"],
        ["design.fontSet", "toolbar.designFonts"],
        ["design.paragraphSpacing", "toolbar.designParagraphSpacing"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("design", id, commandKey, {
        docxImport: "preserved",
        offlineDocxExport: "fallback",
        ...visibleMatureWebContract(id),
      }),
    ),
    ...(
      [
        ["design.effects", "toolbar.designEffects"],
        ["design.pictureWatermark", "toolbar.pictureWatermark"],
      ] as const
    ).map(([id, commandKey]) => pendingRibbonLeaf("design", id, commandKey)),
    ...(
      [
        ["layout.margins", "toolbar.pageMargins"],
        ["layout.orientation", "toolbar.pageOrientation"],
        ["layout.pageSize", "toolbar.pageSize"],
        ["layout.columns", "toolbar.columns"],
        ["layout.breaks", "toolbar.breaks"],
        ["layout.writingMode", "toolbar.verticalText"],
        ["layout.pageBorder", "toolbar.pageBorder"],
        ["layout.paragraphIndent", "toolbar.paragraphDialog"],
        ["layout.paragraphSpacing", "toolbar.paragraphDialog"],
      ] as const
    ).map(([id, commandKey]) => ribbonLeaf("layout", id, commandKey)),
    ribbonLeaf("layout", "layout.pageColor", "toolbar.pageColor", {
      // Microsoft's "Differences between using a document in the browser and
      // in Word" lists Page color as supported in Word for the web (watermark
      // and theme stay desktop-only). The Web ribbon has no Design tab, so the
      // live control sits on Layout > Page Background beside Page Border; the
      // `design.pageColor` leaf keeps describing the desktop Design-tab
      // placement, exactly like `layout.pageBorder` / `design.pageBorder`.
      webEvidence:
        "ms-support-differences-browser-vs-word-3e863ce3:page-color",
    }),
    ribbonLeaf("layout", "layout.columns.unequal", "toolbar.columnsUnequal", {
      contextPredicate: "exactSection",
      reasonKey: "document.capability.exactSectionModelRequired",
    }),
    ribbonLeaf(
      "layout",
      "layout.section.verticalAlignment",
      "toolbar.sectionVerticalAlignment",
    ),
    ribbonLeaf("layout", "layout.hyphenation.auto", "toolbar.hyphenationAuto", {
      ...hiddenMatureWebContract("layout.hyphenation.auto"),
    }),
    ribbonLeaf("layout", "layout.object.wrap", "toolbar.objectWrap", {
      contextPredicate: "selection",
      // Tight/Through polygons and absolute offsets intentionally collapse to
      // the supported square-side subset on import.
      docxImport: "collapsed",
    }),
    ribbonLeaf("layout", "layout.object.layer", "toolbar.objectLayer", {
      // Native WPS shapes own strict `wp:anchor/@relativeHeight` through both
      // import/export paths and original-package patchback. Images retain their
      // independent behind/in-front wrap authority in the same menu.
      authoring: "native",
      contextPredicate: "selection",
    }),
    ribbonLeaf("layout", "layout.object.align", "toolbar.objectAlign", {
      // This is the single inline image's paragraph-level left/center/right
      // projection, not Word's multi-object or page-relative Align geometry.
      authoring: "projection",
      contextPredicate: "selection",
      webEvidence: "office-web-docx-readonly-current-96ff268:picture-context",
    }),
    ribbonLeaf(
      "layout",
      "layout.object.position",
      "toolbar.objectPosition",
      {
        // MRAI desktop extension. Word for the web preserves anchored offsets
        // authored elsewhere but does not expose an equivalent move surface.
        // Our native documentShape subset owns x/y through both DOCX writers.
        contextPredicate: "selection",
        word: "unsupported",
      },
    ),
    ribbonLeaf("layout", "layout.object.size", "toolbar.pageSetupWidth", {
      // Word Web exposes picture size in inches; the existing image/shape
      // writers continue to own the exact internal px/EMU geometry.
      contextPredicate: "selection",
      webEvidence: "office-web-docx-readonly-current-96ff268:picture-context",
    }),
    ribbonLeaf("layout", "layout.object.group", "toolbar.objectGroup", {
      contextPredicate: "selection",
      // Newly-authored native WPS shapes group into a real WordprocessingGroup
      // and round-trip through scratch DOCX. Authenticated imported groups are
      // preserve-only until original-package WPG patchback owns the mutation.
      authoring: "native",
    }),
    ribbonLeaf(
      "layout",
      "layout.object.rotate",
      "toolbar.objectRotate",
      {
        // Images and the MRAI native documentShape subset share one visible
        // Rotate command while retaining their separate persisted attributes.
        // Multi-object/group transforms remain outside this single-selection leaf.
        contextPredicate: "selection",
        webEvidence: "office-web-docx-readonly-current-96ff268:picture-context",
      },
    ),
    ribbonLeaf(
      "references",
      "references.insertCitation",
      "toolbar.insertCitation",
    ),
    ribbonLeaf(
      "references",
      "references.sourceManager",
      "toolbar.sourceManager",
    ),
    ribbonLeaf(
      "references",
      "references.citationStyle",
      "toolbar.citationStyle",
    ),
    ribbonLeaf(
      "references",
      "references.bibliographyInsert",
      "toolbar.bibliography",
    ),
    ribbonLeaf(
      "references",
      "references.bibliographyRemove",
      "toolbar.removeBibliography",
    ),
    ribbonLeaf(
      "references",
      "references.insertFootnote",
      "toolbar.insertFootnote",
    ),
    ribbonLeaf(
      "references",
      "references.insertEndnote",
      "toolbar.insertEndnote",
    ),
    ribbonLeaf(
      "references",
      "references.tocAddText",
      "toolbar.tocAddText",
    ),
    ribbonLeaf("references", "references.tocSettings", "toolbar.tocSettings"),
    ribbonLeaf(
      "references",
      "references.tocUpdate",
      "toolbar.updateTableOfContents",
    ),
    ribbonLeaf(
      "references",
      "references.tocRemove",
      "toolbar.removeTableOfContents",
    ),
    // Captions, Table of Figures, Index and Table of Authorities are desktop
    // Word commands ("Insert a table of figures", "Create and update an
    // index", "Create a table of authorities", Microsoft Support); Word for
    // the web has none of them. The More References group keeps them in the
    // Word-web ribbon as desktop exceptions (the Watermark policy), so they
    // are visible there rather than hidden: a hidden one renders its button
    // but refuses every press with "not in Word for the web".
    ribbonLeaf(
      "references",
      "references.captionSettings",
      "toolbar.captionSettings",
      visibleMatureWebContract("references.captionSettings"),
    ),
    ribbonLeaf(
      "references",
      "references.figureList",
      "toolbar.insertFigureList",
      visibleMatureWebContract("references.figureList"),
    ),
    ribbonLeaf(
      "references",
      "references.tableList",
      "toolbar.insertTableList",
      visibleMatureWebContract("references.tableList"),
    ),
    ribbonLeaf(
      "references",
      "references.generatedListUpdate",
      "toolbar.updateGeneratedList",
      visibleMatureWebContract("references.generatedListUpdate"),
    ),
    ribbonLeaf(
      "references",
      "references.generatedListRemove",
      "toolbar.removeGeneratedList",
      visibleMatureWebContract("references.generatedListRemove"),
    ),
    ribbonLeaf(
      "references",
      "references.crossReferenceSettings",
      "toolbar.crossReferenceSettings",
      visibleMatureWebContract("references.crossReferenceSettings"),
    ),
    ribbonLeaf(
      "references",
      "references.footnoteSettings",
      "toolbar.footnoteSettings",
    ),
    ribbonLeaf(
      "references",
      "references.noteNavigation",
      "spreadsheet.dataFormNext",
      {
        docxImport: "none",
        contextPredicate: "always",
      },
    ),
    ribbonLeaf(
      "references",
      "references.showNotes",
      "presentation.showNotesLabel",
      {
        docxImport: "none",
        contextPredicate: "always",
      },
    ),
    ribbonLeaf(
      "references",
      "references.indexMark",
      "toolbar.markIndexEntry",
      visibleMatureWebContract("references.indexMark"),
    ),
    ribbonLeaf(
      "references",
      "references.indexSettings",
      "toolbar.indexSettings",
      visibleMatureWebContract("references.indexSettings"),
    ),
    ribbonLeaf(
      "references",
      "references.indexUpdate",
      "toolbar.updateIndex",
      visibleMatureWebContract("references.indexUpdate"),
    ),
    ribbonLeaf(
      "references",
      "references.indexRemove",
      "toolbar.removeIndex",
      visibleMatureWebContract("references.indexRemove"),
    ),
    ribbonLeaf(
      "references",
      "references.authorityMark",
      "toolbar.markAuthorityCitation",
      visibleMatureWebContract("references.authorityMark"),
    ),
    ribbonLeaf(
      "references",
      "references.authoritySettings",
      "toolbar.insertAuthorityTable",
      visibleMatureWebContract("references.authoritySettings"),
    ),
    ribbonLeaf(
      "references",
      "references.authorityUpdate",
      "toolbar.updateAuthorityTable",
      visibleMatureWebContract("references.authorityUpdate"),
    ),
    ribbonLeaf(
      "references",
      "references.authorityRemove",
      "toolbar.removeAuthorityTable",
      visibleMatureWebContract("references.authorityRemove"),
    ),
    ribbonLeaf(
      "mailings",
      "mailings.recipients",
      "toolbar.mailMergeRecipients",
      {
        authoring: "projection",
        docxImport: "collapsed",
        offlineDocxExport: "fallback",
        canonicalDocxExport: "fallback",
        word: "editable",
        wps: "editable",
        reasonKey: "document.capability.mailMergeWorkflowMetadata",
      },
    ),
    ribbonLeaf("mailings", "mailings.mergeField", "toolbar.insertMergeField"),
    ...(
      [
        ["mailings.envelopes", "toolbar.mailingsEnvelopes"],
        ["mailings.labels", "toolbar.mailingsLabels"],
        ["mailings.mergeType", "toolbar.mailingsMergeType"],
        ["mailings.filterSort", "toolbar.mailingsFilterSort"],
        ["mailings.addressBlock", "toolbar.mailingsAddressBlock"],
        ["mailings.greetingLine", "toolbar.mailingsGreetingLine"],
        ["mailings.matchFields", "toolbar.mailingsMatchFields"],
        ["mailings.rules", "toolbar.mailingsRules"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("mailings", id, commandKey, {
        authoring: "projection",
        docxImport: "collapsed",
        offlineDocxExport: "fallback",
        canonicalDocxExport: "fallback",
        word: "editable",
        wps: "editable",
        reasonKey: "document.capability.mailMergeWorkflowMetadata",
      }),
    ),
    ribbonLeaf("mailings", "mailings.preview", "toolbar.mailingsPreview", {
      authoring: "projection",
      docxImport: "collapsed",
      offlineDocxExport: "fallback",
      canonicalDocxExport: "fallback",
      word: "editable",
      wps: "editable",
      reasonKey: "document.capability.mailMergeWorkflowMetadata",
    }),
    ribbonLeaf(
      "mailings",
      "mailings.finishDocx",
      "toolbar.mailingsFinishDocx",
      {
        authoring: "projection",
        docxImport: "collapsed",
        offlineDocxExport: "fallback",
        canonicalDocxExport: "fallback",
        word: "editable",
        wps: "editable",
        reasonKey: "document.capability.mailMergeWorkflowMetadata",
      },
    ),
    ribbonLeaf("mailings", "mailings.sendEmail", "toolbar.mailingsSendEmail", {
      authoring: "external",
      docxImport: "none",
      offlineDocxExport: "reject",
      canonicalDocxExport: "reject",
      word: "unsupported",
      wps: "unsupported",
      contextPredicate: "online",
      reasonKey: "document.capability.connectorOptional",
    }),
    ...(
      [
        ["review.trackChanges", "toolbar.suggesting"],
        ["review.acceptAll", "toolbar.acceptAllChanges"],
        ["review.comment", "toolbar.addComment"],
      ] as const
    ).map(([id, commandKey]) => ribbonLeaf("review", id, commandKey)),
    ...(
      [
        ["review.commentsPane", "comments.title"],
        ["review.changesPane", "trackChanges.title"],
        ["review.previousChange", "toolbar.previousChange"],
        ["review.nextChange", "toolbar.nextChange"],
        ["review.markupFilter", "toolbar.markupFilter"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("review", id, commandKey, {
        // Pane visibility, change navigation, and markup presentation are
        // session-only review actions. They remain available in read-only
        // documents and never claim a persisted DOCX mutation.
        docxImport: "none",
        contextPredicate: "always",
      }),
    ),
    ...(
      [
        ["review.acceptOne", "toolbar.acceptChange"],
        ["review.rejectOne", "toolbar.rejectChange"],
      ] as const
    ).map(([id, commandKey]) => ribbonLeaf("review", id, commandKey)),
    // Word for the web has no Compare; Word desktop has Review > Compare >
    // Compare Documents ("Compare and merge two versions of a document",
    // Microsoft Support). MRAI built the panel, so it stays reachable in the
    // web ribbon as a desktop exception, the policy of Watermark and the More
    // References group (DOCS-DIALOG-AUDIT 13/14).
    ribbonLeaf("review", "review.compare", "toolbar.compareDocuments", {
      ...visibleMatureWebContract("review.compare"),
    }),
    ribbonLeaf("review", "review.protectDocument", "toolbar.restrictEditing", {
      // The adapter keeps this available only when the editor can lock OR the
      // current native lock can be removed. The catalog cannot use `editable`
      // here because protection itself makes the editor read-only.
      contextPredicate: "always",
      ...hiddenMatureWebContract("review.protectDocument"),
    }),
    ribbonLeaf("review", "review.accessibility", "toolbar.accessibilityCheck", {
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("review", "review.wordCount", "toolbar.readabilityTitle", {
      authoring: "projection",
      docxImport: "none",
      word: "visible",
      wps: "visible",
      contextPredicate: "always",
    }),
    ribbonLeaf("review", "review.readAloud", "toolbar.readAloud", {
      // Browser speech synthesis is a local review projection. It deliberately
      // has no DOCX import/model claim and is feature-detected by the adapter.
      authoring: "projection",
      docxImport: "none",
      word: "visible",
      wps: "visible",
      contextPredicate: "always",
      ...visibleMatureWebContract("review.readAloud"),
    }),
    ribbonLeaf("review", "review.language", "toolbar.proofingLanguage"),
    ribbonLeaf("review", "review.lockTracking", "toolbar.lockTracking"),
    pendingRibbonLeaf("review", "review.thesaurus", "toolbar.thesaurus"),
    pendingRibbonLeaf("review", "review.translate", "toolbar.translate", {
      ...providerBoundWebContract("review.translate"),
    }),
    ...(
      [
        ["view.page", "toolbar.pageView"],
        ["view.sheetStack", "toolbar.sheetStack"],
        ["view.outlinePane", "outline.title"],
        ["view.stylesPane", "toolbar.stylesPane"],
        ["view.backlinksPane", "toolbar.backlinks"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("view", id, commandKey, {
        docxImport: "none",
        contextPredicate: "always",
        ...(["view.sheetStack", "view.stylesPane", "view.backlinksPane"].includes(
          id,
        )
          ? visibleMatureWebContract(id)
          : {}),
      }),
    ),
    ribbonLeaf("view", "view.zoom100", "toolbar.zoom100", {
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.onePage", "toolbar.onePage", {
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.multiplePages", "toolbar.multiplePages", {
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.pageWidth", "toolbar.pageWidth", {
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.focus", "toolbar.focusMode", {
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.immersiveReader", "toolbar.immersiveReader", {
      // Immersive Reader is a session-only, read-only projection of the live
      // Docs model. It deliberately has no serialized view state or DOCX
      // import claim; the adapter gates protected documents and owns focus.
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.reading", "toolbar.readingView", {
      // Word for the web's Reading View (View > Reading View; Alt+Win, W, F):
      // the same document rendered read-only until "Edit document" returns to
      // Editing View. MRAI projects it by holding the live editor
      // non-editable for the session; it owns no view state and never
      // touches the document, so it must stay executable in a read-only
      // editor (that is how the reader leaves the view again).
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
      webEvidence:
        "ms-support-use-a-screen-reader-to-explore-and-navigate-word:web-reading-view",
    }),
    ribbonLeaf("view", "view.navigationPane", "toolbar.navigationPane", {
      // The native navigation pane projects live headings, pages, and search
      // results. It is a session view and never mutates or serializes content.
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.ruler", "toolbar.ruler", {
      // Ruler visibility is session-only. Marker edits route through the
      // existing paragraph indent/tab-stop commands and their native DOCX I/O.
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.web", "toolbar.webView", {
      // Web Layout is a session-only continuous projection. It parks only
      // browser-derived structural pagination and never serializes view state
      // or claims a DOCX representation.
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
    }),
    ribbonLeaf("view", "view.outline", "toolbar.outlineView", {
      // Outline is an editable, session-only structural projection of the
      // mounted document. Heading and paragraph outline levels remain the
      // authored model; paper chrome and derived pagination are only parked.
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
      ...hiddenMatureWebContract("view.outline"),
    }),
    ribbonLeaf("view", "view.draft", "toolbar.draftView", {
      // Draft is a session-only body editing projection. It keeps the authored
      // model and text formatting while hiding paper, pagination, and band
      // decoration; export remains owned by the unchanged document model.
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
      ...hiddenMatureWebContract("view.draft"),
    }),
    ribbonLeaf("view", "view.sideBySide", "toolbar.sideBySide", {
      // The notebook window manager already supports independent duplicate
      // views of one page. This command opens/reuses that exact page identity
      // and gives both windows bounded, session-only viewport geometry.
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
      ...hiddenMatureWebContract("view.sideBySide"),
    }),
    ribbonLeaf(
      "view",
      "view.synchronousScrolling",
      "toolbar.synchronousScrolling",
      {
        // Two registered document scrollports share only bounded vertical
        // progress. The session retires on close/page change and never writes
        // the document or a serialized DOCX view setting.
        authoring: "projection",
        docxImport: "none",
        contextPredicate: "always",
        ...hiddenMatureWebContract("view.synchronousScrolling"),
      },
    ),
    ribbonLeaf("view", "view.split", "toolbar.splitView", {
      // Two EditorViews share the exact primary transaction stream while
      // keeping independent scroll/selection projections. Split state is local
      // to the window and never enters the document or DOCX package.
      authoring: "projection",
      docxImport: "none",
      contextPredicate: "always",
      ...hiddenMatureWebContract("view.split"),
    }),
    ...(
      [
        ["view.newWindow", "toolbar.newWindow"],
        ["view.arrangeAll", "toolbar.arrangeAll"],
        ["view.switchWindows", "toolbar.switchWindows"],
      ] as const
    ).map(([id, commandKey]) =>
      ribbonLeaf("view", id, commandKey, {
        // Workspace window commands alter only session geometry and focus.
        // They never mutate the document or serialize a DOCX view setting.
        authoring: "projection",
        docxImport: "none",
        contextPredicate: "always",
        ...hiddenMatureWebContract(id),
      }),
    ),
  ]);

export function findDocumentFeatureCapability(
  id: string,
): DocumentFeatureCapability | null {
  return (
    DOCUMENT_FEATURE_CAPABILITY_REGISTRY.find((entry) => entry.id === id) ??
    null
  );
}

export const DOCUMENT_FEATURE_CAPABILITIES = Object.freeze({
  atoms: DOCUMENT_ATOM_CAPABILITIES,
  documentAttributes: DOCUMENT_ATTRIBUTE_CAPABILITIES,
  leaves: DOCUMENT_FEATURE_CAPABILITY_REGISTRY,
});

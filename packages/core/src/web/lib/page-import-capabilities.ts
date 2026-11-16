/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { CODE_LANGUAGE_OPTIONS } from "./code-runner/languages";

/**
 * Client-side page-import capability catalog.
 *
 * This catalog owns picker hints and client routing only. The native `accept`
 * attribute is not a security boundary: drag/drop can bypass it and every
 * upload still goes through the server's declaration, signature, quota, ACL,
 * and adapter validation before it is persisted.
 */
export type PageImportCapabilityId =
  | "docx"
  | "xlsx"
  | "pptx"
  | "pdf"
  | "ipynb"
  | "markdown"
  | "html"
  | "delimited"
  | "code"
  | "text";

export interface PageImportFileDescriptor {
  readonly name: string;
  readonly type?: string;
}

export interface PageImportCapability {
  readonly id: PageImportCapabilityId;
  readonly extensions: readonly string[];
  readonly mediaTypes: readonly string[];
  /** Whether the generic page-import pickers may advertise this capability. */
  readonly guidedPicker: boolean;
  /** Whether a guest editor can materialize the file without server storage. */
  readonly guestInline: boolean;
}

export const TEXT_PAGE_IMPORT_MAX_BYTES = 2 * 1024 * 1024;

const CODE_EXTENSIONS = CODE_LANGUAGE_OPTIONS.flatMap(
  (option) => option.extensions,
);

export const PAGE_IMPORT_CAPABILITIES: readonly PageImportCapability[] = [
  {
    id: "docx",
    extensions: [".docx"],
    mediaTypes: [
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "docx",
    extensions: [".docm"],
    mediaTypes: ["application/vnd.ms-word.document.macroenabled.12"],
    guidedPicker: true,
    // Macro-enabled packages require authenticated original-package custody.
    // A local guest conversion would silently discard the VBA relationship.
    guestInline: false,
  },
  {
    id: "xlsx",
    extensions: [".xlsx"],
    mediaTypes: [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "xlsx",
    extensions: [".xlsm", ".xltx", ".xltm"],
    mediaTypes: [
      "application/vnd.ms-excel.sheet.macroenabled.12",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
      "application/vnd.ms-excel.template.macroenabled.12",
    ],
    guidedPicker: true,
    // A macro-enabled workbook keeps its VBA project (never run) and downloads
    // as a .xlsm; a template opens as a workbook made from it. Both under
    // original-package custody: a local guest reading would drop the project.
    guestInline: false,
  },
  {
    id: "xlsx",
    extensions: [".xls"],
    mediaTypes: ["application/vnd.ms-excel"],
    guidedPicker: true,
    // A legacy BIFF8 workbook has no browser reader: the server converts it to
    // .xlsx through LibreOffice under original-package custody, the way Excel
    // Online converts a .xls before editing. Guests cannot take that path.
    guestInline: false,
  },
  {
    id: "pptx",
    extensions: [".pptx"],
    mediaTypes: [
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "pptx",
    extensions: [".pptm", ".potm", ".ppsm"],
    mediaTypes: [
      "application/vnd.ms-powerpoint.presentation.macroenabled.12",
      "application/vnd.ms-powerpoint.template.macroenabled.12",
      "application/vnd.ms-powerpoint.slideshow.macroenabled.12",
    ],
    guidedPicker: true,
    // A macro-enabled deck is imported as a deck under original-package
    // custody, its VBA project kept (never run) and its download a .pptm.
    // A local guest conversion would silently discard the project.
    guestInline: false,
  },
  {
    id: "pptx",
    extensions: [".potx", ".ppsx"],
    mediaTypes: [
      "application/vnd.openxmlformats-officedocument.presentationml.template",
      "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
    ],
    guidedPicker: true,
    // PowerPoint's Design Template and Show: the presentation's own package,
    // read on the server as the deck it holds (a new presentation from the
    // template, as File > New makes one; a show opened for editing, as File >
    // Open does) under original-package custody.
    guestInline: false,
  },
  {
    id: "pptx",
    // A PowerPoint 97-2003 Show (.pps) or Template (.pot) is the same
    // compound document as a .ppt, converted the same way.
    extensions: [".ppt", ".pps", ".pot"],
    mediaTypes: ["application/vnd.ms-powerpoint"],
    guidedPicker: true,
    // A PowerPoint 97-2003 deck has no browser reader: the server converts it
    // to .pptx through LibreOffice under original-package custody, with the
    // deck's own theme and table fills put back (`pptx/legacy_ppt`).
    guestInline: false,
  },
  // Other office formats Office itself opens: no browser reader, so the
  // server converts each to the package its page edits through LibreOffice
  // under original-package custody (`file_adapters/office_conversion`).
  {
    id: "xlsx",
    // An Excel 2003 XML Spreadsheet is a .xml like any XML: it is not listed
    // here (a plain .xml stays a code page) and is routed by its head
    // (`isSpreadsheetMl2003Head`) to the server, which converts it.
    extensions: [".xlsb", ".ods", ".fods", ".ots", ".sxc", ".wk1", ".wks", ".xlt", ".et", ".ett", ".slk", ".dif", ".dbf", ".numbers"],
    mediaTypes: [
      "application/vnd.ms-excel.sheet.binary.macroenabled.12",
      "application/vnd.oasis.opendocument.spreadsheet",
      "application/vnd.oasis.opendocument.spreadsheet-flat-xml",
      "application/vnd.oasis.opendocument.spreadsheet-template",
      "application/vnd.sun.xml.calc",
      "application/vnd.lotus-1-2-3",
      "application/x-sylk",
      "application/x-dif",
      "application/x-dbf",
      "application/vnd.apple.numbers",
    ],
    guidedPicker: true,
    guestInline: false,
  },
  {
    id: "docx",
    extensions: [".doc", ".odt", ".rtf"],
    mediaTypes: [
      "application/msword",
      "application/vnd.oasis.opendocument.text",
      "application/rtf",
      "text/rtf",
    ],
    guidedPicker: true,
    guestInline: false,
  },
  {
    // Word's other documents and templates, each read on the server into the
    // package its page edits: a Word 97-2003 template (.dot) as the .doc it
    // is; WPS Writer (.wps/.wpt, a Word 97-2003 binary) and Microsoft Works
    // (.wps); OpenDocument Text Template, Flat OpenDocument Text and
    // OpenOffice.org 1.x Writer; and a Word template (.dotx/.dotm), edited as
    // the document File > New makes of it (a .dotm as a .docm, its macros
    // kept and never run); and Word's Single File Web Page (.mht/.mhtml).
    id: "docx",
    extensions: [".dot", ".wps", ".wpt", ".ott", ".fodt", ".sxw", ".dotx", ".dotm", ".mht", ".mhtml"],
    mediaTypes: [
      "application/vnd.ms-works",
      "application/wps-office.wps",
      "application/kswps",
      "application/vnd.oasis.opendocument.text-template",
      "application/vnd.oasis.opendocument.text-flat-xml",
      "application/vnd.sun.xml.writer",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
      "application/vnd.ms-word.template.macroenabled.12",
      "application/x-mimearchive",
      "multipart/related",
    ],
    guidedPicker: true,
    guestInline: false,
  },
  {
    // Impress's and WPS's decks: OpenDocument Presentation, its template
    // (a new presentation, as Impress makes one), Flat OpenDocument
    // Presentation, OpenOffice.org 1.x Impress, and WPS Presentation and its
    // template (a PowerPoint 97-2003 binary, converted as a .ppt).
    id: "pptx",
    extensions: [".odp", ".otp", ".fodp", ".sxi", ".dps", ".dpt"],
    mediaTypes: [
      "application/vnd.oasis.opendocument.presentation",
      "application/vnd.oasis.opendocument.presentation-template",
      "application/vnd.oasis.opendocument.presentation-flat-xml",
      "application/vnd.sun.xml.impress",
      "application/wps-office.dps",
      "application/wps-office.dpt",
    ],
    guidedPicker: true,
    guestInline: false,
  },
  {
    id: "pdf",
    extensions: [".pdf"],
    mediaTypes: ["application/pdf"],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "ipynb",
    extensions: [".ipynb"],
    mediaTypes: ["application/x-ipynb+json"],
    // Retained only so attachment indexing can classify existing files. Page
    // creation no longer advertises or materializes Jupyter notebooks.
    guidedPicker: false,
    guestInline: false,
  },
  {
    id: "markdown",
    extensions: [".md", ".markdown"],
    mediaTypes: ["text/markdown", "text/x-markdown"],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "html",
    extensions: [".html", ".htm"],
    mediaTypes: ["text/html", "application/xhtml+xml"],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "delimited",
    extensions: [".csv", ".tsv", ".psv"],
    mediaTypes: ["text/csv", "text/tab-separated-values"],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "code",
    extensions: CODE_EXTENSIONS,
    mediaTypes: [
      "application/json",
      "application/yaml",
      "application/x-yaml",
    ],
    guidedPicker: true,
    guestInline: true,
  },
  {
    id: "text",
    // .tex is LaTeX source and imports as a Code page (CODE_EXTENSIONS).
    extensions: [".txt", ".log", ".rst"],
    mediaTypes: ["text/plain", "text/*"],
    guidedPicker: true,
    guestInline: true,
  },
] as const;

function unique(values: readonly string[]): string[] {
  return Array.from(new Set(values));
}

/** Native file-picker hint generated from the same rows client routing reads. */
export const GUIDED_PAGE_IMPORT_ACCEPT = unique(
  PAGE_IMPORT_CAPABILITIES.filter((row) => row.guidedPicker).flatMap((row) => [
    ...row.extensions,
    ...row.mediaTypes,
  ]),
).join(",");

/** Picker hint for the local-only guest editor (no attachment persistence). */
export const GUEST_INLINE_PAGE_IMPORT_ACCEPT = unique(
  PAGE_IMPORT_CAPABILITIES.filter((row) => row.guestInline).flatMap((row) => [
    ...row.extensions,
    ...row.mediaTypes,
  ]),
).join(",");

export function pageImportExtension(name: string): string {
  const normalized = name.toLowerCase();
  const index = normalized.lastIndexOf(".");
  return index >= 0 ? normalized.slice(index) : "";
}

export function pageImportMediaType(type: string | undefined): string {
  return (type || "").split(";", 1)[0].trim().toLowerCase();
}

const SPREADSHEETML_2003_NAMESPACE = "urn:schemas-microsoft-com:office:spreadsheet";
const SPREADSHEETML_2003_ROOT = /<(?:[A-Za-z_][\w.-]*:)?Workbook[\s>/]/;

/**
 * Whether the start of a .xml file (its first 256 KB, decoded as UTF-8) is an
 * Excel 2003 XML Spreadsheet, as the server judges it
 * (`office_conversion.is_spreadsheetml_2003`): XML whose root is
 * SpreadsheetML's `Workbook`. Such a file is imported as a workbook (source
 * kind `xlsx`); any other .xml keeps its own route.
 */
export function isSpreadsheetMl2003Head(head: string): boolean {
  const body = head.replace(/^\uFEFF/, "").trimStart();
  if (!body.startsWith("<") || body.includes("\u0000")) return false;
  return body.includes(SPREADSHEETML_2003_NAMESPACE) && SPREADSHEETML_2003_ROOT.test(body);
}

export function resolvePageImportCapability(
  file: PageImportFileDescriptor,
): PageImportCapability | null {
  const extension = pageImportExtension(file.name);
  if (extension) {
    const byExtension = PAGE_IMPORT_CAPABILITIES.find((row) =>
      row.extensions.includes(extension),
    );
    if (byExtension) return byExtension;
  }

  const mediaType = pageImportMediaType(file.type);
  if (!mediaType) return null;
  const byMediaType = PAGE_IMPORT_CAPABILITIES.find((row) =>
    row.mediaTypes.includes(mediaType),
  );
  if (byMediaType) return byMediaType;
  if (mediaType.startsWith("text/")) {
    return (
      PAGE_IMPORT_CAPABILITIES.find((row) => row.id === "text") ?? null
    );
  }
  return null;
}

export function isTextPageImportFile(
  file: PageImportFileDescriptor,
): boolean {
  const id = resolvePageImportCapability(file)?.id;
  return (
    id === "markdown" ||
    id === "html" ||
    id === "delimited" ||
    id === "code" ||
    id === "text"
  );
}

export function isMarkdownPageImportFile(
  file: PageImportFileDescriptor,
): boolean {
  return resolvePageImportCapability(file)?.id === "markdown";
}

export function isHtmlPageImportFile(
  file: PageImportFileDescriptor,
): boolean {
  return resolvePageImportCapability(file)?.id === "html";
}

/**
 * Files imported only through server conversion, by the package their page
 * edits: legacy .xls/.ppt, and the formats `file_adapters/office_conversion`
 * converts (Excel Binary, OpenDocument, Word 97-2003, RTF).
 */
const CONVERTED_IMPORT_EXTENSIONS: Readonly<Record<string, "xlsx" | "docx" | "pptx">> = {
  ".xls": "xlsx",
  ".xlsb": "xlsx",
  ".ods": "xlsx",
  ".fods": "xlsx",
  ".ots": "xlsx",
  ".sxc": "xlsx",
  ".wk1": "xlsx",
  ".wks": "xlsx",
  ".xlt": "xlsx",
  ".et": "xlsx",
  ".ett": "xlsx",
  ".slk": "xlsx",
  ".dif": "xlsx",
  ".dbf": "xlsx",
  ".numbers": "xlsx",
  ".doc": "docx",
  ".dot": "docx",
  ".odt": "docx",
  ".rtf": "docx",
  ".wps": "docx",
  ".wpt": "docx",
  ".ott": "docx",
  ".fodt": "docx",
  ".sxw": "docx",
  ".dotx": "docx",
  ".dotm": "docx",
  ".mht": "docx",
  ".mhtml": "docx",
  ".ppt": "pptx",
  ".pps": "pptx",
  ".pot": "pptx",
  ".odp": "pptx",
  ".otp": "pptx",
  ".fodp": "pptx",
  ".sxi": "pptx",
  ".dps": "pptx",
  ".dpt": "pptx",
};
const CONVERTED_IMPORT_MEDIA_TYPES: Readonly<Record<string, "xlsx" | "docx" | "pptx">> = {
  "application/vnd.ms-excel": "xlsx",
  "application/vnd.ms-excel.sheet.binary.macroenabled.12": "xlsx",
  "application/vnd.oasis.opendocument.spreadsheet": "xlsx",
  "application/vnd.oasis.opendocument.spreadsheet-flat-xml": "xlsx",
  "application/vnd.oasis.opendocument.spreadsheet-template": "xlsx",
  "application/vnd.sun.xml.calc": "xlsx",
  "application/vnd.lotus-1-2-3": "xlsx",
  "application/x-sylk": "xlsx",
  "application/x-dif": "xlsx",
  "application/x-dbf": "xlsx",
  "application/vnd.apple.numbers": "xlsx",
  "application/msword": "docx",
  "application/vnd.oasis.opendocument.text": "docx",
  "application/rtf": "docx",
  "text/rtf": "docx",
  "application/vnd.ms-works": "docx",
  "application/wps-office.wps": "docx",
  "application/kswps": "docx",
  "application/vnd.oasis.opendocument.text-template": "docx",
  "application/vnd.oasis.opendocument.text-flat-xml": "docx",
  "application/vnd.sun.xml.writer": "docx",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template": "docx",
  "application/vnd.ms-word.template.macroenabled.12": "docx",
  "application/x-mimearchive": "docx",
  "multipart/related": "docx",
  "application/vnd.ms-powerpoint": "pptx",
  "application/vnd.oasis.opendocument.presentation": "pptx",
  "application/vnd.oasis.opendocument.presentation-template": "pptx",
  "application/vnd.oasis.opendocument.presentation-flat-xml": "pptx",
  "application/vnd.sun.xml.impress": "pptx",
  "application/wps-office.dps": "pptx",
  "application/wps-office.dpt": "pptx",
};

/** The package a server-converted file is edited as, or null for any other file. */
export function convertedImportTarget(
  file: PageImportFileDescriptor,
): "xlsx" | "docx" | "pptx" | null {
  const byExtension = CONVERTED_IMPORT_EXTENSIONS[pageImportExtension(file.name)];
  if (byExtension) return byExtension;
  return CONVERTED_IMPORT_MEDIA_TYPES[pageImportMediaType(file.type)] ?? null;
}

/** A deck in another format (.ppt/.pps/.pot, .odp/.otp/.fodp, .sxi, .dps/.dpt):
 * imported only through server conversion. */
export function isLegacyPresentationImportFile(
  file: PageImportFileDescriptor,
): boolean {
  return convertedImportTarget(file) === "pptx";
}

/** A workbook in another format (.xls, .xlsb, .ods, .fods): imported only
 * through server conversion. */
export function isLegacySpreadsheetImportFile(
  file: PageImportFileDescriptor,
): boolean {
  return convertedImportTarget(file) === "xlsx";
}

/** A document in another format (.doc, .odt, .rtf): imported only through
 * server conversion. */
export function isLegacyDocumentImportFile(
  file: PageImportFileDescriptor,
): boolean {
  return convertedImportTarget(file) === "docx";
}

/**
 * A file the server converts from a format that is neither the package its
 * page edits nor Office's own legacy binary of it (OpenDocument, Excel
 * Binary, Word 97-2003, RTF): its import says that it was converted, as
 * Office says when it opens one.
 */
export function convertedImportSourceLabel(
  file: PageImportFileDescriptor,
): string | null {
  const extension = pageImportExtension(file.name);
  // Excel's and PowerPoint's own legacy binaries (and the BIFF template, the
  // 97-2003 show and template) say nothing.
  if ([".xls", ".xlt", ".ppt", ".pps", ".pot"].includes(extension)) return null;
  if (extension && CONVERTED_IMPORT_EXTENSIONS[extension]) {
    return extension.slice(1).toUpperCase();
  }
  return null;
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Print preflight (WPS PDF ▸ 打印预检).
 *
 * The point of a preflight is to say what will go wrong at the printer while
 * there is still time to fix it: a font nobody embedded, an image that looks
 * fine on screen and muddy on paper, a hairline the press cannot hold, content
 * that runs past the printable area, pages that are not all the same size, and
 * annotations the author can see but the printer will not print.
 *
 * Pure: a description of the document in, ranked findings out. Every finding
 * names the page and says what was measured, so the panel never has to guess
 * at wording and a fix can be checked by re-running the same scan.
 */

export type PreflightSeverity = "error" | "warning" | "info";

export type PreflightRule =
  | "font-not-embedded"
  | "image-below-print-resolution"
  | "hairline-stroke"
  | "content-outside-printable-area"
  | "mixed-page-sizes"
  | "annotation-will-not-print"
  | "no-pages";

export interface PreflightPage {
  /** 1-based, as the panel shows it. */
  number: number;
  widthPt: number;
  heightPt: number;
  /** Printable inset on each side, in points. */
  marginPt?: number;
  /** Furthest extent of drawn content from the page edges, in points. */
  contentInsetPt?: number;
  images?: readonly { name: string; effectiveDpi: number }[];
  /** Stroke widths used on the page, in points. */
  strokeWidthsPt?: readonly number[];
  /** Annotations that are visible on screen but flagged not to print. */
  nonPrintingAnnotations?: number;
}

export interface PreflightDocument {
  pages: readonly PreflightPage[];
  /** Fonts the document uses, with whether the file carries the outlines. */
  fonts?: readonly { name: string; embedded: boolean }[];
  /** Below this an image is called out. 150 is the usual floor for print. */
  minimumDpi?: number;
  /** Strokes thinner than this may disappear on press. */
  hairlinePt?: number;
}

export interface PreflightFinding {
  rule: PreflightRule;
  severity: PreflightSeverity;
  /** 0 for a document-wide finding. */
  page: number;
  /** What was measured, for the panel's detail line. */
  detail: string;
}

export const DEFAULT_MINIMUM_DPI = 150;
export const DEFAULT_HAIRLINE_PT = 0.25;

export const PDF_PRINT_INVENTORY_MAX_PAGES = 10_000;
export const PDF_PRINT_PAGE_MAX_POINTS = 1_000_000;

export type PdfPrintPageRotation = 0 | 90 | 180 | 270;

export interface PdfPrintPageInventoryInput {
  /** Final 1-based output position, not a source-page id. */
  page: number;
  /** Unrotated final CropBox width and height, in PDF points. */
  widthPt: number;
  heightPt: number;
  /** Final effective /Rotate, normalized to a clockwise quarter-turn. */
  rotation: number;
}

export interface PdfPrintPageInventoryEntry {
  readonly page: number;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly rotation: PdfPrintPageRotation;
  readonly effectiveWidthPt: number;
  readonly effectiveHeightPt: number;
}

export interface PdfPrintPageSizeGroup {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly pages: readonly number[];
}

export interface PdfPrintPageInventory {
  readonly pages: readonly PdfPrintPageInventoryEntry[];
  readonly sizes: readonly PdfPrintPageSizeGroup[];
  /** Stable exact-CAS identity over every final page box and rotation. */
  readonly fingerprint: string;
}

function canonicalPrintPoint(value: unknown): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value <= 0 ||
    value > PDF_PRINT_PAGE_MAX_POINTS
  ) {
    throw new Error("Invalid PDF print page size");
  }
  // Normalize only parser/serializer float noise. The fingerprint still keeps
  // six decimal places (far below a printer pixel) so distinct page boxes are
  // never merged at the two-decimal precision used only by the visible label.
  const canonical = Math.round(value * 1_000_000) / 1_000_000;
  if (canonical <= 0) throw new Error("Invalid PDF print page size");
  return canonical;
}

function canonicalPrintRotation(value: unknown): PdfPrintPageRotation {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error("Invalid PDF print page rotation");
  }
  const integer = Math.trunc(value);
  if (integer !== value || integer % 90 !== 0) {
    throw new Error("Invalid PDF print page rotation");
  }
  return (((integer % 360) + 360) % 360) as PdfPrintPageRotation;
}

/**
 * Build the one authoritative page-box inventory shared by the visible
 * preflight and the final-byte witness.
 *
 * The array is deliberately strict: it is either a complete, consecutive
 * 1..N output inventory or it throws. Filtering malformed pages would turn a
 * partial scan into permission to print a different document.
 */
export function createPdfPrintPageInventory(
  raw: unknown,
): PdfPrintPageInventory {
  if (!Array.isArray(raw) || raw.length > PDF_PRINT_INVENTORY_MAX_PAGES) {
    throw new Error("Invalid PDF print page inventory");
  }

  const pages: PdfPrintPageInventoryEntry[] = [];
  const seen = new Set<number>();
  for (let index = 0; index < raw.length; index += 1) {
    const candidate = raw[index];
    if (!candidate || typeof candidate !== "object") {
      throw new Error("Invalid PDF print page inventory");
    }
    const input = candidate as Partial<PdfPrintPageInventoryInput>;
    const page = input.page;
    if (
      typeof page !== "number" ||
      !Number.isSafeInteger(page) ||
      page !== index + 1 ||
      seen.has(page)
    ) {
      throw new Error("Invalid PDF print page identity");
    }
    seen.add(page);
    const widthPt = canonicalPrintPoint(input.widthPt);
    const heightPt = canonicalPrintPoint(input.heightPt);
    const rotation = canonicalPrintRotation(input.rotation);
    const swapsAxes = rotation === 90 || rotation === 270;
    pages.push(
      Object.freeze({
        page,
        widthPt,
        heightPt,
        rotation,
        effectiveWidthPt: swapsAxes ? heightPt : widthPt,
        effectiveHeightPt: swapsAxes ? widthPt : heightPt,
      }),
    );
  }

  const sizeGroups = new Map<
    string,
    { widthPt: number; heightPt: number; pages: number[] }
  >();
  for (const page of pages) {
    const key = `${page.effectiveWidthPt}x${page.effectiveHeightPt}`;
    const existing = sizeGroups.get(key);
    if (existing) {
      existing.pages.push(page.page);
    } else {
      sizeGroups.set(key, {
        widthPt: page.effectiveWidthPt,
        heightPt: page.effectiveHeightPt,
        pages: [page.page],
      });
    }
  }
  const sizes = [...sizeGroups.values()].map((group) =>
    Object.freeze({
      widthPt: group.widthPt,
      heightPt: group.heightPt,
      pages: Object.freeze([...group.pages]),
    }),
  );
  const frozenPages = Object.freeze([...pages]);
  return Object.freeze({
    pages: frozenPages,
    sizes: Object.freeze(sizes),
    fingerprint: JSON.stringify(
      frozenPages.map((page) => [
        page.page,
        page.widthPt,
        page.heightPt,
        page.rotation,
      ]),
    ),
  });
}

export function pdfPrintPageInventoriesMatch(
  expected: PdfPrintPageInventory,
  current: PdfPrintPageInventory,
): boolean {
  return (
    expected.pages.length === current.pages.length &&
    expected.fingerprint === current.fingerprint
  );
}

/** Run the checks a printer would run. */
export function preflight(document: PreflightDocument): PreflightFinding[] {
  const findings: PreflightFinding[] = [];
  const minimumDpi = document.minimumDpi ?? DEFAULT_MINIMUM_DPI;
  const hairline = document.hairlinePt ?? DEFAULT_HAIRLINE_PT;

  if (document.pages.length === 0) {
    return [
      { rule: "no-pages", severity: "error", page: 0, detail: "0 pages" },
    ];
  }

  for (const font of document.fonts ?? []) {
    if (!font.embedded) {
      findings.push({
        rule: "font-not-embedded",
        severity: "error",
        page: 0,
        detail: font.name,
      });
    }
  }

  // Every page the same size is what a printer expects; report the odd ones.
  const sizeKey = (page: PreflightPage) =>
    `${Math.round(page.widthPt)}×${Math.round(page.heightPt)}`;
  const sizes = new Map<string, number[]>();
  for (const page of document.pages) {
    const key = sizeKey(page);
    sizes.set(key, [...(sizes.get(key) ?? []), page.number]);
  }
  if (sizes.size > 1) {
    const summary = [...sizes.entries()]
      .map(([size, pages]) => `${size}pt ×${pages.length}`)
      .join(", ");
    findings.push({
      rule: "mixed-page-sizes",
      severity: "warning",
      page: 0,
      detail: summary,
    });
  }

  for (const page of document.pages) {
    for (const image of page.images ?? []) {
      if (image.effectiveDpi < minimumDpi) {
        findings.push({
          rule: "image-below-print-resolution",
          severity: image.effectiveDpi < minimumDpi / 2 ? "error" : "warning",
          page: page.number,
          detail: `${image.name} @ ${Math.round(image.effectiveDpi)} dpi`,
        });
      }
    }
    for (const width of page.strokeWidthsPt ?? []) {
      if (width > 0 && width < hairline) {
        findings.push({
          rule: "hairline-stroke",
          severity: "warning",
          page: page.number,
          detail: `${width}pt`,
        });
        break;
      }
    }
    if (
      page.marginPt !== undefined &&
      page.contentInsetPt !== undefined &&
      page.contentInsetPt < page.marginPt
    ) {
      findings.push({
        rule: "content-outside-printable-area",
        severity: "error",
        page: page.number,
        detail: `${page.contentInsetPt}pt < ${page.marginPt}pt`,
      });
    }
    if ((page.nonPrintingAnnotations ?? 0) > 0) {
      findings.push({
        rule: "annotation-will-not-print",
        severity: "info",
        page: page.number,
        detail: `${page.nonPrintingAnnotations}`,
      });
    }
  }

  const rank: Record<PreflightSeverity, number> = {
    error: 0,
    warning: 1,
    info: 2,
  };
  return findings.sort(
    (a, b) => rank[a.severity] - rank[b.severity] || a.page - b.page,
  );
}

/** True when nothing found would stop the job. */
export function readyToPrint(findings: readonly PreflightFinding[]): boolean {
  return !findings.some((finding) => finding.severity === "error");
}

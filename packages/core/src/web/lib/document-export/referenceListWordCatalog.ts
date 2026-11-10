/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import catalog from "./referenceListWordCatalog.json";

export const INDEX_FORMATS = [
  "template",
  "classic",
  "fancy",
  "modern",
  "bulleted",
  "formal",
  "simple",
] as const;
export type IndexFormat = (typeof INDEX_FORMATS)[number];

export const TOF_FORMATS = [
  "template",
  "classic",
  "distinctive",
  "centered",
  "formal",
  "simple",
] as const;
export type TofFormat = (typeof TOF_FORMATS)[number];

export type ReferenceListTabLeader = "dot" | "hyphen" | "underscore" | "none";

interface ReferenceListCatalogShape {
  index: Record<
    string,
    {
      headings: boolean;
      defaultTabLeader: ReferenceListTabLeader;
    }
  >;
  tableOfFigures: Record<
    string,
    {
      defaultTabLeader: ReferenceListTabLeader;
      defaultRightAlignPageNumbers: boolean;
    }
  >;
}

const CATALOG = catalog as unknown as ReferenceListCatalogShape;

export function isIndexFormat(value: unknown): value is IndexFormat {
  return (
    typeof value === "string" &&
    (INDEX_FORMATS as readonly string[]).includes(value)
  );
}

export function isTofFormat(value: unknown): value is TofFormat {
  return (
    typeof value === "string" &&
    (TOF_FORMATS as readonly string[]).includes(value)
  );
}

/** Whether a format puts letter headings between alphabetic groups. */
export function indexFormatHasHeadings(format: IndexFormat): boolean {
  return CATALOG.index[format]?.headings === true;
}

/** The tab leader Word's dialog preselects for an index format. */
export function indexFormatDefaultTabLeader(
  format: IndexFormat,
): ReferenceListTabLeader {
  return CATALOG.index[format]?.defaultTabLeader ?? "dot";
}

/** The leader and alignment Word's dialog preselects for a figures format. */
export function tofFormatDefaults(format: TofFormat): {
  tabLeader: ReferenceListTabLeader;
  rightAlignPageNumbers: boolean;
} {
  const entry = CATALOG.tableOfFigures[format] ?? CATALOG.tableOfFigures.template;
  return {
    tabLeader: entry.defaultTabLeader,
    rightAlignPageNumbers: entry.defaultRightAlignPageNumbers,
  };
}

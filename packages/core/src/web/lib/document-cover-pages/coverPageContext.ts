/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What one document's cover shows: the page it sits on, the palette its
 * theme colours resolve against, and the property values its fields print.
 * Pure (no editor imports) so the editor atom, the gallery and the DOCX
 * exporter all read a cover the same way.
 */

import {
  documentMetadataValue,
  normalizeDocumentMetadata,
} from "../../components/console/editor/native-document/documentMetadata";
import {
  normalizePageOrientation,
  pageGeometryTwips,
} from "../../components/console/editor/native-document/pageSetup";
import { documentThemeColorsForExport } from "../office-table-style-word";

import {
  COVER_FIELD_METADATA_KEY,
  coverSceneFromDefinition,
  documentCoverDefinition,
  normalizeCoverScene,
  type CoverFieldKey,
  type CoverScene,
  type CoverThemeColors,
} from "./coverPages";

const EMU_PER_TWIP = 635;

export interface DocumentCoverPageContext {
  page: { widthEmu: number; heightEmu: number };
  margins: { left: number; top: number; right: number; bottom: number };
  theme: CoverThemeColors;
  fontFamily: string | undefined;
  values: Partial<Record<CoverFieldKey, string>>;
}

/** Read one document's cover context from its root attrs. */
export function documentCoverPageContext(
  attrs: Record<string, unknown>,
): DocumentCoverPageContext {
  const geometry = pageGeometryTwips(
    attrs.pageSize,
    attrs.pageMargins,
    attrs.customPageSize,
    attrs.customPageMargins,
    normalizePageOrientation(attrs.pageOrientation),
  );
  const metadata = normalizeDocumentMetadata(attrs.documentMetadata);
  const values: Partial<Record<CoverFieldKey, string>> = {};
  for (const [field, key] of Object.entries(COVER_FIELD_METADATA_KEY) as Array<
    [CoverFieldKey, (typeof COVER_FIELD_METADATA_KEY)[CoverFieldKey]]
  >) {
    const value = documentMetadataValue(metadata, key);
    if (value) values[field] = value;
  }
  // The one document theme the DOCX theme part carries, so the cover paints
  // here as Word draws it (imported theme, Design colour set, Office theme).
  const theme = documentThemeColorsForExport(attrs);
  const fontFamily =
    typeof attrs.defaultFontFamily === "string" && attrs.defaultFontFamily
      ? attrs.defaultFontFamily
      : undefined;
  return {
    page: {
      widthEmu: geometry.width * EMU_PER_TWIP,
      heightEmu: geometry.height * EMU_PER_TWIP,
    },
    margins: {
      left: geometry.left * EMU_PER_TWIP,
      top: geometry.top * EMU_PER_TWIP,
      right: geometry.right * EMU_PER_TWIP,
      bottom: geometry.bottom * EMU_PER_TWIP,
    },
    theme: {
      accents: [
        theme.accent1,
        theme.accent2,
        theme.accent3,
        theme.accent4,
        theme.accent5,
        theme.accent6,
      ],
      base: {
        dk1: theme.text1,
        lt1: theme.background1,
        dk2: theme.text2,
        lt2: theme.background2,
      },
    },
    fontFamily,
    values,
  };
}

/**
 * The scene one cover atom draws: an imported cover carries its own scene;
 * an MRAI cover is drawn from the catalog on this document's page and theme.
 */
export function documentCoverPageScene(
  attrs: Record<string, unknown>,
  context: DocumentCoverPageContext,
): CoverScene | null {
  const imported = normalizeCoverScene(attrs.scene);
  if (imported) return imported;
  const definition = documentCoverDefinition(attrs.coverId);
  if (!definition) return null;
  return coverSceneFromDefinition(
    definition,
    context.page,
    context.theme,
    typeof attrs.locale === "string" ? attrs.locale : "en",
  );
}

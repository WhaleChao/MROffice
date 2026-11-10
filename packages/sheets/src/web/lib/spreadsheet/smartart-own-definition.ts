/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The layout definition a preserved sheet diagram converts from when MRAI
 * stocks no definition of its layout: its own part in the workbook's package,
 * by urn and digest (`parts.layout`, `fingerprints.layout`, SMARTART-GAP.md
 * §8.6). `null` for a layout MRAI owns, or a view that names no part.
 *
 * Kept apart from `./native-smartart-convert` (server-side: the shared
 * layout engine, the package's XML), which re-exports it: the worksheet's
 * object context asks it in the browser, and importing it from there put
 * the engine host's modules in the public bundle.
 */
import { nativeLayoutToCatalogId } from "@mrnote/core/web/lib/office-smartart/native-layouts.ts";

type UnknownRecord = Record<string, unknown>;

function record(value: unknown): UnknownRecord | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as UnknownRecord)
    : null;
}

export function spreadsheetSmartArtOwnDefinitionRef(
  object: unknown,
): { readonly urn: string; readonly sha256: string } | null {
  const view = record(record(record(object)?.source)?.nativeSmartArtView);
  const urn = view?.layoutIdentity;
  const part = record(view?.parts)?.layout;
  const sha256 = record(view?.fingerprints)?.layout;
  if (
    typeof urn !== "string" ||
    !urn ||
    nativeLayoutToCatalogId(urn) !== null ||
    typeof part !== "string" ||
    !/^xl\/diagrams\/layout[0-9]*\.xml$/.test(part) ||
    typeof sha256 !== "string" ||
    !/^[0-9a-f]{64}$/.test(sha256)
  ) {
    return null;
  }
  return { urn, sha256 };
}


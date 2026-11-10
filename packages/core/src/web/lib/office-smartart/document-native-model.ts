/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the Docs host reads off a diagram before converting it: the native view
 * a preserved group carries, the catalog layout its Office layout imports as,
 * and whether its pictures allow it. The conversion itself
 * (`./document-native-import`) runs on the server; the browser asks for it
 * through `./remote`.
 */

import { officeSmartArtLayout } from "./catalog";
import { smartArtPictureLayoutConvertible } from "./native-convert-model";
import { nativeLayoutToCatalogId } from "./native-layouts";

/** The schema the API's preserve-only projection declares about itself. */
export const NATIVE_SMARTART_VIEW_SCHEMA = "office-smartart-native-v1";

/** The five DiagramML parts one diagram is made of, in the host's own naming. */
export type NativeSmartArtPartRole =
  | "data"
  | "layout"
  | "style"
  | "colors"
  | "drawing";

/**
 * What `POST /api/v1/office/smartart/native-projection` returns for one
 * diagram package. Only the fields the Docs host reads are declared; anything
 * else the API adds travels through `nativeSmartArtView` untouched.
 */
export interface NativeSmartArtProjectionResponse {
  readonly schemaVersion?: unknown;
  readonly layoutIdentity?: unknown;
  readonly members?: unknown;
  readonly textEditable?: unknown;
  readonly receipt?: unknown;
  readonly dataXml?: unknown;
  readonly widthEmu?: unknown;
  readonly heightEmu?: unknown;
  readonly unprojectedNodes?: unknown;
  readonly parts?: unknown;
  readonly fingerprints?: unknown;
  readonly smartArtStyle?: unknown;
  readonly scene3d?: unknown;
  /** Why the Docs group model refused this diagram, when it did. */
  readonly documentFallbackRefusal?: unknown;
  /** The diagram drawn as a picture, for exactly that case. */
  readonly documentFallbackImage?: unknown;
}


export interface DocumentNativeSmartArtFrame {
  readonly widthEmu: number;
  readonly heightEmu: number;
}


/**
 * The catalog layout a Word diagram imports as, or `null` when the Docs host
 * keeps it preserved.
 *
 * One reason a diagram stays preserved at this step: MRAI does not own its
 * Office layout. The four PICTURE layouts DO import — a group member holds no
 * picture, so each picture point becomes the neutral slot the worksheet host
 * draws for it (`documentSmartArtGroupMembers`) — but only while the diagram's
 * points carry no picture of their own; see
 * {@link documentSmartArtPictureLayoutConvertible}, which
 * {@link convertDocumentSmartArt} applies before converting anything.
 */
export function documentSmartArtCatalogLayoutId(
  layoutIdentity: unknown,
): string | null {
  const layoutId = nativeLayoutToCatalogId(layoutIdentity);
  const entry = layoutId ? officeSmartArtLayout(layoutId) : null;
  return entry ? entry.id : null;
}

/**
 * May this outline become a Docs diagram of THIS layout?
 *
 * Fail closed on the one loss converting a picture layout could cause: a
 * picture point whose picture the Docs group cannot hold would come back as an
 * empty slot, and the diagram would have lost it. Such a diagram stays
 * preserved — Office's own drawing, pictures and all — exactly as before.
 *
 * A Docs member CAN hold one now (`DocumentShapeGroupMember.src`), so the test
 * is no longer "does this diagram have pictures" but "did the caller bring
 * them": a point with a picture is convertible when `images` resolves its
 * relationship to a payload the member model accepts. A caller that reads no
 * media at all passes none and gets the old answer exactly — which is what
 * keeps a host that has not been taught to read them from silently dropping
 * pictures instead of preserving the diagram.
 */
export const documentSmartArtPictureLayoutConvertible =
  smartArtPictureLayoutConvertible;


/** The native view a preserved Docs SmartArt group carries, or `null`. */
export function documentNativeSmartArtView(
  attrs: unknown,
): {
  readonly layoutIdentity: string | null;
  readonly dataXml: string;
  readonly smartArtStyle?: unknown;
  /** WordArt each data point's words wear (by the point's `modelId`), as the
   *  drawing part's label runs state it: the conversion keeps it. */
  readonly itemTextPaint?: Readonly<Record<string, unknown>>;
  /** The host theme's fill and effect style lists, as the import read them (`hostTheme`). */
  readonly hostTheme?: { readonly effects?: string; readonly xml: string };
  /**
   * The diagram's layout part in the document's own package, by address and
   * digest (`parts.layout`, `fingerprints.layout`): a layout MRAI stocks no
   * definition of is laid out from it (SMARTART-GAP.md §8.6).
   */
  readonly layoutDefinition?: { readonly part: string; readonly sha256: string };
} | null {
  const source =
    attrs && typeof attrs === "object"
      ? ((attrs as Record<string, unknown>).source as Record<string, unknown> | null)
      : null;
  const view =
    source && typeof source === "object"
      ? (source.nativeSmartArtView as Record<string, unknown> | undefined)
      : undefined;
  if (
    !view ||
    typeof view !== "object" ||
    view.schemaVersion !== NATIVE_SMARTART_VIEW_SCHEMA ||
    typeof view.dataXml !== "string" ||
    !view.dataXml
  ) {
    return null;
  }
  const itemTextPaint: Record<string, unknown> = {};
  for (const member of Array.isArray(view.members) ? view.members : []) {
    if (!member || typeof member !== "object") continue;
    const record = member as { role?: unknown; dataPointIds?: unknown; textStyle?: { textPaint?: unknown } };
    const paint = record.role === "text" ? record.textStyle?.textPaint : undefined;
    if (!paint || typeof paint !== "object" || !Array.isArray(record.dataPointIds)) continue;
    for (const id of record.dataPointIds) {
      if (typeof id === "string" && id) itemTextPaint[id] = paint;
    }
  }
  return {
    layoutIdentity:
      typeof view.layoutIdentity === "string" ? view.layoutIdentity : null,
    dataXml: view.dataXml,
    smartArtStyle: view.smartArtStyle,
    ...(Object.keys(itemTextPaint).length ? { itemTextPaint } : {}),
    ...(nativeSmartArtHostTheme(view.hostTheme) ? { hostTheme: nativeSmartArtHostTheme(view.hostTheme)! } : {}),
    ...(() => {
      const part = (view.parts as Record<string, unknown> | undefined)?.layout;
      const sha256 = (view.fingerprints as Record<string, unknown> | undefined)?.layout;
      return typeof part === "string" &&
        /^word\/diagrams\/layout[0-9]*\.xml$/.test(part) &&
        typeof sha256 === "string" &&
        /^[0-9a-f]{64}$/.test(sha256)
        ? { layoutDefinition: { part, sha256 } }
        : {};
    })(),
  };
}

/**
 * The file's own theme lists a Docs diagram is drawn through: its view's
 * `hostTheme` -- the preserved diagram's, or the one its conversion keeps on
 * the converted group's receipt (`source.nativeSmartArtView`), so every
 * rebuild after the conversion draws the gradients and shadows it drew.
 */
export function documentSmartArtNodeHostTheme(
  attrs: unknown,
): { readonly effects?: string; readonly xml: string } | undefined {
  const source =
    attrs && typeof attrs === "object" ? (attrs as Record<string, unknown>).source : null;
  const view =
    source && typeof source === "object" && !Array.isArray(source)
      ? (source as Record<string, unknown>).nativeSmartArtView
      : null;
  return view && typeof view === "object" && !Array.isArray(view)
    ? nativeSmartArtHostTheme((view as Record<string, unknown>).hostTheme)
    : undefined;
}

/** A view's `hostTheme` (the import's reading of the host theme's style lists), or undefined. */
export function nativeSmartArtHostTheme(value: unknown): { readonly effects?: string; readonly xml: string } | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  if (typeof record.xml !== "string" || !record.xml || record.xml.length > 48_000) return undefined;
  return {
    ...(typeof record.effects === "string" && record.effects.length <= 80 ? { effects: record.effects } : {}),
    xml: record.xml,
  };
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * MRAI's SmartArt identity record as the browser reads it: the vocabulary
 * (member roles, item ids and indexes, levels, picture slots, data point ids)
 * and the reader of the JSON record a projection ships
 * (`readSmartArtIdentityRecord`). The record's OOXML spelling -- the
 * `a:extLst` / `dgm:extLst` extension the hosts write and read back -- is
 * `./smartart-identity-ext`, which runs only on the server (the engine host,
 * the exports and the importers); the browser never parses a package part.
 */

import {
  DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS,
  DOCUMENT_SMARTART_MAX_LEVEL,
  type DocumentShapeSmartArtRole,
} from "../document-shapes";
import { isSmartArtStyleId, type SmartArtStyleId } from "./styles-model";

export interface SmartArtIdentityMember {
  readonly smartArtRole: DocumentShapeSmartArtRole;
  readonly smartArtItemId?: string;
  readonly smartArtItemIndex?: number;
  readonly smartArtLevel?: number;
  /**
   * The diagram's PICTURE SLOT (`picture="1"`). Neither host's shape model
   * holds a picture — a Word group member is a `wps:wsp`, a sheet member an
   * `xdr:sp` — so a picture layout's point is written as a plain shape and only
   * this marker says what it is; without it a reimported picture layout comes
   * back as a diagram of anonymous boxes. Only ever on a `node`, and optional
   * everywhere: a host that projects no picture slots never sets it.
   */
  readonly smartArtPicturePlaceholder?: true;
  /**
   * The `modelId` of the DATA POINT this member records (`pointId`).
   *
   * Only a record written into a diagram's own data part carries it — a host's
   * shape-group record has no data model to point into — and only since the
   * writer started keying its records by it. It is what lets a record survive
   * Office EDITING the diagram: an outline item's id IS its point's `modelId`
   * (`./native-outline`), so a member can be resolved to the item it belongs to
   * by name instead of by position, and a point Office inserted in the MIDDLE
   * no longer shifts every later member onto the wrong item.
   *
   * Absent means the record predates this and stays positional, all or nothing:
   * a file MRAI already wrote must not start reopening differently.
   */
  readonly smartArtPointId?: string;
}

/**
 * The longest `pointId` this vocabulary will read back. A data model's
 * `modelId` is a GUID in every file Office writes and a derived GUID-shaped
 * digest in every file MRAI writes; the bound is the one the slide's own
 * structure receipts already hold point ids to, so one spelling is refused
 * everywhere rather than in one reader.
 */
export const MAX_SMARTART_POINT_ID_LENGTH = 128;

export interface SmartArtIdentity {
  readonly kind: string;
  readonly groupId: string;
  /**
   * Which of the fourteen SmartArt Styles the diagram wears
   * (`docs/office-graphics/SMARTART-STYLES.md` §1). Optional, and the only
   * attribute a host may omit: ordinary shapes carry a style's paint but
   * cannot SAY which style produced it, and nothing may derive a style from
   * how a member happens to look. A record without it reopens as the default.
   */
  readonly smartArtStyle?: SmartArtStyleId;
  readonly members: readonly SmartArtIdentityMember[];
}

/**
 * A `pointId` this vocabulary accepts, or `null` when the record does not carry
 * one. `undefined` for a stamp it cannot prove — an empty or over-long id is a
 * half-written record, and the whole identity then fails closed like every
 * other unprovable attribute here.
 */
export function smartArtPointIdAttribute(value: string | null | undefined): string | null | undefined {
  if (value === null || value === undefined) return null;
  if (typeof value !== "string" || !value || value.length > MAX_SMARTART_POINT_ID_LENGTH) {
    return undefined;
  }
  return value;
}

/**
 * The identity record as a projection can SHIP it, read back — the record's own
 * attribute vocabulary in JSON, one entry per member.
 *
 * For the host whose projection cannot hand the browser the diagram's data
 * part: a slide's preserved-diagram view is repeated on every projected member,
 * and the data part is a whole XML document, where this is the same handful of
 * strings the record already spells. `smart_art_identity_record` in
 * `office_smartart_identity_ext.py` writes it, from the SAME reader the server
 * uses on the part itself, so the browser reads a record it never re-parses and
 * the two hosts that DO hold the part keep reading it from the part.
 *
 * Fails closed on anything it cannot prove, exactly like the XML reader: the
 * caller then converts by the layout urn alone.
 */
export function readSmartArtIdentityRecord(
  value: unknown,
): SmartArtIdentity | null {
  const record = value as Record<string, unknown> | null;
  if (!record || typeof record !== "object" || Array.isArray(record)) return null;
  const { kind, groupId, style, members } = record;
  if (
    typeof kind !== "string" ||
    !kind ||
    typeof groupId !== "string" ||
    !groupId ||
    (style !== undefined && !isSmartArtStyleId(style)) ||
    !Array.isArray(members) ||
    members.length === 0 ||
    members.length > DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS
  ) {
    return null;
  }
  const read: SmartArtIdentityMember[] = [];
  for (const entry of members) {
    const member = entry as Record<string, unknown> | null;
    if (!member || typeof member !== "object" || Array.isArray(member)) {
      return null;
    }
    const role = member.role;
    if (role !== "node" && role !== "label" && role !== "connector") return null;
    const picture = member.picture;
    if (picture !== undefined && (picture !== true || role !== "node")) {
      return null;
    }
    const itemId = member.itemId;
    const itemIndex = member.itemIndex;
    const level = member.level;
    // A record's `pointId` is the data point it is ABOUT. A connector is
    // MRAI's own drawing and belongs to no point, so one there is a stamp this
    // reader cannot prove — exactly like an item on a connector.
    const pointId = smartArtPointIdAttribute(
      member.pointId === undefined ? null : (member.pointId as string),
    );
    if (pointId === undefined) return null;
    if (role === "connector") {
      if (
        itemId !== undefined ||
        itemIndex !== undefined ||
        level !== undefined ||
        pointId !== null
      ) {
        return null;
      }
      read.push({ smartArtRole: role });
      continue;
    }
    if (
      typeof itemId !== "string" ||
      !itemId ||
      !Number.isInteger(itemIndex) ||
      (itemIndex as number) < 0 ||
      (itemIndex as number) >= DOCUMENT_SMARTART_MAX_IDENTITY_ITEMS ||
      !Number.isInteger(level) ||
      (level as number) < 0 ||
      (level as number) > DOCUMENT_SMARTART_MAX_LEVEL
    ) {
      return null;
    }
    read.push({
      smartArtRole: role,
      smartArtItemId: itemId,
      smartArtItemIndex: itemIndex as number,
      smartArtLevel: level as number,
      ...(picture === true ? { smartArtPicturePlaceholder: true as const } : {}),
      ...(pointId === null ? {} : { smartArtPointId: pointId }),
    });
  }
  return {
    kind,
    groupId,
    ...(style === undefined ? {} : { smartArtStyle: style as SmartArtStyleId }),
    members: read,
  };
}

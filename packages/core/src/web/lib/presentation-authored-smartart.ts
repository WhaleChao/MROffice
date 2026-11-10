/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import { readSmartArtOutline } from "./presentation-smartart";

const record = (value: unknown): Record<string, unknown> | null => value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;

/** A source-verified MRAI author group can use its own whole-group writer.
 * Ordinary native groups and DiagramML do not gain per-shape editing here. */
export function isVerifiedAuthoredSmartArtGroup(blocks: readonly PresentationBlock[]): boolean {
  if (!blocks.length || blocks.length > 256 || new Set(blocks.map(b => b.id)).size !== blocks.length) return false;
  const groupId = blocks[0].groupId;
  if (typeof groupId !== "string" || !groupId || blocks.some(b => b.groupId !== groupId)) return false;
  const first = record(record(blocks[0].source)?.nativeAuthoredSmartArt);
  if (!first || first.schemaVersion !== "pptx-authored-smartart-v1" ||
      Object.keys(first).sort().join(",") !== "drawingSha256,groupShapeId,schemaVersion,slidePart" ||
      !/^ppt\/slides\/slide[1-9][0-9]*\.xml$/.test(String(first.slidePart)) ||
      !/^shp[1-9][0-9]*$/.test(String(first.groupShapeId)) || !/^[a-f0-9]{64}$/.test(String(first.drawingSha256))) return false;
  const groupAuthority = record(record(blocks[0].source)?.nativeGroupAuthority);
  if (!Array.isArray(groupAuthority?.members) || groupAuthority.members.length !== blocks.length) return false;
  const expectedIds = groupAuthority.members.map(member => record(member)?.shapeId);
  const actualIds = blocks.map(block => record(block.source)?.shapeId);
  if (new Set(expectedIds).size !== blocks.length || expectedIds.some(id => typeof id !== "string" || !actualIds.includes(id))) return false;
  for (const block of blocks) {
    const source = record(block.source), receipt = record(source?.nativeAuthoredSmartArt), authority = record(source?.nativeGroupAuthority);
    if (!receipt || !authority || Object.keys(first).some(key => first[key] !== receipt[key]) ||
        Object.keys(receipt).length !== 4 || authority.groupShapeId !== first.groupShapeId || authority.slidePart !== first.slidePart) return false;
  }
  return readSmartArtOutline(blocks, groupId) !== null;
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Commands for source-owned native SmartArt; never regenerate ordinary shapes. */
import type { NativeSmartArtGroup } from "@mrnote/core/web/lib/presentation-native-smartart.ts";

type NativeParts = Record<"data" | "layout" | "style" | "colors" | "drawing", string>;
type Authority = { schemaVersion: "pptx-smartart-edit-v2"; parts: NativeParts; fingerprints: NativeParts };
type Identity = { slidePart: string; frameShapeId: string; drawingAuthority: Authority };
export interface NativeSmartArtFlatStructureReceipt extends Identity {
  schemaVersion: "pptx-smartart-structure-v1";
  points: Array<{ modelId: string; text: string }>;
  minPoints: 2;
  maxPoints: 20;
  /** Receipts written before Move Up / Move Down name the first two only. */
  operations: ["insert-after", "delete"] | ["insert-after", "delete", "move-up", "move-down"];
}
export interface NativeSmartArtHierarchyReceipt extends Identity {
  schemaVersion: "pptx-smartart-hierarchy-v1";
  points: Array<{ modelId: string; text: string; level: number; canPromote: boolean; canDemote: boolean }>;
  operations: ["promote", "demote"];
}
export type NativeSmartArtStructureReceipt = NativeSmartArtFlatStructureReceipt | NativeSmartArtHierarchyReceipt;
export type NativeSmartArtStructureOperation = "insert-after" | "delete" | "move-up" | "move-down" | "promote" | "demote";
type FlatCommand =
  | { operation: "insert-after"; pointId: string; newPointId: string; text: string }
  | { operation: "delete" | "move-up" | "move-down"; pointId: string };
type HierarchyCommand = { operation: "promote" | "demote"; pointId: string };
export type NativeSmartArtStructureCommand = FlatCommand | HierarchyCommand;
export type NativeSmartArtStructureRequest = Identity & (
  | ({ schemaVersion: "pptx-smartart-structure-v1" } & FlatCommand)
  | ({ schemaVersion: "pptx-smartart-hierarchy-v1" } & HierarchyCommand)
);
const record = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;
const pointId = (value: unknown): value is string => typeof value === "string" && /^[A-Za-z0-9{}_-]{1,128}$/.test(value);
const textValid = (value: unknown): value is string => typeof value === "string" && value.length <= 2000 &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/u.test(value) &&
  !Array.from(value).some(c => c.length === 1 && /[\uD800-\uDFFF]/u.test(c));
const parts = ["data", "layout", "style", "colors", "drawing"] as const;

function hierarchyCan(points: NativeSmartArtHierarchyReceipt["points"], operation: "promote" | "demote", index: number): boolean {
  const level = points[index]?.level;
  if (index < 0 || typeof level !== "number" || !Number.isInteger(level) || level < 1 || level > 12) return false;
  if (operation === "promote") return level >= 2;
  let sibling = false;
  for (let i = index - 1; i >= 0; i--) {
    if (points[i].level < level) break;
    if (points[i].level === level) { sibling = true; break; }
  }
  // Demote moves the node alone (its children keep their level), as the
  // server's writer does: only the node itself must stay within 12.
  return sibling && level < 12;
}

/** Availability consumes a validated profile and never enables another schema's writer. */
export function canNativeSmartArtStructureOperation(
  profile: NativeSmartArtStructureReceipt | null | undefined,
  operation: NativeSmartArtStructureOperation,
  selectedPointId: string | null | undefined,
): boolean {
  if (!profile || !pointId(selectedPointId)) return false;
  const index = profile.points.findIndex(p => p.modelId === selectedPointId);
  if (index < 0) return false;
  if (profile.schemaVersion === "pptx-smartart-structure-v1") {
    // Move Up / Move Down swap the point with its neighbour, as the Text
    // Pane's order is the diagram's: the first cannot rise, the last fall.
    const moves = (profile.operations as readonly string[]).includes(operation);
    if (operation === "move-up") return moves && index > 0;
    if (operation === "move-down") return moves && index < profile.points.length - 1;
    return operation === "insert-after" ? profile.points.length < profile.maxPoints :
      operation === "delete" && profile.points.length > profile.minPoints;
  }
  if (profile.schemaVersion !== "pptx-smartart-hierarchy-v1" || (operation !== "promote" && operation !== "demote")) return false;
  const point = profile.points[index];
  return (operation === "promote" ? point.canPromote === true : point.canDemote === true) && hierarchyCan(profile.points, operation, index);
}

export function readNativeSmartArtStructure(group: NativeSmartArtGroup | null, ephemeralProfile?: unknown): NativeSmartArtStructureReceipt | null {
  if (!group || !group.members.length) return null;
  const source = record(group.members[0].source),
    receipt = record(ephemeralProfile === undefined ? source?.nativeSmartArtStructure : ephemeralProfile);
  const flat = receipt?.schemaVersion === "pptx-smartart-structure-v1";
  const hierarchy = receipt?.schemaVersion === "pptx-smartart-hierarchy-v1";
  if (!receipt || (!flat && !hierarchy) ||
    Object.keys(receipt).some(k => !["schemaVersion", "slidePart", "frameShapeId", "drawingAuthority", "points", "operations", ...(flat ? ["minPoints", "maxPoints"] : [])].includes(k)) ||
    typeof receipt.slidePart !== "string" || !/^ppt\/slides\/slide[1-9][0-9]*\.xml$/.test(receipt.slidePart) ||
    typeof receipt.frameShapeId !== "string" || !/^shp[1-9][0-9]*$/.test(receipt.frameShapeId) ||
    receipt.slidePart !== group.receipt.slidePart || receipt.frameShapeId !== group.receipt.frameShapeId ||
    flat && (receipt.minPoints !== 2 || receipt.maxPoints !== 20) ||
    !(flat
      ? [["insert-after", "delete"], ["insert-after", "delete", "move-up", "move-down"]]
      : [["promote", "demote"]]
    ).some((operations) => JSON.stringify(receipt.operations) === JSON.stringify(operations))) return null;
  const authority = record(receipt.drawingAuthority), paths = record(authority?.parts), hashes = record(authority?.fingerprints);
  const groupPaths = record(group.receipt.parts), groupHashes = record(group.receipt.fingerprints);
  if (!authority || authority.schemaVersion !== "pptx-smartart-edit-v2" || Object.keys(authority).some(k => !["schemaVersion", "parts", "fingerprints"].includes(k)) ||
    !paths || !hashes || Object.keys(paths).length !== 5 || Object.keys(hashes).length !== 5 ||
    parts.some(k => typeof paths[k] !== "string" || !/^ppt\/diagrams\/[A-Za-z0-9_-]+\.xml$/.test(paths[k]) ||
      typeof hashes[k] !== "string" || !/^[a-f0-9]{64}$/.test(hashes[k]) ||
      paths[k] !== groupPaths?.[k] || hashes[k] !== groupHashes?.[k]) || new Set(Object.values(paths)).size !== 5) return null;
  if (!Array.isArray(receipt.points) || receipt.points.length < 2 || receipt.points.length > 20 || receipt.points.length !== group.outline.length) return null;
  const seen = new Set<string>();
  for (let index = 0; index < receipt.points.length; index++) {
    const point = record(receipt.points[index]), live = group.outline[index];
    if (!point || Object.keys(point).some(k => !["modelId", "text", ...(hierarchy ? ["level", "canPromote", "canDemote"] : [])].includes(k)) || !pointId(point.modelId) ||
      !textValid(point.text) || seen.has(point.modelId) || live.id !== point.modelId || live.text !== point.text) return null;
    if (flat && live.level !== 0) return null;
    if (hierarchy && (typeof point.level !== "number" || !Number.isInteger(point.level) || point.level !== live.level ||
      point.level < (index === 0 ? 0 : 1) || point.level > 12 || index === 0 && point.level !== 0 ||
      index > 0 && point.level > group.outline[index - 1].level + 1 ||
      typeof point.canPromote !== "boolean" || typeof point.canDemote !== "boolean")) return null;
    seen.add(point.modelId);
  }
  if (hierarchy) {
    const points = receipt.points as NativeSmartArtHierarchyReceipt["points"];
    if (points.some((p, i) => p.canPromote && !hierarchyCan(points, "promote", i) || p.canDemote && !hierarchyCan(points, "demote", i))) return null;
  }
  if (ephemeralProfile === undefined && group.members.some(member => JSON.stringify(record(member.source)?.nativeSmartArtStructure) !== JSON.stringify(receipt))) return null;
  return receipt as unknown as NativeSmartArtStructureReceipt;
}

/** Flush text before structural requests; the server owns the new projection. */
export function buildNativeSmartArtStructureRequest(
  group: NativeSmartArtGroup | null,
  command: NativeSmartArtStructureCommand,
  ephemeralProfile?: unknown,
): NativeSmartArtStructureRequest | null {
  const receipt = readNativeSmartArtStructure(group, ephemeralProfile);
  if (!receipt || !canNativeSmartArtStructureOperation(receipt, command.operation, command.pointId)) return null;
  const identity = { slidePart: receipt.slidePart, frameShapeId: receipt.frameShapeId,
    drawingAuthority: structuredClone(receipt.drawingAuthority) };
  if (receipt.schemaVersion === "pptx-smartart-hierarchy-v1") {
    if (command.operation !== "promote" && command.operation !== "demote") return null;
    return { ...identity, schemaVersion: receipt.schemaVersion, operation: command.operation, pointId: command.pointId };
  }
  if (command.operation === "insert-after") {
    if (!pointId(command.newPointId) || receipt.points.some(p => p.modelId === command.newPointId) || !textValid(command.text)) return null;
    return { ...identity, schemaVersion: receipt.schemaVersion, operation: command.operation,
      pointId: command.pointId, newPointId: command.newPointId, text: command.text };
  }
  return command.operation === "delete" || command.operation === "move-up" || command.operation === "move-down"
    ? { ...identity, schemaVersion: receipt.schemaVersion, operation: command.operation, pointId: command.pointId }
    : null;
}

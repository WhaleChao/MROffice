/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PRESERVED sheet SmartArt object, read: its members (the cached drawing),
 * its text points and its data part. Pure; shared by the worksheet editor
 * (`./native-smartart`) and the server's conversion (`./native-smartart-convert`).
 */

/** Cap on members one preserved diagram may draw (mirrors the zod schema). */
export const MAX_NATIVE_SMARTART_MEMBERS = 400;

/** Cap on one member's text, the same bound the patch model enforces. */
export const MAX_NATIVE_SMARTART_TEXT_LENGTH = 2_000;

type UnknownRecord = Record<string, unknown>;

function record(value: unknown): UnknownRecord | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as UnknownRecord)
    : null;
}

function modelId(value: unknown): value is string {
  return typeof value === "string" && /^[A-Za-z0-9{}_-]{1,128}$/.test(value);
}

/** One drawn member of a preserved diagram. */
export interface NativeSmartArtMember {
  memberId: string;
  role: "shape" | "text";
  zIndex: number;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
  preset?: string;
  customPath?: UnknownRecord;
  lineEnds?: UnknownRecord;
  paintOpacity?: UnknownRecord;
  strokeWidthEmu?: number;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  text?: string;
  textStyle?: UnknownRecord;
  /** Data points this member's text is bound to; exactly one is editable. */
  dataPointIds: string[];
  /**
   * The `pres` point its `dsp:sp` draws. A shape member and a text member
   * with the same one are ONE shape on the canvas (Office draws them as one
   * `dsp:sp`), selected, moved and resized together.
   */
  modelId?: string;
}

/** One editable text point of a preserved diagram. */
export interface NativeSmartArtTextPoint {
  modelId: string;
  text: string;
  memberId: string;
}

/** A preserved sheet diagram, validated. */
export interface SpreadsheetNativeSmartArt {
  objectId: string;
  layoutIdentity: string;
  widthEmu: number;
  heightEmu: number;
  members: NativeSmartArtMember[];
  /** Text points the receipt admits, in drawing order. */
  points: NativeSmartArtTextPoint[];
  /** The data part, which the conversion reads the outline from. */
  dataXml: string;
  /** False when the projection could not build an edit receipt. */
  textEditable: boolean;
}

export function finiteOr(value: unknown, fallback = 0): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function normalizeMember(
  value: unknown,
  index: number,
): NativeSmartArtMember | null {
  const raw = record(value);
  if (!raw) return null;
  const role = raw.role === "text" ? "text" : "shape";
  const dataPointIds = Array.isArray(raw.dataPointIds)
    ? raw.dataPointIds.filter(modelId).slice(0, 64)
    : [];
  return {
    memberId:
      typeof raw.memberId === "string" && raw.memberId
        ? raw.memberId
        : `m${index}`,
    role,
    zIndex: finiteOr(raw.zIndex, index),
    left: finiteOr(raw.left),
    top: finiteOr(raw.top),
    width: Math.max(0, finiteOr(raw.width)),
    height: Math.max(0, finiteOr(raw.height)),
    rotation: finiteOr(raw.rotation),
    ...(typeof raw.preset === "string" ? { preset: raw.preset } : {}),
    ...(record(raw.customPath) ? { customPath: record(raw.customPath)! } : {}),
    ...(record(raw.lineEnds) ? { lineEnds: record(raw.lineEnds)! } : {}),
    ...(record(raw.paintOpacity)
      ? { paintOpacity: record(raw.paintOpacity)! }
      : {}),
    ...(typeof raw.strokeWidthEmu === "number"
      ? { strokeWidthEmu: raw.strokeWidthEmu }
      : {}),
    ...(typeof raw.fill === "string" ? { fill: raw.fill } : {}),
    ...(typeof raw.stroke === "string" ? { stroke: raw.stroke } : {}),
    ...(typeof raw.strokeWidth === "number"
      ? { strokeWidth: raw.strokeWidth }
      : {}),
    ...(role === "text"
      ? { text: typeof raw.text === "string" ? raw.text : "" }
      : {}),
    ...(record(raw.textStyle) ? { textStyle: record(raw.textStyle)! } : {}),
    dataPointIds,
    ...(modelId(raw.modelId) ? { modelId: raw.modelId } : {}),
  };
}

/** True for a sheet object that is a preserved native SmartArt diagram. */
export function isSpreadsheetNativeSmartArtObject(
  object: unknown,
): object is UnknownRecord {
  const raw = record(object);
  return (
    !!raw &&
    raw.type === "smartart" &&
    record(raw.source)?.nodeType === "nativeSmartArtGroup"
  );
}

/**
 * Read one preserved diagram, or `null` when the object is not a complete,
 * self-consistent native SmartArt projection. Fail-closed: a diagram this
 * reader cannot prove is drawn as nothing rather than as a guess.
 */
export function readSpreadsheetNativeSmartArt(
  object: unknown,
): SpreadsheetNativeSmartArt | null {
  const raw = record(object);
  if (!isSpreadsheetNativeSmartArtObject(raw)) return null;
  const source = record(raw.source);
  const view = record(source?.nativeSmartArtView);
  const receipt = record(source?.nativeSmartArtEdit);
  const rawMembers = Array.isArray(raw.members) ? raw.members : null;
  if (
    !view ||
    view.schemaVersion !== "office-smartart-native-v1" ||
    typeof view.layoutIdentity !== "string" ||
    typeof view.dataXml !== "string" ||
    !rawMembers ||
    rawMembers.length === 0 ||
    rawMembers.length > MAX_NATIVE_SMARTART_MEMBERS ||
    typeof raw.widthEmu !== "number" ||
    typeof raw.heightEmu !== "number" ||
    raw.widthEmu < 1 ||
    raw.heightEmu < 1
  ) {
    return null;
  }
  const members: NativeSmartArtMember[] = [];
  for (const [index, entry] of rawMembers.entries()) {
    const member = normalizeMember(entry, index);
    if (!member) return null;
    members.push(member);
  }
  members.sort((left, right) => left.zIndex - right.zIndex);
  // Only the points the RECEIPT names can be written back, and only a member
  // bound to exactly one point owns an editable string: a member covering
  // several points is one native text body the writer cannot split.
  const admitted = new Set<string>();
  const receiptPoints = Array.isArray(receipt?.points) ? receipt!.points : [];
  for (const entry of receiptPoints) {
    const point = record(entry);
    if (point && modelId(point.modelId)) admitted.add(point.modelId);
  }
  const points: NativeSmartArtTextPoint[] = [];
  const seen = new Set<string>();
  for (const member of members) {
    if (member.role !== "text" || member.dataPointIds.length !== 1) continue;
    const id = member.dataPointIds[0];
    if (!admitted.has(id) || seen.has(id)) continue;
    seen.add(id);
    points.push({ modelId: id, text: member.text ?? "", memberId: member.memberId });
  }
  return {
    objectId: String(raw.id ?? ""),
    layoutIdentity: view.layoutIdentity,
    widthEmu: raw.widthEmu,
    heightEmu: raw.heightEmu,
    members,
    points,
    dataXml: view.dataXml,
    textEditable: view.textEditable === true && points.length > 0,
  };
}

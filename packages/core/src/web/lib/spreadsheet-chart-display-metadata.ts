/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet } from "./api";
import { getPage } from "./notebook-sdk";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";
import { normalizeChartVisibility, type ChartVisibility } from "./spreadsheet/chart-visibility";

type Json = Record<string, unknown>;
const record = (raw: unknown): Json | undefined => raw && typeof raw === "object" && !Array.isArray(raw) ? raw as Json : undefined;
const digest = (raw: unknown): raw is string => typeof raw === "string" && /^[a-f0-9]{64}$/.test(raw);
export interface ChartDisplayMetadata {
  pageId: string;
  version: number;
  contentSha256: string;
  source: { resourceId: string; resourceVersionId: string | null; packageSha256: string };
  charts: { sheetId: string; objectId: string; source: Json; objectSha256: string; chartVisibility: ChartVisibility; savedObject: Json }[];
}

/** The server digest binds its verified snapshot; a second ordinary page read
 * checks that the exact canonical snapshot captured by this client is current.
 * No reserialization/hash dialect is substituted for a stored content hash. */
export function readChartDisplayMetadata(raw: unknown, pageId: string, version: number, savedContent: unknown, readback: Awaited<ReturnType<typeof getPage>>): ChartDisplayMetadata | undefined {
  const value = record(raw), source = record(value?.source), content = record(savedContent);
  if (!value || !source || value.state !== "resolved" || value.page_id !== pageId || value.content_version !== version || !digest(value.content_sha256) || !digest(source.package_sha256) || typeof source.resource_id !== "string" || !source.resource_id || (source.resource_version_id !== null && typeof source.resource_version_id !== "string") || !Array.isArray(value.charts) || value.charts.length > 200 || !Array.isArray(content?.sheets) || readback.id !== pageId || readback.content_version !== version || !equal(readback.content_json, savedContent)) return undefined;
  const charts: ChartDisplayMetadata["charts"] = [];
  for (const rawEntry of value.charts) {
    const entry = record(rawEntry), binding = record(entry?.source), policy = normalizeChartVisibility(entry?.chartVisibility);
    if (!entry || !binding || !policy || !digest(entry.object_sha256) || typeof entry.sheet_id !== "string" || typeof entry.object_id !== "string" || !["part", "drawingPart", "relationshipId"].every(key => typeof binding[key] === "string" && binding[key])) return undefined;
    const owners = content.sheets.map(record).filter(sheet => sheet?.id === entry.sheet_id);
    const objects = owners.length === 1 && Array.isArray(owners[0]?.objects) ? owners[0].objects.map(record).filter(obj => obj?.id === entry.object_id) : [];
    const object = objects[0], originalSource = record(object?.source);
    if (objects.length !== 1 || !object || !originalSource || "chartVisibility" in object || !["part", "drawingPart", "relationshipId"].every(key => originalSource[key] === binding[key])) return undefined;
    if (charts.some(chart => chart.sheetId === entry.sheet_id && chart.objectId === entry.object_id)) return undefined;
    charts.push({ sheetId: entry.sheet_id, objectId: entry.object_id, source: binding, objectSha256: entry.object_sha256, chartVisibility: policy, savedObject: object });
  }
  return { pageId, version, contentSha256: value.content_sha256, source: { resourceId: source.resource_id, resourceVersionId: source.resource_version_id as string | null, packageSha256: source.package_sha256 }, charts };
}

export async function loadChartDisplayMetadata(pageId: string, version: number, savedContent: unknown, signal: AbortSignal): Promise<ChartDisplayMetadata | undefined> {
  const [metadata, readback] = await Promise.all([
    apiGet<unknown>(`/api/v1/pages/${encodeURIComponent(pageId)}/chart-display-metadata`, { signal }),
    getPage(pageId, { signal }),
  ]);
  return signal.aborted ? undefined : readChartDisplayMetadata(metadata, pageId, version, savedContent, readback);
}

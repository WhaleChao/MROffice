"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useSyncExternalStore } from "react";

import { apiDelete, apiGet, apiPatch, apiPost, isApiRequestError } from "../api";
import { subscribeAuthState } from "../auth-state";
import type { ChartBlockAttrsV2 } from "./catalog";
import {
  MAX_CHART_TEMPLATE_NAME_LENGTH,
  MAX_OFFICE_CHART_TEMPLATES,
  officeChartTemplateFromSavedChart,
  type OfficeChartTemplate,
} from "./chart-templates";

/**
 * The user's Templates folder for charts (Office keeps `.crtx` files in the
 * user's own Templates folder, offered by Insert Chart and Change Chart Type
 * in every application): `/api/v1/chart-templates`. One copy per tab, shared
 * by every Docs, Sheets and Slides window, read when a dialog first needs it.
 * Manage Templates renames, deletes, downloads (`.crtx`) and imports here.
 */

const URL = "/api/v1/chart-templates";

export interface OfficeChartTemplateFolder {
  readonly status: "idle" | "loading" | "ready" | "failed";
  readonly templates: readonly OfficeChartTemplate[];
  readonly limit: number;
}

export type OfficeChartTemplateSaveResult = "ok" | "limit" | "tooLarge" | "invalid" | "failed";

const EMPTY: OfficeChartTemplateFolder = { status: "idle", templates: [], limit: MAX_OFFICE_CHART_TEMPLATES };
let folder: OfficeChartTemplateFolder = EMPTY;
let generation = 0;
const listeners = new Set<() => void>();

function emit(next: OfficeChartTemplateFolder): void {
  folder = next;
  for (const listener of listeners) listener();
}

interface SavedChartWire {
  id?: unknown;
  name?: unknown;
  attrs?: unknown;
}

function templateFromWire(value: unknown): OfficeChartTemplate | null {
  if (!value || typeof value !== "object") return null;
  const wire = value as SavedChartWire;
  if (typeof wire.id !== "string" || !wire.id || typeof wire.name !== "string") return null;
  return officeChartTemplateFromSavedChart({ id: wire.id, name: wire.name, attrs: wire.attrs });
}

/** The folder as the server answers it; junk entries are skipped. */
export function officeChartTemplatesFromWire(raw: unknown): OfficeChartTemplate[] {
  const list = raw && typeof raw === "object" ? (raw as { templates?: unknown }).templates : null;
  if (!Array.isArray(list)) return [];
  const out: OfficeChartTemplate[] = [];
  const ids = new Set<string>();
  for (const entry of list) {
    const template = templateFromWire(entry);
    if (!template || ids.has(template.id)) continue;
    ids.add(template.id);
    out.push(template);
  }
  return out;
}

/** Reads the folder (again when `force`, or after a failed read). */
export async function loadOfficeChartTemplates(force = false): Promise<void> {
  if (folder.status === "loading" || (folder.status === "ready" && !force)) return;
  const asked = generation;
  emit({ ...folder, status: "loading" });
  try {
    const raw = await apiGet<unknown>(URL, { cache: "no-store" });
    if (asked !== generation) return;
    const limit = (raw as { limit?: unknown } | null)?.limit;
    emit({
      status: "ready",
      templates: officeChartTemplatesFromWire(raw),
      limit: typeof limit === "number" && limit >= 1 ? Math.floor(limit) : MAX_OFFICE_CHART_TEMPLATES,
    });
  } catch {
    if (asked === generation) emit({ ...folder, status: "failed" });
  }
}

/** Save as Template: the chart, named, into the folder. */
export async function saveOfficeChartTemplate(
  name: string,
  attrs: ChartBlockAttrsV2,
): Promise<OfficeChartTemplateSaveResult> {
  try {
    const raw = await apiPost<unknown>(URL, { name: name.trim(), attrs });
    const template = templateFromWire(raw);
    if (!template) return "failed";
    if (folder.status === "ready") {
      emit({ ...folder, templates: [...folder.templates.filter((entry) => entry.id !== template.id), template] });
    } else {
      void loadOfficeChartTemplates(true);
    }
    return "ok";
  } catch (error) {
    if (isApiRequestError(error)) {
      if (error.code === "chart_template_limit" || error.status === 409) return "limit";
      if (error.code === "chart_template_too_large" || error.status === 413) return "tooLarge";
      if (error.code === "chart_template_invalid") return "invalid";
    }
    return "failed";
  }
}

/** Removes one template from the folder; false when the server refused. */
export async function deleteOfficeChartTemplate(id: string): Promise<boolean> {
  try {
    await apiDelete<unknown>(`${URL}/${encodeURIComponent(id)}`);
  } catch (error) {
    // Already gone (another device): the folder drops it all the same.
    if (!(isApiRequestError(error) && error.status === 404)) return false;
  }
  emit({ ...folder, templates: folder.templates.filter((entry) => entry.id !== id) });
  return true;
}

/** Manage Templates > Rename; false when the server refused. */
export async function renameOfficeChartTemplate(id: string, name: string): Promise<boolean> {
  const clean = name.trim();
  if (!clean) return false;
  try {
    const raw = await apiPatch<unknown>(`${URL}/${encodeURIComponent(id)}`, { name: clean });
    const template = templateFromWire(raw);
    if (!template) return false;
    emit({ ...folder, templates: folder.templates.map((entry) => (entry.id === id ? template : entry)) });
    return true;
  } catch {
    return false;
  }
}

interface CrtxFileWire {
  filename?: unknown;
  media_type?: unknown;
  content_base64?: unknown;
}

/** Manage Templates > Download: the template as Excel's `.crtx` file. */
export async function downloadOfficeChartTemplate(id: string): Promise<boolean> {
  try {
    const [raw, { downloadBlobFile }] = await Promise.all([
      apiGet<CrtxFileWire>(`${URL}/${encodeURIComponent(id)}/crtx`, { cache: "no-store" }),
      import("../download-file"),
    ]);
    if (typeof raw?.filename !== "string" || typeof raw.content_base64 !== "string") return false;
    const binary = atob(raw.content_base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    const type = typeof raw.media_type === "string" ? raw.media_type : "application/vnd.ms-excel.chart-template";
    return downloadBlobFile(raw.filename, new Blob([bytes], { type }));
  } catch {
    return false;
  }
}

export type OfficeChartTemplateImportResult = OfficeChartTemplateSaveResult | "unreadable";

/** The largest `.crtx` the import sends (the API's bound). */
export const MAX_OFFICE_CHART_TEMPLATE_FILE_BYTES = 2 * 1024 * 1024;

function base64FromBytes(bytes: Uint8Array): string {
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary);
}

/**
 * Manage Templates > Import: a `.crtx` file joins the folder, named after
 * the file as Office lists a template by its file name.
 */
export async function importOfficeChartTemplate(file: File): Promise<OfficeChartTemplateImportResult> {
  if (file.size > MAX_OFFICE_CHART_TEMPLATE_FILE_BYTES) return "tooLarge";
  const name = file.name.replace(/\.crtx$/i, "").trim() || "Chart template";
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const raw = await apiPost<unknown>(`${URL}/import`, {
      name: name.slice(0, MAX_CHART_TEMPLATE_NAME_LENGTH),
      content_base64: base64FromBytes(bytes),
    });
    const template = templateFromWire(raw);
    if (!template) return "unreadable";
    if (folder.status === "ready") {
      emit({ ...folder, templates: [...folder.templates.filter((entry) => entry.id !== template.id), template] });
    } else {
      void loadOfficeChartTemplates(true);
    }
    return "ok";
  } catch (error) {
    if (isApiRequestError(error)) {
      if (error.code === "chart_template_limit" || error.status === 409) return "limit";
      if (error.code === "chart_template_too_large" || error.status === 413) return "tooLarge";
      if (error.code === "chart_template_file_invalid" || error.code === "chart_template_invalid") return "unreadable";
    }
    return "failed";
  }
}

/** A template of the folder as last read (a pick made in a dialog). */
export function officeChartTemplateById(id: string): OfficeChartTemplate | undefined {
  return folder.templates.find((entry) => entry.id === id);
}

/**
 * The folder's Manage Templates commands, as the All Charts dialog takes
 * them (`OfficeAllChartsTemplates`): the same in every editor.
 */
export const officeUserChartTemplateCommands = {
  onDelete: (template: OfficeChartTemplate) => deleteOfficeChartTemplate(template.id),
  onRename: (template: OfficeChartTemplate, name: string) => renameOfficeChartTemplate(template.id, name),
  onDownload: (template: OfficeChartTemplate) => downloadOfficeChartTemplate(template.id),
  onImport: (file: File) => importOfficeChartTemplate(file),
} as const;

if (typeof window !== "undefined") {
  // Another account in this tab: its folder is not this one's.
  subscribeAuthState(() => {
    generation += 1;
    emit(EMPTY);
  });
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** The folder, read once `active` (a dialog that shows it is open). */
export function useOfficeChartTemplates(active: boolean): OfficeChartTemplateFolder {
  const snapshot = useSyncExternalStore(subscribe, () => folder, () => EMPTY);
  useEffect(() => {
    if (active && (snapshot.status === "idle" || snapshot.status === "failed")) void loadOfficeChartTemplates();
    // Only a dialog opening (or the folder reset) asks again.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active, snapshot.status === "idle"]);
  return snapshot;
}

/** Tests: forget the folder. */
export function resetOfficeChartTemplatesForTests(): void {
  generation += 1;
  folder = EMPTY;
}

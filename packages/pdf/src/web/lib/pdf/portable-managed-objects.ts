/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PDFArray, PDFBool, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFString } from "pdf-lib";

import { parseManagedNativeTableShadow, type ManagedNativeTableRecord } from "./managed-table";
import type { PdfChart } from "./managed-chart";

const MAX_OBJECTS = 200;
const ID = /^[A-Za-z0-9._~%-]{1,378}$/;
const TOKEN = /^[A-Za-z0-9_-]{4,400000}$/;

export interface PortableManagedObjects {
  readonly tables: readonly ManagedNativeTableRecord[];
  readonly charts: readonly PdfChart[];
}

const text = (value: unknown): string | null =>
  value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : null;

function fromBase64Url(token: string): string | null {
  try {
    const binary = atob(token.replaceAll("-", "+").replaceAll("_", "/") + "=".repeat((4 - (token.length % 4)) % 4));
    return new TextDecoder("utf-8", { fatal: true }).decode(Uint8Array.from(binary, (ch) => ch.charCodeAt(0)));
  } catch {
    return null;
  }
}

function boxToken(value: number): string {
  return String(Number(Math.min(1, Math.max(0, value)).toFixed(6)));
}

/** The page's tables and charts (1-based `pageNumber`) as records; none
 * when the page holds none or the file cannot be read. */
export async function readPortableManagedObjects(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PortableManagedObjects> {
  const pdf = await PDFDocument.load(sourceBytes, { updateMetadata: false, ignoreEncryption: false });
  if (pdf.isEncrypted || !Number.isSafeInteger(pageNumber) || pageNumber < 1 || pageNumber > pdf.getPageCount()) {
    return { tables: [], charts: [] };
  }
  const page = pdf.getPage(pageNumber - 1);
  const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
  if (!annots) return { tables: [], charts: [] };
  const crop = page.getCropBox();
  const left = crop.x, bottom = crop.y, width = crop.width, height = crop.height;
  const right = left + width, top = bottom + height;
  const rotation = ((page.getRotation().angle % 360) + 360) % 360;
  if (!(width > 0 && height > 0) || ![0, 90, 180, 270].includes(rotation)) return { tables: [], charts: [] };
  const fraction = (x: number, y: number): [number, number] => {
    if (rotation === 90) return [(y - bottom) / height, (x - left) / width];
    if (rotation === 180) return [(right - x) / width, (y - bottom) / height];
    if (rotation === 270) return [(top - y) / height, (right - x) / width];
    return [(x - left) / width, (top - y) / height];
  };
  const tables: ManagedNativeTableRecord[] = [];
  const charts: PdfChart[] = [];
  const seen = new Set<string>();
  for (let index = 0; index < annots.size() && tables.length + charts.length < MAX_OBJECTS; index++) {
    const dict = pdf.context.lookup(annots.get(index));
    if (!(dict instanceof PDFDict) || dict.get(PDFName.of("MRNoteManaged")) !== PDFBool.True) continue;
    const kind = dict.get(PDFName.of("MRNoteTable")) === PDFBool.True ? "table"
      : dict.get(PDFName.of("MRNoteChart")) === PDFBool.True ? "chart" : null;
    if (!kind) continue;
    const name = text(dict.get(PDFName.of("NM")));
    const token = text(dict.get(PDFName.of(kind === "table" ? "MRNoteTableModel" : "MRNoteChartModel")));
    const rect = dict.lookupMaybe(PDFName.of("Rect"), PDFArray);
    const prefix = `MRNote:${kind}.`;
    if (!name?.startsWith(prefix) || !ID.test(name.slice(prefix.length)) || !token || !TOKEN.test(token) ||
      !rect || rect.size() !== 4) continue;
    // A page copied twice holds one object under one name twice: the first is
    // the object, the copy stays as the file draws it.
    if (seen.has(name)) continue;
    seen.add(name);
    const values = [0, 1, 2, 3].map((at) => rect.lookup(at));
    if (!values.every((value) => value instanceof PDFNumber)) continue;
    const [x0, y0, x1, y1] = values.map((value) => (value as PDFNumber).asNumber());
    const corners = [fraction(x0, y0), fraction(x1, y0), fraction(x1, y1), fraction(x0, y1)];
    const fx = Math.min(...corners.map(([x]) => x)), fy = Math.min(...corners.map(([, y]) => y));
    const fw = Math.max(...corners.map(([x]) => x)) - fx, fh = Math.max(...corners.map(([, y]) => y)) - fy;
    if (![fx, fy, fw, fh].every(Number.isFinite) || fx < -1e-6 || fy < -1e-6 || fw <= 0 || fh <= 0 ||
      fx + fw > 1 + 1e-3 || fy + fh > 1 + 1e-3) continue;
    const box = [fx, fy, Math.min(fw, 1 - fx), Math.min(fh, 1 - fy)];
    const encodedId = name.slice(prefix.length);
    if (kind === "table") {
      const record = parseManagedNativeTableShadow({
        page: pageNumber,
        text: `[[MRNOTE_NATIVE_TABLE_V1|upsert|table.${encodedId}|auto|${box.map(boxToken).join(",")}|${token}]]`,
      });
      if (record && !record.deleted) tables.push(record);
      continue;
    }
    const model = fromBase64Url(token);
    let attrs: unknown = null;
    try {
      attrs = model === null ? null : JSON.parse(model);
    } catch {
      attrs = null;
    }
    let id: string;
    try {
      id = decodeURIComponent(encodedId);
    } catch {
      continue;
    }
    if (!id || !attrs || typeof attrs !== "object" || Array.isArray(attrs)) continue;
    charts.push({
      id, page: pageNumber, x: box[0], y: box[1], width: box[2], height: box[3],
      attrs: attrs as Record<string, unknown>, svg: null,
    });
  }
  return { tables, charts };
}

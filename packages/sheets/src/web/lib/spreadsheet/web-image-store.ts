/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pictures Excel's IMAGE function names, as this browser has them.
 *
 * A picture is never loaded from its source by the page: the app's CSP admits
 * no third-party image host, and a collaborator opening a shared sheet would
 * otherwise hand that host their address. The page's server fetches it
 * (`POST /api/v1/office/sheets/web-images`, SSRF-guarded, redirects refused)
 * and answers the picture as an embedded raster, or the error Excel's IMAGE
 * shows for that source: #VALUE! (not https, not a supported picture),
 * #CONNECT! (unreachable, or anything but 200), #BLOCKED! (a non-public
 * address or a redirect).
 *
 * One store per browser tab: an entry is "loading" (the cell shows #BUSY!, as
 * Excel does while it downloads), the picture, or its error. Listeners hear
 * every settled fetch, so a window recalculates and redraws.
 */

import { MAX_CALCULATION_WEB_IMAGES, MAX_CALCULATION_WEB_IMAGE_SOURCE } from "@mrnote/core/web/lib/spreadsheet/background-calculation-protocol.ts";
import { ApiRequestError, apiPost } from "@mrnote/core/web/lib/api.ts";
import type { WebImageStatus } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { sanitizeImageSource } from "@mrnote/core/web/lib/spreadsheet/sheet-image.ts";

export const SPREADSHEET_WEB_IMAGE_PATH = "/api/v1/office/sheets/web-images";

export interface SpreadsheetWebImagePicture {
  /** An embedded raster data URL. */
  readonly src: string;
  /** The picture's own size in pixels (sizing 2 draws it at this size). */
  readonly width: number;
  readonly height: number;
}

type Entry =
  | { readonly status: "loading" }
  | { readonly status: "ready"; readonly picture: SpreadsheetWebImagePicture }
  | { readonly status: "#CONNECT!" | "#BLOCKED!" | "#VALUE!" };

type ImageError = Extract<Entry, { status: "#CONNECT!" | "#BLOCKED!" | "#VALUE!" }>["status"];

type Fetcher = (source: string) => Promise<SpreadsheetWebImagePicture | ImageError>;

/** The most sources one tab keeps; the oldest settled ones go first. */
const MAX_ENTRIES = 400;
const entries = new Map<string, Entry>();
const listeners = new Set<() => void>();
let revision = 0;

async function fetchFromServer(source: string): Promise<SpreadsheetWebImagePicture | ImageError> {
  try {
    const response = await apiPost<{ src?: unknown; width?: unknown; height?: unknown }>(
      SPREADSHEET_WEB_IMAGE_PATH,
      { url: source },
    );
    const src = sanitizeImageSource(response.src);
    const width = Number(response.width);
    const height = Number(response.height);
    if (!src || !(width > 0) || !(height > 0)) return "#VALUE!";
    return { src, width, height };
  } catch (error) {
    const details =
      error instanceof ApiRequestError
        ? (error.details as { code?: unknown; reason?: unknown; host?: unknown } | undefined)
        : undefined;
    const code = details?.code;
    if (code === "#BLOCKED!" && details?.reason === "host_not_approved" && typeof details.host === "string") {
      noteHostQuestion(source, details.host);
    }
    if (code === "#VALUE!" || code === "#BLOCKED!" || code === "#CONNECT!") return code;
    // The server itself unreachable: Excel's "check the internet connection".
    return "#CONNECT!";
  }
}

let fetcher: Fetcher = fetchFromServer;

function settle(source: string, entry: Entry): void {
  entries.delete(source);
  entries.set(source, entry);
  while (entries.size > MAX_ENTRIES) {
    const oldest = [...entries.entries()].find(([, value]) => value.status !== "loading");
    if (!oldest) break;
    entries.delete(oldest[0]);
  }
  revision += 1;
  for (const listener of [...listeners]) listener();
}

/** Start fetching a source's picture unless this tab already has (or is getting) it. */
export function requestSpreadsheetWebImage(source: string): void {
  if (entries.has(source)) return;
  entries.set(source, { status: "loading" });
  void fetcher(source).then(
    (result) =>
      settle(
        source,
        typeof result === "string" ? { status: result } : { status: "ready", picture: result },
      ),
    () => settle(source, { status: "#CONNECT!" }),
  );
}

/**
 * The engine's reading of a source ({@link WebImageStatus}); an unknown source
 * starts its fetch and reads as loading.
 */
export function spreadsheetWebImageStatus(source: string): WebImageStatus {
  const entry = entries.get(source);
  if (!entry) {
    requestSpreadsheetWebImage(source);
    return "loading";
  }
  return entry.status === "ready" ? "ready" : entry.status;
}

/** The picture of a source when this tab has it. */
export function spreadsheetWebImagePicture(source: string): SpreadsheetWebImagePicture | null {
  const entry = entries.get(source);
  return entry?.status === "ready" ? entry.picture : null;
}

/**
 * The sources this tab asked for and was refused, with Excel's IMAGE error
 * for each, in a stable order: what the calculation session is told so that
 * the cell and everything reading it take the error (IFERROR, ISERROR), not
 * only the cell's own display. A loading or shown picture is not listed: the
 * session answers those with the picture value, which is how this tab learns
 * a source to fetch in the first place. Bounded as the session bounds it.
 */
export function spreadsheetWebImageRefusals(): Array<readonly [string, ImageError]> {
  const refused: Array<readonly [string, ImageError]> = [];
  for (const [source, entry] of entries) {
    if (entry.status === "loading" || entry.status === "ready") continue;
    if (source.length > MAX_CALCULATION_WEB_IMAGE_SOURCE) continue;
    refused.push([source, entry.status]);
  }
  refused.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
  return refused.slice(0, MAX_CALCULATION_WEB_IMAGES);
}

/** Hear every settled fetch; returns the unsubscribe. */
export function subscribeSpreadsheetWebImages(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Changes whenever a fetch settles (a memo dependency for the evaluator). */
export function spreadsheetWebImageRevision(): number {
  return revision;
}

// ── Per-host consent ─────────────────────────────────────────────────────
// An IMAGE source can be computed from cells (and the AI writes formulas), so
// its URL can carry workbook text. The server answers #BLOCKED! for a host the
// member has not allowed; the sheet asks once per host per tab, as Excel asks
// before turning on access to linked images. "Allow" is recorded on the server
// (per member + host); "Don't allow" lasts for this tab.

export const SPREADSHEET_WEB_IMAGE_HOSTS_PATH = "/api/v1/office/sheets/web-images/hosts";

/** Host -> the refused sources that name it. */
const hostSources = new Map<string, Set<string>>();
const deniedHosts = new Set<string>();
let hostQuestions: readonly string[] = [];

function noteHostQuestion(source: string, host: string): void {
  const sources = hostSources.get(host) ?? new Set<string>();
  sources.add(source);
  hostSources.set(host, sources);
  if (!deniedHosts.has(host) && !hostQuestions.includes(host)) hostQuestions = [...hostQuestions, host];
}

/** The hosts this tab is waiting for the member to allow or refuse, in order. */
export function spreadsheetWebImageHostQuestions(): readonly string[] {
  return hostQuestions;
}

function dropQuestion(host: string): void {
  hostQuestions = hostQuestions.filter((candidate) => candidate !== host);
}

/** The member allows pictures from ``host``: record it, then fetch its sources again. */
export async function allowSpreadsheetWebImageHost(host: string): Promise<void> {
  await apiPost(SPREADSHEET_WEB_IMAGE_HOSTS_PATH, { host });
  dropQuestion(host);
  for (const source of hostSources.get(host) ?? []) entries.delete(source);
  hostSources.delete(host);
  revision += 1;
  for (const listener of [...listeners]) listener();
}

/** The member refuses ``host`` for this tab: its cells keep #BLOCKED!. */
export function denySpreadsheetWebImageHost(host: string): void {
  deniedHosts.add(host);
  dropQuestion(host);
  revision += 1;
  for (const listener of [...listeners]) listener();
}

/**
 * A printed workbook's pictures (the print page holds no session to ask the
 * server with): the server fetched the IMAGE sources it may, under the
 * printing member's approved hosts, and hands each picture or its error in;
 * any other source reads as Excel's #BLOCKED!. Forgets every entry.
 */
export function answerSpreadsheetWebImagesFrom(
  answers: Readonly<Record<string, SpreadsheetWebImagePicture | ImageError>>,
): void {
  entries.clear();
  revision += 1;
  fetcher = async (source) => {
    const answer = Object.prototype.hasOwnProperty.call(answers, source) ? answers[source] : undefined;
    if (typeof answer === "string") return answer === "#CONNECT!" || answer === "#VALUE!" ? answer : "#BLOCKED!";
    const src = sanitizeImageSource(answer?.src);
    return src && Number(answer?.width) > 0 && Number(answer?.height) > 0
      ? { src, width: Number(answer!.width), height: Number(answer!.height) }
      : "#BLOCKED!";
  };
}

/** Tests: replace the server fetch and forget every entry. */
export function resetSpreadsheetWebImagesForTests(next?: Fetcher): void {
  entries.clear();
  hostSources.clear();
  deniedHosts.clear();
  hostQuestions = [];
  revision = 0;
  fetcher = next ?? fetchFromServer;
}

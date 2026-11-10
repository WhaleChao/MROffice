"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * LibreOffice in MRNote Classic (docs/desktop-classic.md). Old Office
 * formats (.doc/.xls/.ppt), OpenDocument and RTF files, EMF/WMF pictures and
 * OpenDocument exports need it; the app ships none. Where the API answers
 * `libreoffice_required`, the page offers "Download and enable": the app's
 * shell fetches the official release into its own folder
 * (`window.mraiDesktop.libreOffice`, apps/desktop src/classic/libreoffice.ts),
 * shows the progress itself, and the API uses the copy from its next
 * conversion on, without a restart; the operation that failed is then tried
 * again.
 */

import { useCallback, useEffect, useState } from "react";

import { apiGet, isApiRequestError } from "./api";
import { IS_DESKTOP_CLASSIC } from "./desktop-classic";

export type LibreOfficePhase = "space" | "download" | "verify" | "unpack" | "check" | "done";

/** What the shell offers and holds (src/classic/libreoffice.ts). */
export interface DesktopLibreOfficeStatus {
  offered: { version: string; downloadBytes: number } | null;
  installed: { version: string; bytes: number; installedAt: string } | null;
  installing: { phase: LibreOfficePhase; received: number; total: number } | null;
  space: { need: number; free: number | null } | null;
}

export type DesktopLibreOfficeOutcome = { ok: true; version: string } | { ok: false; cancelled?: boolean; error?: string };

/** Which LibreOffice the API converts with now (GET /api/v1/desktop/libreoffice). */
export interface LibreOfficeInUse {
  available: boolean;
  source: "app" | "installed" | "path" | null;
  path: string | null;
}

interface LibreOfficeBridge {
  status: () => Promise<DesktopLibreOfficeStatus | null>;
  install: () => Promise<DesktopLibreOfficeOutcome>;
  remove: () => Promise<{ ok: boolean; error?: string }>;
  onChange: (listener: () => void) => () => void;
}

/** The shell's bridge, in MRNote Classic builds that have it. */
export function desktopLibreOffice(): LibreOfficeBridge | null {
  if (!IS_DESKTOP_CLASSIC || typeof window === "undefined") return null;
  const bridge = (window as unknown as { mraiDesktop?: { classic?: boolean; libreOffice?: LibreOfficeBridge } }).mraiDesktop;
  return bridge?.classic && bridge.libreOffice?.install ? bridge.libreOffice : null;
}

/** The API said this needs LibreOffice: the error, or what it wraps, carries
 * its code (an ApiRequestError, or an import's SourceImportError). */
export function isLibreOfficeRequired(error: unknown): boolean {
  const seen = new Set<unknown>();
  let current: unknown = error;
  while (current && typeof current === "object" && !seen.has(current)) {
    seen.add(current);
    const code = isApiRequestError(current) ? current.code : (current as { code?: unknown }).code;
    if (code === "libreoffice_required") return true;
    current = (current as { cause?: unknown }).cause;
  }
  return false;
}

/** The app can fetch it for this failure: Classic, a shell that can, and the API asked for it. */
export function canEnableLibreOfficeFor(error: unknown): boolean {
  return desktopLibreOffice() !== null && isLibreOfficeRequired(error);
}

let running: Promise<DesktopLibreOfficeOutcome> | null = null;

/** Download and enable (one at a time; a second ask joins the first). */
export function enableDesktopLibreOffice(): Promise<DesktopLibreOfficeOutcome> {
  const bridge = desktopLibreOffice();
  if (!bridge) return Promise.resolve({ ok: false });
  if (!running) {
    running = bridge
      .install()
      .catch((error: unknown) => ({ ok: false as const, error: error instanceof Error ? error.message : String(error) }))
      .finally(() => {
        running = null;
      });
  }
  return running;
}

export async function libreOfficeInUse(): Promise<LibreOfficeInUse | null> {
  if (!IS_DESKTOP_CLASSIC) return null;
  try {
    return await apiGet<LibreOfficeInUse>("/api/v1/desktop/libreoffice");
  } catch {
    return null;
  }
}

/** "300 MB": the download's size, rounded as people say it. */
export function formatDownloadSize(bytes: number | null | undefined): string {
  const value = bytes && bytes > 0 ? bytes : 300e6;
  if (value >= 1e9) return `${(value / 1e9).toFixed(1)} GB`;
  return `${Math.round(value / 1e7) * 10} MB`;
}

/** The shell's status and the API's answer, kept current while mounted. */
export function useDesktopLibreOffice(): {
  supported: boolean;
  status: DesktopLibreOfficeStatus | null;
  inUse: LibreOfficeInUse | null;
  refresh: () => Promise<void>;
} {
  const [status, setStatus] = useState<DesktopLibreOfficeStatus | null>(null);
  const [inUse, setInUse] = useState<LibreOfficeInUse | null>(null);
  const supported = desktopLibreOffice() !== null;
  const refresh = useCallback(async () => {
    const bridge = desktopLibreOffice();
    if (!bridge) return;
    const [shell, api] = await Promise.all([bridge.status().catch(() => null), libreOfficeInUse()]);
    setStatus(shell);
    setInUse(api);
  }, []);
  useEffect(() => {
    const bridge = desktopLibreOffice();
    if (!bridge) return;
    void refresh();
    return bridge.onChange(() => void refresh());
  }, [refresh]);
  return { supported, status, inUse, refresh };
}

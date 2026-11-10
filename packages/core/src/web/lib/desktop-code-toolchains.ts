"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * MRNote Classic (docs/desktop-classic.md) runs compiled languages and LaTeX
 * on the toolchains installed on the person's own computer (the local API's
 * app/desktop/code_runtime.py). When one is missing, the Code window says
 * what to install instead of "the runtime is unavailable"
 * (`GET /api/v1/desktop/code-toolchains`).
 */

import { useEffect, useState } from "react";

import { apiGet } from "./api";
import { IS_DESKTOP_CLASSIC } from "./desktop-classic";

interface ToolchainRow {
  readonly language: string;
  readonly available: boolean;
  readonly install: string;
}

let cached: Promise<Map<string, ToolchainRow>> | null = null;
let cachedAt = 0;
/** A toolchain installed while the app runs is found on the next look. */
const STALE_MS = 30_000;

function load(): Promise<Map<string, ToolchainRow>> {
  if (!cached || Date.now() - cachedAt > STALE_MS) {
    cachedAt = Date.now();
    cached = apiGet<{ languages: ToolchainRow[] }>("/api/v1/desktop/code-toolchains")
      .then((answer) => new Map(answer.languages.map((row) => [row.language, row])))
      .catch(() => new Map<string, ToolchainRow>());
  }
  return cached;
}

/** What to install to run `language` here, or null (not the desktop app,
 * not a compiled language, or already installed). */
export function useDesktopToolchainInstallHint(language: string, unavailable: boolean): string | null {
  const [hint, setHint] = useState<{ language: string; text: string | null } | null>(null);
  useEffect(() => {
    if (!IS_DESKTOP_CLASSIC || !unavailable) return;
    let active = true;
    void load().then((rows) => {
      const row = rows.get(language);
      if (active) setHint({ language, text: row && !row.available && row.install ? row.install : null });
    });
    return () => {
      active = false;
    };
  }, [language, unavailable]);
  if (!IS_DESKTOP_CLASSIC || !unavailable || hint?.language !== language) return null;
  return hint.text;
}

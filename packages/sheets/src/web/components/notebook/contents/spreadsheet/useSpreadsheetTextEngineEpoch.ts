"use client";
/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";
import { onTextEngineChange } from "@mrnote/core/web/lib/text-engine/measure.ts";
import { spreadsheetTextEngineEpoch, spreadsheetTextEngineOn } from "@mrnote/core/web/lib/spreadsheet/spreadsheet-text-engine.ts";

const subscribe = (listener: () => void) => (spreadsheetTextEngineOn() ? onTextEngineChange(listener) : () => {});

/**
 * Renders the grid again when the text engine's faces land, so every cell
 * measured while they were loading (spill, "####", Shrink to Fit) is decided
 * again from the engine -- the same decision in every browser, whichever
 * order the faces arrived in.
 */
export function useSpreadsheetTextEngineEpoch(): string {
  return useSyncExternalStore(subscribe, spreadsheetTextEngineEpoch, () => "");
}

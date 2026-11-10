"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { createContext, useContext } from "react";

import type { NotebookInteriorMode } from "../../lib/notebook-interior-mode";

/**
 * The notebook's interior design as the notebook route decided it
 * (NotebookInteriorMode.tsx holds the provider). Its own module, so the
 * canvas, the tray and the sidebar read it without the provider's imports.
 */
export interface NotebookInteriorState {
  /** What the interior shows now: the resolved mode, or until it is known
   * the mode this notebook last showed here (the workbench the first time). */
  readonly mode: NotebookInteriorMode;
  /** The mode is decided (plan and ownership are known). */
  readonly resolved: boolean;
  /** This person may switch this notebook between the designs. */
  readonly canSwitch: boolean;
  /** Asking for the workbench here opens the upgrade panel. */
  readonly workbenchNeedsUpgrade: boolean;
  

  readonly workbenchNeedsLargerScreen: boolean;
  /** Ask for a design: switches it (and remembers the choice), or explains
   * the upgrade when the workbench is not in this person's plan, in a panel
   * beside `anchor` (the switch that asked). */
  readonly requestMode: (mode: NotebookInteriorMode, anchor?: HTMLElement | null) => void;
  /** This is the person's own notebook and their plan holds only them:
   * sharing it opens the upgrade panel (the server refuses it anyway). */
  readonly sharingNeedsUpgrade: boolean;
  /** Someone else's notebook whose owner's plan no longer includes sharing:
   * this member's editing role reads only (the read-only chrome says why). */
  readonly memberAccessCapped: boolean;
  /** Explain a paid feature in the floating panel beside `anchor`. */
  readonly requestUpgrade: (reason: NotebookUpgradeReason, anchor?: HTMLElement | null) => void;
}

/** What a free plan asked for: the workbench design, or sharing. */
/** What the floating panel explains: a paid feature, or (on a phone) that the
 * workbench needs a larger screen. */
export type NotebookUpgradeReason = "workbench" | "share" | "largerScreen";

export const WORKBENCH_ONLY: NotebookInteriorState = {
  mode: "workbench",
  resolved: true,
  canSwitch: false,
  workbenchNeedsUpgrade: false,
  workbenchNeedsLargerScreen: false,
  requestMode: () => {},
  sharingNeedsUpgrade: false,
  memberAccessCapped: false,
  requestUpgrade: () => {},
};

export const NotebookInteriorContext = createContext<NotebookInteriorState>(WORKBENCH_ONLY);

/** The interior design of the notebook this surface is in; the workbench
 * outside a notebook route (previews, the dashboard). */
export function useNotebookInterior(): NotebookInteriorState {
  return useContext(NotebookInteriorContext);
}


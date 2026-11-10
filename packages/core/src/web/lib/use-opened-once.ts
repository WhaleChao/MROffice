"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useState } from "react";

/**
 * Whether a dialog mounted closed with its editor should render: not
 * until it is first opened, and from then on (for its closing animation
 * and the fields it keeps). An editor renders its closed dialogs again
 * with every change of the window; one never opened then builds nothing.
 */
export function useOpenedOnce(open: boolean): boolean {
  const [everOpened, setEverOpened] = useState(open);
  if (open && !everOpened) setEverOpened(true);
  return open || everOpened;
}

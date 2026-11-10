/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { KeyboardEvent } from "react";

/**
 * Returns an `onKeyDown` handler that fires `handler` on Enter or Space.
 *
 * Use together with `role="button"` + `tabIndex={0}` to make a non-button
 * element (a clickable row / card / list item) keyboard-operable, mirroring
 * its `onClick`. This matches the established WorkPacketsListWindow row
 * pattern and fixes WCAG 2.1.1 (Keyboard) on mouse-only selection surfaces.
 */
export function onActivateKey(handler: () => void) {
  return (event: KeyboardEvent) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      handler();
    }
  };
}

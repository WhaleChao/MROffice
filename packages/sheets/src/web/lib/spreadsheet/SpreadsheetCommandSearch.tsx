"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet's name for the shared title-bar command search
 * (components/office-ribbon/OfficeCommandSearch): the same palette every
 * editor's Search button (Alt+Q) opens, with the worksheet's own sentences
 * and test ids as its defaults.
 */
export {
  OfficeCommandSearch as SpreadsheetCommandSearch,
  type OfficeCommandSearchItem as SpreadsheetCommandSearchItem,
} from "@mrnote/core/web/components/office-ribbon/OfficeCommandSearch.tsx";

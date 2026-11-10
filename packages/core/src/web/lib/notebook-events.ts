/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/notebook-events.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export const DECK_PART_READY_EVENT: any = "mrn:deck-part-ready";
export type DECK_PART_READY_EVENT = any;
export const DECK_PART_ROOM_NOT_SAVED_EVENT: any = "mrn:deck-part-room-not-saved";
export type DECK_PART_ROOM_NOT_SAVED_EVENT = any;
export const DECK_SLIDES_CHANGED_EVENT: any = "mrn:deck-slides-changed";
export type DECK_SLIDES_CHANGED_EVENT = any;
export const NOTEBOOK_PAGES_CHANGED_EVENT: any = "mrnote:notebook-pages-changed";
export type NOTEBOOK_PAGES_CHANGED_EVENT = any;
export const NOTEBOOK_PAGE_RENAMED_EVENT: any = "mrnote:notebook-page-renamed";
export type NOTEBOOK_PAGE_RENAMED_EVENT = any;
export const PAGE_PICTURE_NOT_DRAWN_EVENT: any = "mrn:page-picture-not-drawn";
export type PAGE_PICTURE_NOT_DRAWN_EVENT = any;
export const PAGE_REMOTE_SAVED_EVENT: any = "mrn:page-remote-saved";
export type PAGE_REMOTE_SAVED_EVENT = any;
export const dispatchNotebookPagesChanged: any = inertFunction("lib/notebook-events.ts#dispatchNotebookPagesChanged", "void");
export type dispatchNotebookPagesChanged = any;
export const dispatchNotebookProposalsConfirmed: any = inertFunction("lib/notebook-events.ts#dispatchNotebookProposalsConfirmed", "void");
export type dispatchNotebookProposalsConfirmed = any;
export const dispatchNotebookStorageChanged: any = inertFunction("lib/notebook-events.ts#dispatchNotebookStorageChanged", "void");
export type dispatchNotebookStorageChanged = any;
export const dispatchNotebooksChanged: any = inertFunction("lib/notebook-events.ts#dispatchNotebooksChanged", "void");
export type dispatchNotebooksChanged = any;
export const pageRenameFromEvent: any = inertFunction("lib/notebook-events.ts#pageRenameFromEvent", "null");
export type pageRenameFromEvent = any;

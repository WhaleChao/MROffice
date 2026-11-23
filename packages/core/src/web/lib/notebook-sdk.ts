/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/notebook-sdk.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export type DeckPartRoomUnsaved = any;
export type DocxPublishedReadbackRead = any;
export type EditablePagePatchback = any;
export type NotebookInfo = any;
export type PageExportFidelityRead = any;
export type PageExportNoticeRead = any;
export type PageItem = any;
export type PageSaveHoldItem = any;
export type PageType = any;
export type PageVersionItem = any;
export type PdfPatchbackIncompleteMarker = any;
export type PdfPatchbackRepairAuth = any;
export type PdfRestoreDurableContentRequest = any;
export const createNotebook: any = inertFunction("lib/notebook-sdk.ts#createNotebook", "promise");
export type createNotebook = any;
export const createPage: any = inertFunction("lib/notebook-sdk.ts#createPage", "promise");
export type createPage = any;
export const createPageSnapshot: any = inertFunction("lib/notebook-sdk.ts#createPageSnapshot", "promise");
export type createPageSnapshot = any;
export const deletePage: any = inertFunction("lib/notebook-sdk.ts#deletePage", "promise-void");
export type deletePage = any;
export const deletePageVersion: any = inertFunction("lib/notebook-sdk.ts#deletePageVersion", "promise-void");
export type deletePageVersion = any;
export const downloadPageSaveHoldPptx: any = inertFunction("lib/notebook-sdk.ts#downloadPageSaveHoldPptx", "promise");
export type downloadPageSaveHoldPptx = any;
export const getPage: any = inertFunction("lib/notebook-sdk.ts#getPage", "promise");
export type getPage = any;
export const getPageFontBindings: any = inertFunction("lib/notebook-sdk.ts#getPageFontBindings", "promise");
export type getPageFontBindings = any;
export const getPageVersion: any = inertFunction("lib/notebook-sdk.ts#getPageVersion", "promise");
export type getPageVersion = any;
export const listNotebooks: any = inertFunction("lib/notebook-sdk.ts#listNotebooks", "promise");
export type listNotebooks = any;
export const listPageSaveHolds: any = inertFunction("lib/notebook-sdk.ts#listPageSaveHolds", "promise");
export type listPageSaveHolds = any;
export const listPageVersions: any = inertFunction("lib/notebook-sdk.ts#listPageVersions", "promise");
export type listPageVersions = any;
export const listPages: any = inertFunction("lib/notebook-sdk.ts#listPages", "promise");
export type listPages = any;
export const notebookSDK: any = inertValue("lib/notebook-sdk.ts#notebookSDK");
export type notebookSDK = any;
export const readNotebookPages: any = inertFunction("lib/notebook-sdk.ts#readNotebookPages", "promise");
export type readNotebookPages = any;
export const readPageSaveHold: any = inertFunction("lib/notebook-sdk.ts#readPageSaveHold", "promise");
export type readPageSaveHold = any;
export const restorePageVersion: any = inertFunction("lib/notebook-sdk.ts#restorePageVersion", "promise");
export type restorePageVersion = any;
export const settledEditorPageRead: any = inertFunction("lib/notebook-sdk.ts#settledEditorPageRead", "null");
export type settledEditorPageRead = any;
export const updatePage: any = inertFunction("lib/notebook-sdk.ts#updatePage", "promise");
export type updatePage = any;

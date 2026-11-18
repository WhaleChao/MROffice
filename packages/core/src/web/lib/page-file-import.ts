/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/page-file-import.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export type AttachmentImportDescriptor = any;
export const IPYNB_PAGE_IMPORT_RETIRED_MESSAGE: any = "Jupyter notebook (.ipynb) import is no longer supported. Import a .py file into a Code page instead.";
export type IPYNB_PAGE_IMPORT_RETIRED_MESSAGE = any;
export const LONG_DECK_SOURCE_ROAD_SLIDES: any = 12;
export type LONG_DECK_SOURCE_ROAD_SLIDES = any;
export const LONG_WORD_SOURCE_ROAD_XML_BYTES: any = inertValue("lib/page-file-import.ts#LONG_WORD_SOURCE_ROAD_XML_BYTES");
export type LONG_WORD_SOURCE_ROAD_XML_BYTES = any;
export const PAGE_IMPORT_BATCH_MANIFEST_SAFE_BYTES: any = inertValue("lib/page-file-import.ts#PAGE_IMPORT_BATCH_MANIFEST_SAFE_BYTES");
export type PAGE_IMPORT_BATCH_MANIFEST_SAFE_BYTES = any;
export const PAGE_IMPORT_BATCH_MAX_FILES: any = 20;
export type PAGE_IMPORT_BATCH_MAX_FILES = any;
export const PAGE_IMPORT_BATCH_MAX_TOTAL_BYTES: any = inertValue("lib/page-file-import.ts#PAGE_IMPORT_BATCH_MAX_TOTAL_BYTES");
export type PAGE_IMPORT_BATCH_MAX_TOTAL_BYTES = any;
export const PAGE_IMPORT_SELECTION_MAX_FILES: any = 200;
export type PAGE_IMPORT_SELECTION_MAX_FILES = any;
export type PageFileImportReceipt = any;
export type PageFileImportWarning = any;
export type PageImportPreflightOptions = any;
export const assertPageImportSelectionBounds: any = inertFunction("lib/page-file-import.ts#assertPageImportSelectionBounds", "void");
export type assertPageImportSelectionBounds = any;
export const assertPageImportSelectionCount: any = inertFunction("lib/page-file-import.ts#assertPageImportSelectionCount", "void");
export type assertPageImportSelectionCount = any;
export const checkOpenedFilesAhead: any = inertFunction("lib/page-file-import.ts#checkOpenedFilesAhead", "void");
export type checkOpenedFilesAhead = any;
export const createPageFromAttachment: any = inertFunction("lib/page-file-import.ts#createPageFromAttachment", "promise");
export type createPageFromAttachment = any;
export const createPageFromFile: any = inertFunction("lib/page-file-import.ts#createPageFromFile", "promise");
export type createPageFromFile = any;
export const createPageFromFileWithReceipt: any = inertFunction("lib/page-file-import.ts#createPageFromFileWithReceipt", "promise");
export type createPageFromFileWithReceipt = any;
export const createPageImportReceipts: any = inertFunction("lib/page-file-import.ts#createPageImportReceipts", "promise");
export type createPageImportReceipts = any;
export const createPagesFromFilesAtomically: any = inertFunction("lib/page-file-import.ts#createPagesFromFilesAtomically", "promise");
export type createPagesFromFilesAtomically = any;
export const createPdfPageFromFile: any = inertFunction("lib/page-file-import.ts#createPdfPageFromFile", "promise");
export type createPdfPageFromFile = any;
export const getPreparedEditableDocument: any = inertFunction("lib/page-file-import.ts#getPreparedEditableDocument", "promise");
export type getPreparedEditableDocument = any;
export const isNativeSourceImportFile: any = inertFunction("lib/page-file-import.ts#isNativeSourceImportFile", "boolean");
export type isNativeSourceImportFile = any;
export const isPdfPasswordRequiredImportError: any = inertValue("lib/page-file-import.ts#isPdfPasswordRequiredImportError");
export type isPdfPasswordRequiredImportError = any;
export const isWordXmlHead: any = inertFunction("lib/page-file-import.ts#isWordXmlHead", "boolean");
export type isWordXmlHead = any;
export const needsSourceImport: any = inertFunction("lib/page-file-import.ts#needsSourceImport", "boolean");
export type needsSourceImport = any;
export const noteWordXmlFiles: any = inertFunction("lib/page-file-import.ts#noteWordXmlFiles", "promise-void");
export type noteWordXmlFiles = any;
export const partitionPageImportFilesIntoAtomicUnits: any = inertFunction("lib/page-file-import.ts#partitionPageImportFilesIntoAtomicUnits", "array");
export type partitionPageImportFilesIntoAtomicUnits = any;
export const precheckImportFile: any = inertFunction("lib/page-file-import.ts#precheckImportFile", "boolean");
export type precheckImportFile = any;
export const predictedImportedPage: any = inertFunction("lib/page-file-import.ts#predictedImportedPage", "unknown");
export type predictedImportedPage = any;
export const preflightPageImportFiles: any = inertFunction("lib/page-file-import.ts#preflightPageImportFiles", "promise-void");
export type preflightPageImportFiles = any;
export const presentationSlideCount: any = inertFunction("lib/page-file-import.ts#presentationSlideCount", "number");
export type presentationSlideCount = any;
export const settleSourceImportRoad: any = inertFunction("lib/page-file-import.ts#settleSourceImportRoad", "promise");
export type settleSourceImportRoad = any;

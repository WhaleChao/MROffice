/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/resources-sdk.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export type PdfConversionExportFormat = any;
export type PdfConversionJob = any;
export type PdfConversionJobError = any;
export type ResourceRead = any;
export type ResourceVersionRead = any;
export const cancelPdfConversionJob: any = inertFunction("lib/resources-sdk.ts#cancelPdfConversionJob", "promise");
export type cancelPdfConversionJob = any;
export const createPdfConversionJob: any = inertFunction("lib/resources-sdk.ts#createPdfConversionJob", "promise");
export type createPdfConversionJob = any;
export const createPdfProjection: any = inertFunction("lib/resources-sdk.ts#createPdfProjection", "promise");
export type createPdfProjection = any;
export const getPdfConversionJob: any = inertFunction("lib/resources-sdk.ts#getPdfConversionJob", "promise");
export type getPdfConversionJob = any;
export const getResource: any = inertFunction("lib/resources-sdk.ts#getResource", "promise");
export type getResource = any;
export const getResourceEditablePageSourceSha256: any = inertFunction("lib/resources-sdk.ts#getResourceEditablePageSourceSha256", "promise");
export type getResourceEditablePageSourceSha256 = any;
export const getResourceVersionById: any = inertFunction("lib/resources-sdk.ts#getResourceVersionById", "promise");
export type getResourceVersionById = any;
export const getResourceVersionDownload: any = inertFunction("lib/resources-sdk.ts#getResourceVersionDownload", "promise");
export type getResourceVersionDownload = any;
export const getResourceVersions: any = inertFunction("lib/resources-sdk.ts#getResourceVersions", "promise");
export type getResourceVersions = any;
export const listPdfConversionJobs: any = inertFunction("lib/resources-sdk.ts#listPdfConversionJobs", "promise");
export type listPdfConversionJobs = any;
export const listResources: any = inertFunction("lib/resources-sdk.ts#listResources", "promise");
export type listResources = any;
export const listResourcesAll: any = inertFunction("lib/resources-sdk.ts#listResourcesAll", "promise");
export type listResourcesAll = any;
export const preparePdfConversion: any = inertFunction("lib/resources-sdk.ts#preparePdfConversion", "promise");
export type preparePdfConversion = any;

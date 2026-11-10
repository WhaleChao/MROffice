/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/api.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export const AUTH_SESSION_EXPIRED_EVENT: any = "mrnote:auth-session-expired";
export type AUTH_SESSION_EXPIRED_EVENT = any;
export const AUTH_SESSION_HELD_EVENT: any = "mrnote:auth-session-held";
export type AUTH_SESSION_HELD_EVENT = any;
export const AUTH_SESSION_RESTORED_EVENT: any = "mrnote:auth-session-restored";
export type AUTH_SESSION_RESTORED_EVENT = any;
export const ApiRequestError: any = inertClass("lib/api.ts#ApiRequestError", true);
export type ApiRequestError = any;
export type LocalApiAnswerer = any;
export const accessGateErrorFor: any = inertFunction("lib/api.ts#accessGateErrorFor", "null");
export type accessGateErrorFor = any;
export const activeWorkspaceHeaders: any = inertFunction("lib/api.ts#activeWorkspaceHeaders", "unknown");
export type activeWorkspaceHeaders = any;
export const answerApiRequestLocally: any = inertFunction("lib/api.ts#answerApiRequestLocally", "promise");
export type answerApiRequestLocally = any;
export const apiDelete: any = inertFunction("lib/api.ts#apiDelete", "promise");
export type apiDelete = any;
export const apiFetchRaw: any = inertFunction("lib/api.ts#apiFetchRaw", "promise");
export type apiFetchRaw = any;
export const apiGet: any = inertFunction("lib/api.ts#apiGet", "promise");
export type apiGet = any;
export const apiPatch: any = inertFunction("lib/api.ts#apiPatch", "promise");
export type apiPatch = any;
export const apiPost: any = inertFunction("lib/api.ts#apiPost", "promise");
export type apiPost = any;
export const apiPostFormData: any = inertFunction("lib/api.ts#apiPostFormData", "promise");
export type apiPostFormData = any;
export const apiPut: any = inertFunction("lib/api.ts#apiPut", "promise");
export type apiPut = any;
export const buildPresignedUploadInit: any = inertFunction("lib/api.ts#buildPresignedUploadInit", "unknown");
export type buildPresignedUploadInit = any;
export const buildStreamPostHeaders: any = inertFunction("lib/api.ts#buildStreamPostHeaders", "promise");
export type buildStreamPostHeaders = any;
export const clearWorkspaceId: any = inertFunction("lib/api.ts#clearWorkspaceId", "void");
export type clearWorkspaceId = any;
export const confirmAuthSessionRestored: any = inertFunction("lib/api.ts#confirmAuthSessionRestored", "promise");
export type confirmAuthSessionRestored = any;
export const createApiResponseError: any = inertFunction("lib/api.ts#createApiResponseError", "unknown");
export type createApiResponseError = any;
export const fetchWorkspaceScoped: any = inertFunction("lib/api.ts#fetchWorkspaceScoped", "promise");
export type fetchWorkspaceScoped = any;
export const handleStreamUnauthorized: any = inertFunction("lib/api.ts#handleStreamUnauthorized", "void");
export type handleStreamUnauthorized = any;
export const invalidateCsrfCache: any = inertFunction("lib/api.ts#invalidateCsrfCache", "void");
export type invalidateCsrfCache = any;
export const isApiRequestError: any = inertFunction("lib/api.ts#isApiRequestError", "boolean");
export type isApiRequestError = any;
export const isAuthSessionHeld: any = inertFunction("lib/api.ts#isAuthSessionHeld", "boolean");
export type isAuthSessionHeld = any;
export const lastSessionRefreshAt: any = inertFunction("lib/api.ts#lastSessionRefreshAt", "number");
export type lastSessionRefreshAt = any;
export const localizedApiErrorCodeMessage: any = inertFunction("lib/api.ts#localizedApiErrorCodeMessage", "null");
export type localizedApiErrorCodeMessage = any;
export const logout: any = inertFunction("lib/api.ts#logout", "promise");
export type logout = any;
export const notebookServiceWorkspaceId: any = inertFunction("lib/api.ts#notebookServiceWorkspaceId", "null");
export type notebookServiceWorkspaceId = any;
export const persistWorkspaceId: any = inertFunction("lib/api.ts#persistWorkspaceId", "void");
export type persistWorkspaceId = any;
export const personalWorkspaceHeaders: any = inertFunction("lib/api.ts#personalWorkspaceHeaders", "unknown");
export type personalWorkspaceHeaders = any;
export const personalWorkspaceId: any = inertFunction("lib/api.ts#personalWorkspaceId", "null");
export type personalWorkspaceId = any;
export const preparedJsonBytesOf: any = inertFunction("lib/api.ts#preparedJsonBytesOf", "undefined");
export type preparedJsonBytesOf = any;
export const readAvailableApiCsrfToken: any = inertFunction("lib/api.ts#readAvailableApiCsrfToken", "null");
export type readAvailableApiCsrfToken = any;
export const refreshAuthSession: any = inertFunction("lib/api.ts#refreshAuthSession", "promise");
export type refreshAuthSession = any;
export const registerLocalApiAnswerer: any = inertFunction("lib/api.ts#registerLocalApiAnswerer", "unsubscribe");
export type registerLocalApiAnswerer = any;
export const registerUnsavedWorkHold: any = inertFunction("lib/api.ts#registerUnsavedWorkHold", "unsubscribe");
export type registerUnsavedWorkHold = any;
export const resolveApiWorkspaceId: any = inertFunction("lib/api.ts#resolveApiWorkspaceId", "null");
export type resolveApiWorkspaceId = any;
export const uploadToPresignedUrl: any = inertFunction("lib/api.ts#uploadToPresignedUrl", "promise");
export type uploadToPresignedUrl = any;
export const withPreparedJsonBytes: any = inertFunction("lib/api.ts#withPreparedJsonBytes", "unknown");
export type withPreparedJsonBytes = any;

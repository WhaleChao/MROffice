/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/collab/live-transform.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export const LIVE_TRANSFORM_FIELD: any = "xf";
export type LIVE_TRANSFORM_FIELD = any;
export type LiveTransformFrame = any;
export type LiveTransformObject = any;
export type LiveTransformPublisher = any;
export const createLiveTransformPublisher: any = inertFunction("lib/collab/live-transform.ts#createLiveTransformPublisher", "unknown");
export type createLiveTransformPublisher = any;
export const createLiveTransformTracker: any = inertFunction("lib/collab/live-transform.ts#createLiveTransformTracker", "unsubscribe");
export type createLiveTransformTracker = any;
export const decodeLiveTransform: any = inertFunction("lib/collab/live-transform.ts#decodeLiveTransform", "null");
export type decodeLiveTransform = any;
export const dropLiveTransform: any = inertFunction("lib/collab/live-transform.ts#dropLiveTransform", "void");
export type dropLiveTransform = any;
export const liveTransformKey: any = inertFunction("lib/collab/live-transform.ts#liveTransformKey", "string");
export type liveTransformKey = any;
export const liveTransforms: any = inertValue("lib/collab/live-transform.ts#liveTransforms");
export type liveTransforms = any;
export const noteCommittedTransform: any = inertFunction("lib/collab/live-transform.ts#noteCommittedTransform", "void");
export type noteCommittedTransform = any;
export const shortestTurn: any = inertFunction("lib/collab/live-transform.ts#shortestTurn", "number");
export type shortestTurn = any;
export const subscribeLiveTransformSurface: any = inertFunction("lib/collab/live-transform.ts#subscribeLiveTransformSurface", "unsubscribe");
export type subscribeLiveTransformSurface = any;
export const useLiveTransform: any = inertHook("lib/collab/live-transform.ts#useLiveTransform", "null");
export type useLiveTransform = any;
export const useLiveTransformTracker: any = inertHook("lib/collab/live-transform.ts#useLiveTransformTracker", "void");
export type useLiveTransformTracker = any;

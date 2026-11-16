/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// This module (MR Note's AI) is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export type AiReviewChange = any;
export type AiReviewDecision = any;
export type AiReviewGroup = any;
export type AiReviewKeyKind = any;
export const OPEN_AI_REVIEW_EVENT: any = "mrnote:open-ai-review";
export type OPEN_AI_REVIEW_EVENT = any;
export type OpenAiReviewDetail = any;
export const aiEditReviewEnabled: any = inertFunction("cut#aiEditReviewEnabled", "boolean");
export type aiEditReviewEnabled = any;
export const decideAiEditReview: any = inertFunction("cut#decideAiEditReview", "promise");
export type decideAiEditReview = any;
export const fetchAiEditReview: any = inertFunction("cut#fetchAiEditReview", "promise");
export type fetchAiEditReview = any;
export const takePendingAiReviewOpen: any = inertFunction("cut#takePendingAiReviewOpen", "null");
export type takePendingAiReviewOpen = any;

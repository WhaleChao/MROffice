/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink to Math / Ink to Text, the modes' reader: the strokes of one
 * piece of writing as one PNG (the Math Input Control's own rasteriser) sent
 * to the server, which answers LaTeX (`/ink-math`) or the written words
 * (`/ink-text`, `apps/api/app/routers/ink_math.py`). The server reads the
 * picture with its built-in recognizers (as Office reads ink with its own;
 * nothing is sent on, no consent to ask) unless it is configured to use the
 * workspace's AI vision model instead, which the capability answer says
 * (`engine`, `needsAcknowledgement`). One capability answer covers both modes.
 *
 * Ink to Text names the writer's language (BCP 47), as Office reads ink in
 * the editing language: the server starts with that language's recognizer
 * (Korean, Thai, the Latin languages' accents) and still tries the others
 * when it is unsure, so writing in another language reads too.
 */
import { apiGet, apiPostFormData } from "./api";
import { ensureDesktopOndemand } from "./desktop-ondemand";
import {
  rasterizeInkMathStrokes,
  recognizeInkMath,
  type InkMathCapability,
  type InkMathPoint,
} from "./office-ink-math";
import {
  officeInkReadingAlternatives,
  usableOfficeInkReading,
  type OfficeInkRecognitionMode,
} from "./office-ink-conversion";

export interface InkTextRecognition {
  readonly text: string;
  readonly alternatives: readonly string[];
  readonly confidence: number;
  readonly modelId: string;
  readonly policyVersion: number;
  readonly policyMarkerSha256: string;
}

export async function recognizeInkText(
  pageId: string,
  image: Blob,
  options: {
    readonly privacyAcknowledged: boolean;
    readonly signal?: AbortSignal;
    /** The writer's language (BCP 47). */
    readonly language?: string | null;
  },
): Promise<InkTextRecognition> {
  const form = new FormData();
  form.append("file", image, "ink-text.png");
  form.append("privacy_acknowledged", options.privacyAcknowledged ? "true" : "false");
  if (options.language) form.append("language", options.language);
  return apiPostFormData<InkTextRecognition>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/ink-text/recognize`,
    form,
    { signal: options.signal },
  );
}

const CAPABILITY_TTL_MS = 60_000;
const capabilityCache = new Map<string, { at: number; answer: Promise<InkMathCapability> }>();

const UNAVAILABLE: InkMathCapability = Object.freeze({
  available: false,
  reasonCode: "unavailable",
  needsAcknowledgement: false,
  modelId: null,
  residencyRegion: null,
  maxImageBytes: 0,
  acceptedMediaTypes: [],
});

/**
 * Whether the page's workspace lets handwriting be read (its vision
 * authority), asked once a minute per page however often a mode is toggled.
 */
export function loadOfficeInkRecognitionCapability(pageId: string): Promise<InkMathCapability> {
  const now = Date.now();
  const cached = capabilityCache.get(pageId);
  if (cached && now - cached.at < CAPABILITY_TTL_MS) return cached.answer;
  // The desktop app fetches the readers' models after installing
  // (lib/desktop-ondemand): asked for here, the first use waits for them with
  // the download's progress showing; one that cannot come now is asked again
  // next time.
  let fetched = true;
  const answer = ensureDesktopOndemand("ink-models")
    .catch(() => {
      fetched = false;
    })
    .then(() => apiGet<InkMathCapability>(`/api/v1/pages/${encodeURIComponent(pageId)}/ink-text/capability`))
    .then((value) => {
      if (!fetched) capabilityCache.delete(pageId);
      return value && typeof value.available === "boolean" ? value : UNAVAILABLE;
    })
    .catch(() => {
      capabilityCache.delete(pageId);
      return UNAVAILABLE;
    });
  capabilityCache.set(pageId, { at: now, answer });
  return answer;
}

/** Forget the cached capabilities (tests; a workspace policy change). */
export function resetOfficeInkRecognitionCapabilities(): void {
  capabilityCache.clear();
}

/** The reading's codes that mean "this is not writing", which leave the ink as ink silently. */
const UNREAD_CODES = new Set(["ink_math_unrecognized", "ink_text_unrecognized"]);

export type OfficeInkReadingOutcome =
  | {
      readonly kind: "read";
      readonly value: string;
      /** The recogniser's other readings, likeliest first (the correction menu's). */
      readonly alternatives: readonly string[];
    }
  | { readonly kind: "unread" }
  | { readonly kind: "failed"; readonly code: string | null };

/**
 * Read one piece of writing: LaTeX for Ink to Math, the words for Ink to
 * Text, with the recogniser's other readings. `unread` when the strokes say
 * nothing (the model's empty answer); `failed` when the service could not
 * answer. An unsure reading is a reading.
 */
export async function readOfficeInkStrokes(
  pageId: string,
  mode: OfficeInkRecognitionMode,
  strokes: readonly (readonly InkMathPoint[])[],
  options: {
    readonly signal?: AbortSignal;
    readonly lineWidth?: number;
    /** Ink to Text: the writer's language (BCP 47). */
    readonly language?: string | null;
  } = {},
): Promise<OfficeInkReadingOutcome> {
  const image = await rasterizeInkMathStrokes(strokes, { lineWidth: options.lineWidth });
  if (!image) return { kind: "unread" };
  try {
    if (mode === "math") {
      const answer = await recognizeInkMath(pageId, image, { privacyAcknowledged: true, signal: options.signal });
      const value = usableOfficeInkReading({ value: answer.latex, confidence: answer.confidence });
      return value
        ? { kind: "read", value, alternatives: officeInkReadingAlternatives(value, answer.alternatives) }
        : { kind: "unread" };
    }
    const answer = await recognizeInkText(pageId, image, {
      privacyAcknowledged: true,
      signal: options.signal,
      language: options.language,
    });
    const value = usableOfficeInkReading({ value: answer.text, confidence: answer.confidence });
    return value
      ? { kind: "read", value, alternatives: officeInkReadingAlternatives(value, answer.alternatives) }
      : { kind: "unread" };
  } catch (error) {
    const code = (error as { code?: unknown } | null)?.code;
    if (typeof code === "string" && UNREAD_CODES.has(code)) return { kind: "unread" };
    if ((error as { name?: unknown } | null)?.name === "AbortError") return { kind: "unread" };
    return { kind: "failed", code: typeof code === "string" ? code : null };
  }
}

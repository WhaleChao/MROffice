/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF *visual* signature overlays (Acrobat-style "Sign" parity).
 *
 * A signature annotation is a single box on a page (page-relative *percentage*
 * coordinates [0..100], like image overlays / form-field rects) holding EITHER:
 *
 *   - a DRAWN signature: a raster image data URL captured from a `<canvas>` ink
 *     pad (`kind: "drawn"`, `image`),
 *   - an uploaded IMAGE signature (`kind: "image"`, `image`), or
 *   - a TYPED signature: a short line of user text rendered in a styled script
 *     font (`kind: "typed"`, `text` + `font`).
 *
 * This is a VISUAL signature only — it is NOT a cryptographic / digital
 * certificate signature and makes no authenticity / non-repudiation claim. (A
 * future step could add real PAdES/PKCS#7 signing; that is deliberately out of
 * scope here.) The signature is stored as a SEPARATE `pdf_signatures` annotation
 * type, additive to the existing `pdf_highlight` / `pdf_note` / `pdf_ink`
 * annotations and to the `pdf_form_values` map — old PDFs with none of these
 * keep loading unchanged, and a signature carries no AcroForm form-fill value
 * (signature `Sig` widgets are recognized by lib/pdf/acroform.ts only so the
 * viewer can offer a "Sign here" affordance at the field rect).
 *
 * Everything here has no DOM / React dependency so it can be unit-tested in
 * isolation and reused by the viewer (and a potential backend mirror). It
 * mirrors `sticky-notes.ts` / `ink.ts`:
 *
 *   - the persisted annotation shape + its sanitizer ({@link sanitizeSignaturePayload}),
 *   - position/size clamping (keep the box on-page, bound its dimensions),
 *   - a STRICT data-URL guard for the drawn image (PNG/JPEG raster only — never
 *     SVG, which is markup / a script-injection vector — plus a decoded-byte
 *     cap) so a crafted payload can't smuggle an unbounded blob or active
 *     content into an `<img src>`,
 *   - typed-text normalization (collapse whitespace, strip control chars, cap
 *     length). The text is ALWAYS rendered as a plain text node, never
 *     `innerHTML`, so this is normalization + bounding, not escaping,
 *   - an allow-listed `font` set so a crafted payload can't inject arbitrary CSS
 *     into the inline `font-family`,
 *   - `buildSignaturePayload` to produce exactly what we POST (persisted bytes
 *     match what we render).
 */

/** Discriminates drawn, uploaded-image, and typed visual signatures. */
export type SignatureKind = "drawn" | "image" | "typed";

/**
 * Allowed font-family keys for a TYPED signature. Kept as an allow-list (like
 * the note / ink color allow-lists) so a crafted payload can't inject arbitrary
 * CSS into the inline `font-family`. The VALUE is a safe, self-contained CSS
 * font stack; the KEY is what is persisted + chosen in the UI.
 */
export const SIGNATURE_FONTS: Readonly<Record<string, string>> = {
  // A cursive script (closest to a handwritten signature look).
  cursive: '"Segoe Script", "Brush Script MT", cursive',
  // A serif (formal print signature).
  serif: 'Georgia, "Times New Roman", serif',
  // A sans-serif (clean print signature).
  sans: '"Helvetica Neue", Arial, sans-serif',
  // A monospace (typewriter look).
  mono: '"Courier New", Courier, monospace',
};

/** Persisted font keys, in display order. */
export const SIGNATURE_FONT_KEYS: readonly string[] =
  Object.keys(SIGNATURE_FONTS);

/** Default typed-signature font when a payload omits / supplies an unknown one. */
export const DEFAULT_SIGNATURE_FONT = "cursive";

/** Max length of a typed signature's text after normalization. */
export const SIGNATURE_MAX_TEXT = 120;

/**
 * Max number of signatures kept per page. Mirrors {@link
 * PDF_INK_MAX_STROKES_PER_PAGE} in spirit (a generous-but-bounded per-page cap)
 * but smaller — a page rarely needs many signatures.
 */
export const SIGNATURE_MAX_PER_PAGE = 50;

/** Min / max box size in page-percentage units (resolution-independent). */
export const SIGNATURE_MIN_SIZE = 1;
export const SIGNATURE_MAX_SIZE = 100;
/** Default box size for a freshly placed signature (page-percentage units). */
export const DEFAULT_SIGNATURE_WIDTH = 24;
export const DEFAULT_SIGNATURE_HEIGHT = 8;

/** Normalize one visual-signature object rotation to a clockwise integer. */
export function normalizeSignatureObjectRotation(value: unknown): number {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  const normalized = Math.round(numeric) % 360;
  return normalized < 0 ? normalized + 360 : normalized;
}

/**
 * Max decoded bytes for a DRAWN signature image (1 MB). A canvas signature is a
 * small line drawing, so this is generous; it bounds the persisted blob so a
 * crafted payload can't smuggle a multi-megabyte image. Mirrors the cover-image
 * decoded-byte guard in lib/notebook-cover.ts.
 */
export const SIGNATURE_MAX_IMAGE_BYTES = 1024 * 1024;

/**
 * Strict allow-list for a DRAWN signature's data URL: base64 PNG or JPEG only.
 * SVG / webp / gif are deliberately excluded — SVG is markup (a script-injection
 * vector) and the canvas only ever produces PNG/JPEG via `toDataURL`. Mirrors
 * SAFE_COVER_DATA_IMAGE_PATTERN (narrowed to the two raster types a signature
 * pad emits).
 */
const SAFE_SIGNATURE_IMAGE_PATTERN =
  /^data:image\/(?<mime>png|jpeg);base64,(?<payload>[A-Za-z0-9+/]+={0,2})$/i;

/**
 * Strip C0 control characters EXCEPT nothing (a typed signature is a single
 * line, so tabs / newlines are collapsed to spaces too) plus DEL (U+007F).
 * Built from char codes so no literal control byte lives in the source.
 */
const SIGNATURE_CONTROL_CHARS_REGEX = new RegExp(
  `[${String.fromCharCode(0)}-${String.fromCharCode(31)}${String.fromCharCode(127)}]`,
  "g",
);

export interface SignatureAnnotation {
  id: string;
  /** 1-based page number. */
  page: number;
  /** Box top-left + size in page-relative percent [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** "drawn" canvas, uploaded "image", or "typed" font text. */
  kind: SignatureKind;
  /**
   * For a DRAWN signature: a sanitized PNG/JPEG base64 data URL (render as an
   * `<img src>`). Empty string for a typed signature.
   */
  image: string;
  /**
   * For a TYPED signature: the normalized signature text. ALWAYS rendered as a
   * plain text node, never `innerHTML`. Empty string for a drawn signature.
   */
  text: string;
  /**
   * For a TYPED signature: the font-family KEY (one of {@link
   * SIGNATURE_FONT_KEYS}). Resolve to a CSS stack via {@link
   * resolveSignatureFont}. Always {@link DEFAULT_SIGNATURE_FONT} for a drawn
   * signature.
   */
  font: string;
  /**
   * Optional strict PNG/JPEG raster mirrored into the original PDF's native
   * `/Stamp` appearance. Drawn signatures reuse `image`; typed signatures keep
   * a browser-rendered snapshot of the selected allow-listed font. Legacy
   * reader-only rows omit it and therefore remain explicit flatten-only edits.
   */
  nativeAppearanceImage?: string;
  /** Clockwise visual-object rotation; unrelated to certificate signing. */
  objectRotation?: number;
}

function isFinitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

/** Clamp a single coordinate value into the on-page percentage range. */
function clampPercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  if (value < 0) return 0;
  if (value > 100) return 100;
  return value;
}

/** Clamp a box dimension into [{@link SIGNATURE_MIN_SIZE}, {@link SIGNATURE_MAX_SIZE}]. */
function clampSize(value: number, fallback: number): number {
  if (!Number.isFinite(value)) return fallback;
  if (value < SIGNATURE_MIN_SIZE) return SIGNATURE_MIN_SIZE;
  if (value > SIGNATURE_MAX_SIZE) return SIGNATURE_MAX_SIZE;
  return value;
}

/**
 * Clamp a signature box so its visible, optionally rotated bounds stay fully
 * on the page. Preserve the requested size and shift the centre inward; only
 * scale when a rotated box is physically too large to fit anywhere.
 */
export function clampSignatureBox(
  box: {
    x: number;
    y: number;
    width: number;
    height: number;
  },
  objectRotation = 0,
): { x: number; y: number; width: number; height: number } {
  const requestedWidth = clampSize(box.width, DEFAULT_SIGNATURE_WIDTH);
  const requestedHeight = clampSize(box.height, DEFAULT_SIGNATURE_HEIGHT);
  const rotation = normalizeSignatureObjectRotation(objectRotation);
  const radians = (rotation * Math.PI) / 180;
  const rawCosine = Math.abs(Math.cos(radians));
  const rawSine = Math.abs(Math.sin(radians));
  const cosine = rawCosine < 1e-10 ? 0 : rawCosine;
  const sine = rawSine < 1e-10 ? 0 : rawSine;
  const requestedVisualWidth = requestedWidth * cosine + requestedHeight * sine;
  const requestedVisualHeight =
    requestedWidth * sine + requestedHeight * cosine;
  const fitScale = Math.min(
    1,
    100 / requestedVisualWidth,
    100 / requestedVisualHeight,
  );
  const width = requestedWidth * fitScale;
  const height = requestedHeight * fitScale;
  const visualWidth = width * cosine + height * sine;
  const visualHeight = width * sine + height * cosine;
  const horizontalInset = Math.max(width, visualWidth) / 2;
  const verticalInset = Math.max(height, visualHeight) / 2;
  const requestedCenterX = clampPercent(box.x) + requestedWidth / 2;
  const requestedCenterY = clampPercent(box.y) + requestedHeight / 2;
  const centerX = Math.min(
    100 - horizontalInset,
    Math.max(horizontalInset, requestedCenterX),
  );
  const centerY = Math.min(
    100 - verticalInset,
    Math.max(verticalInset, requestedCenterY),
  );
  return {
    x: centerX - width / 2,
    y: centerY - height / 2,
    width,
    height,
  };
}

/** True when the value is a non-empty member of the font allow-list. */
export function isSignatureFont(value: unknown): value is string {
  return (
    typeof value === "string" &&
    Object.prototype.hasOwnProperty.call(SIGNATURE_FONTS, value)
  );
}

/** Resolve a font KEY to its safe CSS font stack (default when unknown). */
export function resolveSignatureFont(key: unknown): string {
  return isSignatureFont(key)
    ? SIGNATURE_FONTS[key]
    : SIGNATURE_FONTS[DEFAULT_SIGNATURE_FONT];
}

/**
 * Normalize a typed signature's text: coerce to string, strip control chars +
 * DEL, collapse ALL whitespace runs to single spaces (a signature is one line),
 * trim, and hard-cap length. NOT HTML escaping — the renderer must place this
 * into a text node, never `innerHTML`. Returns "" for non-string / empty input.
 */
export function normalizeSignatureText(value: unknown): string {
  if (typeof value !== "string") return "";
  const cleaned = value
    .replace(SIGNATURE_CONTROL_CHARS_REGEX, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.slice(0, SIGNATURE_MAX_TEXT);
}

/**
 * MIME types accepted for an UPLOADED signature image. Mirrors the data-URL
 * allow-list ({@link SAFE_SIGNATURE_IMAGE_PATTERN}) — PNG / JPEG raster only.
 * SVG / gif / webp are excluded for the same script-injection / unbounded-blob
 * reasons; the uploaded file is also re-validated as a PNG/JPEG data URL by
 * {@link sanitizeSignatureImage} after it is read, so this is a fast first-pass
 * reject, not the sole guard.
 */
export const SIGNATURE_UPLOAD_MIME_TYPES: readonly string[] = [
  "image/png",
  "image/jpeg",
];

/**
 * Result of validating a user-picked signature image FILE before reading it. A
 * pure, DOM-free check of the file's declared `type` + `size` against the same
 * allow-list / byte cap the persisted data URL must satisfy — so an oversized or
 * non-PNG/JPEG pick is rejected with a stable reason BEFORE the bytes are read,
 * without ever decoding it. `ok: false` carries a machine-readable `reason`
 * (`"type"` | `"size"` | `"empty"`) the UI maps to a localized message.
 */
export type SignatureImageFileCheck =
  { ok: true } | { ok: false; reason: "type" | "size" | "empty" };

/**
 * Validate an uploaded signature image file by its metadata only (no decode):
 *  - `type` must be in {@link SIGNATURE_UPLOAD_MIME_TYPES} (PNG / JPEG),
 *  - `size` must be > 0 and ≤ {@link SIGNATURE_MAX_IMAGE_BYTES}.
 *
 * Accepts the structural subset of `File` it reads (`type` + `size`) so it can be
 * unit-tested without a real `File`. This bounds the upload up front; the bytes
 * are still re-validated as a safe data URL by {@link sanitizeSignatureImage}
 * after they are read, so a spoofed `type` cannot smuggle non-raster content
 * through to the `<img src>`.
 */
export function validateSignatureImageFile(file: {
  type?: unknown;
  size?: unknown;
}): SignatureImageFileCheck {
  const type = typeof file?.type === "string" ? file.type.toLowerCase() : "";
  if (!SIGNATURE_UPLOAD_MIME_TYPES.includes(type)) {
    return { ok: false, reason: "type" };
  }
  const size = Number(file?.size);
  if (!Number.isFinite(size) || size <= 0) {
    return { ok: false, reason: "empty" };
  }
  if (size > SIGNATURE_MAX_IMAGE_BYTES) {
    return { ok: false, reason: "size" };
  }
  return { ok: true };
}

/**
 * Validate a DRAWN signature's data URL: must be a base64 PNG/JPEG (never SVG /
 * other) within the decoded-byte cap. Returns the trimmed URL or null. This is
 * the single guard that decides whether a string is safe to feed an `<img src>`.
 */
export function sanitizeSignatureImage(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  const match = SAFE_SIGNATURE_IMAGE_PATTERN.exec(trimmed);
  const mime = match?.groups?.mime?.toLowerCase();
  const payload = match?.groups?.payload;
  if (!mime || !payload || payload.length % 4 !== 0) return null;
  let decoded: string;
  try {
    decoded = globalThis.atob(payload);
  } catch {
    return null;
  }
  if (!decoded || decoded.length > SIGNATURE_MAX_IMAGE_BYTES) return null;
  const isPng =
    mime === "png" &&
    decoded.length >= 8 &&
    decoded.charCodeAt(0) === 0x89 &&
    decoded.slice(1, 4) === "PNG" &&
    decoded.charCodeAt(4) === 0x0d &&
    decoded.charCodeAt(5) === 0x0a &&
    decoded.charCodeAt(6) === 0x1a &&
    decoded.charCodeAt(7) === 0x0a;
  const isJpeg =
    mime === "jpeg" &&
    decoded.length >= 3 &&
    decoded.charCodeAt(0) === 0xff &&
    decoded.charCodeAt(1) === 0xd8 &&
    decoded.charCodeAt(2) === 0xff;
  if (!isPng && !isJpeg) return null;
  return trimmed;
}

/**
 * Validate + normalize a persisted `pdf_signatures` payload into a usable
 * annotation, or return null when it is structurally invalid (mirrors
 * `sanitizeNotePayload` / `sanitizeInkPayload`). Out-of-range positions are
 * rejected (returns null) rather than silently snapped; the box size is clamped.
 * A "drawn" payload needs a safe image; a "typed" payload needs non-empty text;
 * a payload that ends up with neither usable content is rejected.
 */
export function sanitizeSignaturePayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): SignatureAnnotation | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const record = payload as Record<string, unknown>;
  const page = Number(record.page);
  if (!Number.isInteger(page) || page < 1 || page > 10000) return null;
  const { x, y } = record;
  if (!isFinitePercent(x) || !isFinitePercent(y)) return null;

  // A signature is either drawn (image) or typed (text). Default unknown/missing
  // kind from whichever content is present so a legacy/partial payload still
  // resolves; if neither is usable, reject.
  const rawKind = record.kind;
  const image = sanitizeSignatureImage(record.image);
  const text = normalizeSignatureText(record.text);
  let kind: SignatureKind;
  if (rawKind === "drawn") kind = "drawn";
  else if (rawKind === "image") kind = "image";
  else if (rawKind === "typed") kind = "typed";
  else kind = image ? "drawn" : "typed";

  if ((kind === "drawn" || kind === "image") && !image) return null;
  if (kind === "typed" && !text) return null;

  const objectRotation = normalizeSignatureObjectRotation(
    record.objectRotation,
  );
  const box = clampSignatureBox(
    {
      x,
      y,
      width: Number(record.width),
      height: Number(record.height),
    },
    objectRotation,
  );
  const font = isSignatureFont(record.font)
    ? record.font
    : DEFAULT_SIGNATURE_FONT;
  const nativeAppearanceImage =
    record.nativeAppearanceImage === undefined
      ? undefined
      : sanitizeSignatureImage(record.nativeAppearanceImage);
  if (record.nativeAppearanceImage !== undefined && !nativeAppearanceImage) {
    return null;
  }
  return {
    id,
    page,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    kind,
    // Only carry the content for the resolved kind so a typed payload can't also
    // smuggle a (separately rendered) image and vice-versa.
    image: kind === "drawn" || kind === "image" ? (image as string) : "",
    text: kind === "typed" ? text : "",
    font,
    ...(nativeAppearanceImage ? { nativeAppearanceImage } : {}),
    ...(objectRotation === 0 ? {} : { objectRotation }),
  };
}

/**
 * Build the payload to POST for a new signature. Position/size are clamped, the
 * image guarded (drawn) and text normalized (typed), the font forced onto the
 * allow-list — so what we persist matches what we render. Returns null when the
 * input has no usable content for its kind (so a blank signature is never
 * persisted).
 */
export function buildSignaturePayload(input: {
  page: number;
  x: number;
  y: number;
  width?: number;
  height?: number;
  kind: SignatureKind;
  image?: string;
  text?: string;
  font?: string;
  nativeAppearanceImage?: string;
  objectRotation?: number;
}): {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  kind: SignatureKind;
  image: string;
  text: string;
  font: string;
  nativeAppearanceImage?: string;
  objectRotation?: number;
} | null {
  const objectRotation = normalizeSignatureObjectRotation(input.objectRotation);
  const box = clampSignatureBox(
    {
      x: input.x,
      y: input.y,
      width: input.width ?? DEFAULT_SIGNATURE_WIDTH,
      height: input.height ?? DEFAULT_SIGNATURE_HEIGHT,
    },
    objectRotation,
  );
  if (input.kind === "drawn" || input.kind === "image") {
    const image = sanitizeSignatureImage(input.image);
    if (!image) return null;
    const nativeAppearanceImage =
      input.nativeAppearanceImage === undefined
        ? undefined
        : sanitizeSignatureImage(input.nativeAppearanceImage);
    if (input.nativeAppearanceImage !== undefined && !nativeAppearanceImage) {
      return null;
    }
    return {
      page: input.page,
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      kind: input.kind,
      image,
      text: "",
      font: DEFAULT_SIGNATURE_FONT,
      ...(nativeAppearanceImage ? { nativeAppearanceImage } : {}),
      ...(objectRotation === 0 ? {} : { objectRotation }),
    };
  }
  const text = normalizeSignatureText(input.text);
  if (!text) return null;
  const nativeAppearanceImage =
    input.nativeAppearanceImage === undefined
      ? undefined
      : sanitizeSignatureImage(input.nativeAppearanceImage);
  if (input.nativeAppearanceImage !== undefined && !nativeAppearanceImage) {
    return null;
  }
  return {
    page: input.page,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    kind: "typed",
    image: "",
    text,
    font: isSignatureFont(input.font) ? input.font : DEFAULT_SIGNATURE_FONT,
    ...(nativeAppearanceImage ? { nativeAppearanceImage } : {}),
    ...(objectRotation === 0 ? {} : { objectRotation }),
  };
}

/**
 * Cap the number of signatures kept for a single page to {@link
 * SIGNATURE_MAX_PER_PAGE}, dropping the OLDEST first (the tail — most recent —
 * is what the user just added). Returns a new array; input untouched. Mirrors
 * `capStrokesPerPage`.
 */
export function capSignaturesPerPage<T>(
  signatures: ReadonlyArray<T>,
  max = SIGNATURE_MAX_PER_PAGE,
): T[] {
  if (!Array.isArray(signatures)) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : SIGNATURE_MAX_PER_PAGE;
  if (signatures.length <= cap) return signatures.slice();
  return signatures.slice(signatures.length - cap);
}

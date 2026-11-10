/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PDF editor production feature flags.
 *
 * Accepted desktop paths are on by default. An exact `0` is the deploy-time
 * kill switch; other values must not accidentally hide production-ready PDF
 * capabilities.
 */

/**
 * Coordinates cancellation and release of pdf.js canvas/text render resources
 * across zoom, rotation, layer, document, and page-lifecycle changes.
 */
export function pdfRenderLifecycleV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_RENDER_LIFECYCLE_V2 !== "0";
}

/**
 * Normalizes box-overlay pointer deltas against the live rendered page extent
 * so notebook-canvas CSS zoom cannot skew persisted geometry.
 */
export function pdfOverlayCoordinatesV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_OVERLAY_COORDINATES_V2 !== "0";
}

/**
 * Hands PageWindow mutations the server-decrypted session bytes used by the
 * viewer, then rejects any still-encrypted fallback before pdf-lib can edit it.
 */
export function pdfPageWindowMutationGuardV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_PAGE_WINDOW_MUTATION_GUARD_V2 !== "0";
}

/**
 * Removes unreachable source-page objects after redaction raster replacement
 * and refuses ambiguous AcroForm/shared-reference residue.
 *
 * This safety boundary has graduated to the production default. An explicit
 * `0` disables Apply Redactions completely; callers must never treat the kill
 * switch as permission to fall back to the legacy best-effort writer.
 */
export function pdfRedactionReachabilityV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_REDACTION_REACHABILITY_V2 !== "0";
}

/**
 * Refuses primary-export form edits before pdf-lib can delete XFA or create an
 * empty AcroForm for stale overlay values.
 *
 * This safety boundary is the production default. An explicit `0` retains the
 * deploy-time rollback switch without weakening any downstream mutation guard.
 */
export function pdfAcroFormMutationGuardV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_ACROFORM_MUTATION_GUARD_V2 !== "0";
}

/**
 * Keeps mixed-run caret placement on grapheme boundaries and prevents a
 * single-line text box from growing when the user deletes content.
 *
 * This path graduated after its focused DOM, geometry, style-run, reflow, and
 * export regression suites passed. It is now the production default; `0` is
 * retained as a deploy-time kill switch for rollback without rewriting edits.
 */
export function pdfTextEditStabilityV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_TEXT_EDIT_STABILITY_V2 !== "0";
}

/**
 * Emits the imported FreeText V2 patch protocol whose server path validates,
 * preserves, and reapplies a source annotation's `/RD` inner rectangle.
 */
export function pdfFreeTextRectangleDifferencesV2Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_FREETEXT_RD_V2 !== "0";
}

/**
 * Routes every in-scope PDF byte output through the typed P0.03 save policy.
 * The accepted policy is now the production default; `0` remains the exact
 * deploy-time kill switch for the legacy rollback path.
 */
export function pdfSavePolicyV1Enabled(): boolean {
  return process.env.NEXT_PUBLIC_PDF_SAVE_POLICY_V1 !== "0";
}


/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PDF.js's legacy build, as `loadCompatiblePdfJs` (lib/pdf/pdfjs-runtime.ts)
 * loads it on an engine that lacks what the modern build calls.
 *
 * Only the loader imports this module. jsdom is such an engine, so the unit
 * tests' setup (tests/unit/setup-pdfjs-build.ts) resolves it to whatever
 * "pdfjs-dist" is in the test file: a suite that mocks "pdfjs-dist" gets the
 * same mock from the loader, while a test that imports
 * "pdfjs-dist/legacy/build/pdf.mjs" itself still runs the real legacy build.
 */
export * from "pdfjs-dist/legacy/build/pdf.mjs";

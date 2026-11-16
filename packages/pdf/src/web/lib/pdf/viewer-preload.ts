/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ComponentType } from "react";

type PdfViewerModule = typeof import("../../components/notebook/contents/PDFViewer");

let loading: Promise<PdfViewerModule> | null = null;
let loaded: PdfViewerModule["default"] | null = null;

/** The reader's module: the load already on its way, else a new one. A
 * failed load is not kept (the next caller asks again). */
export function loadPdfViewerModule(): Promise<PdfViewerModule> {
  if (loading) return loading;
  const load = import("../../components/notebook/contents/PDFViewer");
  loading = load;
  load.then(
    (module) => { loaded = module.default; },
    () => { if (loading === load) loading = null; },
  );
  return load;
}

/** The reader, once its code has arrived; null until then. */
export function loadedPdfViewer(): ComponentType<never> | null {
  return loaded as unknown as ComponentType<never> | null;
}

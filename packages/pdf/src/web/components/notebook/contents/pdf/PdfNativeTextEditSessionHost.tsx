"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  forwardRef,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import PdfNativeListTextEditDialog, {
  type PdfNativeListTextEditDialogProps,
} from "./PdfNativeListTextEditDialog";
import { markPdfTextLatency } from "../../../../lib/pdf/text-edit-latency";

type Opening = Pick<
  PdfNativeListTextEditDialogProps,
  "originalText" | "initialLayout" | "initialCaretOffset" | "initialDraftText"
> & { sessionKey: string };

export interface PdfNativeTextEditSessionHandle {
  begin: (opening: Opening) => void;
  /** The proof for this opening failed for good: the session stops waiting
   * (no longer busy), shows `message`, and keeps the draft editable. */
  fail: (message: string) => void;
  retire: () => void;
  isActive: () => boolean;
}

const rejectUnpreparedWrite = async () => false;

/** Keep the first caret paint local, as other editors do for their input
 * surfaces. The provisional session has no lease/proof and cannot write. The
 * parent adopts the SAME dialog after its first paint; it must not replace
 * text, selection, IME or geometry with the original snapshot on adoption. */
const PdfNativeTextEditSessionHost = forwardRef<
  PdfNativeTextEditSessionHandle,
  PdfNativeListTextEditDialogProps & { scopeKey: string; sessionKey: string }
>(function PdfNativeTextEditSessionHost(
  { scopeKey, sessionKey, ...props },
  ref,
) {
  const [opening, setOpening] = useState<
    (Opening & { scopeKey: string; failure?: string }) | null
  >(null);
  const openingRef = useRef<(Opening & { scopeKey: string; failure?: string }) | null>(null);
  const liveScopeRef = useRef({ scopeKey, open: props.open });
  useLayoutEffect(() => { liveScopeRef.current = { scopeKey, open: props.open }; }, [scopeKey, props.open]);
  useImperativeHandle(
    ref,
    () => ({
      begin: (next) => { const value = { ...next, scopeKey }; openingRef.current = value; setOpening(value); },
      fail: (message) => {
        const current = openingRef.current;
        if (!current) return;
        const value = { ...current, failure: message };
        openingRef.current = value;
        setOpening(value);
      },
      retire: () => { openingRef.current = null; setOpening(null); },
      isActive: () => liveScopeRef.current.open || openingRef.current?.scopeKey === liveScopeRef.current.scopeKey,
    }),
    [scopeKey],
  );
  useLayoutEffect(() => {
    if (props.open || opening?.scopeKey !== scopeKey) { openingRef.current = null; setOpening(null); }
  }, [props.open, opening?.scopeKey, scopeKey]);

  const provisional =
    !props.open && opening?.scopeKey === scopeKey ? opening : null;
  // (Development: an editor given another key while open is a new editor --
  // its caret and keys start again -- and says which key moved.)
  const shownKeyRef = useRef<string | null>(null);
  const shownKey = props.open || provisional ? `${provisional ? "p" : "o"}|${provisional?.sessionKey ?? sessionKey}` : null;
  if (process.env.NODE_ENV !== "production" && shownKey !== shownKeyRef.current) {
    if (shownKeyRef.current && shownKey) markPdfTextLatency("editor.rekeyed", {
      from: shownKeyRef.current.replace(/[0-9a-f]{40,}/g, (sha) => sha.slice(0, 8)),
      to: shownKey.replace(/[0-9a-f]{40,}/g, (sha) => sha.slice(0, 8)), scopeMoved: opening !== null && opening.scopeKey !== scopeKey });
    shownKeyRef.current = shownKey;
  }
  return (
    <PdfNativeListTextEditDialog
      {...props}
      {...(provisional
        ? {
            originalText: provisional.originalText,
            initialLayout: provisional.initialLayout,
            initialCaretOffset: provisional.initialCaretOffset,
            initialDraftText: provisional.initialDraftText ?? null,
            open: true,
            preparationPending: provisional.failure === undefined,
            preparationFailed: provisional.failure !== undefined,
            propertyProof: null,
            sourceGeometry: undefined,
            sourceAppearance: undefined,
            paragraphColors: undefined,
            initialEditingSurface: null,
            initialEditingSurfacePrepared: false,
            error: provisional.failure ?? null,
            applyDisabledReason: provisional.failure ?? props.labels.previewing,
            // Publishing the initial draft to PageWindow would synchronously render
            // the whole reader before this caret paints. Adoption publishes it once.
            onDraftStateChange: undefined,
            onApply: rejectUnpreparedWrite,
          }
        : {})}
      key={provisional?.sessionKey ?? sessionKey}
    />
  );
});

export default PdfNativeTextEditSessionHost;

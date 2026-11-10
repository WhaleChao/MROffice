"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type RefObject,
} from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import {
  analyzeSpreadsheetPicture,
  readClipboardPicture,
  type SpreadsheetDataFromPictureCapability,
  type SpreadsheetDataFromPictureResult,
  type SpreadsheetDataFromPictureSource,
} from "../../../lib/spreadsheet/data-from-picture";

export interface SpreadsheetDataFromPictureDialogProps {
  readonly open: boolean;
  readonly locale: string;
  readonly pageId: string;
  readonly source: SpreadsheetDataFromPictureSource;
  readonly capability: SpreadsheetDataFromPictureCapability;
  readonly returnFocusRef: RefObject<HTMLElement | null>;
  readonly onCancel: () => void;
  readonly onCommit: (
    rows: readonly (readonly string[])[],
    result: SpreadsheetDataFromPictureResult,
  ) => boolean;
}

function pictureError(error: unknown, t: (key: string) => string): string {
  const code = isApiRequestError(error) ? error.code : undefined;
  if (
    code === "workspace_ai_egress_policy_unconfigured" ||
    code === "workspace_ai_egress_policy_denied" ||
    code === "workspace_ai_egress_workload_denied"
  ) {
    return t("spreadsheet.surfaceCopy.workspacePolicyDoesNotAuthorizeSendingThisImageToAn");
  }
  if (
    code === "workspace_ai_egress_credentials_unconfigured" ||
    code === "workspace_ai_egress_endpoint_unconfigured" ||
    code === "workspace_ai_egress_endpoint_configuration_stale"
  ) {
    return t("spreadsheet.surfaceCopy.theWorkspaceDedicatedAIProviderIsNotConfigured");
  }
  if (
    code === "workspace_ai_egress_policy_stale" ||
    code === "spreadsheet_data_from_picture_page_stale"
  ) {
    return t("spreadsheet.surfaceCopy.theSpreadsheetOrWorkspacePolicyChangedCloseAndTryAgain");
  }
  if (code?.startsWith("spreadsheet_data_from_picture_image_")) {
    return t("spreadsheet.surfaceCopy.chooseAPNGJPEGOrWebPImageNoLargerThan");
  }
  if (
    error instanceof Error &&
    ["clipboard_image_unavailable", "clipboard_image_missing"].includes(
      error.message,
    )
  ) {
    return t("spreadsheet.surfaceCopy.noReadableImageIsAvailableOnTheClipboard");
  }
  return t("spreadsheet.surfaceCopy.thePictureCouldNotBeAnalyzedTryAgain");
}

export function SpreadsheetDataFromPictureDialog({
  open,
  pageId,
  source,
  capability,
  returnFocusRef,
  onCancel,
  onCommit,
}: SpreadsheetDataFromPictureDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const copy = useMemo(
    () => ({
      title: t("spreadsheet.surfaceCopy.dataFromPicture"),
      description: t("spreadsheet.surfaceCopy.chooseAPictureAndReviewTheRecognizedTableThePicture"),
      privacy: t("spreadsheet.surfaceCopy.iHavePermissionToProcessThisPictureAndAgreeTo"),
      chooseFile: t("spreadsheet.surfaceCopy.choosePicture"),
      readClipboard: t("spreadsheet.surfaceCopy.readClipboardPicture"),
      analyze: t("spreadsheet.surfaceCopy.analyzingPicture"),
      insert: t("spreadsheet.surfaceCopy.insertData"),
      cancel: t("spreadsheet.surfaceCopy.cancel"),
      review: t("spreadsheet.surfaceCopy.review"),
      empty: t("spreadsheet.surfaceCopy.chooseAPictureToReviewAndEditTheRecognizedCells"),
      privacyRequired: t("spreadsheet.surfaceCopy.confirmTheImagePermissionAndPrivacyNoticeFirst"),
      stale: t("spreadsheet.surfaceCopy.theSpreadsheetOrSelectionChangedCloseAndStartAgain"),
    }),
    [t],
  );
  const [privacyAcknowledged, setPrivacyAcknowledged] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] =
    useState<SpreadsheetDataFromPictureResult | null>(null);
  const [draftRows, setDraftRows] = useState<string[][]>([]);
  const controllerRef = useRef<AbortController | null>(null);
  const privacyRef = useRef<HTMLInputElement | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (open) return;
    controllerRef.current?.abort();
    controllerRef.current = null;
    setPrivacyAcknowledged(false);
    setBusy(false);
    setError(null);
    setResult(null);
    setDraftRows([]);
  }, [open]);

  useEffect(
    () => () => {
      controllerRef.current?.abort();
    },
    [],
  );

  const analyze = useCallback(
    async (file: File) => {
      if (!privacyAcknowledged) {
        setError(copy.privacyRequired);
        privacyRef.current?.focus();
        return;
      }
      controllerRef.current?.abort();
      const controller = new AbortController();
      controllerRef.current = controller;
      setBusy(true);
      setError(null);
      setResult(null);
      setDraftRows([]);
      try {
        const next = await analyzeSpreadsheetPicture(
          pageId,
          {
            file,
            source,
            privacyAcknowledged: true,
            capability,
          },
          controller.signal,
        );
        if (controller.signal.aborted) return;
        setResult(next);
        setDraftRows(next.rows.map((row) => row.map((cell) => cell.text)));
      } catch (nextError) {
        if (!controller.signal.aborted) {
          setError(pictureError(nextError, t));
        }
      } finally {
        if (!controller.signal.aborted) setBusy(false);
      }
    },
    [
      capability,
      copy.privacyRequired,
      pageId,
      privacyAcknowledged,
      source,
      t,
    ],
  );

  const chooseFile = useCallback(() => {
    if (!privacyAcknowledged) {
      setError(copy.privacyRequired);
      privacyRef.current?.focus();
      return;
    }
    fileRef.current?.click();
  }, [copy.privacyRequired, privacyAcknowledged]);

  const chooseClipboard = useCallback(async () => {
    if (!privacyAcknowledged) {
      setError(copy.privacyRequired);
      privacyRef.current?.focus();
      return;
    }
    try {
      await analyze(await readClipboardPicture());
    } catch (nextError) {
      setError(pictureError(nextError, t));
    }
  }, [analyze, copy.privacyRequired, privacyAcknowledged, t]);

  const handleFile = useCallback(
    (event: ChangeEvent<HTMLInputElement>) => {
      const file = event.currentTarget.files?.[0];
      event.currentTarget.value = "";
      if (file) void analyze(file);
    },
    [analyze],
  );

  const insert = useCallback(() => {
    if (!result || !draftRows.length) return;
    if (!onCommit(draftRows, result)) {
      setError(copy.stale);
      return;
    }
    onCancel();
  }, [copy.stale, draftRows, onCancel, onCommit, result]);

  const sourceAction =
    source === "file" ? (
      <>
        {null}
        {null}
      </>
    ) : (
      null
    );

  return (
    <OfficeDialogShell
      open={open}
      title={copy.title}
      description={copy.description}
      defaultActionLabel={copy.insert}
      cancelActionLabel={copy.cancel}
      defaultActionDisabled={!result || draftRows.length === 0}
      busy={busy}
      busyActionLabel={copy.analyze}
      onDefaultAction={insert}
      onCancel={onCancel}
      initialFocusRef={privacyRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="spreadsheet-data-from-picture"
      className="sm:max-w-[860px]"
    >
      <label className="flex items-start gap-2">
        {null}
        <span>{copy.privacy}</span>
      </label>
      <div className="flex items-center gap-2">{sourceAction}</div>
      {error ? (
        null
      ) : null}
      {result ? (
        <div className={`${officeDialogParts.well} max-h-[420px] overflow-auto`}>
          {null}
        </div>
      ) : (
        <p className={officeDialogParts.message}>{copy.empty}</p>
      )}
      {result?.warnings.map((warning, index) => (
        <p key={index} className={officeDialogParts.hint}>
          {warning}
        </p>
      ))}
    </OfficeDialogShell>
  );
}

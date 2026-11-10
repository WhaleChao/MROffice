"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { containDialogTab } from "@mrnote/core/web/lib/dialog-tab-containment.ts";
import {
  officeDialogClassName,
  officeDialogParts as parts,
  useOfficeDialogChrome,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { Files, X } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import LegacyMailMergeRecipientsDialog from "@mrnote/core/web/components/console/editor/MailMergeRecipientsDialog.tsx";
import {
  MAX_MAIL_MERGE_FINISH_RECORDS,
  planMailMergeFinish,
  type MailMergeData,
  type MailMergeFinishSelection,
} from "@mrnote/core/web/lib/document/mailMerge.ts";
import type { MailMergeDialogAction } from "@mrnote/core/web/components/console/editor/native-document/mergeFieldExtension.tsx";

export interface TrustedMailMergePrintJob {
  trusted: true;
  submit: (recordNumbers: readonly number[]) => void | Promise<void>;
}

export interface MailMergeEmailConnector {
  connectorId: string;
  sendAfterConfirmation: (
    recordNumbers: readonly number[],
    signal?: AbortSignal,
  ) => void | Promise<void>;
}

interface MailMergeRecipientsDialogProps {
  data: MailMergeData;
  previewRecord: number;
  /** Template/data writes retire in readonly; finish/export remains available. */
  canEdit?: boolean;
  initialAction?: MailMergeDialogAction;
  onApplyData: (data: MailMergeData) => void;
  onPreview: (recordNumber: number) => void;
  onFinish: (recordNumbers: readonly number[]) => void | Promise<void>;
  onInsertAddressBlock: (data: MailMergeData) => void;
  onInsertGreetingLine: (data: MailMergeData) => void;
  onClose: () => void;
  /** Absent until a browser/server print job has a trusted completion contract. */
  printJob?: TrustedMailMergePrintJob | null;
  /** Absent until a real email connector is configured for this document. */
  emailConnector?: MailMergeEmailConnector | null;
}

function finishCopy(
  t: (key: string, values?: Record<string, string | number>) => string,
) {
  const k = (leaf: string) => t(`mailMerge.finishDialog.${leaf}`);
  return {
    back: k("back"),
    mergeRecords: k("mergeRecords"),
    all: k("all"),
    current: k("current"),
    currentUnavailable: k("currentUnavailable"),
    range: k("range"),
    from: k("from"),
    to: k("to"),
    ready: (count: number) => t("mailMerge.finishDialog.ready", { count }),
    empty: k("empty"),
    invalid: k("invalid"),
    overLimit: t("mailMerge.finishDialog.overLimit", {
      max: String(MAX_MAIL_MERGE_FINISH_RECORDS),
    }),
    editDocuments: k("editDocuments"),
    printDocuments: k("printDocuments"),
    printUnavailable: k("printUnavailable"),
    sendEmail: k("sendEmail"),
    emailUnavailable: k("emailUnavailable"),
    emailReadOnly: k("emailReadOnly"),
    emailConfirm: k("emailConfirm"),
    confirmSend: k("confirmSend"),
    cancel: k("cancel"),
    working: k("working"),
    failed: k("failed"),
  };
}

function finishNumber(value: string): number {
  const trimmed = value.trim();
  if (!/^\d+$/.test(trimmed)) return Number.NaN;
  return Number(trimmed);
}

export default function MailMergeRecipientsDialog({
  data,
  previewRecord,
  canEdit = true,
  initialAction = "recipients",
  onApplyData,
  onPreview,
  onFinish,
  onInsertAddressBlock,
  onInsertGreetingLine,
  onClose,
  printJob = null,
  emailConnector = null,
}: MailMergeRecipientsDialogProps) {
  const t = useTranslations("console-notebooks");
  const chromeTone = useOfficeDialogChrome();
  const copy = useMemo(() => finishCopy(t), [t]);
  const [finishOpen, setFinishOpen] = useState(
    () => initialAction === "finishDocx",
  );
  const [mode, setMode] = useState<MailMergeFinishSelection["mode"]>("all");
  const [from, setFrom] = useState("1");
  const eligibleCount = useMemo(
    () => planMailMergeFinish(data, { mode: "all" }).eligibleCount,
    [data],
  );
  const [to, setTo] = useState(() => String(Math.max(eligibleCount, 1)));
  const [busyAction, setBusyAction] = useState<
    "documents" | "print" | "email" | null
  >(null);
  const [actionError, setActionError] = useState("");
  const [emailConfirmationOpen, setEmailConfirmationOpen] = useState(false);
  const finishScopeRef = useRef<HTMLInputElement | null>(null);
  const actionGenerationRef = useRef(0);
  const activeActionRef = useRef<{
    generation: number;
    action: "documents" | "print" | "email";
    controller: AbortController;
  } | null>(null);
  const mountedRef = useRef(true);
  const retireActiveAction = useCallback(
    (action?: "documents" | "print" | "email") => {
      const active = activeActionRef.current;
      if (!active || (action && active.action !== action)) return;
      activeActionRef.current = null;
      active.controller.abort();
      if (mountedRef.current) setBusyAction(null);
    },
    [],
  );
  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      activeActionRef.current?.controller.abort();
      activeActionRef.current = null;
    };
  }, []);
  useEffect(() => {
    if (!finishOpen) return;
    let focusFrame = 0;
    const settleFrame = window.requestAnimationFrame(() => {
      focusFrame = window.requestAnimationFrame(() => {
        finishScopeRef.current?.focus({ preventScroll: true });
      });
    });
    return () => {
      window.cancelAnimationFrame(settleFrame);
      window.cancelAnimationFrame(focusFrame);
    };
  }, [finishOpen]);
  useEffect(() => {
    if (canEdit) return;
    setEmailConfirmationOpen(false);
    retireActiveAction("email");
  }, [canEdit, retireActiveAction]);

  const currentPlan = useMemo(
    () =>
      planMailMergeFinish(data, {
        mode: "current",
        currentRecordNumber: previewRecord,
      }),
    [data, previewRecord],
  );
  const selection = useMemo<MailMergeFinishSelection>(() => {
    if (mode === "current") {
      return { mode, currentRecordNumber: previewRecord };
    }
    if (mode === "range") {
      return { mode, from: finishNumber(from), to: finishNumber(to) };
    }
    return { mode };
  }, [from, mode, previewRecord, to]);
  const plan = useMemo(
    () => planMailMergeFinish(data, selection),
    [data, selection],
  );

  const chooseMode = (nextMode: MailMergeFinishSelection["mode"]) => {
    setMode(nextMode);
    setActionError("");
    setEmailConfirmationOpen(false);
  };
  const runAction = async (
    action: "documents" | "print" | "email",
    callback: (
      recordNumbers: readonly number[],
      signal?: AbortSignal,
    ) => void | Promise<void>,
  ) => {
    if (
      plan.status !== "ready" ||
      busyAction !== null ||
      (action === "email" && !canEdit)
    ) {
      return;
    }
    const generation = actionGenerationRef.current + 1;
    actionGenerationRef.current = generation;
    const controller = new AbortController();
    const authority = { generation, action, controller };
    activeActionRef.current = authority;
    setBusyAction(action);
    setActionError("");
    try {
      if (action === "email") {
        await callback(plan.recordNumbers, controller.signal);
      } else {
        await callback(plan.recordNumbers);
      }
      if (
        !mountedRef.current ||
        controller.signal.aborted ||
        activeActionRef.current !== authority
      ) {
        return;
      }
      if (action === "email") setEmailConfirmationOpen(false);
    } catch {
      if (
        mountedRef.current &&
        !controller.signal.aborted &&
        activeActionRef.current === authority
      ) {
        setActionError(copy.failed);
      }
    } finally {
      if (activeActionRef.current === authority) {
        activeActionRef.current = null;
        if (mountedRef.current) setBusyAction(null);
      }
    }
  };
  const closeDialog = () => {
    retireActiveAction();
    onClose();
  };
  const returnToRecipients = () => {
    retireActiveAction();
    setFinishOpen(false);
  };

  if (!finishOpen) {
    return (
      <LegacyMailMergeRecipientsDialog
        data={data}
        previewRecord={previewRecord}
        canEdit={canEdit}
        initialAction={initialAction}
        onApplyData={onApplyData}
        onPreview={onPreview}
        onFinish={() => {
          setMode("all");
          setFrom("1");
          setTo(String(Math.max(eligibleCount, 1)));
          setActionError("");
          setEmailConfirmationOpen(false);
          setFinishOpen(true);
        }}
        onInsertAddressBlock={onInsertAddressBlock}
        onInsertGreetingLine={onInsertGreetingLine}
        onClose={onClose}
      />
    );
  }

  const statusMessage =
    plan.status === "ready"
      ? copy.ready(plan.recordNumbers.length)
      : plan.status === "over-limit"
        ? copy.overLimit
        : plan.status === "invalid"
          ? copy.invalid
          : copy.empty;
  const canRun = plan.status === "ready" && busyAction === null;
  const printReady = printJob?.trusted === true;
  const emailReady = Boolean(
    emailConnector?.connectorId.trim() &&
    typeof emailConnector.sendAfterConfirmation === "function",
  );
  const emailAuthorized = canEdit && emailReady;

  return (
    <div
      className={`mail-merge-dialog ${officeDialogClassName({ tone: chromeTone ?? "document" })}`}
      role="dialog"
      aria-modal="true"
      aria-label={t("mailMerge.finishMerge")}
      data-testid="mail-merge-finish-dialog"
      onKeyDown={(event) => {
        // aria-modal: Tab and Shift+Tab wrap inside, never reach the page.
        if (containDialogTab(event, event.currentTarget)) return;
        if (event.key !== "Escape") return;
        event.preventDefault();
        event.stopPropagation();
        if (emailConfirmationOpen) {
          retireActiveAction("email");
          setEmailConfirmationOpen(false);
          return;
        }
        closeDialog();
      }}
    >
      <div className={`mail-merge-dialog__head ${parts.header}`}>
        <span className={`mail-merge-dialog__heading ${parts.title}`}>
          <Files size={16} aria-hidden="true" />
          {t("mailMerge.finishMerge")}
        </span>
        <button
          type="button"
          onClick={closeDialog}
          className={`mail-merge-dialog__close ${parts.close}`}
          aria-label={t("mailMerge.close")}
          style={{ minWidth: 44, minHeight: 44 }}
        >
          <X size={16} aria-hidden="true" />
        </button>
      </div>

      <button
        type="button"
        className={`mail-merge-dialog__nav-btn ${parts.button}`}
        onClick={returnToRecipients}
        style={{ minHeight: 44 }}
        data-testid="mail-merge-finish-back"
      >
        <DocumentDialogIcon icon="nav.arrow.left" />
        {copy.back}
      </button>

      <fieldset
        className={`mail-merge-dialog__section ${parts.group}`}
        style={{ marginInline: 0 }}
      >
        <legend>{copy.mergeRecords}</legend>
        <div className="mail-merge-dialog__checks">
          <label style={{ minHeight: 44 }}>
            <input
              ref={finishScopeRef}
              type="radio"
              name="mail-merge-finish-scope"
              checked={mode === "all"}
              onChange={() => chooseMode("all")}
              data-testid="mail-merge-scope-all"
            />
            <DocumentDialogIcon icon="file.multiple" />
            {copy.all}
          </label>
          <label style={{ minHeight: 44 }}>
            <input
              type="radio"
              name="mail-merge-finish-scope"
              checked={mode === "current"}
              disabled={currentPlan.status !== "ready"}
              onChange={() => chooseMode("current")}
              aria-describedby={
                currentPlan.status === "ready"
                  ? undefined
                  : "mail-merge-current-unavailable"
              }
              data-testid="mail-merge-scope-current"
            />
            <DocumentDialogIcon icon="file.document" />
            {copy.current}
          </label>
          <label style={{ minHeight: 44 }}>
            <input
              type="radio"
              name="mail-merge-finish-scope"
              checked={mode === "range"}
              onChange={() => chooseMode("range")}
              data-testid="mail-merge-scope-range"
            />
            <DocumentDialogIcon icon="calendar.range" />
            {copy.range}
          </label>
        </div>
        {currentPlan.status !== "ready" ? (
          <p
            id="mail-merge-current-unavailable"
            className={`mail-merge-dialog__hint ${parts.hint}`}
          >
            {copy.currentUnavailable}
          </p>
        ) : null}
        {mode === "range" ? (
          <div className="mail-merge-dialog__grid">
            <label>
              <span className={`mail-merge-dialog__label ${parts.label}`}>{copy.from}</span>
              <input
                type="number"
                inputMode="numeric"
                min={1}
                max={eligibleCount}
                step={1}
                value={from}
                onChange={(event) => {
                  setFrom(event.target.value);
                  setActionError("");
                  setEmailConfirmationOpen(false);
                }}
                aria-invalid={plan.status === "invalid"}
                aria-describedby="mail-merge-finish-status"
                style={{ minHeight: 44 }}
                data-testid="mail-merge-range-from"
              />
            </label>
            <label>
              <span className={`mail-merge-dialog__label ${parts.label}`}>{copy.to}</span>
              <input
                type="number"
                inputMode="numeric"
                min={1}
                max={eligibleCount}
                step={1}
                value={to}
                onChange={(event) => {
                  setTo(event.target.value);
                  setActionError("");
                  setEmailConfirmationOpen(false);
                }}
                aria-invalid={plan.status === "invalid"}
                aria-describedby="mail-merge-finish-status"
                style={{ minHeight: 44 }}
                data-testid="mail-merge-range-to"
              />
            </label>
          </div>
        ) : null}
      </fieldset>

      <p
        id="mail-merge-finish-status"
        className={`mail-merge-dialog__hint ${parts.hint}`}
        role={plan.status === "ready" ? "status" : "alert"}
        data-testid="mail-merge-finish-status"
      >
        {statusMessage}
      </p>
      {actionError ? (
        <p className={`mail-merge-dialog__hint ${parts.hint}`} role="alert">
          {actionError}
        </p>
      ) : null}

      <div className="mail-merge-dialog__actions">
        <button
          type="button"
          className={`mail-merge-dialog__finish ${parts.primaryButton}`}
          disabled={!canRun}
          onClick={() => void runAction("documents", onFinish)}
          style={{ minHeight: 44 }}
          data-testid="mail-merge-edit-documents"
        >
          <DocumentDialogIcon icon="notebook.edit" />
          {busyAction === "documents" ? copy.working : copy.editDocuments}
        </button>
        <button
          type="button"
          className={`mail-merge-dialog__nav-btn ${parts.button}`}
          disabled={!canRun || !printReady}
          onClick={() => printJob && void runAction("print", printJob.submit)}
          aria-describedby={printReady ? undefined : "mail-merge-print-reason"}
          title={printReady ? undefined : copy.printUnavailable}
          style={{ minHeight: 44 }}
          data-testid="mail-merge-print-documents"
        >
          <DocumentDialogIcon icon="file.print" />
          {busyAction === "print" ? copy.working : copy.printDocuments}
        </button>
        <button
          type="button"
          className={`mail-merge-dialog__nav-btn ${parts.button}`}
          disabled={!canRun || !emailAuthorized}
          onClick={() => {
            if (!canEdit) return;
            setEmailConfirmationOpen(true);
          }}
          aria-describedby={
            canEdit
              ? emailReady
                ? undefined
                : "mail-merge-email-reason"
              : "mail-merge-email-readonly-reason"
          }
          title={
            canEdit
              ? emailReady
                ? undefined
                : copy.emailUnavailable
              : copy.emailReadOnly
          }
          style={{ minHeight: 44 }}
          data-testid="mail-merge-send-email"
        >
          <DocumentDialogIcon icon="mail.envelope" />
          {copy.sendEmail}
        </button>
      </div>
      {!printReady ? (
        <p id="mail-merge-print-reason" className={`mail-merge-dialog__hint ${parts.hint}`}>
          {copy.printUnavailable}
        </p>
      ) : null}
      {canEdit && !emailReady ? (
        <p id="mail-merge-email-reason" className={`mail-merge-dialog__hint ${parts.hint}`}>
          {copy.emailUnavailable}
        </p>
      ) : null}
      {!canEdit ? (
        <p
          id="mail-merge-email-readonly-reason"
          className={`mail-merge-dialog__hint ${parts.hint}`}
        >
          {copy.emailReadOnly}
        </p>
      ) : null}

      {emailConfirmationOpen && emailConnector && canEdit ? (
        <div
          className={`mail-merge-dialog__section ${parts.group}`}
          role="alertdialog"
          aria-label={copy.sendEmail}
          data-testid="mail-merge-email-confirmation"
        >
          <p className={`mail-merge-dialog__hint ${parts.hint}`}>{copy.emailConfirm}</p>
          <div className="mail-merge-dialog__actions">
            <button
              type="button"
              className={`mail-merge-dialog__finish ${parts.primaryButton}`}
              disabled={!canRun || !canEdit}
              onClick={() =>
                void runAction("email", emailConnector.sendAfterConfirmation)
              }
              style={{ minHeight: 44 }}
              data-testid="mail-merge-confirm-send-email"
            >
              {busyAction === "email" ? copy.working : copy.confirmSend}
            </button>
            <button
              type="button"
              className={`mail-merge-dialog__nav-btn ${parts.button}`}
              disabled={busyAction !== null}
              onClick={() => setEmailConfirmationOpen(false)}
              style={{ minHeight: 44 }}
            >
              {copy.cancel}
            </button>
          </div>
        </div>
      ) : null}
    </div>
  );
}

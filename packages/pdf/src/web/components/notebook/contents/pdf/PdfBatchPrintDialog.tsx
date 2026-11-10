"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ChevronDown,
  ChevronUp,
  FilePlus2,
  Printer,
  Trash2,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type RefObject,
} from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  inspectPdfBatchPrintItem,
  parsePdfBatchPrintRange,
  PDF_BATCH_PRINT_MAX_AGGREGATE_BYTES,
  PDF_BATCH_PRINT_MAX_FILES,
  PDF_BATCH_PRINT_MAX_SOURCE_BYTES,
  releasePdfBatchPrintItems,
  type PdfBatchPrintItem,
} from "../../../../lib/pdf/batch-print";

interface PdfBatchPrintDialogProps {
  readonly open: boolean;
  readonly seed: PdfBatchPrintItem | null;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onCancel: () => void;
  readonly onPrint: (items: readonly PdfBatchPrintItem[]) => Promise<boolean>;
}

/** A PDF starts with "%PDF-" within its first 1024 bytes (ISO 32000-1
 * §7.5.2 lets readers accept leading junk that far). */
function hasPdfHeader(bytes: Uint8Array): boolean {
  const head = bytes.subarray(0, Math.min(bytes.length, 1024 + 5));
  for (let index = 0; index + 5 <= head.length; index += 1) {
    if (
      head[index] === 0x25 &&
      head[index + 1] === 0x50 &&
      head[index + 2] === 0x44 &&
      head[index + 3] === 0x46 &&
      head[index + 4] === 0x2d
    ) {
      return true;
    }
  }
  return false;
}

function PdfBatchPrintDialogSession({
  seed,
  returnFocusRef,
  onCancel,
  onPrint,
}: Omit<PdfBatchPrintDialogProps, "open">): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const [items, setItems] = useState<readonly PdfBatchPrintItem[]>(
    seed ? [seed] : [],
  );
  const [busy, setBusy] = useState(false);
  const [reading, setReading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const generationRef = useRef(1);
  const itemsRef = useRef(items);
  itemsRef.current = items;

  // The queue's bytes are wiped when the dialog really goes away. React's
  // StrictMode (every `next dev` run) unmounts and mounts once more at
  // first render with the same state, so a wipe in the cleanup itself
  // zeroed the current PDF before it was ever printed ("No PDF header
  // found"). The wipe waits a task, and a mount that follows cancels it.
  const releaseTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(() => {
    if (releaseTimerRef.current !== null) {
      clearTimeout(releaseTimerRef.current);
      releaseTimerRef.current = null;
    }
    return () => {
      generationRef.current += 1;
      releaseTimerRef.current = setTimeout(() => {
        releaseTimerRef.current = null;
        releasePdfBatchPrintItems(itemsRef.current);
        itemsRef.current = [];
      }, 0);
    };
  }, []);

  const rangeErrors = useMemo(
    () =>
      items.map((item) => {
        try {
          parsePdfBatchPrintRange(item.range, item.source.pageCount);
          return null;
        } catch (rangeError) {
          return rangeError instanceof Error
            ? rangeError.message
            : "Invalid page range";
        }
      }),
    [items],
  );
  const selectedPageCount = useMemo(
    () =>
      items.reduce((total, item, index) => {
        if (rangeErrors[index]) return total;
        return (
          total +
          parsePdfBatchPrintRange(item.range, item.source.pageCount).length
        );
      }, 0),
    [items, rangeErrors],
  );
  const invalid =
    items.length === 0 || rangeErrors.some(Boolean) || selectedPageCount < 1;

  const close = () => {
    if (busy) return;
    generationRef.current += 1;
    onCancel();
  };

  const pickFiles = async (event: ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.currentTarget.files ?? []);
    event.currentTarget.value = "";
    if (files.length === 0 || reading || busy) return;
    if (items.length + files.length > PDF_BATCH_PRINT_MAX_FILES) {
      setError(
        t("pdf.batchPrint.tooManyFiles", { max: PDF_BATCH_PRINT_MAX_FILES }),
      );
      return;
    }
    let declaredBytes = items.reduce(
      (total, item) => total + item.bytes.byteLength,
      0,
    );
    // capability:label-only — input validation on files the user just picked
    // from their disk for a local batch print. None is a workspace file, so
    // there is no id for useFileCapability to resolve. A file is judged by
    // its content, not its name: a real PDF without ".pdf" is a PDF, and a
    
    for (const file of files) {
      if (
        !Number.isSafeInteger(file.size) ||
        file.size < 1 ||
        file.size > PDF_BATCH_PRINT_MAX_SOURCE_BYTES
      ) {
        setError(t("pdf.batchPrint.fileInvalid"));
        return;
      }
      declaredBytes += file.size;
      if (declaredBytes > PDF_BATCH_PRINT_MAX_AGGREGATE_BYTES) {
        setError(t("pdf.batchPrint.totalTooLarge"));
        return;
      }
    }
    const generation = ++generationRef.current;
    setReading(true);
    setError(null);
    const added: PdfBatchPrintItem[] = [];
    try {
      for (let index = 0; index < files.length; index += 1) {
        const file = files[index];
        const bytes = new Uint8Array(await file.arrayBuffer());
        if (generationRef.current !== generation) {
          bytes.fill(0);
          throw new Error("Batch print file selection changed");
        }
        if (!hasPdfHeader(bytes)) {
          bytes.fill(0);
          throw new Error(t("pdf.batchPrint.notPdf", { name: file.name }));
        }
        added.push(
          await inspectPdfBatchPrintItem({
            id: `local:${generation}:${index}:${file.name}`,
            name: file.name,
            bytes,
            dataItemId: `local-pdf:batch-print:${generation}:${index}`,
            byteGenerationKey: `batch-print-picker:${generation}:${index}`,
            documentRevision: generation,
          }),
        );
        bytes.fill(0);
      }
      if (generationRef.current !== generation) {
        releasePdfBatchPrintItems(added);
        return;
      }
      setItems((current) => [...current, ...added]);
    } catch (readError) {
      releasePdfBatchPrintItems(added);
      if (generationRef.current === generation) {
        setError(
          readError instanceof Error
            ? readError.message
            : t("pdf.batchPrint.readFailed"),
        );
      }
    } finally {
      if (generationRef.current === generation) setReading(false);
    }
  };

  const remove = (index: number) => {
    if (busy || reading) return;
    setItems((current) => {
      const target = current[index];
      if (target) target.bytes.fill(0);
      return current.filter((_, itemIndex) => itemIndex !== index);
    });
  };

  const move = (index: number, direction: -1 | 1) => {
    if (busy || reading) return;
    setItems((current) => {
      const destination = index + direction;
      if (destination < 0 || destination >= current.length) return current;
      const next = [...current];
      [next[index], next[destination]] = [next[destination], next[index]];
      return next;
    });
  };

  const updateRange = (index: number, range: string) => {
    setItems((current) =>
      current.map((item, itemIndex) =>
        itemIndex === index ? Object.freeze({ ...item, range }) : item,
      ),
    );
  };

  const submit = async () => {
    if (busy || reading || invalid) return;
    const snapshot = itemsRef.current;
    setBusy(true);
    setError(null);
    try {
      if (await onPrint(snapshot)) onCancel();
      else {
        setError(t("pdf.batchPrint.startFailed"));
      }
    } catch {
      // Ranges are checked before this; what fails past them (a reader, a
      // size limit, the print frame) reads as the print not starting.
      setError(t("pdf.batchPrint.startFailed"));
    } finally {
      setBusy(false);
    }
  };

  return (
    <OfficeDialogShell
      open
      title={t("pdf.batchPrint.title")}
      description={t("pdf.batchPrint.description")}
      defaultActionLabel={t("pdf.batchPrint.printQueue")}
      cancelActionLabel={t("pdf.batchPrint.cancel")}
      defaultActionDisabled={invalid || reading || busy}
      busy={busy}
      busyActionLabel={t("pdf.batchPrint.preparing")}
      initialFocus="cancel"
      returnFocusRef={returnFocusRef}
      onDefaultAction={() => void submit()}
      onCancel={close}
      testIdPrefix="pdf-batch-print"
      className="sm:max-w-[760px]"
    >
      <div className="grid gap-3">
        <input
          ref={fileInputRef}
          type="file"
          accept=".pdf,application/pdf"
          multiple
          className="hidden"
          aria-hidden="true"
          tabIndex={-1}
          data-testid="pdf-batch-print-file-input"
          onChange={(event) => void pickFiles(event)}
        />
        <div className="flex flex-wrap items-center justify-between gap-2">
          <p className={`${parts.hint} m-0`} role="status">
            {t("pdf.batchPrint.queueSummary", {
              files: items.length,
              pages: selectedPageCount,
            })}
          </p>
          <button
            type="button"
            className={`${parts.compactButton} flex items-center gap-2`}
            disabled={
              busy || reading || items.length >= PDF_BATCH_PRINT_MAX_FILES
            }
            onClick={() => fileInputRef.current?.click()}
            data-testid="pdf-batch-print-add-files"
          >
            <FilePlus2 size={18} aria-hidden />
            {reading
              ? t("pdf.batchPrint.inspecting")
              : t("pdf.batchPrint.addPdfs")}
          </button>
        </div>

        <ol
          className="m-0 grid max-h-[360px] list-none gap-2 overflow-y-auto p-0 pe-1"
          aria-label={t("pdf.batchPrint.printQueue")}
          data-testid="pdf-batch-print-queue"
        >
          {items.map((item, index) => (
            <li
              key={item.id}
              className={`${parts.well} grid gap-2 sm:grid-cols-[minmax(0,1fr)_180px_auto] sm:items-center`}
              data-testid={`pdf-batch-print-item-${index}`}
            >
              <div className="min-w-0">
                <p
                  className="m-0 truncate text-sm font-medium"
                  title={item.name}
                >
                  {index + 1}. {item.name}
                </p>
                <p className={`${parts.hint} m-0`}>
                  {t("pdf.batchPrint.sourceVerified", {
                    count: item.source.pageCount,
                  })}
                </p>
              </div>
              <label className={parts.field}>
                {t("pdf.batchPrint.pageRange")}
                <input
                  type="text"
                  value={item.range}
                  placeholder={t("pdf.batchPrint.pageRangePlaceholder")}
                  disabled={busy || reading}
                  aria-invalid={Boolean(rangeErrors[index])}
                  aria-describedby={
                    rangeErrors[index]
                      ? `pdf-batch-print-range-error-${index}`
                      : undefined
                  }
                  data-testid={`pdf-batch-print-range-${index}`}
                  onChange={(event) =>
                    updateRange(index, event.currentTarget.value)
                  }
                />
                {rangeErrors[index] ? (
                  <span
                    id={`pdf-batch-print-range-error-${index}`}
                    className={parts.error}
                    role="alert"
                  >
                    {t("pdf.batchPrint.pageRangeInvalid")}
                  </span>
                ) : null}
              </label>
              <div className="flex items-center justify-end gap-1">
                <button
                  type="button"
                  className={`${parts.iconButton} flex items-center justify-center`}
                  aria-label={t("pdf.batchPrint.moveUp")}
                  disabled={busy || reading || index === 0}
                  onClick={() => move(index, -1)}
                  data-testid={`pdf-batch-print-move-up-${index}`}
                >
                  <ChevronUp size={18} aria-hidden />
                </button>
                <button
                  type="button"
                  className={`${parts.iconButton} flex items-center justify-center`}
                  aria-label={t("pdf.batchPrint.moveDown")}
                  disabled={busy || reading || index === items.length - 1}
                  onClick={() => move(index, 1)}
                  data-testid={`pdf-batch-print-move-down-${index}`}
                >
                  <ChevronDown size={18} aria-hidden />
                </button>
                <button
                  type="button"
                  className={`${parts.compactButton} flex items-center justify-center`} data-tone="danger"
                  aria-label={t("pdf.batchPrint.remove")}
                  disabled={busy || reading}
                  onClick={() => remove(index)}
                  data-testid={`pdf-batch-print-remove-${index}`}
                >
                  <Trash2 size={18} aria-hidden />
                </button>
              </div>
            </li>
          ))}
        </ol>

        {items.length === 0 ? (
          <div className={`${parts.well} ${parts.hint} grid min-h-24 place-items-center text-center`}>
            <span>{t("pdf.batchPrint.emptyQueue")}</span>
          </div>
        ) : null}
        {error ? (
          <p
            className={`${parts.message} m-0`}
            role="alert"
            data-testid="pdf-batch-print-error"
          >
            {error}
          </p>
        ) : null}
        <p className={`${parts.hint} m-0 flex items-center gap-2`}>
          <Printer size={16} aria-hidden />
          {t("pdf.batchPrint.privacyHint")}
        </p>
      </div>
    </OfficeDialogShell>
  );
}

export default function PdfBatchPrintDialog(
  props: PdfBatchPrintDialogProps,
): React.JSX.Element | null {
  if (!props.open || !props.seed) return null;
  return (
    <PdfBatchPrintDialogSession
      key={`${props.seed.id}:${props.open ? "open" : "closed"}`}
      seed={props.seed}
      returnFocusRef={props.returnFocusRef}
      onCancel={props.onCancel}
      onPrint={props.onPrint}
    />
  );
}

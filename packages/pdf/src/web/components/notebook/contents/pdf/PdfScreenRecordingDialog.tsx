"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Check } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useState, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

export interface PdfScreenRecordingDialogProps {
  readonly open: boolean;
  readonly busy: boolean;
  readonly supported: boolean;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onStart: (includeSharedAudio: boolean) => void;
  readonly onCancel: () => void;
}

export default function PdfScreenRecordingDialog({
  open,
  busy,
  supported,
  returnFocusRef,
  onStart,
  onCancel,
}: PdfScreenRecordingDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const [includeSharedAudio, setIncludeSharedAudio] = useState(true);

  return (
    <OfficeDialogShell
      open={open}
      title={t("pdf.screenRecording.title")}
      description={t("pdf.screenRecording.description")}
      defaultActionLabel={t("pdf.screenRecording.start")}
      cancelActionLabel={t("common.cancel")}
      busyActionLabel={t("pdf.screenRecording.busy")}
      busy={busy}
      defaultActionDisabled={!supported}
      onDefaultAction={() => onStart(includeSharedAudio)}
      onCancel={onCancel}
      initialFocus="cancel"
      returnFocusRef={returnFocusRef}
      testIdPrefix="pdf-screen-recording"
    >
      <label className={`${parts.check} cursor-pointer min-h-11`}>
        <span className="relative mt-0.5 size-4 shrink-0">
          <input
            type="checkbox"
            checked={includeSharedAudio}
            disabled={busy || !supported}
            onChange={(event) => setIncludeSharedAudio(event.target.checked)}
            data-testid="pdf-screen-recording-audio"
            className={`${parts.markInput} absolute -left-3 -top-3.5 size-11 cursor-pointer`}
          />
          <span
            aria-hidden
            className={`${parts.mark} pointer-events-none absolute inset-0 size-4`}
          />
          <Check
            aria-hidden
            className="pointer-events-none absolute left-1/2 top-1/2 size-3 -translate-x-1/2 -translate-y-1/2"
          />
        </span>
        <span>
          <strong>
            {t("pdf.screenRecording.audio")}
          </strong>
          <span className={`${parts.hint} mt-1 block`}>
            {t("pdf.screenRecording.audioHint")}
          </span>
        </span>
      </label>

      <div className={parts.group}>
        {t("pdf.screenRecording.note")}
      </div>

      {!supported ? (
        <p
          role="alert"
          data-testid="pdf-screen-recording-unavailable"
          className={`${parts.error} m-0`}
        >
          {t("pdf.screenRecording.unsupported")}
        </p>
      ) : null}
    </OfficeDialogShell>
  );
}

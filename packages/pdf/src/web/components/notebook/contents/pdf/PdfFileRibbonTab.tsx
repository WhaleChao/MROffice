"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  Download,
  Files,
  Info,
  Printer,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useRef } from "react";

import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";

interface PdfFileCommand {
  readonly label: string;
  readonly disabled?: boolean;
  readonly disabledReason?: string;
  readonly busy?: boolean;
  readonly onInvoke?: (trigger: HTMLButtonElement | null) => void;
  readonly testId: string;
}

interface PdfFileRibbonTabProps {
  readonly visible: boolean;
  readonly ribbon?: string;
  readonly error?: { readonly title: string; readonly message: string } | null;
  readonly outputGroupLabel: string;
  readonly documentGroupLabel: string;
  readonly print: PdfFileCommand;
  readonly batchPrint: PdfFileCommand;
  readonly downloadOriginal: PdfFileCommand;
  readonly properties: PdfFileCommand;
}

function commandAccessibleLabel(command: PdfFileCommand): string {
  if (command.busy) return command.label;
  if ((command.disabled || !command.onInvoke) && command.disabledReason) {
    return `${command.label} — ${command.disabledReason}`;
  }
  return command.label;
}

export default function PdfFileRibbonTab({
  visible,
  ribbon = "file",
  error,
  outputGroupLabel,
  documentGroupLabel,
  print,
  batchPrint,
  downloadOriginal,
  properties,
}: PdfFileRibbonTabProps): React.JSX.Element | null {
  const printRef = useRef<HTMLButtonElement>(null);
  const batchPrintRef = useRef<HTMLButtonElement>(null);
  const downloadRef = useRef<HTMLButtonElement>(null);
  const propertiesRef = useRef<HTMLButtonElement>(null);
  if (!visible) return null;
  return (
    <>
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={outputGroupLabel}
        size="content"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommand
            buttonRef={printRef}
            disabled={print.disabled || print.busy || !print.onInvoke}
            displayLabel={print.label}
            icon={<Printer size={24} aria-hidden />}
            label={commandAccessibleLabel(print)}
            onClick={() => print.onInvoke?.(printRef.current)}
            testId={print.testId}
            variant="large"
          />
          <OfficeHomeCommand
            buttonRef={batchPrintRef}
            disabled={
              batchPrint.disabled || batchPrint.busy || !batchPrint.onInvoke
            }
            displayLabel={batchPrint.label}
            icon={<Files size={24} aria-hidden />}
            label={commandAccessibleLabel(batchPrint)}
            onClick={() => batchPrint.onInvoke?.(batchPrintRef.current)}
            testId={batchPrint.testId}
            variant="large"
          />
        </OfficeHomeCommandRow>
        {error ? (
          <span
            className="pdf-viewer__flatten-error"
            role="alert"
            title={error.title}
            data-testid="pdf-file-output-error"
          >
            {error.message}
          </span>
        ) : null}
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup
        ribbon={ribbon}
        label={documentGroupLabel}
        size="content"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommand
            buttonRef={downloadRef}
            disabled={
              downloadOriginal.disabled ||
              downloadOriginal.busy ||
              !downloadOriginal.onInvoke
            }
            displayLabel={downloadOriginal.label}
            icon={<Download size={24} aria-hidden />}
            label={commandAccessibleLabel(downloadOriginal)}
            onClick={() => downloadOriginal.onInvoke?.(downloadRef.current)}
            testId={downloadOriginal.testId}
            variant="large"
          />
          <OfficeHomeCommand
            buttonRef={propertiesRef}
            disabled={
              properties.disabled || properties.busy || !properties.onInvoke
            }
            displayLabel={properties.label}
            icon={<Info size={24} aria-hidden />}
            label={commandAccessibleLabel(properties)}
            onClick={() => properties.onInvoke?.(propertiesRef.current)}
            testId={properties.testId}
            variant="large"
          />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
    </>
  );
}

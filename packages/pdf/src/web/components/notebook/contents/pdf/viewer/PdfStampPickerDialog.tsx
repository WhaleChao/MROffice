"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import insertRibbonStyles from "../PdfInsertRibbonTab.module.css";
import { useRef, type Dispatch, type SetStateAction, type RefObject } from "react";
import {
  Dialog,
  DialogContent,
  DialogTitle,
  DialogClose,
  DialogDescription,
} from "@mrnote/core/web/components/ui/dialog.tsx";
import { PDF_POPOVER_CHROME } from "./popoverChrome";
import {
  officeDialogParts,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  OfficeGalleryMenuGrid,
  OfficeGalleryMenuTileBody,
  officeGalleryMenuStyles,
} from "@mrnote/core/web/components/office-ribbon/OfficeGalleryMenu.tsx";
import {
  X,
  Image as ImageIcon,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { STAMP_PRESETS } from "../../../../../lib/pdf/stamps-annotation";
import {
  SIGNATURE_MAX_IMAGE_BYTES,
  SIGNATURE_UPLOAD_MIME_TYPES,
  validateSignatureImageFile,
  sanitizeSignatureImage,
} from "@mrnote/core/web/lib/pdf/signatures.ts";
import { fitOfficePictureBlob } from "@mrnote/core/web/lib/office-picture-envelope.ts";
import { pdfStampPictureForWriter } from "../../../../../lib/pdf/stamp-picture-check";
import type { PdfManagedStampCommandOrigin } from "./readerStamp";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfStampPickerDialogProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  stampPickerOpen: boolean;
  setStampPickerOpen: Dispatch<SetStateAction<boolean>>;
  setStampPickerImageOnly: Dispatch<SetStateAction<boolean>>;
  stampPickerImageOnly: boolean;
  stampPickerReturnFocusRef: RefObject<HTMLButtonElement | null>;
  setArmedStamp: Dispatch<SetStateAction<{ kind: "preset"; preset: string; origin: PdfManagedStampCommandOrigin; } | { kind: "image"; image: string; aspectRatio: number; origin: PdfManagedStampCommandOrigin; } | null>>;
  stampPickerOrigin: PdfManagedStampCommandOrigin;
  setStampUploadError: Dispatch<SetStateAction<boolean>>;
  stampUploadError: boolean;
  stampUploadRunRef: RefObject<number>;
}

export function PdfStampPickerDialog({
  t,
  stampPickerOpen,
  setStampPickerOpen,
  setStampPickerImageOnly,
  stampPickerImageOnly,
  stampPickerReturnFocusRef,
  setArmedStamp,
  stampPickerOrigin,
  setStampUploadError,
  stampUploadError,
  stampUploadRunRef,
}: PdfStampPickerDialogProps) {
  // The upload entry is a button: a <label> around a hidden file input took
  // no keyboard focus, so Tab never reached it (user audit PDF 0014).
  const stampUploadInputRef = useRef<HTMLInputElement | null>(null);
  return (
    stampPickerOpen ? (
      <Dialog
        open={stampPickerOpen}
        onOpenChange={(open) => {
          if (!open) stampUploadRunRef.current += 1;
          setStampPickerOpen(open);
          if (!open) {
            setStampPickerImageOnly(false);
            setStampUploadError(false);
          }
        }}
      >
        <DialogContent
          skin="none"
          className={`${PDF_POPOVER_CHROME} ${insertRibbonStyles.decorationDialog}`}
          showCloseButton={false}
          data-pdf-ribbon-surface="annotate"
          data-testid="pdf-stamp-picker"
          onCloseAutoFocus={(event) => {
            event.preventDefault();
            const trigger = stampPickerReturnFocusRef.current;
            stampPickerReturnFocusRef.current = null;
            trigger?.focus();
          }}
        >
          <div className={officeDialogParts.header}>
            <DialogTitle className={officeDialogParts.title}>
              {stampPickerImageOnly
                ? t("pdf.insertImage")
                : t("pdfStampTool.pickerLabel")}
            </DialogTitle>
            <DialogClose
              className={officeDialogParts.close}
              aria-label={t("common.close")}
            >
              <X size={14} aria-hidden="true" />
            </DialogClose>
          </div>
          <DialogDescription className="sr-only">
            {stampPickerImageOnly
              ? t("pdf.insertImage")
              : t("pdfStampTool.pickerLabel")}
          </DialogDescription>
          <div
            className={insertRibbonStyles.decorationForm}
            role="group"
            aria-label={t("pdfStampTool.pickerLabel")}
            style={{ display: "flex", flexDirection: "column", gap: 8 }}
          >
            {

}
            {stampPickerImageOnly ? null : (
              <OfficeGalleryMenuGrid columns={2} tileWidth={132} testId="pdf-stamp-preset-grid">
                {STAMP_PRESETS.map((preset) => (
                  <button
                    key={preset.key}
                    type="button"
                    className={officeGalleryMenuStyles.tile}
                    data-testid="pdf-stamp-preset"
                    data-preset={preset.key}
                    aria-label={preset.text}
                    title={preset.text}
                    onClick={() => {
                      stampUploadRunRef.current += 1;
                      setArmedStamp({
                        kind: "preset",
                        preset: preset.key,
                        origin: stampPickerOrigin,
                      });
                      setStampPickerOpen(false);
                    }}
                  >
                    <OfficeGalleryMenuTileBody
                      thumbnail={
                        <span
                          style={{
                            display: "inline-block",
                            padding: "3px 8px",
                            fontSize: 12,
                            fontWeight: 700,
                            whiteSpace: "nowrap",
                            color: preset.color,
                            border: `2px solid ${preset.color}`,
                            borderRadius: 6,
                          }}
                        >
                          {preset.text}
                        </span>
                      }
                      name={preset.text}
                    />
                  </button>
                ))}
              </OfficeGalleryMenuGrid>
            )}
            <button
              type="button"
              className={officeDialogParts.compactButton}
              data-testid="pdf-stamp-upload-button"
              onClick={() => stampUploadInputRef.current?.click()}
              style={{ alignSelf: "flex-start" }}
            >
              <ImageIcon size={13} aria-hidden />
              {t("pdfStampTool.upload")}
            </button>
              <input
                ref={stampUploadInputRef}
                type="file"
                accept={SIGNATURE_UPLOAD_MIME_TYPES.join(",")}
                data-testid="pdf-stamp-upload"
                tabIndex={-1}
                aria-hidden
                style={{ display: "none" }}
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  event.target.value = "";
                  if (!file) return;
                  const stampUploadRun = ++stampUploadRunRef.current;
                  const check = validateSignatureImageFile(file);
                  // A picture past the stamp's 1 MB is fitted to it (decoded
                  // from the file and re-encoded smaller), as a picture an
                  // Excel sheet or a slide holds inline is, not refused.
                  const fitted =
                    !check.ok && check.reason === "size"
                      ? fitOfficePictureBlob(
                          file,
                          file.type.toLowerCase() === "image/png" ? "image/png" : "image/jpeg",
                          undefined,
                          Math.floor(SIGNATURE_MAX_IMAGE_BYTES / 3) * 4 + 32,
                        )
                      : null;
                  if (!check.ok && !fitted) {
                    setStampUploadError(true);
                    return;
                  }
                  type FittedReader = {
                    result: unknown;
                    onload: (() => void) | null;
                    onerror: (() => void) | null;
                    readAsDataURL: (blob: Blob) => void;
                  };
                  const reader: FittedReader | FileReader = fitted
                    ? {
                        result: null,
                        onload: null,
                        onerror: null,
                        readAsDataURL(this: FittedReader) {
                          void fitted.then(
                            (dataUrl) => {
                              this.result = dataUrl;
                              if (dataUrl) this.onload?.();
                              else this.onerror?.();
                            },
                            () => this.onerror?.(),
                          );
                        },
                      }
                    : new FileReader();
                  reader.onload = () => {
                    if (stampUploadRunRef.current !== stampUploadRun) return;
                    const image = sanitizeSignatureImage(reader.result);
                    if (!image) {
                      setStampUploadError(true);
                      return;
                    }
                    const preview = new Image();
                    preview.onload = () => {
                      if (stampUploadRunRef.current !== stampUploadRun)
                        return;
                      const aspectRatio =
                        preview.naturalWidth / preview.naturalHeight;
                      if (!Number.isFinite(aspectRatio) || aspectRatio <= 0) {
                        setStampUploadError(true);
                        return;
                      }
                      // The picture as the PDF's writer takes it: a file a
                      // browser draws though its data is damaged or cut
                      // short (or larger than the writer reads) is put on
                      // the page as the browser draws it, encoded again;
                      // taken as it was, the writer refused its row, the
                      // page's PDF was not written and every later write
                      // was refused with it (stamp-picture-check.ts).
                      void pdfStampPictureForWriter({
                        dataUrl: image,
                        width: preview.naturalWidth,
                        height: preview.naturalHeight,
                        maxBytes: SIGNATURE_MAX_IMAGE_BYTES,
                      }).then((taken) => {
                        if (stampUploadRunRef.current !== stampUploadRun)
                          return;
                        const picture = taken
                          ? sanitizeSignatureImage(taken)
                          : null;
                        if (!picture) {
                          setStampUploadError(true);
                          return;
                        }
                        setStampUploadError(false);
                        setArmedStamp({
                          kind: "image",
                          image: picture,
                          aspectRatio,
                          origin: stampPickerOrigin,
                        });
                        setStampPickerImageOnly(false);
                        setStampPickerOpen(false);
                      });
                    };
                    preview.onerror = () => {
                      if (stampUploadRunRef.current !== stampUploadRun)
                        return;
                      setStampUploadError(true);
                    };
                    preview.src = image;
                  };
                  reader.onerror = () => {
                    if (stampUploadRunRef.current !== stampUploadRun) return;
                    setStampUploadError(true);
                  };
                  reader.readAsDataURL(file);
                }}
              />
            {stampUploadError ? (
              <span
                role="alert"
                data-testid="pdf-stamp-upload-error"
                className={officeDialogParts.error}
              >
                {t("pdfStampTool.uploadError")}
              </span>
            ) : null}
          </div>
        </DialogContent>
      </Dialog>
    ) : null
  );
}

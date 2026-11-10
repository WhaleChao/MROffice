"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Picture Format tab while a background-removal or erase session runs: its
 * pens, Discard All and Keep Changes, drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { OfficeHomeRibbonGroup, OfficeHomeCommandRow, OfficeHomeCommand } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { BackgroundRemovalSegment, BackgroundRemovalSpec } from "@mrnote/core/web/lib/office-picture-background-removal.ts";
import type { OfficePictureEditPen } from "@mrnote/core/web/components/office-ribbon/OfficePictureEditOverlay.tsx";
import type { Dispatch, SetStateAction } from "react";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationPictureEditSessionRibbonGroupCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  pictureEditSession: { blockId: string; slideId: string; mode: "removal" | "erase"; baseSrc: string; originalSrc?: string | null; spec: BackgroundRemovalSpec; pen: OfficePictureEditPen; eraseSegments: BackgroundRemovalSegment[]; busy: boolean; };
  setPictureEditSession: Dispatch<SetStateAction<{ blockId: string; slideId: string; mode: "removal" | "erase"; baseSrc: string; originalSrc?: string | null; spec: BackgroundRemovalSpec; pen: OfficePictureEditPen; eraseSegments: BackgroundRemovalSegment[]; busy: boolean; } | null>>;
  discardPictureEdit: () => void;
  keepPictureEdit: () => Promise<void>;
}

export function PresentationPictureEditSessionRibbonGroup({ ctl }: { ctl: PresentationPictureEditSessionRibbonGroupCtl }) {
  const {
    t,
    pictureEditSession,
    setPictureEditSession,
    discardPictureEdit,
    keepPictureEdit,
  } = ctl;
  return (
                <OfficeHomeRibbonGroup
                  label={t(
                    pictureEditSession.mode === "removal"
                      ? "presentation.image.backgroundRemovalGroup"
                      : "presentation.image.erase",
                  )}
                  size="standard"
                  testId="presentation-ribbon-picture-edit-session"
                >
                  <OfficeHomeCommandRow>
                    {pictureEditSession.mode === "removal" ? (
                      <>
                        <OfficeHomeCommand
                          variant="large"
                          active={pictureEditSession.pen === "keep"}
                          icon={<OfficeSemanticIcon semanticId="draw.pen" size={24} />}
                          label={t("presentation.image.markKeep")}
                          onClick={() =>
                            setPictureEditSession((current) =>
                              current ? { ...current, pen: current.pen === "keep" ? null : "keep" } : current,
                            )
                          }
                          testId="presentation-ribbon-picture-mark-keep"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          active={pictureEditSession.pen === "remove"}
                          icon={<OfficeSemanticIcon semanticId="draw.eraser.stroke" size={24} />}
                          label={t("presentation.image.markRemove")}
                          onClick={() =>
                            setPictureEditSession((current) =>
                              current ? { ...current, pen: current.pen === "remove" ? null : "remove" } : current,
                            )
                          }
                          testId="presentation-ribbon-picture-mark-remove"
                        />
                      </>
                    ) : null}
                    <OfficeHomeCommand
                      variant="large"
                      icon={<OfficeSemanticIcon semanticId="image.removeBackground" size={24} />}
                      label={t("presentation.image.discardAll")}
                      onClick={discardPictureEdit}
                      testId="presentation-ribbon-picture-discard"
                    />
                    <OfficeHomeCommand
                      variant="large"
                      disabled={
                        pictureEditSession.busy ||
                        (pictureEditSession.mode === "erase" &&
                          pictureEditSession.eraseSegments.length === 0)
                      }
                      icon={<OfficeSemanticIcon semanticId="image.edit" size={24} />}
                      label={t("presentation.image.keepChanges")}
                      onClick={() => {
                        void keepPictureEdit();
                      }}
                      testId="presentation-ribbon-picture-keep"
                    />
                  </OfficeHomeCommandRow>
                </OfficeHomeRibbonGroup>
  );
}

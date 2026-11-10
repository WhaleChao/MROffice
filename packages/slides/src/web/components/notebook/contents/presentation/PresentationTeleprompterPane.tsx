"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The teleprompter pane that scrolls the speaker notes, rendered while it is
 * open.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the render
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  StickyNote,
  X,
  Pause,
  Play,
  Minus,
  Plus,
  RotateCcw,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { TELEPROMPTER_FONT_SIZE_MIN, TELEPROMPTER_FONT_SIZE_STEP, TELEPROMPTER_FONT_SIZE_MAX } from "../../../../lib/presentation/window-constants";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationTeleprompterPaneCtl {
  pageId: string;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  activeIndex: number;
  recordStudioOpen: boolean;
  teleprompterRunning: boolean;
  setTeleprompterRunning: Dispatch<SetStateAction<boolean>>;
  teleprompterSpeed: number;
  setTeleprompterSpeed: Dispatch<SetStateAction<number>>;
  teleprompterFontSize: number;
  setTeleprompterFontSize: Dispatch<SetStateAction<number>>;
  teleprompterPanelRef: RefObject<HTMLElement | null>;
  teleprompterScrollRef: RefObject<HTMLDivElement | null>;
  activeTeleprompterNotes: string;
  closeTeleprompter: () => void;
}

export function PresentationTeleprompterPane({ ctl }: { ctl: PresentationTeleprompterPaneCtl }) {
  const {
    pageId,
    t,
    activeIndex,
    recordStudioOpen,
    teleprompterRunning,
    setTeleprompterRunning,
    teleprompterSpeed,
    setTeleprompterSpeed,
    teleprompterFontSize,
    setTeleprompterFontSize,
    teleprompterPanelRef,
    teleprompterScrollRef,
    activeTeleprompterNotes,
    closeTeleprompter,
  } = ctl;
  return (
            <aside
              ref={teleprompterPanelRef}
              className={presentationChrome.teleprompterPanel}
              role="dialog"
              aria-modal="false"
              aria-labelledby={`presentation-teleprompter-title-${pageId}`}
              tabIndex={-1}
              data-running={teleprompterRunning ? "true" : "false"}
              data-speed={teleprompterSpeed}
              data-testid="presentation-record-teleprompter-panel"
              onKeyDown={(event) => {
                if (event.key !== "Escape") return;
                event.preventDefault();
                event.stopPropagation();
                closeTeleprompter();
              }}
            >
              <header className={presentationChrome.teleprompterHeader}>
                <span className={presentationChrome.teleprompterHeading}>
                  <StickyNote size={18} aria-hidden="true" />
                  <span>
                    <strong id={`presentation-teleprompter-title-${pageId}`}>
                      {t("presentation.record.teleprompter")}
                    </strong>
                    <small>
                      {t("presentation.zoom.slideLabel", {
                        index: activeIndex + 1,
                      })}
                    </small>
                  </span>
                </span>
                <button
                  type="button"
                  className={presentationChrome.teleprompterIconButton}
                  onClick={closeTeleprompter}
                  aria-label={t("presentation.record.teleprompterClose")}
                  data-testid="presentation-record-teleprompter-close"
                >
                  <X size={18} aria-hidden="true" />
                </button>
              </header>
              <div
                ref={teleprompterScrollRef}
                className={presentationChrome.teleprompterBody}
                tabIndex={0}
                style={{ fontSize: teleprompterFontSize }}
                data-empty={!activeTeleprompterNotes.trim() || undefined}
                data-font-size={teleprompterFontSize}
                data-testid="presentation-record-teleprompter-notes"
              >
                {activeTeleprompterNotes.trim() ? (
                  activeTeleprompterNotes
                ) : (
                  <p className={presentationChrome.teleprompterEmpty}>
                    {t("presentation.record.teleprompterEmpty")}
                  </p>
                )}
              </div>
              <footer className={presentationChrome.teleprompterControls}>
                <button
                  type="button"
                  className={presentationChrome.teleprompterAction}
                  disabled={!activeTeleprompterNotes.trim()}
                  aria-pressed={teleprompterRunning}
                  onClick={() =>
                    setTeleprompterRunning((running) => !running)
                  }
                  data-testid="presentation-record-teleprompter-play"
                >
                  {teleprompterRunning ? (
                    <Pause size={16} aria-hidden="true" />
                  ) : (
                    <Play size={16} aria-hidden="true" />
                  )}
                  {t(
                    teleprompterRunning
                      ? "presentation.record.teleprompterPause"
                      : "presentation.record.teleprompterPlay",
                  )}
                </button>
                <label className={presentationChrome.teleprompterSpeed}>
                  <span>{t("presentation.record.teleprompterSpeed")}</span>
                  <input
                    type="range"
                    min={1}
                    max={5}
                    step={1}
                    value={teleprompterSpeed}
                    aria-valuetext={t(
                      "presentation.record.teleprompterSpeedValue",
                      { value: teleprompterSpeed },
                    )}
                    onChange={(event) =>
                      setTeleprompterSpeed(
                        Math.min(5, Math.max(1, Number(event.target.value))),
                      )
                    }
                    data-testid="presentation-record-teleprompter-speed"
                  />
                </label>
                {recordStudioOpen ? (
                  <>
                    <button
                      type="button"
                      className={presentationChrome.teleprompterAction}
                      data-control-id="PPTCTL.RECORD_STUDIO.NOTES.ZOOM_OUT"
                      data-testid="presentation-record-teleprompter-zoom-out"
                      disabled={
                        teleprompterFontSize <= TELEPROMPTER_FONT_SIZE_MIN
                      }
                      onClick={() =>
                        setTeleprompterFontSize((size) =>
                          Math.max(
                            TELEPROMPTER_FONT_SIZE_MIN,
                            size - TELEPROMPTER_FONT_SIZE_STEP,
                          ),
                        )
                      }
                    >
                      <Minus size={16} aria-hidden="true" />
                      {t("presentation.record.teleprompterZoomOut")}
                    </button>
                    <button
                      type="button"
                      className={presentationChrome.teleprompterAction}
                      data-control-id="PPTCTL.RECORD_STUDIO.NOTES.ZOOM_IN"
                      data-testid="presentation-record-teleprompter-zoom-in"
                      disabled={
                        teleprompterFontSize >= TELEPROMPTER_FONT_SIZE_MAX
                      }
                      onClick={() =>
                        setTeleprompterFontSize((size) =>
                          Math.min(
                            TELEPROMPTER_FONT_SIZE_MAX,
                            size + TELEPROMPTER_FONT_SIZE_STEP,
                          ),
                        )
                      }
                    >
                      <Plus size={16} aria-hidden="true" />
                      {t("presentation.record.teleprompterZoomIn")}
                    </button>
                  </>
                ) : null}
                <button
                  type="button"
                  className={presentationChrome.teleprompterAction}
                  onClick={() => {
                    setTeleprompterRunning(false);
                    if (teleprompterScrollRef.current) {
                      teleprompterScrollRef.current.scrollTop = 0;
                    }
                  }}
                  data-testid="presentation-record-teleprompter-reset"
                >
                  <RotateCcw size={16} aria-hidden="true" />
                  {t("presentation.record.teleprompterReset")}
                </button>
              </footer>
            </aside>
  );
}

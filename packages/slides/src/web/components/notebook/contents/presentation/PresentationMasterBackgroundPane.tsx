"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Background for the slide master of an Office theme deck, in Slide
 * Master view (PPT-B030): PowerPoint's pane with the master selected, whose
 * fill every slide without one of its own shows. Solid, gradient and pattern
 * fills (the fields are the slide pane's own, `PresentationBackgroundFillFields`)
 * and Reset Background, which gives the master the theme's ground back. A
 * picture fill is not offered: the download cannot write one on the master
 * yet, and the pane does not pretend it can. Gradient and pattern follow
 * desktop PowerPoint's pane (E036), not the web app's Solid-only master.
 */

import { useState } from "react";

import { Grid3x3, Layers, RotateCcw, Square, X } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import type { OfficeColorPalettePopoverLabels } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import {
  normalizeBgGradient,
  normalizeBgPatternFill,
  type BgGradient,
  type SlideBgPattern,
} from "@mrnote/core/web/lib/presentation-slide-bg.ts";
import type { OfficeMasterBackground } from "../../../../lib/presentation-office-master";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";

import { PresentationBackgroundFillFields } from "./PresentationBackgroundFillFields";

type MasterFillKind = "solid" | "gradient" | "pattern";

export interface PresentationMasterBackgroundPaneProps {
  readonly background: string | undefined;
  readonly backgroundGradient: unknown;
  readonly backgroundPattern: unknown;
  readonly disabled: boolean;
  readonly paletteLabels: OfficeColorPalettePopoverLabels;
  readonly onFill: (fill: OfficeMasterBackground) => void;
  readonly onReset: () => void;
  readonly onClose: () => void;
  readonly t: (key: string) => string;
}

function currentKind(gradient: unknown, pattern: unknown): MasterFillKind {
  if (normalizeBgGradient(gradient)) return "gradient";
  if (pattern && typeof pattern === "object") return "pattern";
  return "solid";
}

export function PresentationMasterBackgroundPane({
  background,
  backgroundGradient,
  backgroundPattern,
  disabled,
  paletteLabels,
  onFill,
  onReset,
  onClose,
  t,
}: PresentationMasterBackgroundPaneProps) {
  const [kind, setKind] = useState<MasterFillKind>(() =>
    currentKind(backgroundGradient, backgroundPattern),
  );
  const [recents, setRecents] = useState<string[]>([]);
  const ground = background ?? "#ffffff";
  return (
    <aside
      className={presentationChrome.taskPane}
      role="dialog"
      aria-modal="false"
      aria-labelledby="presentation-master-background-pane-title"
      data-testid="presentation-master-background-pane"
      onKeyDown={(event) => {
        if (event.key !== "Escape") return;
        event.preventDefault();
        event.stopPropagation();
        onClose();
      }}
    >
      <header className={presentationChrome.taskPaneHeader}>
        <strong id="presentation-master-background-pane-title">
          {t("presentation.backgroundPane.title")}
        </strong>
        <button
          type="button"
          className={presentationChrome.taskPaneClose}
          onClick={onClose}
          aria-label={t("common.close")}
          title={t("common.close")}
          data-testid="presentation-master-background-pane-close"
        >
          <X size={16} aria-hidden="true" />
        </button>
      </header>
      <section className={presentationChrome.taskPaneSection}>
        <span className={presentationChrome.taskPaneSectionLabel}>{t("presentation.bgFillLabel")}</span>
        <div
          className={presentationChrome.backgroundKindGrid}
          role="radiogroup"
          aria-label={t("presentation.bgFillLabel")}
        >
          {(["solid", "gradient", "pattern"] as const).map((option) => (
            <button
              key={option}
              type="button"
              role="radio"
              aria-checked={kind === option}
              disabled={disabled}
              className={presentationChrome.backgroundKindButton}
              data-selected={kind === option ? "true" : "false"}
              data-testid={`presentation-master-background-${option}`}
              onClick={() => {
                setKind(option);
                // A kind is a fill: picking one gives the master that fill at
                // once, from the ground it has now, as the slide pane does.
                if (option === "solid") onFill({ background: ground });
                else if (option === "gradient") {
                  onFill({
                    background: ground,
                    backgroundGradient:
                      normalizeBgGradient(backgroundGradient) ?? { from: ground, to: "#d9d9d9", angle: 90 },
                  });
                } else {
                  onFill({
                    background: ground,
                    backgroundPattern: normalizeBgPatternFill(backgroundPattern) ?? {
                      pattern: "dots",
                      color: "#7f7f7f",
                      background: ground,
                    },
                  });
                }
              }}
            >
              {option === "solid" ? (
                <Square size={18} aria-hidden="true" />
              ) : option === "gradient" ? (
                <Layers size={18} aria-hidden="true" />
              ) : (
                <Grid3x3 size={18} aria-hidden="true" />
              )}
              <span>{t(`presentation.backgroundPane.fill.${option}`)}</span>
            </button>
          ))}
        </div>
      </section>
      <section className={presentationChrome.taskPaneSection}>
        <PresentationBackgroundFillFields
          kind={kind}
          backgroundColor={ground}
          backgroundTransparency={undefined}
          backgroundGradient={backgroundGradient}
          backgroundPattern={backgroundPattern}
          disabled={disabled}
          recents={recents}
          paletteLabels={paletteLabels}
          onSolidColor={(color) => onFill({ background: color })}
          // The master's fill carries no transparency (PptxThemeMasterBackground).
          onTransparency={() => {}}
          onGradient={(gradient: BgGradient) => onFill({ background: ground, backgroundGradient: gradient })}
          onPattern={(patch: { pattern?: SlideBgPattern; color?: string; background?: string }) =>
            onFill({
              background: ground,
              backgroundPattern: {
                ...(normalizeBgPatternFill(backgroundPattern) ?? {
                  pattern: "dots",
                  color: "#7f7f7f",
                  background: ground,
                }),
                ...patch,
              },
            })
          }
          onRecentColor={(color) =>
            setRecents((current) => [color, ...current.filter((value) => value !== color)].slice(0, 8))
          }
          t={t}
        />
      </section>
      <footer className={presentationChrome.taskPaneSection}>
        <button
          type="button"
          className={presentationChrome.taskPaneButton}
          disabled={disabled}
          onClick={() => {
            setKind("solid");
            onReset();
          }}
          data-testid="presentation-master-background-reset"
        >
          <RotateCcw size={16} aria-hidden="true" />
          {t("presentation.backgroundMenu.reset")}
        </button>
      </footer>
    </aside>
  );
}

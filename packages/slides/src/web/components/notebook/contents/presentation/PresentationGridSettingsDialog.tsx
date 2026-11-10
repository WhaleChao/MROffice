"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef, useState, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  formatPresentationMeasurement,
  parsePresentationMeasurement,
  PRESENTATION_INCH_MARK,
  type PresentationMeasurementUnit,
} from "@mrnote/core/web/lib/presentation-measurement.ts";
import {
  PRESENTATION_GRID_SPACING_PRESETS,
  presentationGridSpacingEmu,
  presentationGridSpacingPreset,
  presentationViewFlag,
  writePresentationGridDefaults,
  type PresentationViewSettingFlag,
  type PresentationViewSettings,
} from "@mrnote/core/web/lib/presentation-grid.ts";

/** The Spacing box's bounds, EMU: 0.01" to 2" (PowerPoint's list runs from
 * 1/24" to 2"; a smaller typed step is kept as typed). */
export const GRID_SETTINGS_MIN_SPACING_EMU = 9_144;
export const GRID_SETTINGS_MAX_SPACING_EMU = 1_828_800;

const FLAG_SECTIONS: ReadonlyArray<{
  readonly titleKey: string;
  readonly flags: ReadonlyArray<{ readonly flag: PresentationViewSettingFlag; readonly labelKey: string }>;
  readonly spacing?: boolean;
}> = [
  {
    titleKey: "presentation.gridSettings.snapTo",
    flags: [
      { flag: "snapToGrid", labelKey: "presentation.gridSettings.snapToGrid" },
      { flag: "snapToObjects", labelKey: "presentation.gridSettings.snapToObjects" },
    ],
  },
  {
    titleKey: "presentation.gridSettings.gridSection",
    spacing: true,
    flags: [{ flag: "showGrid", labelKey: "presentation.gridSettings.showGrid" }],
  },
  {
    titleKey: "presentation.gridSettings.guideSection",
    flags: [
      { flag: "showGuides", labelKey: "presentation.gridSettings.showGuides" },
      { flag: "smartGuides", labelKey: "presentation.gridSettings.smartGuides" },
    ],
  },
];

/**
 * PowerPoint's Grid Settings dialog (Windows: View > Show dialog launcher,
 * "Grid and Guides"; Mac: View > Guides > Grid Settings…; also right-click
 * > Grid and Guides). Microsoft Support, "Work with gridlines and use
 * snap-to-grid in PowerPoint":
 *
 * - Snap to: "Snap objects to grid" (to the nearest grid intersection) and
 *   "Snap objects to other objects" (to the edges and centres of the other
 *   objects and the slide), independent of each other;
 * - Grid settings: Spacing (the list from 1/24" to 2", or a typed length in
 *   the locale's unit) and "Display grid on screen";
 * - Guide settings: "Display drawing guides on screen" and "Display smart
 *   guides when shapes are aligned" (the guides that help align and space
 *   an object in relation to others).
 *
 * The options are the deck's `viewSettings` (`lib/presentation-grid`); OK
 * writes them all at once, Cancel none. Showing the grid never changes
 * snapping. Holding Alt (Command on a Mac) while dragging overrides every
 * snap option. "Set as Default" stores the dialog's options, as they stand,
 * as the user's default for the presentations opened afterwards (an option
 * a deck does not state itself takes it: `withPresentationGridDefaults`);
 * it changes nothing in this deck until OK.
 */
export function PresentationGridSettingsDialog({
  open,
  settings,
  unit,
  returnFocusRef,
  onApply,
  onCancel,
}: {
  readonly open: boolean;
  readonly settings: PresentationViewSettings | undefined;
  readonly unit: PresentationMeasurementUnit;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  /** Only the options the dialog changed. */
  readonly onApply: (patch: Partial<PresentationViewSettings>) => void;
  readonly onCancel: () => void;
}): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const [flags, setFlags] = useState<Record<PresentationViewSettingFlag, boolean>>(() => ({
    snapToGrid: presentationViewFlag(settings, "snapToGrid"),
    snapToObjects: presentationViewFlag(settings, "snapToObjects"),
    showGrid: presentationViewFlag(settings, "showGrid"),
    showGuides: presentationViewFlag(settings, "showGuides"),
    smartGuides: presentationViewFlag(settings, "smartGuides"),
  }));
  const startSpacing = presentationGridSpacingEmu(settings);
  const unitSuffix = unit === "in" ? PRESENTATION_INCH_MARK : ` ${t("presentation.columnsDialog.cm")}`;
  const formatSpacing = (emu: number) => `${formatPresentationMeasurement(emu, unit, 3)}${unitSuffix}`;
  // The spacing in exact EMU while the box shows a value the dialog itself
  // put there (the deck's own, or a preset picked from the list): its text
  // is rounded for display and must never be read back, or OK would move a
  // 1/12" grid (76 200) to "0.212 cm" (76 320). Typing in the box drops it.
  const [exactSpacing, setExactSpacing] = useState<number | null>(startSpacing);
  const [spacingDraft, setSpacingDraft] = useState(() => {
    const start = presentationGridSpacingPreset(startSpacing);
    return start ? start.label : formatSpacing(startSpacing);
  });
  const firstRef = useRef<HTMLInputElement | null>(null);
  const parsed = parsePresentationMeasurement(spacingDraft, unit);
  const typedPreset = PRESENTATION_GRID_SPACING_PRESETS.find((entry) => entry.label === spacingDraft.trim());
  const spacingEmu =
    exactSpacing ?? (typedPreset ? typedPreset.emu : parsed === null ? null : Math.round(parsed));
  const spacingValid =
    spacingEmu !== null &&
    spacingEmu >= GRID_SETTINGS_MIN_SPACING_EMU &&
    spacingEmu <= GRID_SETTINGS_MAX_SPACING_EMU;
  const [defaultStored, setDefaultStored] = useState(false);
  const presetValue = spacingEmu === null ? "" : (presentationGridSpacingPreset(spacingEmu)?.emu ?? "");
  return (
    <OfficeDialogShell
      open={open}
      description=""
      title={t("presentation.gridSettings.title")}
      defaultActionLabel={t("presentation.columnsDialog.ok")}
      cancelActionLabel={t("toolbar.cancel")}
      defaultActionDisabled={!spacingValid}
      onDefaultAction={() => {
        if (!spacingValid || spacingEmu === null) return;
        const patch: { -readonly [K in keyof PresentationViewSettings]: PresentationViewSettings[K] } = {};
        for (const section of FLAG_SECTIONS) {
          for (const { flag } of section.flags) {
            if (flags[flag] !== presentationViewFlag(settings, flag)) patch[flag] = flags[flag];
          }
        }
        if (spacingEmu !== startSpacing) patch.gridSpacingEmu = spacingEmu;
        onApply(patch);
      }}
      onCancel={onCancel}
      footerStart={
        <button
          type="button"
          className={officeDialogParts.button}
          disabled={!spacingValid}
          aria-pressed={defaultStored}
          onClick={() => {
            if (!spacingValid || spacingEmu === null) return;
            setDefaultStored(writePresentationGridDefaults({ ...flags, gridSpacingEmu: spacingEmu }));
          }}
          data-testid="presentation-grid-settings-set-default"
        >
          {t("presentation.gridSettings.setAsDefault")}
        </button>
      }
      initialFocusRef={firstRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-grid-settings"
      className="sm:max-w-[400px]"
    >
      <div className="grid gap-4">
        {FLAG_SECTIONS.map((section, sectionIndex) => (
          <fieldset key={section.titleKey} className={officeDialogParts.group}>
            <legend className={officeDialogParts.sectionTitle}>{t(section.titleKey)}</legend>
            {section.spacing ? (
              <div className={officeDialogParts.row}>
                <label className={officeDialogParts.field}>
                  <span className={officeDialogParts.label}>{t("presentation.gridSettings.spacing")}</span>
                  <select
                    value={presetValue}
                    onChange={(event) => {
                      const entry = PRESENTATION_GRID_SPACING_PRESETS.find(
                        (candidate) => String(candidate.emu) === event.target.value,
                      );
                      if (entry) {
                        setDefaultStored(false);
                        setExactSpacing(entry.emu);
                        setSpacingDraft(entry.label);
                      }
                    }}
                    data-testid="presentation-grid-settings-spacing-preset"
                  >
                    <option value="" disabled>
                      {t("presentation.gridSettings.custom")}
                    </option>
                    {PRESENTATION_GRID_SPACING_PRESETS.map((entry) => (
                      <option key={entry.emu} value={entry.emu}>
                        {entry.label}
                      </option>
                    ))}
                  </select>
                </label>
                <label className={officeDialogParts.field}>
                  <span className={officeDialogParts.label}>{t("presentation.gridSettings.spacingValue")}</span>
                  <input
                    type="text"
                    inputMode="decimal"
                    value={spacingDraft}
                    aria-invalid={!spacingValid}
                    onChange={(event) => {
                      setDefaultStored(false);
                      setExactSpacing(null);
                      setSpacingDraft(event.target.value);
                    }}
                    data-testid="presentation-grid-settings-spacing"
                  />
                </label>
              </div>
            ) : null}
            {section.flags.map(({ flag, labelKey }, flagIndex) => (
              <label key={flag} className={officeDialogParts.check}>
                <input
                  ref={sectionIndex === 0 && flagIndex === 0 ? firstRef : undefined}
                  type="checkbox"
                  checked={flags[flag]}
                  onChange={(event) => {
                    const checked = event.target.checked;
                    setDefaultStored(false);
                    setFlags((current) => ({ ...current, [flag]: checked }));
                  }}
                  data-testid={`presentation-grid-settings-${flag}`}
                />
                <span>{t(labelKey)}</span>
              </label>
            ))}
          </fieldset>
        ))}
      </div>
    </OfficeDialogShell>
  );
}

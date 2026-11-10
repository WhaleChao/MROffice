"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Bullets and Numbering dialog (Home > Bullets ▾ / Numbering ▾ >
 * Bullets and Numbering...; PPT-A039, PPT-A040). Bulleted and Numbered pages,
 * each the library's None plus seven items; Size (25-400 % of text), Color
 * (Automatic follows the text) and, on the Numbered page, Start at. What OK
 * writes is decided by `lib/presentation-bullet-library`.
 *
 * Picture... and Customize... (a glyph from the Symbol dialog) are shown
 * disabled with the reason: a picture bullet is kept and drawn when a file
 * brings one, but not authored here yet.
 */

import { useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeColorPalettePanel } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePanel.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  BULLET_SIZE_PERCENT_BOUNDS,
  BULLET_START_AT_BOUNDS,
  bulletsNumberingResult,
  type BulletsNumberingProblem,
  type BulletsNumberingResult,
  type BulletsNumberingSource,
  type PresentationListLibraryItem,
} from "@mrnote/core/web/lib/presentation-bullet-library.ts";

export interface PresentationBulletsNumberingDialogProps {
  readonly open: boolean;
  readonly source: BulletsNumberingSource;
  readonly bulletLibrary: readonly PresentationListLibraryItem[];
  readonly numberLibrary: readonly PresentationListLibraryItem[];
  /** The theme's colour grid, as the window's other colour menus use it. */
  readonly themeColumns?: readonly (readonly string[])[];
  readonly onApply: (result: BulletsNumberingResult) => boolean;
  readonly onCancel: () => void;
}

const TEST_ID = "presentation-bullets-dialog";

export function PresentationBulletsNumberingDialog(
  props: PresentationBulletsNumberingDialogProps,
): React.JSX.Element | null {
  if (!props.open) return null;
  return <PresentationBulletsNumberingDialogBody {...props} />;
}

function PresentationBulletsNumberingDialogBody({
  source,
  bulletLibrary,
  numberLibrary,
  themeColumns,
  onApply,
  onCancel,
}: PresentationBulletsNumberingDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const word = (key: string, values?: Record<string, string | number>) =>
    t(key as Parameters<typeof t>[0], values as never);
  const [page, setPage] = useState<"bullet" | "number">(source.kind);
  const [itemId, setItemId] = useState<string | null>(source.itemId ?? null);
  const [size, setSize] = useState(String(source.sizePercent));
  const [color, setColor] = useState<string | null>(source.color ?? null);
  const [startAt, setStartAt] = useState(String(source.startAt));
  const [paletteOpen, setPaletteOpen] = useState(false);
  const [problem, setProblem] = useState<BulletsNumberingProblem | null>(null);
  const firstTileRef = useRef<HTMLButtonElement | null>(null);

  const library = page === "bullet" ? bulletLibrary : numberLibrary;
  // A tile picked on the other page does not carry over: None is selected.
  const selected = library.some((item) => item.id === itemId) ? itemId : null;

  const apply = () => {
    const item = library.find((candidate) => candidate.id === selected) ?? null;
    const result = bulletsNumberingResult({ item, size, color, startAt });
    if ("problem" in result) {
      setProblem(result.problem);
      return;
    }
    if (onApply(result)) onCancel();
  };

  const problemMessage =
    problem === null
      ? null
      : word("presentation.paragraphDialog.invalidValue", {
          field:
            problem === "size"
              ? word("presentation.bulletsDialog.size")
              : word("presentation.bulletsDialog.startAt"),
          minimum: problem === "size" ? BULLET_SIZE_PERCENT_BOUNDS[0] : BULLET_START_AT_BOUNDS[0],
          maximum: problem === "size" ? BULLET_SIZE_PERCENT_BOUNDS[1] : BULLET_START_AT_BOUNDS[1],
        });

  const tile = (item: PresentationListLibraryItem | null, index: number) => {
    const id = item?.id ?? "none";
    const active = (item?.id ?? null) === selected;
    return (
      <button
        key={id}
        ref={index === 0 ? firstTileRef : undefined}
        type="button"
        role="radio"
        aria-checked={active}
        aria-label={word(`presentation.bulletsLibrary.${id}`)}
        title={word(`presentation.bulletsLibrary.${id}`)}
        data-testid={`${TEST_ID}-item-${id}`}
        className="office-dialog__tile"
        style={{
          display: "grid",
          placeItems: "center",
          height: 72,
          border: active ? "2px solid var(--office-accent, #c43e1c)" : "1px solid var(--office-border, #c8c6c4)",
          // The dialog's own well (white in light, graphite in dark), so the
          // glyph in the dialog's ink reads in both.
          background: "var(--office-surface, var(--office-dialog-surface, #fff))",
          fontSize: item?.kind === "number" ? 13 : 24,
          borderRadius: 4,
        }}
        onClick={() => {
          setProblem(null);
          setItemId(item?.id ?? null);
        }}
      >
        {item ? <span aria-hidden>{item.preview}</span> : <span aria-hidden>{word("toolbar.listNone")}</span>}
      </button>
    );
  };

  const unavailable = word("presentation.bulletsDialog.pictureUnavailable");

  return (
    <OfficeDialogShell
      open
      testIdPrefix={TEST_ID}
      title={word("presentation.bulletsDialog.title")}
      description=""
      defaultActionLabel={word("toolbar.dialogOk")}
      defaultActionTestId={`${TEST_ID}-apply`}
      cancelActionLabel={word("toolbar.pageSetupCancel")}
      cancelActionTestId={`${TEST_ID}-cancel`}
      overlayTestId={`${TEST_ID}-layer`}
      onDefaultAction={apply}
      onCancel={onCancel}
      initialFocusRef={firstTileRef}
      className="sm:max-w-[520px]"
      footerStart={
        problemMessage ? (
          <span role="alert" className="office-dialog__error" data-testid={`${TEST_ID}-problem`}>
            {problemMessage}
          </span>
        ) : undefined
      }
      tabStrip={
        <div className="office-dialog__tabs" role="tablist" aria-label={word("presentation.bulletsDialog.title")}>
          {(["bullet", "number"] as const).map((kind) => (
            <button
              key={kind}
              type="button"
              role="tab"
              aria-selected={page === kind}
              className="office-dialog__tab"
              data-testid={`${TEST_ID}-tab-${kind}`}
              onClick={() => {
                setProblem(null);
                setPage(kind);
              }}
            >
              {word(kind === "bullet" ? "presentation.bulletsDialog.bulleted" : "presentation.bulletsDialog.numbered")}
            </button>
          ))}
        </div>
      }
    >
      <div className="grid gap-3">
        <div role="radiogroup" className="grid grid-cols-4 gap-2" data-testid={`${TEST_ID}-library`}>
          {[null, ...library].map((item, index) => tile(item, index))}
        </div>
        <div className="grid grid-cols-3 items-end gap-3">
          <label className={officeDialogParts.field}>
            <span>{word("presentation.bulletsDialog.size")}</span>
            <span className={officeDialogParts.measure}>
              <input
                inputMode="numeric"
                data-testid={`${TEST_ID}-size`}
                value={size}
                onChange={(event) => {
                  setProblem(null);
                  setSize(event.target.value);
                }}
              />
              <span>{word("presentation.bulletsDialog.percentOfText")}</span>
            </span>
          </label>
          <div className={officeDialogParts.field}>
            <span>{word("presentation.bulletsDialog.color")}</span>
            <button
              type="button"
              aria-expanded={paletteOpen}
              data-testid={`${TEST_ID}-color`}
              onClick={() => setPaletteOpen((value) => !value)}
              className="flex h-8 items-center gap-2 rounded border px-2"
            >
              <span
                aria-hidden
                style={{
                  width: 18,
                  height: 12,
                  border: "1px solid #8a8886",
                  background: color ?? "transparent",
                }}
              />
              <span>{color ?? word("toolbar.textColorAutomatic")}</span>
            </button>
          </div>
          {page === "number" ? (
            <label className={officeDialogParts.field}>
              <span>{word("presentation.bulletsDialog.startAt")}</span>
              <input
                inputMode="numeric"
                data-testid={`${TEST_ID}-start-at`}
                value={startAt}
                onChange={(event) => {
                  setProblem(null);
                  setStartAt(event.target.value);
                }}
              />
            </label>
          ) : (
            <div className="flex gap-2" title={unavailable}>
              <button type="button" disabled data-testid={`${TEST_ID}-picture`}>
                {word("presentation.bulletsDialog.picture")}
              </button>
              <button type="button" disabled data-testid={`${TEST_ID}-customize`}>
                {word("presentation.bulletsDialog.customize")}
              </button>
            </div>
          )}
        </div>
        {paletteOpen ? (
          <OfficeColorPalettePanel
            value={color ?? undefined}
            themeColumns={themeColumns}
            testIdPrefix={`${TEST_ID}-palette`}
            automaticSelected={color === null}
            onAutomatic={() => {
              setColor(null);
              setPaletteOpen(false);
            }}
            onPick={(next) => {
              setColor(next);
              setPaletteOpen(false);
            }}
            labels={{
              automatic: word("toolbar.textColorAutomatic"),
              themeColors: word("toolbar.themeColors"),
              standardColors: word("toolbar.standardColors"),
              moreColors: word("toolbar.textColorCustom"),
            }}
          />
        ) : null}
      </div>
    </OfficeDialogShell>
  );
}

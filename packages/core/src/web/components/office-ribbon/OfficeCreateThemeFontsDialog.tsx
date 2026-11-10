"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useId, useMemo, useRef, useState, type CSSProperties, type RefObject } from "react";
import { useTranslations } from "../../../host/next-intl";

import { OfficeDialogShell } from "./OfficeDialogShell";
import { officeDialogParts } from "./officeDialogChrome";
import { officeFontFamilies } from "../../lib/office-fonts/substitutes";
import {
  compareFontLabelsEastAsianFirst,
  officeFamilyDisplayNames,
  officeFamilyFromDisplayName,
} from "../../lib/office-fonts/localized-names";

/** The four faces and the name Create New Theme Fonts saves. */
export interface OfficeThemeFontsDraft {
  readonly headingLatin: string;
  readonly bodyLatin: string;
  readonly headingEastAsian?: string;
  readonly bodyEastAsian?: string;
  readonly name: string;
}

const MAX_FACE = 120;
const MAX_NAME = 80;
const UNSAFE = /[<>&\u0000-\u001f\u007f]/;

function face(value: string): string | null {
  const trimmed = officeFamilyFromDisplayName(value.trim());
  return trimmed && trimmed.length <= MAX_FACE && !UNSAFE.test(trimmed) ? trimmed : null;
}

/**
 * Office's Create New Theme Fonts dialog (PowerPoint's Design > Variants >
 * Fonts > Customize Fonts..., Word's Design > Fonts, Excel's Page Layout >
 * Fonts; Microsoft Support "Create new theme fonts"): a Heading font and a
 * Body font for Latin text, and -- where East Asian text is on -- a Heading
 * font and a Body font for Asian text, each with its Sample, and the Name the
 * pair is saved under ("Custom 1"). A face is typed or picked: Office keeps a
 * name it does not have installed and draws it with its substitute, so a
 * typed name is kept as typed (a Chinese name such as 黑体 is the family it
 * stands for, SimHei). Save hands the draft to the host, which writes it as
 * its theme's font scheme.
 */
export function OfficeCreateThemeFontsDialog({
  open,
  initial,
  showEastAsian,
  locale,
  returnFocusRef,
  onSave,
  onCancel,
  testIdPrefix = "office-theme-fonts",
}: {
  readonly open: boolean;
  /** The faces the dialog opens with: the theme's current scheme. */
  readonly initial: OfficeThemeFontsDraft;
  readonly showEastAsian: boolean;
  readonly locale?: string;
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onSave: (draft: OfficeThemeFontsDraft) => void;
  readonly onCancel: () => void;
  readonly testIdPrefix?: string;
}): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const displayNames = useMemo(() => officeFamilyDisplayNames(locale), [locale]);
  const shown = (family: string | undefined) => (family ? displayNames[family] ?? family : "");
  const [headingLatin, setHeadingLatin] = useState(shown(initial.headingLatin));
  const [bodyLatin, setBodyLatin] = useState(shown(initial.bodyLatin));
  const [headingEastAsian, setHeadingEastAsian] = useState(shown(initial.headingEastAsian));
  const [bodyEastAsian, setBodyEastAsian] = useState(shown(initial.bodyEastAsian));
  const [name, setName] = useState(initial.name);
  const firstRef = useRef<HTMLInputElement | null>(null);
  const listId = useId();
  // Reopened, the dialog starts again from the theme's scheme.
  useEffect(() => {
    if (!open) return;
    setHeadingLatin(shown(initial.headingLatin));
    setBodyLatin(shown(initial.bodyLatin));
    setHeadingEastAsian(shown(initial.headingEastAsian));
    setBodyEastAsian(shown(initial.bodyEastAsian));
    setName(initial.name);
    // `shown` reads `displayNames`, which `locale` fixes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, initial, displayNames]);
  const families = useMemo(
    () =>
      officeFontFamilies()
        .map((family) => displayNames[family] ?? family)
        .sort(compareFontLabelsEastAsianFirst),
    [displayNames],
  );

  const heading = face(headingLatin);
  const body = face(bodyLatin);
  const headingAsian = showEastAsian && headingEastAsian.trim() ? face(headingEastAsian) : undefined;
  const bodyAsian = showEastAsian && bodyEastAsian.trim() ? face(bodyEastAsian) : undefined;
  const trimmedName = name.trim();
  const valid =
    heading !== null &&
    body !== null &&
    headingAsian !== null &&
    bodyAsian !== null &&
    trimmedName.length > 0 &&
    trimmedName.length <= MAX_NAME &&
    !/[\u0000-\u001f\u007f]/.test(trimmedName);

  const row = (
    label: string,
    value: string,
    set: (value: string) => void,
    testId: string,
    ref?: RefObject<HTMLInputElement | null>,
  ) => (
    <label className={officeDialogParts.field}>
      <span className={officeDialogParts.label}>{label}</span>
      <input
        ref={ref}
        type="text"
        list={listId}
        autoComplete="off"
        spellCheck={false}
        maxLength={MAX_FACE}
        value={value}
        onChange={(event) => set(event.target.value)}
        data-testid={`${testIdPrefix}-${testId}`}
      />
    </label>
  );
  // Office's Sample: the heading face over the body face, in the script
  // the column sets.
  const sample = (headingFace: string | null | undefined, bodyFace: string | null | undefined, asian: boolean, testId: string) => (
    <div
      className="rounded border border-[var(--office-dialog-border,#d0d0d0)] bg-white px-3 py-2 text-black"
      aria-label={t("officeThemeFonts.sample")}
      data-testid={`${testIdPrefix}-${testId}`}
    >
      <div className="text-[11px] text-neutral-500">{t("officeThemeFonts.sample")}</div>
      <div
        className="truncate text-[20px] leading-tight"
        style={{ fontFamily: headingFace ? `"${headingFace}"` : undefined } as CSSProperties}
      >
        {asian ? t("officeThemeFonts.sampleHeadingAsian") : t("officeThemeFonts.sampleHeading")}
      </div>
      <div
        className="truncate text-[12px] leading-snug"
        style={{ fontFamily: bodyFace ? `"${bodyFace}"` : undefined } as CSSProperties}
      >
        {asian ? t("officeThemeFonts.sampleBodyAsian") : t("officeThemeFonts.sampleBody")}
      </div>
    </div>
  );

  return (
    <OfficeDialogShell
      open={open}
      description=""
      title={t("officeThemeFonts.title")}
      defaultActionLabel={t("officeThemeFonts.save")}
      cancelActionLabel={t("toolbar.cancel")}
      defaultActionDisabled={!valid}
      onDefaultAction={() => {
        if (!valid || !heading || !body) return;
        onSave({
          headingLatin: heading,
          bodyLatin: body,
          ...(headingAsian ? { headingEastAsian: headingAsian } : {}),
          ...(bodyAsian ? { bodyEastAsian: bodyAsian } : {}),
          name: trimmedName,
        });
      }}
      onCancel={onCancel}
      initialFocusRef={firstRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix={testIdPrefix}
      className="sm:max-w-[520px]"
    >
      <datalist id={listId}>
        {families.map((family) => (
          <option key={family} value={family} />
        ))}
      </datalist>
      <div className="grid gap-4">
        <fieldset className="grid gap-2" data-testid={`${testIdPrefix}-latin`}>
          {showEastAsian ? (
            <legend className="mb-1 text-xs font-medium">{t("officeThemeFonts.latinText")}</legend>
          ) : null}
          <div className="grid gap-3 sm:grid-cols-[1fr_180px]">
            <div className="grid gap-2">
              {row(t("officeThemeFonts.headingFont"), headingLatin, setHeadingLatin, "heading", firstRef)}
              {row(t("officeThemeFonts.bodyFont"), bodyLatin, setBodyLatin, "body")}
            </div>
            {sample(heading, body, false, "sample")}
          </div>
        </fieldset>
        {showEastAsian ? (
          <fieldset className="grid gap-2" data-testid={`${testIdPrefix}-east-asian`}>
            <legend className="mb-1 text-xs font-medium">{t("officeThemeFonts.asianText")}</legend>
            <div className="grid gap-3 sm:grid-cols-[1fr_180px]">
              <div className="grid gap-2">
                {row(t("officeThemeFonts.headingFontAsian"), headingEastAsian, setHeadingEastAsian, "heading-east-asian")}
                {row(t("officeThemeFonts.bodyFontAsian"), bodyEastAsian, setBodyEastAsian, "body-east-asian")}
              </div>
              {sample(headingAsian ?? heading, bodyAsian ?? body, true, "sample-east-asian")}
            </div>
          </fieldset>
        ) : null}
        <label className={officeDialogParts.field}>
          <span className={officeDialogParts.label}>{t("officeThemeFonts.name")}</span>
          <input
            type="text"
            maxLength={MAX_NAME}
            value={name}
            onChange={(event) => setName(event.target.value)}
            data-testid={`${testIdPrefix}-name`}
          />
        </label>
      </div>
    </OfficeDialogShell>
  );
}

/**
 * The name Create New Theme Fonts proposes: "Custom 1", or the next number
 * after the custom names the host already holds.
 */
export function nextCustomThemeFontsName(prefix: string, taken: readonly (string | undefined)[]): string {
  let highest = 0;
  for (const value of taken) {
    if (!value || !value.startsWith(prefix + " ")) continue;
    const number = Number(value.slice(prefix.length + 1));
    if (Number.isInteger(number) && number > highest) highest = number;
  }
  return `${prefix} ${highest + 1}`;
}

"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef, useState } from "react";
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { EMU_PER_TWIP } from "@mrnote/core/web/lib/office-units.ts";
import {
  formatPresentationMeasurement,
  parsePresentationMeasurement,
  presentationMeasurementUnit,
} from "@mrnote/core/web/lib/presentation-measurement.ts";

import {
  WORD_DEFAULT_HYPHENATION_ZONE_TWIPS,
  type HyphenationOptions,
} from "@mrnote/core/web/components/console/editor/native-document/documentHyphenation.ts";

/**
 * Word's Layout > Hyphenation > Hyphenation Options dialog (Microsoft
 * Support, "Hyphenate text in Word"): Automatically hyphenate document,
 * Hyphenate words in CAPS, Hyphenation zone (0.25" by default) and Limit
 * consecutive hyphens to (No limit by default). OK writes them all at once;
 * Cancel leaves them as they were. Word's Manual... button is not offered
 * (manual hyphenation is not modelled).
 */

const MAX_ZONE_TWIPS = 31_680;
const MAX_LIMIT = 32_767;

export interface DocumentHyphenationOptionsDialogProps {
  readonly automatic: boolean;
  readonly options: HyphenationOptions | null;
  readonly canEdit: boolean;
  readonly onApply: (automatic: boolean, options: HyphenationOptions | null) => void;
  readonly onCancel: () => void;
  readonly returnFocusTarget?: () => HTMLElement | null;
}

export function DocumentHyphenationOptionsDialog({
  automatic: initialAutomatic,
  options,
  canEdit,
  onApply,
  onCancel,
  returnFocusTarget,
}: DocumentHyphenationOptionsDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const unit = presentationMeasurementUnit(useLocale());
  const [automatic, setAutomatic] = useState(initialAutomatic);
  const [caps, setCaps] = useState(options?.hyphenateCaps !== false);
  const initialZone = options?.zoneTwips ?? WORD_DEFAULT_HYPHENATION_ZONE_TWIPS;
  const [zoneDraft, setZoneDraft] = useState(() =>
    formatPresentationMeasurement(initialZone * EMU_PER_TWIP, unit, 2),
  );
  const [limitDraft, setLimitDraft] = useState(() =>
    options?.consecutiveLimit ? String(options.consecutiveLimit) : "",
  );
  const firstRef = useRef<HTMLInputElement | null>(null);

  const zoneEmu = parsePresentationMeasurement(zoneDraft, unit);
  const zoneTwips = zoneEmu === null ? null : Math.round(zoneEmu / EMU_PER_TWIP);
  const zoneValid = zoneTwips !== null && zoneTwips > 0 && zoneTwips <= MAX_ZONE_TWIPS;
  const limitText = limitDraft.trim();
  const noLimit = limitText === "" || limitText === t("toolbar.hyphenationNoLimit");
  const limit = noLimit ? 0 : Number(limitText);
  const limitValid = noLimit || (Number.isInteger(limit) && limit >= 0 && limit <= MAX_LIMIT);

  const apply = () => {
    if (!zoneValid || !limitValid || zoneTwips === null) return;
    // A setting the document did not state and that is still Word's default
    // stays unstated, so an untouched OK writes nothing new.
    const next: { zoneTwips?: number; consecutiveLimit?: number; hyphenateCaps?: false } = {};
    if (options?.zoneTwips !== undefined || zoneTwips !== WORD_DEFAULT_HYPHENATION_ZONE_TWIPS) next.zoneTwips = zoneTwips;
    if (options?.consecutiveLimit !== undefined || limit !== 0) next.consecutiveLimit = limit;
    if (!caps) next.hyphenateCaps = false;
    onApply(automatic, Object.keys(next).length > 0 ? next : null);
  };

  return (
    <OfficeDialogShell
      open
      title={t("toolbar.hyphenation")}
      description={<span className="sr-only">{t("toolbar.hyphenationOptions")}</span>}
      defaultActionLabel={t("toolbar.dialogOk")}
      cancelActionLabel={t("toolbar.dialogCancel")}
      defaultActionDisabled={!canEdit || !zoneValid || !limitValid}
      onDefaultAction={apply}
      onCancel={onCancel}
      initialFocusRef={firstRef}
      returnFocusTarget={returnFocusTarget}
      testIdPrefix="document-hyphenation-options"
      defaultActionTestId="document-hyphenation-options-apply"
      cancelActionTestId="document-hyphenation-options-cancel"
      className="w-[min(400px,calc(100vw-32px))]"
    >
      <section className={parts.group}>
        <label className={parts.field}>
          <input
            ref={firstRef}
            type="checkbox"
            checked={automatic}
            disabled={!canEdit}
            data-testid="document-hyphenation-options-automatic"
            onChange={(event) => setAutomatic(event.target.checked)}
          />
          <span>{t("toolbar.hyphenationAuto")}</span>
        </label>
        <label className={parts.field}>
          <input
            type="checkbox"
            checked={caps}
            disabled={!canEdit}
            data-testid="document-hyphenation-options-caps"
            onChange={(event) => setCaps(event.target.checked)}
          />
          <span>{t("toolbar.hyphenationCaps")}</span>
        </label>
        <div className="grid grid-cols-[max-content_minmax(0,1fr)] items-center gap-x-3 gap-y-2">
          <label className="contents">
            <span className={parts.label}>{t("toolbar.hyphenationZone")}</span>
            <input
              type="text"
              value={zoneDraft}
              disabled={!canEdit}
              aria-invalid={!zoneValid || undefined}
              data-testid="document-hyphenation-options-zone"
              onChange={(event) => setZoneDraft(event.target.value)}
            />
          </label>
          <label className="contents">
            <span className={parts.label}>{t("toolbar.hyphenationLimit")}</span>
            <input
              type="text"
              inputMode="numeric"
              value={limitDraft}
              placeholder={t("toolbar.hyphenationNoLimit")}
              disabled={!canEdit}
              aria-invalid={!limitValid || undefined}
              data-testid="document-hyphenation-options-limit"
              onChange={(event) => setLimitDraft(event.target.value)}
            />
          </label>
        </div>
      </section>
    </OfficeDialogShell>
  );
}

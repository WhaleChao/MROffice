"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Master Layout dialog (Slide Master view ▸ Master Layout ▸
 * Master Layout, on the slide master; PPT-B030): the master's placeholders --
 * Title, Text, Date, Slide Number, Footer -- as check boxes. A placeholder
 * the master still has is checked and unavailable; checking a deleted one
 * puts it back on OK (Microsoft Support, "Restore placeholders on the slide
 * master"). Nothing here deletes: that is selecting the placeholder on the
 * master and pressing Delete.
 */

import { useRef, useState, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  OFFICE_MASTER_PLACEHOLDER_FAMILIES,
  type OfficeMasterPlaceholderFamily,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";

export interface PresentationMasterLayoutDialogLabels {
  readonly title: string;
  readonly placeholders: string;
  readonly family: Readonly<Record<OfficeMasterPlaceholderFamily, string>>;
  readonly ok: string;
  readonly cancel: string;
}

export function PresentationMasterLayoutDialog({
  open,
  labels,
  deleted,
  onApply,
  onCancel,
  returnFocusRef,
}: {
  open: boolean;
  labels: PresentationMasterLayoutDialogLabels;
  /** The master's placeholders deleted now. */
  deleted: ReadonlySet<OfficeMasterPlaceholderFamily>;
  /** The deleted placeholders checked to come back. */
  onApply: (restore: readonly OfficeMasterPlaceholderFamily[]) => void;
  onCancel: () => void;
  returnFocusRef?: RefObject<HTMLElement | null>;
}) {
  const [restore, setRestore] = useState<ReadonlySet<OfficeMasterPlaceholderFamily>>(() => new Set());
  const firstRef = useRef<HTMLInputElement | null>(null);
  return (
    <OfficeDialogShell
      open={open}
      description=""
      title={labels.title}
      defaultActionLabel={labels.ok}
      cancelActionLabel={labels.cancel}
      onDefaultAction={() =>
        onApply(OFFICE_MASTER_PLACEHOLDER_FAMILIES.filter((family) => restore.has(family)))
      }
      onCancel={onCancel}
      initialFocusRef={firstRef}
      returnFocusRef={returnFocusRef}
      testIdPrefix="presentation-master-layout"
      className="sm:max-w-[320px]"
    >
      <fieldset className={officeDialogParts.field} style={{ border: 0, margin: 0, padding: 0 }}>
        <legend className={officeDialogParts.label}>{labels.placeholders}</legend>
        {OFFICE_MASTER_PLACEHOLDER_FAMILIES.map((family, index) => {
          const present = !deleted.has(family);
          return (
            <label key={family} style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <input
                ref={index === 0 ? firstRef : undefined}
                type="checkbox"
                checked={present || restore.has(family)}
                disabled={present}
                onChange={(event) =>
                  setRestore((current) => {
                    const next = new Set(current);
                    if (event.target.checked) next.add(family);
                    else next.delete(family);
                    return next;
                  })
                }
                data-testid={`presentation-master-layout-${family}`}
              />
              {labels.family[family]}
            </label>
          );
        })}
      </fieldset>
    </OfficeDialogShell>
  );
}

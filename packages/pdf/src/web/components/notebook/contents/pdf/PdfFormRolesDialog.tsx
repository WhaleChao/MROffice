/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 管理角色 (ONLYOFFICE's Manage Roles): the form's filling roles, each
 * renamed, recoloured or removed. A role lives on its fields
 * (lib/pdf/form-roles.ts), so applying the dialog edits every field of a
 * changed role: renamed or recoloured, the fields take the new role; removed,
 * anyone fills them. A role is added where a field takes it (the field's
 * properties), as ONLYOFFICE assigns one.
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { useEffect, useRef, useState } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  isPdfFormRoleName,
  PDF_FORM_ROLE_MAX_NAME,
  PDF_FORM_ROLE_PALETTE,
  type PdfFormRole,
} from "../../../../lib/pdf/form-roles";
import { useOpenedOnce } from "@mrnote/core/web/lib/use-opened-once.ts";

/** What becomes of one role: the role its fields take, or null (anyone). */
export interface PdfFormRoleChange {
  readonly from: string;
  readonly to: PdfFormRole | null;
}

type PdfFormRolesDialogProps = {
  open: boolean;
  roles: readonly PdfFormRole[];
  /** How many fields name each role (by name). */
  fieldCounts: ReadonlyMap<string, number>;
  busy: boolean;
  submitError?: string | null;
  onApply: (changes: readonly PdfFormRoleChange[]) => Promise<boolean>;
  onCancel: () => void;
};

interface RoleDraft {
  readonly from: string;
  name: string;
  color: string;
  removed: boolean;
}

export function PdfFormRolesDialog(props: PdfFormRolesDialogProps) {
  if (!useOpenedOnce(props.open)) return null;
  return <PdfFormRolesDialogBody {...props} />;
}

function PdfFormRolesDialogBody({ open, roles, fieldCounts, busy, submitError = null, onApply, onCancel }:
  PdfFormRolesDialogProps) {
  const t = useTranslations("console-notebooks");
  const firstInputRef = useRef<HTMLInputElement>(null);
  const submitInFlightRef = useRef(false);
  const [drafts, setDrafts] = useState<RoleDraft[]>([]);
  useEffect(() => {
    if (!open) return;
    setDrafts(roles.map((role) => ({ from: role.name, name: role.name, color: role.color, removed: false })));
  }, [open, roles]);

  const kept = drafts.filter((draft) => !draft.removed);
  const names = kept.map((draft) => draft.name.trim());
  const valid = names.every(isPdfFormRoleName) && new Set(names).size === names.length;
  const changes: PdfFormRoleChange[] = drafts.flatMap((draft): PdfFormRoleChange[] => {
    const original = roles.find((role) => role.name === draft.from);
    if (draft.removed) return [{ from: draft.from, to: null }];
    const name = draft.name.trim();
    return original && (original.name !== name || original.color !== draft.color)
      ? [{ from: draft.from, to: { name, color: draft.color } }] : [];
  });
  const canApply = valid && changes.length > 0 && !busy;
  const update = (index: number, patch: Partial<RoleDraft>) =>
    setDrafts((current) => current.map((draft, at) => (at === index ? { ...draft, ...patch } : draft)));

  const apply = async () => {
    if (!canApply || submitInFlightRef.current) return;
    submitInFlightRef.current = true;
    try {
      await onApply(changes);
    } finally {
      submitInFlightRef.current = false;
    }
  };

  return (
    <OfficeDialogShell
      open={open}
      title={t("pdfForm.manageRoles")}
      description={t("pdfForm.manageRolesDescription")}
      defaultActionLabel={t("pdfForm.manageRolesApply")}
      cancelActionLabel={t("pdf.radioGroup.cancel")}
      onDefaultAction={() => void apply()}
      onCancel={onCancel}
      defaultActionDisabled={!canApply}
      busy={busy}
      busyActionLabel={t("pdf.radioGroup.saving")}
      initialFocusRef={firstInputRef}
      testIdPrefix="pdf-form-roles"
    >
      <div className="grid gap-3" data-testid="pdf-form-roles-list">
        {drafts.map((draft, index) => (
          <div key={draft.from} className={`${parts.field} gap-2`} data-testid="pdf-form-role-row"
            data-removed={draft.removed ? "true" : undefined}>
            <div className="flex flex-wrap items-center gap-2">
              <input
                ref={index === 0 ? firstInputRef : undefined}
                value={draft.name}
                maxLength={PDF_FORM_ROLE_MAX_NAME}
                onChange={(event) => update(index, { name: event.target.value })}
                disabled={busy || draft.removed}
                autoComplete="off"
                aria-label={t("pdf.formField.roleName")}
                className="min-w-40 flex-1"
                data-testid="pdf-form-role-name"
              />
              <span className={parts.hint}>
                {t("pdfForm.roleFieldCount", { count: fieldCounts.get(draft.from) ?? 0 })}
              </span>
              <button
                type="button"
                onClick={() => update(index, { removed: !draft.removed })}
                disabled={busy}
                aria-pressed={draft.removed}
                data-testid="pdf-form-role-remove"
              >
                {draft.removed ? t("pdfForm.roleKeep") : t("pdfForm.roleRemove")}
              </button>
            </div>
            <span role="radiogroup" aria-label={t("pdf.formField.roleColor")} className="flex gap-1">
              {PDF_FORM_ROLE_PALETTE.map((color) => (
                <button
                  key={color}
                  type="button"
                  role="radio"
                  aria-checked={draft.color === color}
                  aria-label={color}
                  onClick={() => update(index, { color })}
                  disabled={busy || draft.removed}
                  className="size-6 rounded-full border border-black/20"
                  style={{ background: color, outline: draft.color === color ? "2px solid var(--pdf-chrome-ink, #0f172a)" : undefined, outlineOffset: 1 }}
                  data-testid={`pdf-form-role-color-${color.slice(1)}`}
                />
              ))}
            </span>
          </div>
        ))}
      </div>
      {submitError ? <p role="alert" className="mt-3 text-sm text-red-700">{submitError}</p> : null}
    </OfficeDialogShell>
  );
}

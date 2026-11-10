"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useState } from "react";
import type { Editor } from "@tiptap/core";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type { OfficeSemanticIconId } from "@mrnote/core/web/components/office-ribbon/icons/officeIconRegistry.ts";
import { toast } from "@mrnote/core/web/hooks/use-toast.ts";

import { collectContentControlRanges, type ContentControlRange } from "@mrnote/core/web/components/console/editor/native-document/contentControlExtension.tsx";
import { wordControlPlaceholderText, type ContentControlMeta, type ContentControlType } from "@mrnote/core/web/components/console/editor/native-document/contentControls.ts";
import { DocumentContentControlPropertiesDialog } from "./DocumentContentControlPropertiesDialog";
import { formFieldRanges, missingRequiredFormFields } from "@mrnote/core/web/components/console/editor/native-document/documentFormFill.ts";

/**
 * The Forms tab: Word's Developer > Controls (its content controls, the
 * option button made of a check box control, Properties) and the filling
 * of a form -- Previous Field, Next Field, Clear Fields and Submit, which
 * checks the required fields first (`documentFormFill`).
 */

const CONTROL_BUTTONS: readonly { readonly type: ContentControlType | "radio"; readonly icon: OfficeSemanticIconId }[] = [
  { type: "text", icon: "text.type" },
  { type: "richText", icon: "textBox.insert" },
  { type: "checkbox", icon: "selection.checkbox" },
  { type: "radio", icon: "selection.radio" },
  { type: "dropDownList", icon: "list.bullets" },
  { type: "comboBox", icon: "form.field" },
  { type: "date", icon: "calendar.date" },
  { type: "picture", icon: "image.insert" },
];

/** The group a new option button joins. */
const DEFAULT_RADIO_GROUP = "Group 1";

/** The control the selection is in (its run holds the caret), if any. */
export function controlAtSelection(editor: Editor): ContentControlRange | null {
  const { from, to } = editor.state.selection;
  return (
    collectContentControlRanges(editor.state.doc).find((range) => range.from <= from && to <= range.to) ?? null
  );
}

/** The group a new option button joins: the nearest one before it, else Word's first. */
function radioGroupBefore(editor: Editor): string {
  const at = editor.state.selection.from;
  const before = collectContentControlRanges(editor.state.doc).filter(
    (range) => range.meta.radioGroup && range.to <= at,
  );
  return before[before.length - 1]?.meta.radioGroup || DEFAULT_RADIO_GROUP;
}

export function DocumentFormsRibbon({
  editor,
  onExportFillablePdf,
}: {
  readonly editor: Editor;
  /** Save as a fillable PDF form (the saved page's export); absent: no saved page. */
  readonly onExportFillablePdf?: () => void;
  /** The ribbon tab the groups belong to (the strip filters by it). */
  readonly ribbon?: "forms";
}) {
  const t = useTranslations("console-notebooks");
  const [properties, setProperties] = useState<{ from: number; meta: ContentControlMeta } | null>(null);
  const editable = editor.isEditable;

  const insert = (type: ContentControlType | "radio") => {
    const controlType: ContentControlType = type === "radio" ? "checkbox" : type;
    const prompt = wordControlPlaceholderText(controlType);
    const done = editor
      .chain()
      .focus()
      .insertContentControl({
        type: controlType,
        placeholderText: prompt,
        firstItemText: prompt,
        dateFormat: "M/d/yyyy",
        ...(type === "radio" ? { radioGroup: radioGroupBefore(editor) } : {}),
      })
      .run();
    if (!done) toast({ title: t("forms.insertUnavailable"), variant: "destructive" });
  };

  const openProperties = () => {
    const range = controlAtSelection(editor);
    if (!range) {
      toast({ title: t("forms.noControlSelected") });
      return;
    }
    setProperties({ from: range.from, meta: range.meta });
  };

  const move = (direction: 1 | -1) => {
    if (!formFieldRanges(editor.state.doc).length) {
      toast({ title: t("forms.noFields") });
      return;
    }
    editor.chain().focus().goToFormField(direction).run();
  };

  const submit = () => {
    if (!formFieldRanges(editor.state.doc).length) {
      toast({ title: t("forms.noFields") });
      return;
    }
    const missing = missingRequiredFormFields(editor.state.doc);
    editor.chain().focus().markMissingFormFields(missing).run();
    if (missing.length) {
      toast({
        title: t("forms.submitMissing", { count: missing.length }),
        description: missing
          .slice(0, 5)
          .map((field) => field.label)
          .filter(Boolean)
          .join(t("forms.listSeparator")),
        variant: "destructive",
      });
      return;
    }
    toast({ title: t("forms.submitComplete") });
  };

  return (
    <>
      <OfficeHomeRibbonGroup label={t("forms.group.controls")} ribbon="forms" testId="document-toolbar-forms-controls-group">
        {/* Word's Developer > Controls: the controls as small icons in rows
            of a grid, beside Properties -- the group stands the ribbon's
            height. Eight large buttons in one group wrapped into a column
            taller than the ribbon, and seven of them were cut off below it
            (docx rounds R280-D22). */}
        <OfficeHomeCommandStack rows={3}>
          {[0, 3, 6].map((start) => (
            <OfficeHomeCommandRow key={start} testId={`document-toolbar-forms-controls-row-${start / 3}`}>
              {CONTROL_BUTTONS.slice(start, start + 3).map(({ type, icon }) => (
                <OfficeHomeCommand
                  key={type}
                  controlId={`DOCCTL.RIBBON.FORMS.INSERT.${type.toUpperCase()}`}
                  disabled={!editable}
                  displayLabel={t(`forms.insert.${type}`)}
                  icon={<OfficeSemanticIcon semanticId={icon} size={16} />}
                  label={t(`forms.insert.${type}`)}
                  onClick={() => insert(type)}
                  testId={`document-toolbar-forms-insert-${type}`}
                  variant="icon"
                />
              ))}
            </OfficeHomeCommandRow>
          ))}
        </OfficeHomeCommandStack>
        <OfficeHomeCommandStack rows={3}>
          <OfficeHomeCommand
            controlId="DOCCTL.RIBBON.FORMS.PROPERTIES"
            disabled={!editable}
            displayLabel={t("forms.properties")}
            icon={<OfficeSemanticIcon semanticId="form.properties" size={16} />}
            label={t("forms.properties")}
            onClick={openProperties}
            testId="document-toolbar-forms-properties"
          />
        </OfficeHomeCommandStack>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup label={t("forms.group.fill")} ribbon="forms" testId="document-toolbar-forms-fill-group">
        <OfficeHomeCommandStack>
          <OfficeHomeCommand
            controlId="DOCCTL.RIBBON.FORMS.PREVIOUS"
            displayLabel={t("forms.previous")}
            icon={<OfficeSemanticIcon semanticId="edit.findPrevious" size={16} />}
            label={t("forms.previous")}
            onClick={() => move(-1)}
            testId="document-toolbar-forms-previous"
          />
          <OfficeHomeCommand
            controlId="DOCCTL.RIBBON.FORMS.NEXT"
            displayLabel={t("forms.next")}
            icon={<OfficeSemanticIcon semanticId="edit.findNext" size={16} />}
            label={t("forms.next")}
            onClick={() => move(1)}
            testId="document-toolbar-forms-next"
          />
          <OfficeHomeCommand
            controlId="DOCCTL.RIBBON.FORMS.CLEAR"
            disabled={!editable}
            displayLabel={t("forms.clear")}
            icon={<OfficeSemanticIcon semanticId="format.clear" size={16} />}
            label={t("forms.clear")}
            onClick={() => {
              if (!editor.chain().focus().clearFormFields().run()) toast({ title: t("forms.nothingToClear") });
            }}
            testId="document-toolbar-forms-clear"
          />
        </OfficeHomeCommandStack>
        <OfficeHomeCommand
          controlId="DOCCTL.RIBBON.FORMS.SUBMIT"
          displayLabel={t("forms.submit")}
          icon={<OfficeSemanticIcon semanticId="status.checkDouble" size={24} />}
          label={t("forms.submit")}
          onClick={submit}
          testId="document-toolbar-forms-submit"
          variant="large"
        />
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup label={t("forms.group.save")} ribbon="forms" testId="document-toolbar-forms-save-group">
        <OfficeHomeCommand
          controlId="DOCCTL.RIBBON.FORMS.SAVE_FILLABLE_PDF"
          disabled={!onExportFillablePdf}
          displayLabel={t("forms.exportFillablePdf")}
          icon={<OfficeSemanticIcon semanticId="file.export" size={24} />}
          label={t("forms.exportFillablePdf")}
          onClick={() => onExportFillablePdf?.()}
          testId="document-toolbar-forms-export-fillable-pdf"
          variant="large"
        />
      </OfficeHomeRibbonGroup>
      {properties ? (
        <DocumentContentControlPropertiesDialog
          meta={properties.meta}
          canEdit={editable}
          onApply={(patch) => editor.commands.setContentControlProperties(properties.from, patch)}
          onCancel={() => setProperties(null)}
        />
      ) : null}
    </>
  );
}

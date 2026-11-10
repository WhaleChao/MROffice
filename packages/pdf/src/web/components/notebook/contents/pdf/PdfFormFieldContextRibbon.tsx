"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfFormFieldDialogSection } from "./PdfFormFieldAuthoringDialog";
import {
  MoveHorizontal,
  MoveVertical,
  PenLine,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";

import { useRef, useState } from "react";
import {
  OfficeHomeCheckbox,
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";

import styles from "./PdfFormFieldContextRibbon.module.css";
import type { PdfRibbonTranslator } from "./pdfRibbonTranslator";
import ribbonLayout from "./pdfRibbonLayout.module.css";

export type PdfFormFieldContextKind =
  | "text"
  | "checkbox"
  | "radio"
  | "dropdown"
  | "listbox"
  | "signature"
  | "button";

export interface PdfFormFieldContextRibbonModel {
  name: string;
  kind: PdfFormFieldContextKind;
  required: boolean;
  readOnly: boolean;
  width: number;
  height: number;
  busy: boolean;
  geometryEnabled: boolean;
  geometryDisabledReason?: string;
  appearanceStyle?: {
    backgroundColor: string;
    borderColor: string;
    borderWidth: number;
  };
  appearanceStyleEnabled?: boolean;
  appearanceStyleDisabledReason?: string;
  quickPropertiesEnabled: boolean;
  quickPropertiesDisabledReason?: string;
  propertiesEnabled: boolean;
  propertiesDisabledReason?: string;
  optionsEnabled: boolean;
  optionsDisabledReason?: string;
  actionsEnabled: boolean;
  actionsDisabledReason?: string;
}

export interface PdfFormFieldContextRibbonLabels {
  field: string;
  properties: string;
  appearance: string;
  fillColor: string;
  borderColor: string;
  themeColors: string;
  standardColors: string;
  recentColors: string;
  moreColors: string;
  borderWidth: string;
  options: string;
  positionSize: string;
  width: string;
  height: string;
  actions: string;
  openProperties: string;
  required: string;
  readOnly: string;
  editAppearance: string;
  editOptions: string;
  editPositionSize: string;
  editActions: string;
  exit: string;
  typeLabel: (kind: PdfFormFieldContextKind) => string;
}

export function createPdfFormFieldContextRibbonLabels(
  t: PdfRibbonTranslator,
): PdfFormFieldContextRibbonLabels {
  const typeLabels: Record<PdfFormFieldContextKind, string> = {
    text: t("pdf.formFieldRibbon.typeText"),
    checkbox: t("pdf.formFieldRibbon.typeCheckbox"),
    radio: t("pdf.formFieldRibbon.typeRadio"),
    dropdown: t("pdf.formFieldRibbon.typeDropdown"),
    listbox: t("pdf.formFieldRibbon.typeListbox"),
    signature: t("pdf.formFieldRibbon.typeSignature"),
    button: t("pdf.formFieldRibbon.typeButton"),
  };
  return {
    field: t("pdf.formFieldRibbon.field"),
    properties: t("pdf.formFieldRibbon.properties"),
    appearance: t("pdf.formFieldRibbon.appearance"),
    fillColor: t("pdf.formFieldRibbon.fillColor"),
    borderColor: t("pdf.formFieldRibbon.borderColor"),
    themeColors: t("pdf.formField.themeColors"),
    standardColors: t("pdf.formField.standardColors"),
    recentColors: t("pdf.formField.recentColors"),
    moreColors: t("pdf.formField.moreColors"),
    borderWidth: t("pdf.formFieldRibbon.borderWidth"),
    options: t("pdf.formFieldRibbon.options"),
    positionSize: t("pdf.formFieldRibbon.positionSize"),
    width: t("pdf.formFieldRibbon.width"),
    height: t("pdf.formFieldRibbon.height"),
    actions: t("pdf.formFieldRibbon.actions"),
    openProperties: t("pdf.formFieldRibbon.openProperties"),
    required: t("pdf.formFieldRibbon.required"),
    readOnly: t("pdf.formFieldRibbon.readOnly"),
    editAppearance: t("pdf.formFieldRibbon.editAppearance"),
    editOptions: t("pdf.formFieldRibbon.editOptions"),
    editPositionSize: t("pdf.formFieldRibbon.editPositionSize"),
    editActions: t("pdf.formFieldRibbon.editActions"),
    exit: t("pdf.formFieldRibbon.exit"),
    typeLabel: (kind) => typeLabels[kind],
  };
}

function capabilityTitle(
  enabled: boolean,
  reason?: string,
): string | undefined {
  return enabled ? undefined : reason;
}

export function PdfFormFieldContextRibbon({
  labels,
  model,
  onOpenProperties,
  onOpenOptions,
  onRequiredChange,
  onReadOnlyChange,
  onGeometrySizeChange,
  onAppearanceStyleChange,
  onExit,
}: {
  labels: PdfFormFieldContextRibbonLabels;
  model: PdfFormFieldContextRibbonModel;
  onOpenProperties: (section?: PdfFormFieldDialogSection) => void;
  onOpenOptions: () => void;
  onRequiredChange: (required: boolean) => void;
  onReadOnlyChange: (readOnly: boolean) => void;
  onGeometrySizeChange: (axis: "width" | "height", value: number) => void;
  onAppearanceStyleChange: (change: {
    backgroundColor?: string;
    borderColor?: string;
    borderWidth?: number;
  }) => void;
  onExit: () => void;
}) {
  const propertyDisabled = model.busy || !model.propertiesEnabled;
  const quickPropertyDisabled = model.busy || !model.quickPropertiesEnabled;
  const optionsDisabled = model.busy || !model.optionsEnabled;
  const actionsDisabled = model.busy || !model.actionsEnabled;
  const geometryDisabled = model.busy || !model.geometryEnabled;
  const appearanceStyleDisabled =
    model.busy || model.appearanceStyleEnabled === false;
  const [palette, setPalette] = useState<"fill" | "border" | null>(null);
  const fillAnchorRef = useRef<HTMLButtonElement | null>(null);
  const borderAnchorRef = useRef<HTMLButtonElement | null>(null);
  const paletteLabels = {
    themeColors: labels.themeColors,
    standardColors: labels.standardColors,
    recentColors: labels.recentColors,
    moreColors: labels.moreColors,
  };
  const propertyTitle = capabilityTitle(
    model.propertiesEnabled,
    model.propertiesDisabledReason,
  );
  const quickPropertyTitle = capabilityTitle(
    model.quickPropertiesEnabled,
    model.quickPropertiesDisabledReason,
  );
  const optionsTitle = capabilityTitle(
    model.optionsEnabled,
    model.optionsDisabledReason,
  );
  const actionsTitle = capabilityTitle(
    model.actionsEnabled,
    model.actionsDisabledReason,
  );
  const geometryTitle = capabilityTitle(
    model.geometryEnabled,
    model.geometryDisabledReason,
  );
  const appearanceStyleTitle = capabilityTitle(
    model.appearanceStyleEnabled !== false,
    model.appearanceStyleDisabledReason,
  );
  const commandLabel = (label: string, reason?: string): string =>
    reason ? `${label}: ${reason}` : label;

  return (
    <>
      <OfficeHomeRibbonGroup
        label={labels.field}
        ribbon="form-field-format"
        controlId="PDFCTL.CONTEXT.FORM_FIELD.FIELD"
        className={styles.group}
        testId="pdf-form-context-field-group"
      >
        {/* No identity panel (WPS CT_FormEdit: none); the properties
            dialog names the field. */}
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.FORM_FIELD.PROPERTIES.OPEN"
          disabled={propertyDisabled}
          icon={<OfficeSemanticIcon semanticId="form.properties" size={24} />}
          label={commandLabel(labels.openProperties, propertyTitle)}
          displayLabel={labels.openProperties}
          onClick={() => onOpenProperties("general")}
          testId="pdf-form-context-open-properties"
          variant="large"
        />
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        label={labels.properties}
        ribbon="form-field-format"
        controlId="PDFCTL.CONTEXT.FORM_FIELD.PROPERTIES"
        className={styles.group}
        testId="pdf-form-context-properties-group"
      >
        <OfficeHomeCommandStack className={ribbonLayout.keepsWidth}>
          <OfficeHomeCheckbox
            checked={model.required}
            className={styles.hitTargetCheckbox}
            controlId="PDFCTL.CONTEXT.FORM_FIELD.REQUIRED"
            disabled={quickPropertyDisabled}
            label={labels.required}
            onCheckedChange={onRequiredChange}
            testId="pdf-form-context-required"
            title={quickPropertyTitle}
          />
          <OfficeHomeCheckbox
            checked={model.readOnly}
            className={styles.hitTargetCheckbox}
            controlId="PDFCTL.CONTEXT.FORM_FIELD.READ_ONLY"
            disabled={quickPropertyDisabled}
            label={labels.readOnly}
            onCheckedChange={onReadOnlyChange}
            testId="pdf-form-context-readonly"
            title={quickPropertyTitle}
          />
        </OfficeHomeCommandStack>
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        label={labels.appearance}
        ribbon="form-field-format"
        controlId="PDFCTL.CONTEXT.FORM_FIELD.APPEARANCE"
        className={styles.group}
        testId="pdf-form-context-appearance-group"
      >
        {model.appearanceStyle ? (
          <OfficeHomeCommandRow testId="pdf-form-context-appearance-controls">
            {

}
            {(
              [
                ["fill", labels.fillColor, "shape.fill", model.appearanceStyle.backgroundColor, "pdf-form-context-fill-color", "PDFCTL.CONTEXT.FORM_FIELD.FILL_COLOR"],
                ["border", labels.borderColor, "shape.outline", model.appearanceStyle.borderColor, "pdf-form-context-border-color", "PDFCTL.CONTEXT.FORM_FIELD.BORDER_COLOR"],
              ] as const
            ).map(([which, label, glyph, value, testId, controlId]) => (
              <span key={which} className={styles.paletteHost}>
                <OfficeHomeMenuTrigger
                  ref={which === "fill" ? fillAnchorRef : borderAnchorRef}
                  variant="compact"
                  hasMenuIndicator
                  active={palette === which}
                  aria-expanded={palette === which}
                  aria-haspopup="dialog"
                  controlId={controlId}
                  disabled={appearanceStyleDisabled}
                  disabledReason={appearanceStyleTitle}
                  displayLabel={label}
                  icon={
                    <span className="note-editor-office-toolbar__color-glyph">
                      <OfficeSemanticIcon semanticId={glyph} size={18} />
                      <span
                        className="note-editor-office-toolbar__color-underline"
                        style={{ backgroundColor: value }}
                      />
                    </span>
                  }
                  label={commandLabel(label, appearanceStyleTitle)}
                  onClick={() =>
                    setPalette((current) => (current === which ? null : which))
                  }
                  testId={testId}
                />
                {palette === which ? (
                  <OfficeColorPalettePopover
                    anchorRef={which === "fill" ? fillAnchorRef : borderAnchorRef}
                    labels={paletteLabels}
                    testIdPrefix={`${testId}-palette`}
                    value={value}
                    onDismiss={() => setPalette(null)}
                    onPick={(color) => {
                      setPalette(null);
                      // The palette offers no "no fill" here, so a pick is a colour.
                      if (!color) return;
                      onAppearanceStyleChange(
                        which === "fill" ? { backgroundColor: color } : { borderColor: color },
                      );
                    }}
                  />
                ) : null}
              </span>
            ))}
            <OfficeRibbonNumberField
              controlId="PDFCTL.CONTEXT.FORM_FIELD.BORDER_WIDTH"
              disabled={appearanceStyleDisabled}
              disabledReason={appearanceStyleTitle}
              icon={<PenLine aria-hidden size={18} />}
              label={commandLabel(labels.borderWidth, appearanceStyleTitle)}
              max={12}
              min={0.1}
              onCommit={(raw) => {
                const value = Number.parseFloat(raw);
                if (
                  Number.isFinite(value) &&
                  value >= 0.1 &&
                  value <= 12 &&
                  Math.abs(value - model.appearanceStyle!.borderWidth) >= 0.05
                ) {
                  onAppearanceStyleChange({ borderWidth: value });
                }
              }}
              step={0.1}
              testId="pdf-form-context-border-width"
              unit="pt"
              value={model.appearanceStyle.borderWidth.toFixed(1)}
            />
          </OfficeHomeCommandRow>
        ) : (
          <OfficeHomeCommand
            activation="click"
            controlId="PDFCTL.CONTEXT.FORM_FIELD.APPEARANCE.OPEN"
            disabled={propertyDisabled}
            icon={<OfficeSemanticIcon semanticId="form.appearance" size={24} />}
            label={commandLabel(labels.editAppearance, propertyTitle)}
            displayLabel={labels.editAppearance}
            onClick={() => onOpenProperties("appearance")}
            testId="pdf-form-context-appearance"
          />
        )}
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        label={labels.options}
        ribbon="form-field-format"
        controlId="PDFCTL.CONTEXT.FORM_FIELD.OPTIONS"
        className={styles.group}
        testId="pdf-form-context-options-group"
      >
        <OfficeHomeCommand
          activation="click"
          controlId="PDFCTL.CONTEXT.FORM_FIELD.OPTIONS.OPEN"
          disabled={optionsDisabled}
          icon={<OfficeSemanticIcon semanticId="form.options" size={24} />}
          label={commandLabel(labels.editOptions, optionsTitle)}
          displayLabel={labels.editOptions}
          onClick={onOpenOptions}
          testId="pdf-form-context-options"
          variant="large"
        />
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        label={labels.positionSize}
        ribbon="form-field-format"
        controlId="PDFCTL.CONTEXT.FORM_FIELD.POSITION_SIZE"
        className={styles.group}
        testId="pdf-form-context-position-size-group"
      >
        <OfficeHomeCommandRow testId="pdf-form-context-size-controls">
          <OfficeRibbonNumberField
            controlId="PDFCTL.CONTEXT.FORM_FIELD.WIDTH"
            disabled={geometryDisabled}
            disabledReason={geometryTitle}
            icon={<MoveHorizontal aria-hidden size={18} />}
            label={commandLabel(labels.width, geometryTitle)}
            max={100}
            min={0.1}
            onCommit={(raw) => {
              const value = Number.parseFloat(raw);
              if (
                Number.isFinite(value) &&
                value > 0 &&
                value <= 100 &&
                Math.abs(value - model.width) >= 0.05
              ) {
                onGeometrySizeChange("width", value);
              }
            }}
            step={0.1}
            testId="pdf-form-context-width"
            unit="%"
            value={model.width.toFixed(1)}
          />
          <OfficeRibbonNumberField
            controlId="PDFCTL.CONTEXT.FORM_FIELD.HEIGHT"
            disabled={geometryDisabled}
            disabledReason={geometryTitle}
            icon={<MoveVertical aria-hidden size={18} />}
            label={commandLabel(labels.height, geometryTitle)}
            max={100}
            min={0.1}
            onCommit={(raw) => {
              const value = Number.parseFloat(raw);
              if (
                Number.isFinite(value) &&
                value > 0 &&
                value <= 100 &&
                Math.abs(value - model.height) >= 0.05
              ) {
                onGeometrySizeChange("height", value);
              }
            }}
            step={0.1}
            testId="pdf-form-context-height"
            unit="%"
            value={model.height.toFixed(1)}
          />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>

      <OfficeHomeRibbonGroup
        label={labels.actions}
        ribbon="form-field-format"
        controlId="PDFCTL.CONTEXT.FORM_FIELD.ACTIONS"
        className={styles.group}
        testId="pdf-form-context-actions-group"
      >
        <OfficeHomeCommandRow>
          <OfficeHomeCommand
            activation="click"
            controlId="PDFCTL.CONTEXT.FORM_FIELD.ACTIONS.OPEN"
            disabled={actionsDisabled}
            icon={<OfficeSemanticIcon semanticId="form.actions" size={24} />}
            label={commandLabel(labels.editActions, actionsTitle)}
            displayLabel={labels.editActions}
            onClick={() => onOpenProperties("actions")}
            testId="pdf-form-context-actions"
            variant="large"
          />
          <OfficeHomeCommand
            activation="click"
            controlId="PDFCTL.CONTEXT.FORM_FIELD.EXIT"
            disabled={model.busy}
            displayLabel={labels.exit}
            icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
            label={labels.exit}
            onClick={onExit}
            testId="pdf-form-context-exit"
            variant="large"
          />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
    </>
  );
}

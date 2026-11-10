"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Tools group Office draws in front of the shared Equation tab's own
 * groups: the Equation gallery (the Docs gallery, so every editor offers the
 * same equations under the same names, with Save Selection to Equation
 * Gallery) and Normal Text, which switches the built-up draft's caret between
 * math and ordinary text (MathLive's text mode, `\text{...}` in the LaTeX).
 * Every control is part of the open draft
 * (`VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE`): pressing one never commits it.
 * A slide's and a worksheet's Equation tab draw it with their own control
 * ids and test ids.
 */

import type { EquationLayoutMode } from "../console/editor/extensions/equationCatalog";
import { VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE } from "../console/editor/extensions/visualMathEditor";
import { DocsInsertIcon } from "../console/editor/DocumentInsertRibbonIcons";
import {
  DocumentEquationGalleryMenu,
  type EquationInsertPayload,
} from "../console/editor/native-document/DocumentEquationGalleryMenu";
import styles from "../console/editor/native-document/DocumentEquationContextRibbon.module.css";
import { OfficeHomeCommand, OfficeHomeMenuTrigger } from "../office-ribbon/OfficeHomeRibbon";
import { OfficeRibbonGroup, RibbonHStack } from "../office-ribbon/OfficeRibbonChassis";
import { DropdownMenu, DropdownMenuTrigger } from "../ui/dropdown-menu";

const focusOwner = { [VISUAL_MATH_EDITOR_FOCUS_OWNER_ATTRIBUTE]: "true" } as const;

export interface OfficeEquationToolsGroupProps {
  readonly t: (key: string, values?: Record<string, string>) => string;
  readonly locale: string;
  readonly disabled: boolean;
  /** The editor's layout while a draft is open, else the default. */
  readonly layoutMode: EquationLayoutMode;
  /** The selected equation's LaTeX (Save Selection to Equation Gallery). */
  readonly selectedEquationLatex: string | null;
  /** A gallery equation: its LaTeX, and the whole insertion payload (the
   * document places it as the gallery entry says, inline or display). */
  readonly onInsertEquation: (latex: string, payload: EquationInsertPayload) => void;
  readonly onInsertNewEquation: () => void;
  /** Normal Text is on at the open draft's caret. */
  readonly normalText: boolean;
  readonly onToggleNormalText: () => void;
  /** The host's words for the group and for Normal Text. */
  readonly groupLabel: string;
  readonly equationLabel: string;
  readonly normalTextLabel: string;
  readonly controlIds: { readonly gallery: string; readonly normalText: string };
  /** `<prefix>-ribbon-tools`, `<prefix>-ribbon-gallery`, `<prefix>-normal-text`. */
  readonly testIdPrefix: string;
}

export function OfficeEquationToolsGroup({
  t,
  locale,
  disabled,
  layoutMode,
  selectedEquationLatex,
  onInsertEquation,
  onInsertNewEquation,
  normalText,
  onToggleNormalText,
  groupLabel,
  equationLabel,
  normalTextLabel,
  controlIds,
  testIdPrefix,
}: OfficeEquationToolsGroupProps) {
  return (
    <OfficeRibbonGroup
      ribbon="equation"
      label={groupLabel}
      className={styles.ribbonGroup}
      testId={`${testIdPrefix}-ribbon-tools`}
      dataAttributes={focusOwner}
    >
      <RibbonHStack>
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              {...focusOwner}
              controlId={controlIds.gallery}
              variant="large"
              hasMenuIndicator
              disabled={disabled}
              icon={<DocsInsertIcon name="equation" size={24} />}
              label={equationLabel}
              testId={`${testIdPrefix}-ribbon-gallery`}
            />
          </DropdownMenuTrigger>
          <DocumentEquationGalleryMenu
            t={t as (key: string) => string}
            locale={locale}
            testIdPrefix={`${testIdPrefix}-ribbon`}
            align="start"
            onInsert={(payload) => onInsertEquation(payload.attrs.latex, payload)}
            onInsertNew={onInsertNewEquation}
            selectedEquationLatex={selectedEquationLatex}
          />
        </DropdownMenu>
        {/* A large face beside Equation, as Office draws its Tools group;
          the group carries the draft's focus-owner mark. */}
        <OfficeHomeCommand
          controlId={controlIds.normalText}
          variant="large"
          active={normalText}
          disabled={disabled || layoutMode !== "professional"}
          icon={<span aria-hidden className={styles.normalTextGlyph}>abc</span>}
          label={normalTextLabel}
          onClick={onToggleNormalText}
          testId={`${testIdPrefix}-normal-text`}
        />
      </RibbonHStack>
    </OfficeRibbonGroup>
  );
}

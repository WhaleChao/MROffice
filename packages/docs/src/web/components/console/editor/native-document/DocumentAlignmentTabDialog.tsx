"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's "Alignment Tab" dialog (Header & Footer > Position > Insert
 * Alignment Tab): an alignment (left / centre / right), what it is relative
 * to (margin / indent) and a leader. Word writes it as a positional tab
 * (`w:ptab`); MRAI applies the same geometry as the paragraph's own tab stop
 * -- centre at half the text width, right at the width -- plus a tab, the
 * form every band writer and reader carries and Word reads identically.
 */

import { useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import type { BandRichTabStop } from "@mrnote/core/web/components/console/editor/native-document/headerFooter.ts";

export type AlignmentTabAlignment = "left" | "center" | "right";
export type AlignmentTabRelativeTo = "margin" | "indent";

export interface AlignmentTabChoice {
  alignment: AlignmentTabAlignment;
  relativeTo: AlignmentTabRelativeTo;
  leader: BandRichTabStop["leader"];
}

/**
 * The tab stop an alignment tab resolves to on a line of the given text
 * width (twips), less the paragraph's indents when relative to them. A left
 * alignment tab needs no stop: the tab itself goes to the left edge.
 */
export function alignmentTabStop(
  choice: AlignmentTabChoice,
  textWidthTwips: number,
  indents: { leftTwips: number; rightTwips: number } = { leftTwips: 0, rightTwips: 0 },
): BandRichTabStop | null {
  const width =
    choice.relativeTo === "indent"
      ? textWidthTwips - indents.leftTwips - indents.rightTwips
      : textWidthTwips;
  if (width < 144) return null;
  if (choice.alignment === "center") {
    return { pos: Math.round(width / 2), align: "center", leader: choice.leader };
  }
  if (choice.alignment === "right") {
    return { pos: Math.round(width), align: "right", leader: choice.leader };
  }
  return null;
}

const LEADER_GLYPHS: Record<BandRichTabStop["leader"], string> = {
  none: "",
  dot: "......",
  hyphen: "------",
  underscore: "______",
};

export interface DocumentAlignmentTabDialogProps {
  readonly open: boolean;
  readonly onSubmit: (choice: AlignmentTabChoice) => void;
  readonly onCancel: () => void;
}

export function DocumentAlignmentTabDialog({
  open,
  onSubmit,
  onCancel,
}: DocumentAlignmentTabDialogProps) {
  const t = useTranslations("console-notebooks");
  const [alignment, setAlignment] = useState<AlignmentTabAlignment>("center");
  const [relativeTo, setRelativeTo] = useState<AlignmentTabRelativeTo>("margin");
  const [leader, setLeader] = useState<BandRichTabStop["leader"]>("none");
  return (
    <OfficeDialogShell
      open={open}
      title={t("toolbar.alignmentTabTitle")}
      description={t("toolbar.alignmentTabDescription")}
      defaultActionLabel={t("toolbar.dialogOk")}
      cancelActionLabel={t("toolbar.dialogCancel")}
      onDefaultAction={() => onSubmit({ alignment, relativeTo, leader })}
      onCancel={onCancel}
      testIdPrefix="document-alignment-tab"
    >
      <fieldset className="office-dialog__group">
        <legend>{t("toolbar.alignmentTabAlignment")}</legend>
        {(["left", "center", "right"] as const).map((value) => (
          <label key={value} className="office-dialog__check">
            <input
              type="radio"
              name="alignment-tab-alignment"
              data-testid={`document-alignment-tab-${value}`}
              checked={alignment === value}
              onChange={() => setAlignment(value)}
            />
            <DocumentDialogIcon icon={`paragraph.align.${value}`} />
            {t(`toolbar.alignmentTab_${value}`)}
          </label>
        ))}
      </fieldset>
      <fieldset className="office-dialog__group">
        <legend>{t("toolbar.alignmentTabRelativeTo")}</legend>
        {(["margin", "indent"] as const).map((value) => (
          <label key={value} className="office-dialog__check">
            <input
              type="radio"
              name="alignment-tab-relative"
              data-testid={`document-alignment-tab-relative-${value}`}
              checked={relativeTo === value}
              onChange={() => setRelativeTo(value)}
            />
            <DocumentDialogIcon icon={value === "margin" ? "view.pageGuides" : "paragraph.indent.increase"} />
            {t(`toolbar.alignmentTabRelative_${value}`)}
          </label>
        ))}
      </fieldset>
      <label className="office-dialog__field">
        {t("toolbar.alignmentTabLeader")}
        <select
          data-testid="document-alignment-tab-leader"
          value={leader}
          onChange={(event) => setLeader(event.target.value as BandRichTabStop["leader"])}
        >
          {(["none", "dot", "hyphen", "underscore"] as const).map((value) => (
            <option key={value} value={value}>
              {value === "none" ? t("toolbar.tocTabLeader_none") : LEADER_GLYPHS[value]}
            </option>
          ))}
        </select>
      </label>
    </OfficeDialogShell>
  );
}

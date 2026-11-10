"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/core";
import { Check } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { PointerEventHandler } from "react";

import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { RibbonDropdownMenuContent as DropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { readDocumentTableDesignTarget } from "@mrnote/core/web/lib/document-table-design-authority.ts";
import {
  OFFICE_WORD_TABLE_STYLE_COLORS,
  officeWordTableStyleColor,
  officeWordTableStyleRecolor,
  type OfficeWordTableStyleColor,
} from "@mrnote/core/web/lib/office-table-style-word.ts";

import { documentOfficeTableTheme } from "@mrnote/core/web/components/console/editor/native-document/documentOfficeTableStyles.ts";
import {
  applyDocumentOfficeTableColor,
  nativeDocumentTableColorState,
  readDocumentOfficeTableColorTarget,
  type DocumentOfficeTableColorTarget,
} from "./tableColorButtons";
import {
  setDocumentLocalTableStylePreview,
  setDocumentTableStylePreview,
} from "@mrnote/core/web/components/console/editor/native-document/tableBorders.ts";

/**
 * Table Color for a table wearing an Office gallery style: the gallery's
 * seven colour columns (neutral, accents 1-6) as one control. A colour is the
 * same style in that column (`officeWordTableStyleRecolor`), so the family
 * (Light/Medium/Dark and its number) stays and the choice saves, exports and
 * imports exactly as picking that tile in the gallery does. Word's own Table
 * Design has no such command; there the colour is the gallery column.
 */
export {
  applyDocumentOfficeTableColor,
  readDocumentOfficeTableColorTarget,
  type DocumentOfficeTableColorTarget,
};

function previewColor(
  editor: Editor,
  target: DocumentOfficeTableColorTarget,
  color: OfficeWordTableStyleColor | null,
): void {
  const next = color ? officeWordTableStyleRecolor(target.style, color) : null;
  if (target.kind === "local") {
    setDocumentLocalTableStylePreview(editor, next ? next.slot : undefined);
    return;
  }
  if (next) {
    nativeDocumentTableColorState(editor, next.slot)?.preview();
    return;
  }
  const live = readDocumentTableDesignTarget(editor);
  if (live) setDocumentTableStylePreview(editor, live, null);
}

export function DocumentOfficeTableColorMenu({
  editor,
  target,
  label,
  disabledReason,
  onPointerDown,
}: {
  readonly editor: Editor;
  readonly target: DocumentOfficeTableColorTarget;
  readonly locale: "en" | "zh";
  readonly label: string;
  readonly disabledReason?: string;
  readonly onPointerDown?: PointerEventHandler<HTMLButtonElement>;
}) {
  const theme = documentOfficeTableTheme(editor.state.doc);
  const current = officeWordTableStyleColor(target.style);
  const t = useTranslations("console-notebooks");
  const names = {
    neutral: t("toolbar.tableColorNeutral"),
    accent: (n: number) => t("toolbar.tableColorAccent", { n }),
  };
  return (
    <DropdownMenu
      onOpenChange={(open) => {
        if (!open) previewColor(editor, target, null);
      }}
    >
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          disabled={disabledReason !== undefined}
          disabledReason={disabledReason}
          hasMenuIndicator
          icon={<OfficeSemanticIcon semanticId="table.style.color" size={24} />}
          label={label}
          onPointerDown={onPointerDown}
          testId="document-toolbar-table-color"
          variant="large"
        />
      </DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        data-testid="document-toolbar-table-color-menu"
        sideOffset={6}
      >
        {OFFICE_WORD_TABLE_STYLE_COLORS.map((color) => {
          const next = officeWordTableStyleRecolor(target.style, color);
          const key = color === "neutral" ? "neutral" : `accent${color}`;
          return (
            <DropdownMenuItem
              key={key}
              data-testid={`document-toolbar-table-color-${key}`}
              disabled={!next}
              onFocus={() => previewColor(editor, target, color)}
              onPointerEnter={() => previewColor(editor, target, color)}
              onPointerLeave={() => previewColor(editor, target, null)}
              onSelect={() => {
                previewColor(editor, target, null);
                if (applyDocumentOfficeTableColor(editor, color)) {
                  editor.chain().focus().run();
                }
              }}
            >
              <span
                aria-hidden
                // A theme colour swatch: its real colour in every theme, with
                // a mid-grey edge that reads on the light and the dark menu.
                data-chrome-preview=""
                style={{
                  backgroundColor:
                    color === "neutral" ? theme.text1 : theme[`accent${color}`],
                  border: "1px solid rgb(128 128 128 / 45%)",
                  borderRadius: 3,
                  display: "inline-block",
                  height: 16,
                  width: 16,
                }}
              />
              <span>
                {color === "neutral" ? names.neutral : names.accent(color)}
              </span>
              {current === color ? <Check aria-hidden size={14} /> : null}
            </DropdownMenuItem>
          );
        })}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

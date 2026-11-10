"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Tape, a note-taking app's tape for learning by heart: the pen lays
 * wide striped strips over the words; a tap on a strip peels it back so the
 * words show, another lays it down again (`inkTape`, documentShapeExtension
 * and flipDocumentInkTapeAt). The menu picks the strip's colour and width.
 */

import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeHomeCommand, OfficeHomeLargeSplitLayout, OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { DOCUMENT_INK_TAPE_WIDTHS_PT, type DocumentInkPreset } from "@mrnote/core/web/lib/document-ink.ts";
import { TAPE_INK_EFFECTS, inkEffectPaint, isTapeInkEffect, type TapeInkEffect } from "@mrnote/core/web/lib/office-ink-effects.ts";

/** The pen a Tape tool of this colour and width lays its strips with. */
export function documentInkTapePreset(effect: TapeInkEffect, widthPt: number): DocumentInkPreset {
  return {
    id: `tape-${effect}`,
    tool: "pen",
    color: inkEffectPaint(effect).solid,
    widthPt,
    effect,
    // One width however hard it is pressed: tape is cut, not inked.
    steady: true,
    tape: "covered",
  };
}

/** A strip of the tape as its menu shows it: the colours in its bands. */
export function DocumentInkTapeSwatch({ effect, size = 16 }: { effect: TapeInkEffect; size?: number }) {
  const [a = "#ccc", b = a] = inkEffectPaint(effect).stops;
  return (
    <svg aria-hidden="true" width={size * 1.6} height={size * 0.7} viewBox="0 0 32 14">
      <path d="M1 1h30l-1.5 3 1.5 3-1.5 3 1.5 3H1l1.5-3L1 7l1.5-3Z" fill={a} />
      <path d="M8 1h5l-5 12H3Zm10 0h5l-5 12h-5Z" fill={b} />
    </svg>
  );
}

export function DocumentInkTapeCommand({
  active,
  effect,
  widthPt,
  disabled,
  disabledReason,
  onArm,
  onExit,
}: {
  active: boolean;
  /** The strip the tool lays (the last one picked). */
  effect: TapeInkEffect;
  widthPt: number;
  disabled: boolean;
  disabledReason?: string;
  onArm: (effect: TapeInkEffect, widthPt: number) => void;
  onExit: () => void;
}) {
  const t = useTranslations("console-notebooks");
  return (
    <OfficeHomeLargeSplitLayout
      testId="document-toolbar-ribbon-draw-tape-split"
      primary={
        <OfficeHomeCommand
          active={active}
          controlId="draw.tape"
          disabled={disabled}
          disabledReason={disabledReason}
          icon={<OfficeSemanticIcon semanticId="draw.tape" size={24} />}
          label={t("toolbar.drawTape")}
          onClick={() => (active ? onExit() : onArm(effect, widthPt))}
          testId="document-toolbar-ribbon-draw-tape"
          variant="large"
        />
      }
      trigger={
        <DropdownMenu modal={false}>
          <DropdownMenuTrigger asChild>
            <OfficeHomeMenuTrigger
              controlId="draw.tapeOptions"
              disabled={disabled}
              disabledReason={disabledReason}
              icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
              label={t("toolbar.drawTapeOptions")}
              title={t("toolbar.drawTapeOptions")}
              testId="document-toolbar-ribbon-draw-tape-menu"
              variant="icon"
            />
          </DropdownMenuTrigger>
          <DropdownMenuContent
            align="start"
            className="note-editor-toolbar-menu"
            data-testid="document-toolbar-ribbon-draw-tape-menu-content"
            onCloseAutoFocus={(event) => event.preventDefault()}
          >
            <DropdownMenuLabel>{t("toolbar.drawTapeHint")}</DropdownMenuLabel>
            <DropdownMenuRadioGroup
              value={effect}
              onValueChange={(value) => {
                if (isTapeInkEffect(value)) onArm(value, widthPt);
              }}
            >
              {TAPE_INK_EFFECTS.map((id) => (
                <DropdownMenuRadioItem key={id} value={id} data-testid={`document-toolbar-ribbon-draw-tape-${id}`}>
                  <DocumentInkTapeSwatch effect={id} />
                  {t(`toolbar.drawTapeColor_${id}`)}
                </DropdownMenuRadioItem>
              ))}
            </DropdownMenuRadioGroup>
            <DropdownMenuSeparator />
            <DropdownMenuRadioGroup
              value={String(widthPt)}
              onValueChange={(value) => {
                const next = Number(value);
                if (DOCUMENT_INK_TAPE_WIDTHS_PT.includes(next)) onArm(effect, next);
              }}
            >
              {DOCUMENT_INK_TAPE_WIDTHS_PT.map((width) => (
                <DropdownMenuRadioItem
                  key={width}
                  value={String(width)}
                  data-testid={`document-toolbar-ribbon-draw-tape-width-${width}`}
                >
                  {t(`toolbar.drawTapeWidth_${width}`)}
                </DropdownMenuRadioItem>
              ))}
            </DropdownMenuRadioGroup>
          </DropdownMenuContent>
        </DropdownMenu>
      }
    />
  );
}

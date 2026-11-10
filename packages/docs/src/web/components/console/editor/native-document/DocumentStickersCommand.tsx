"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Stickers: MRAI's own stickers (lib/document-stickers), each put on
 * the page in view as a floating picture (insertPicturesOverPageInView), so
 * it drags, sizes and turns as any picture does.
 */

import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  DOCUMENT_STICKER_IDS,
  documentStickerDataUrl,
  type DocumentStickerId,
} from "@mrnote/core/web/lib/document-stickers.ts";

export function DocumentStickersCommand({
  disabled,
  disabledReason,
  onInsert,
}: {
  disabled: boolean;
  disabledReason?: string;
  onInsert: (id: DocumentStickerId, name: string) => void;
}) {
  const t = useTranslations("console-notebooks");
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          controlId="draw.stickers"
          variant="large"
          hasMenuIndicator
          disabled={disabled}
          disabledReason={disabledReason}
          icon={<OfficeSemanticIcon semanticId="draw.stickers" size={24} />}
          label={t("toolbar.drawStickers")}
          title={t("toolbar.drawStickers")}
          testId="document-toolbar-ribbon-draw-stickers"
        />
      </DropdownMenuTrigger>
      <DropdownMenuContent
        align="start"
        className="note-editor-toolbar-menu"
        data-testid="document-toolbar-ribbon-draw-stickers-menu"
        onCloseAutoFocus={(event) => event.preventDefault()}
      >
        <DropdownMenuLabel>{t("toolbar.drawStickers")}</DropdownMenuLabel>
        <div role="group" style={{ display: "grid", gridTemplateColumns: "repeat(4, 56px)", gap: 4, padding: 4 }}>
          {DOCUMENT_STICKER_IDS.map((id) => {
            const name = t(`toolbar.drawSticker_${id}`);
            return (
              <DropdownMenuItem
                key={id}
                aria-label={name}
                title={name}
                data-testid={`document-toolbar-ribbon-draw-sticker-${id}`}
                style={{ justifyContent: "center", padding: 4, minHeight: 56 }}
                onSelect={() => onInsert(id, name)}
              >
                {/* eslint-disable-next-line @next/next/no-img-element -- an inline SVG data URL, not a remote image */}
                <img alt="" src={documentStickerDataUrl(id, 96)} width={44} height={44} draggable={false} />
              </DropdownMenuItem>
            );
          })}
        </div>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

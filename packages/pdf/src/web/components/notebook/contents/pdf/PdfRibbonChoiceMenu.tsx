"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactNode } from "react";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";





export function PdfRibbonChoiceMenu({
  controlId,
  disabled,
  disabledReason,
  icon,
  label,
  options,
  testId,
  value,
  onPick,
}: {
  controlId: string;
  disabled: boolean;
  disabledReason?: string;
  /** The trigger's glyph when the current option has no preview. */
  icon?: ReactNode;
  label: string;
  options: readonly { value: string; label: string; preview?: ReactNode }[];
  testId: string;
  value: string;
  onPick: (value: string) => void;
}) {
  const current = options.find((option) => option.value === value);
  // Disabled, the reason is the whole accessible name (as the select's was).
  const accessible = disabled && disabledReason ? disabledReason : label;
  return (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          variant="compact"
          hasMenuIndicator
          controlId={controlId}
          data-value={value}
          disabled={disabled}
          disabledReason={disabled ? disabledReason : undefined}
          displayLabel={current?.label ?? value}
          icon={current?.preview ?? icon ?? <OfficeSemanticIcon semanticId="shape.outline" size={16} />}
          label={accessible}
          testId={testId}
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent align="start" aria-label={label} data-testid={`${testId}-menu`}>
        <DropdownMenuRadioGroup value={value}>
          {options.map((option) => (
            <DropdownMenuRadioItem
              key={option.value}
              value={option.value}
              data-testid={`${testId}-option-${option.value}`}
              onSelect={() => onPick(option.value)}
            >
              {option.preview ? (
                <span aria-hidden style={{ display: "inline-flex", width: 40, marginRight: 8 }}>
                  {option.preview}
                </span>
              ) : null}
              {option.label}
            </DropdownMenuRadioItem>
          ))}
        </DropdownMenuRadioGroup>
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
}


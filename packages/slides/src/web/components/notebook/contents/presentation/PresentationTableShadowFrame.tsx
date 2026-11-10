"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ReactNode } from "react";

import { presentationTableShadowLayers } from "@mrnote/core/web/lib/presentation-table-effects.ts";
import { tableStyleLookShadowEffect } from "@mrnote/core/web/lib/presentation-table-style.ts";

/**
 * Table Design > Effects > Shadow when it needs a layer of its own
 * (`presentationTableShadowLayers`): a perspective shadow under the table, an
 * inner shadow over its cells (their fills would hide an inset shadow on the
 * table). An outer shadow is the table frame's own `box-shadow`, so a table
 * without these layers is returned as it is.
 */
export function PresentationTableShadowFrame({
  block,
  children,
}: {
  block: {
    readonly tableEffects?: unknown;
    readonly tableStyleLook?: unknown;
    readonly width?: unknown;
    readonly height?: unknown;
  };
  children: ReactNode;
}) {
  // The table style's background shadow, when scaled or skewed unevenly
  // (readings row 500), is drawn on the layer under the table too.
  const layers = presentationTableShadowLayers(block, tableStyleLookShadowEffect(block.tableStyleLook));
  if (!layers.under && !layers.over) return <>{children}</>;
  return (
    <div
      data-testid="presentation-table-shadow-frame"
      style={{ position: "relative", isolation: "isolate", height: "100%", minHeight: 0 }}
    >
      {layers.under ? <span aria-hidden="true" data-testid="presentation-table-shadow-under" style={layers.under} /> : null}
      {children}
      {layers.over ? <span aria-hidden="true" data-testid="presentation-table-shadow-over" style={layers.over} /> : null}
    </div>
  );
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Link on a selected object and the Action Settings an Action
 * Button carries: an object has ONE click target (`a:hlinkClick`), either a
 * link (`href`: a web or mail address, or `slide://<id>`) or an Action
 * Button's relative jump (`clickAction`). The deck editor's Link field and an
 * editor command both commit what these return.
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { normalizeLinkTooltip } from "@mrnote/core/web/lib/presentation-link-settings.ts";
import {
  isPresentationActionButtonShapeKind,
  type PresentationRelativeClickActionTarget,
} from "@mrnote/core/web/lib/presentation-block-fx.ts";

/**
 * The block linked to `href` (already sanitized), or unlinked when `href`
 * is null. A link replaces an Action Button's Next / Previous / media action,
 * which used to keep firing over it (PPT-B147). Null when nothing changes.
 *
 * `tooltip` is the link's ScreenTip (Insert Hyperlink > ScreenTip...,
 * PPT-B010): stated with the link, or none; left as it is when omitted (an
 * Action Settings target keeps the ScreenTip its link had). Unlinking drops
 * it with the link.
 */
export function blockWithLink<T extends PresentationBlock>(
  block: T,
  href: string | null,
  tooltip?: string | null,
): T | null {
  const next = { ...block } as T & { linkTooltip?: string };
  const currentTooltip = (block as { linkTooltip?: unknown }).linkTooltip;
  if (href) {
    const nextTooltip =
      tooltip === undefined ? normalizeLinkTooltip(currentTooltip) : normalizeLinkTooltip(tooltip);
    if (next.href === href && !("clickAction" in next) && nextTooltip === currentTooltip) {
      return null;
    }
    next.href = href;
    delete (next as { clickAction?: unknown }).clickAction;
    if (nextTooltip) next.linkTooltip = nextTooltip;
    else delete next.linkTooltip;
    return next;
  }
  if (next.href !== undefined || currentTooltip !== undefined) {
    delete next.href;
    delete next.linkTooltip;
    return next;
  }
  return null;
}

/**
 * An Action Button's click action set to a relative jump (Next, Previous,
 * First, Last, Last Viewed), replacing its link; the block itself when it
 * already jumps there. Null for any other object (only an Action Button
 * stores a click action).
 */
export function actionButtonWithRelativeAction<T extends PresentationBlock>(
  block: T,
  target: PresentationRelativeClickActionTarget,
): T | null {
  if (block.type !== "shape" || !isPresentationActionButtonShapeKind((block as { shape?: unknown }).shape)) {
    return null;
  }
  const current = (block as { clickAction?: { kind?: unknown; target?: unknown } }).clickAction;
  if (current?.kind === "relative-slide" && current.target === target && block.href === undefined) return block;
  const next = { ...block, clickAction: { kind: "relative-slide" as const, target } };
  delete next.href;
  delete (next as { linkTooltip?: unknown }).linkTooltip;
  return next;
}

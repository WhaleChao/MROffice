/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Action Settings (Insert > Action; right-click > Action
 * Settings...): a Mouse Click page and a Mouse Over page, each "None" or
 * "Hyperlink to:" a slide, a relative jump or a URL, with "Highlight click" /
 * "Highlight when mouse over" (PPT-B012). The file keeps the click on
 * `p:cNvPr/a:hlinkClick` and the mouse over on `a:hlinkMouseOver`, the same
 * CT_Hyperlink, whose `@highlightClick` is the highlight (ECMA-376
 * §21.1.2.3.5).
 *
 * The page states them as:
 * - click: `href` (web, mail or `slide://<id>`) or `clickAction` (a relative
 *   jump on any object, or an Action Button's media action), `clickHighlight`;
 * - mouse over: `hoverHref` or `hoverAction` (a relative jump),
 *   `hoverHighlight`.
 * Run Program, Run Macro, Object Action and Play Sound have no browser
 * equivalent: the dialog shows them disabled and the server keeps an object
 * that uses one in the package.
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  isPresentationActionButtonShapeKind,
  normalizePresentationClickAction,
  normalizePresentationRelativeClickAction,
  sanitizeBlockHref,
  type PresentationClickAction,
  type PresentationRelativeClickAction,
} from "@mrnote/core/web/lib/presentation-block-fx.ts";

export type PresentationActionTarget =
  | { readonly href: string }
  | { readonly clickAction: PresentationClickAction };

export type PresentationHoverTarget =
  | { readonly href: string }
  | { readonly action: PresentationRelativeClickAction };

export interface PresentationObjectActionSettings {
  readonly click: PresentationActionTarget | null;
  readonly clickHighlight: boolean;
  readonly hover: PresentationHoverTarget | null;
  readonly hoverHighlight: boolean;
}

export const NO_PRESENTATION_ACTION_SETTINGS: PresentationObjectActionSettings = Object.freeze({
  click: null,
  clickHighlight: false,
  hover: null,
  hoverHighlight: false,
});

type ActionFields = {
  href?: unknown;
  clickAction?: unknown;
  clickHighlight?: unknown;
  hoverHref?: unknown;
  hoverAction?: unknown;
  hoverHighlight?: unknown;
  linkTooltip?: unknown;
};

/** What an object's Action Settings say now. */
export function presentationObjectActionSettings(
  block: PresentationBlock | Record<string, unknown>,
): PresentationObjectActionSettings {
  const fields = block as ActionFields;
  const href = sanitizeBlockHref(fields.href);
  const clickAction = normalizePresentationClickAction(fields.clickAction);
  const click: PresentationActionTarget | null = href
    ? { href }
    : clickAction
      ? { clickAction }
      : null;
  const hoverHref = sanitizeBlockHref(fields.hoverHref);
  const hoverAction = normalizePresentationRelativeClickAction(fields.hoverAction);
  const hover: PresentationHoverTarget | null = hoverHref
    ? { href: hoverHref }
    : hoverAction
      ? { action: hoverAction }
      : null;
  return {
    click,
    clickHighlight: click !== null && fields.clickHighlight === true,
    hover,
    hoverHighlight: hover !== null && fields.hoverHighlight === true,
  };
}

/**
 * The block with `settings` as its Action Settings, or null when nothing
 * changes or `settings` is not one this object can carry (a media action on
 * anything but its Movie / Sound Action Button, a slide that is not in the
 * deck). A ScreenTip stays with an unchanged web or slide link.
 */
export function blockWithActionSettings<T extends PresentationBlock>(
  block: T,
  settings: PresentationObjectActionSettings,
  knownSlideIds?: ReadonlySet<string>,
): T | null {
  const next = { ...block } as T & ActionFields;
  const shapeKind = block.type === "shape" ? (block as { shape?: unknown }).shape : undefined;
  const actionButton = isPresentationActionButtonShapeKind(shapeKind);
  const click = settings.click;
  if (click && "href" in click) {
    const href = sanitizeBlockHref(click.href, knownSlideIds);
    if (!href) return null;
    if (href !== block.href) delete next.linkTooltip;
    next.href = href;
    delete next.clickAction;
  } else if (click) {
    const action = normalizePresentationClickAction(click.clickAction);
    if (!action || (action.kind === "play-media" && !actionButton)) return null;
    next.clickAction = action;
    delete next.href;
    delete next.linkTooltip;
  } else {
    delete next.href;
    delete next.clickAction;
    delete next.linkTooltip;
  }
  if (click && settings.clickHighlight) next.clickHighlight = true;
  else delete next.clickHighlight;
  const hover = settings.hover;
  if (hover && "href" in hover) {
    const href = sanitizeBlockHref(hover.href, knownSlideIds);
    if (!href) return null;
    next.hoverHref = href;
    delete next.hoverAction;
  } else if (hover) {
    const action = normalizePresentationRelativeClickAction(hover.action);
    if (!action) return null;
    next.hoverAction = action;
    delete next.hoverHref;
  } else {
    delete next.hoverHref;
    delete next.hoverAction;
  }
  if (hover && settings.hoverHighlight) next.hoverHighlight = true;
  else delete next.hoverHighlight;
  return JSON.stringify(next) === JSON.stringify(block) ? null : next;
}

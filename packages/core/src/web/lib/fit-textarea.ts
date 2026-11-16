/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text boxes as tall as their words, measured without collapsing them.
 * Shared by the AI composers (components/console/ai/composerAssist.ts) and
 * the app's other growing text boxes.
 */

/** How tall ``value`` is in a box like ``el`` but ``width`` px wide: a hidden
 * twin with the box's own font, padding and wrapping, so the answer does not
 * depend on the layout the box is in right now. */
export function heightAtWidth(el: HTMLTextAreaElement, width: number, value: string): number {
  const doc = el.ownerDocument;
  let twin = (el as HTMLTextAreaElement & { __rowTwin?: HTMLTextAreaElement }).__rowTwin;
  if (!twin || !twin.isConnected) {
    twin = doc.createElement("textarea");
    twin.setAttribute("aria-hidden", "true");
    twin.tabIndex = -1;
    twin.rows = 1;
    (el as HTMLTextAreaElement & { __rowTwin?: HTMLTextAreaElement }).__rowTwin = twin;
    doc.body.appendChild(twin);
  }
  const style = doc.defaultView?.getComputedStyle(el);
  if (style) {
    for (const name of [
      "font", "letterSpacing", "wordSpacing", "lineHeight", "textTransform", "textIndent",
      "paddingLeft", "paddingRight", "paddingTop", "paddingBottom", "borderLeftWidth",
      "borderRightWidth", "borderTopWidth", "borderBottomWidth", "boxSizing", "whiteSpace",
      "overflowWrap", "wordBreak", "tabSize",
    ] as const) {
      (twin.style as unknown as Record<string, string>)[name] = style[name];
    }
  }
  Object.assign(twin.style, {
    position: "absolute", top: "0", left: "-10000px", visibility: "hidden", pointerEvents: "none",
    height: "auto", minHeight: "0", maxHeight: "none", overflow: "hidden", width: `${width}px`,
  });
  twin.value = value;
  return twin.scrollHeight;
}











export function fitTextareaToText(
  el: HTMLTextAreaElement,
  { min = 0, max }: { min?: number; max?: number } = {},
): void {
  if (el.clientWidth === 0) return;
  const style = el.ownerDocument.defaultView?.getComputedStyle(el);
  const cap = max ?? (style ? Number.parseFloat(style.maxHeight) : Number.NaN);
  const limit = Number.isFinite(cap) && cap > 0 ? cap : Number.POSITIVE_INFINITY;
  const needed = Math.max(min, heightAtWidth(el, el.clientWidth, el.value));
  const twinHolder = el as HTMLTextAreaElement & { __rowTwin?: HTMLTextAreaElement };
  twinHolder.__rowTwin?.remove();
  delete twinHolder.__rowTwin;
  const next = `${Math.min(needed, limit)}px`;
  if (el.style.height !== next) el.style.height = next;
  const overflow = needed > limit ? "auto" : "hidden";
  if (el.style.overflowY !== overflow) el.style.overflowY = overflow;
}

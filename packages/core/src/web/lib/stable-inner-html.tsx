"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Markup injected with `dangerouslySetInnerHTML`, kept on the page while it is
 * unchanged.
 *
 * React 19 compares the `{ __html }` wrapper by identity, not its string: an
 * inline `dangerouslySetInnerHTML={{ __html: markup }}` hands it a new object
 * on every render, and the element's whole markup is parsed again each time.
 * A deck chart's picture (a full SVG) was parsed five times per click on the
 * chart, eighteen times across the click-click-double-click that opens its
 * title, and every node it held (a title hidden under its editor, a hovered
 * state, a selection mark set on the picture) was thrown away with it.
 */
import { useMemo, type HTMLAttributes, type ReactElement } from "react";

/** The wrapper for `markup`, the same object for as long as the string is. */
export function useInnerHtml(markup: string): { __html: string };
export function useInnerHtml(markup: string | null | undefined): { __html: string } | undefined;
export function useInnerHtml(markup: string | null | undefined): { __html: string } | undefined {
  return useMemo(() => (markup == null ? undefined : { __html: markup }), [markup]);
}

/** A `div` showing `html`, parsed again only when the string changes. */
export function StableHtmlDiv({
  html,
  ...props
}: { html: string } & Omit<HTMLAttributes<HTMLDivElement>, "children" | "dangerouslySetInnerHTML">): ReactElement {
  const inner = useInnerHtml(html);
  return <div {...props} dangerouslySetInnerHTML={inner} />;
}

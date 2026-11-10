"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { placeScrollTop, scrollInMotion } from "../../../lib/scroll-follow";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { JSONContent } from "@tiptap/core";
import { EditorContent, useEditor, type Editor } from "@tiptap/react";
import { common, createLowlight } from "lowlight";
import { useLocale, useTranslations } from "../../../../host/next-intl";

import {
  OfficeTaskPane,
  officeTaskPaneParts,
} from "../../office-ribbon/OfficeTaskPane";
import {
  sourceDocumentPanes,
  syncedScrollTop,
  type CompareResultSources,
  type ScrollAnchor,
  type ScrollView,
  type SourceDocumentRole,
  type SourceDocumentsMode,
} from "../../../lib/document/sourceDocuments";
import { createNativeDocumentPluginRegistry } from "./native-document/pluginRegistry";
import styles from "./SourceDocumentsPane.module.css";

/** The page width the source panes scale to fit (Letter/A4 at 96 dpi). */
const SOURCE_PAGE_WIDTH_PX = 816;

/** The scroll container of the result: the nearest scrolling ancestor. */
function scrollContainerOf(element: HTMLElement | null): HTMLElement | null {
  let node = element?.parentElement ?? null;
  while (node) {
    const overflow = window.getComputedStyle(node).overflowY;
    if (
      (overflow === "auto" || overflow === "scroll") &&
      node.scrollHeight > node.clientHeight + 1
    ) {
      return node;
    }
    node = node.parentElement;
  }
  return (document.scrollingElement as HTMLElement | null) ?? null;
}

/**
 * Every block's box in `scroller`'s scroll coordinates. A transformed editor
 * window draws its boxes scaled; dividing by the scroller's drawn/laid-out
 * ratio puts them back in scrollTop's units.
 */
function scrollView(scroller: HTMLElement): ScrollView {
  const frame = scroller.getBoundingClientRect();
  const scale =
    scroller.offsetHeight > 0 && frame.height > 0
      ? frame.height / scroller.offsetHeight
      : 1;
  const anchors: ScrollAnchor[] = [];
  scroller.querySelectorAll<HTMLElement>("[data-block-id]").forEach((element) => {
    const id = element.getAttribute("data-block-id");
    if (!id) return;
    const box = element.getBoundingClientRect();
    anchors.push({
      id,
      top: (box.top - frame.top) / scale + scroller.scrollTop,
      height: box.height / scale,
    });
  });
  return {
    anchors,
    scrollTop: scroller.scrollTop,
    scrollHeight: scroller.scrollHeight,
    clientHeight: scroller.clientHeight,
  };
}

/** One source document, read-only, laid out as its page and scaled to fit. */
function SourceDocumentCanvas({
  doc,
  label,
  role,
  scrollerRef,
}: {
  readonly doc: JSONContent;
  readonly label: string;
  readonly role: SourceDocumentRole;
  readonly scrollerRef: (element: HTMLDivElement | null) => void;
}) {
  const t = useTranslations("console-notebooks");
  const locale = useLocale();
  const extensions = useMemo(
    () =>
      createNativeDocumentPluginRegistry({
        lowlight: createLowlight(common),
        pageId: `source-document-${role}`,
        placeholder: "",
        translate: (key, values) => t(key, values),
        notebookId: null,
        guestMode: true,
      }).extensions,
    // Rebuild only when the locale changes (useTranslations changes identity
    // on ordinary renders), as the structured preview does.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [locale, role],
  );
  const content = useMemo<JSONContent>(
    () => ({
      ...doc,
      content: [{ type: "pageBlock", content: doc.content ?? [] }],
    }),
    [doc],
  );
  const editor = useEditor(
    {
      immediatelyRender: false,
      editable: false,
      extensions,
      content,
      editorProps: {
        attributes: {
          class: `note-editor-content ${styles.editor}`,
          "aria-label": label,
          "aria-readonly": "true",
        },
      },
    },
    [extensions, content],
  );
  const localRef = useRef<HTMLDivElement | null>(null);
  // One ref function for the canvas's life: a new one each render would
  // detach and re-attach the scroller, re-registering it on every render.
  const setScroller = useCallback(
    (element: HTMLDivElement | null) => {
      localRef.current = element;
      scrollerRef(element);
    },
    [scrollerRef],
  );
  const [zoom, setZoom] = useState(0.4);
  useEffect(() => {
    const element = localRef.current;
    if (!element || typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(() => {
      const width = element.clientWidth;
      if (width > 0) {
        setZoom(Math.min(1, Math.max(0.2, (width - 16) / SOURCE_PAGE_WIDTH_PX)));
      }
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, []);
  return (
    <section
      className={styles.source}
      data-testid={`document-source-${role}`}
      aria-label={label}
    >
      <div className={styles.label}>{label}</div>
      <div
        ref={setScroller}
        className={styles.scroller}
        data-testid={`document-source-${role}-scroller`}
      >
        <div className={styles.canvas} style={{ zoom }}>
          <EditorContent editor={editor} />
        </div>
      </div>
    </section>
  );
}

/**
 * Word's source documents beside a compare/combine result: the original
 * above the revised, read-only, scrolling with the result (and the result
 * with them), as Word's three-pane review window does. Closing the pane is
 * Hide Source Documents.
 */
export default function SourceDocumentsPane({
  editor,
  sources,
  mode,
  onModeChange,
  dockClassName,
}: {
  readonly editor: Editor;
  readonly sources: CompareResultSources;
  readonly mode: SourceDocumentsMode;
  readonly onModeChange: (mode: SourceDocumentsMode) => void;
  readonly dockClassName?: string;
}) {
  const t = useTranslations("console-notebooks");
  const panes = sourceDocumentPanes(mode);
  const scrollersRef = useRef(new Map<SourceDocumentRole, HTMLDivElement>());
  const [scrollerGeneration, setScrollerGeneration] = useState(0);
  const scrollerRefFor = useCallback(
    (role: SourceDocumentRole) => (element: HTMLDivElement | null) => {
      const current = scrollersRef.current.get(role) ?? null;
      if (current === element) return;
      if (element) scrollersRef.current.set(role, element);
      else scrollersRef.current.delete(role);
      setScrollerGeneration((value) => value + 1);
    },
    [],
  );
  const originalRef = useMemo(() => scrollerRefFor("original"), [scrollerRefFor]);
  const revisedRef = useMemo(() => scrollerRefFor("revised"), [scrollerRefFor]);

  // Synchronized scrolling: whichever view the person scrolls drives the
  // others. A position this code set is recognised when its own scroll event
  // arrives and does not drive back (no feedback loop).
  useEffect(() => {
    if (editor.isDestroyed) return;
    const result = scrollContainerOf(editor.view.dom as HTMLElement);
    if (!result) return;
    const views = [result, ...scrollersRef.current.values()];
    if (views.length < 2) return;
    const expected = new Map<HTMLElement, number>();
    let frame: number | null = null;
    let driver: HTMLElement | null = null;
    const align = () => {
      frame = null;
      const from = driver;
      driver = null;
      if (!from || !from.isConnected) return;
      const fromView = scrollView(from);
      for (const target of views) {
        if (target === from || !target.isConnected) continue;
        const top = Math.round(syncedScrollTop(fromView, scrollView(target)));
        if (Math.abs(target.scrollTop - top) < 1) continue;
        // A pane the reader is scrolling keeps their scroll: written from
        // script mid-swipe, it moved from two threads at once and Safari's
        // screen and its hit testing parted (lib/wheel-gesture).
        if (scrollInMotion(target)) continue;
        expected.set(target, top);
        placeScrollTop(target, top);
      }
    };
    const listeners = views.map((view) => {
      const listener = () => {
        const wanted = expected.get(view);
        if (wanted !== undefined) {
          expected.delete(view);
          if (Math.abs(view.scrollTop - wanted) < 2) return;
        }
        driver = view;
        if (frame === null) frame = window.requestAnimationFrame(align);
      };
      view.addEventListener("scroll", listener, { passive: true });
      return { view, listener };
    });
    // Line the panes up with where the result already is.
    driver = result;
    frame = window.requestAnimationFrame(align);
    return () => {
      for (const { view, listener } of listeners) {
        view.removeEventListener("scroll", listener);
      }
      if (frame !== null) window.cancelAnimationFrame(frame);
    };
  }, [editor, scrollerGeneration, mode]);

  const labelFor = (role: SourceDocumentRole) =>
    role === "original"
      ? t("compare.sourceOriginalLabel", { name: sources.original.label })
      : t("compare.sourceRevisedLabel", { name: sources.revised.label });

  return (
    <OfficeTaskPane
      title={t("toolbar.showSourceDocuments")}
      closeLabel={t("toolbar.hideSourceDocuments")}
      onClose={() => onModeChange("none")}
      testIdPrefix="document-source-documents"
      testId="document-source-documents-pane"
      closeTestId="document-source-documents-close"
      dockRight={false}
      className={dockClassName}
      dataAttributes={{ "data-source-documents-mode": mode }}
    >
      <div className={`${officeTaskPaneParts.body} ${styles.body}`}>
        {panes.map((role) => (
          <SourceDocumentCanvas
            key={`${sources.id}-${role}`}
            role={role}
            label={labelFor(role)}
            doc={sources[role].doc as JSONContent}
            scrollerRef={role === "original" ? originalRef : revisedRef}
          />
        ))}
      </div>
    </OfficeTaskPane>
  );
}

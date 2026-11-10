/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The playground: open a file from this computer, edit it in its editor, and
// save it back as a download. No account and no server: everything stays in
// this browser tab.
import { Suspense, lazy, useCallback, useEffect, useMemo, useRef, useState, type ComponentType } from "react";
import { DESKTOP_DOWNLOAD_URL, EDITOR_LOCALES, EditorHostProvider, WEB_APP_URL, browserDownloadHost, noteDesktopFeature, type EditorDefinition, type EditorProps, type LocalFile, type SavedFile } from "@mrnote/core";

import { useToast } from "@mrnote/core/web/hooks/use-toast.ts";

import { EDITOR_LOADERS, type EditorId } from "./editors.generated";

interface OpenTab {
  readonly file: LocalFile;
  readonly editor: EditorId;
  dirty: boolean;
}

const LOCALE_KEY = "mrnote-editors:locale";
function initialLocale(): string {
  try {
    const saved = localStorage.getItem(LOCALE_KEY);
    if (saved && EDITOR_LOCALES.includes(saved)) return saved;
  } catch {
    // storage unavailable
  }
  const browser = navigator.language;
  if (EDITOR_LOCALES.includes(browser)) return browser;
  const base = browser.split("-")[0];
  return EDITOR_LOCALES.includes(base) ? base : "en";
}

const BANNER_KEY = "mrnote-editors:desktop-banner";
/** Office's and PDF's own formats the desktop app opens (the editors here do not). */
const DESKTOP_FORMATS = /\.(docx?|dotx|xlsx?|xlsm|xltx|pptx?|potx|ppsx)$/i;

const definitions = new Map<EditorId, Promise<EditorDefinition>>();
function loadDefinition(id: EditorId): Promise<EditorDefinition> {
  let definition = definitions.get(id);
  if (!definition) {
    definition = EDITOR_LOADERS[id].load();
    definitions.set(id, definition);
  }
  return definition;
}

const components = new Map<EditorId, ComponentType<EditorProps>>();
function editorComponent(id: EditorId): ComponentType<EditorProps> {
  let component = components.get(id);
  if (!component) {
    component = lazy(async () => ({ default: (await loadDefinition(id)).Component }));
    components.set(id, component);
  }
  return component;
}

/** Every editor that opens a file of this name, the one it opens in first:
 * a file several editors take (.html, .md, .txt) opens in the editor of its
 * kind before the code editor, which takes every text file. */
function editorsFor(name: string): EditorId[] {
  const lower = name.toLowerCase();
  const all = (Object.entries(EDITOR_LOADERS) as [EditorId, (typeof EDITOR_LOADERS)[EditorId]][])
    .filter(([, entry]) => entry.extensions.some((extension) => lower.endsWith(extension)))
    .map(([id]) => id);
  return [...all.filter((id) => id !== "code"), ...all.filter((id) => id === "code")];
}

let sequence = 0;
const nextId = () => `${Date.now().toString(36)}-${(sequence += 1)}`;

export function App() {
  const [locale, setLocale] = useState(initialLocale);
  const [tabs, setTabs] = useState<OpenTab[]>([]);
  const [active, setActive] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [alternative, setAlternative] = useState<{ file: LocalFile; editor: EditorId } | null>(null);
  type Link = { label: string; href: string };
  const [noticeLink, setNoticeLink] = useState<{ action: Link; secondary?: Link } | null>(null);
  const desktopNoticeAt = useRef(0);
  const [bannerShown, setBannerShown] = useState(() => {
    try {
      return localStorage.getItem(BANNER_KEY) !== "hidden";
    } catch {
      return true;
    }
  });
  // The editors' messages come here; one about a feature of the MR Note
  // desktop app carries where to get it.
  const host = useMemo(
    () => ({
      ...browserDownloadHost,
      notify({ title, description, action, secondary }: { title?: string; description?: string; action?: Link; secondary?: Link }) {
        if (action) desktopNoticeAt.current = Date.now();
        setAlternative(null);
        setNoticeLink(action ? { action, secondary } : null);
        setNotice([title, description].filter(Boolean).join(" — "));
      },
    }),
    [],
  );

  // The editors' own toasts (a failed insert, a finished export, why a
  // command did not run): MR Note draws them in its app shell; here they go
  // to the notice bar.
  const { toasts } = useToast();
  const shownToast = useRef<string | null>(null);
  useEffect(() => {
    const latest = toasts.find((toast) => toast.open !== false);
    if (!latest || shownToast.current === latest.id) return;
    shownToast.current = latest.id;
    const text = [latest.title, latest.description].filter((part) => typeof part === "string" && part.trim()).join(" — ");
    if (!text) return;
    // the failure that follows "this is in the desktop app" says nothing more
    if (latest.variant === "destructive" && Date.now() - desktopNoticeAt.current < 3000) return;
    setAlternative(null);
    setNoticeLink(null);
    setNotice(text);
  }, [toasts]);

  // Ctrl/Cmd+F opens the open editor's own search where it has one (in MR
  // Note the window's title bar does)
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (!(event.key === "f" || event.key === "F") || !(event.metaKey || event.ctrlKey) || event.altKey || event.shiftKey) return;
      const button = document.querySelector<HTMLElement>('.pg-editor:not([hidden]) [data-testid$="command-search-open"]');
      if (!button) return;
      event.preventDefault();
      event.stopPropagation();
      button.click();
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, []);
  const savers = useRef(new Map<string, () => Promise<SavedFile> | SavedFile>());
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    try {
      localStorage.setItem(LOCALE_KEY, locale);
    } catch {
      // storage unavailable
    }
    document.documentElement.lang = locale;
  }, [locale]);

  const openFile = useCallback((file: LocalFile, chosen?: EditorId) => {
    const editors = editorsFor(file.name);
    const editor = chosen ?? editors[0];
    if (!editor) {
      if (DESKTOP_FORMATS.test(file.name)) noteDesktopFeature(`Opening ${file.name}`);
      else setNotice(`No editor here opens ${file.name}.`);
      return;
    }
    setTabs((current) => [...current, { file, editor, dirty: false }]);
    setActive(file.id);
    const other = editors.find((id) => id !== editor);
    setAlternative(other ? { file: { ...file, id: nextId() }, editor: other } : null);
    setNoticeLink(null);
    setNotice(other ? `Opened ${file.name} in the ${EDITOR_LOADERS[editor].label} editor.` : null);
  }, []);

  const openFromDisk = useCallback(
    async (list: FileList | null) => {
      for (const picked of Array.from(list ?? [])) {
        openFile({ id: nextId(), name: picked.name, bytes: new Uint8Array(await picked.arrayBuffer()) });
      }
    },
    [openFile],
  );

  const newFile = useCallback(
    async (id: EditorId) => {
      const definition = await loadDefinition(id);
      const created = definition.newFile?.();
      if (created) openFile({ id: nextId(), ...created });
    },
    [openFile],
  );

  const save = useCallback(async (tabId: string | null) => {
    if (!tabId) return;
    const saver = savers.current.get(tabId);
    if (!saver) {
      setNotice("This editor has nothing to save yet.");
      return;
    }
    try {
      const saved = await saver();
      const data = saved.bytes instanceof Blob ? saved.bytes : new Blob([saved.bytes as BlobPart], { type: saved.type });
      await browserDownloadHost.saveFile({ name: saved.name, type: saved.type, data });
      setTabs((current) => current.map((tab) => (tab.file.id === tabId ? { ...tab, dirty: false } : tab)));
      setAlternative(null);
      setNoticeLink(null);
      setNotice(`Saved ${saved.name} to your downloads.`);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    }
  }, []);

  const close = useCallback((tabId: string) => {
    savers.current.delete(tabId);
    setTabs((current) => {
      const next = current.filter((tab) => tab.file.id !== tabId);
      setActive((was) => (was === tabId ? next[next.length - 1]?.file.id ?? null : was));
      return next;
    });
  }, []);

  const accept = useMemo(() => Object.values(EDITOR_LOADERS).flatMap((entry) => entry.extensions).join(","), []);

  return (
    <EditorHostProvider host={host} locale={locale} fallback={<div className="pg-loading">Loading…</div>}>
      <div className="pg-shell">
        <header className="pg-bar">
          <a className="pg-brand" href={WEB_APP_URL} target="_blank" rel="noreferrer">
            MR Office <span>Editors</span>
          </a>
          <div className="pg-actions">
            {(Object.keys(EDITOR_LOADERS) as EditorId[]).map((id) =>
              EDITOR_LOADERS[id].canCreate ? (
                <button key={id} type="button" onClick={() => void newFile(id)}>
                  New {EDITOR_LOADERS[id].label}
                </button>
              ) : null,
            )}
            <button type="button" onClick={() => input.current?.click()}>
              Open…
            </button>
            <button type="button" disabled={!active} onClick={() => void save(active)}>
              Save
            </button>
            <input ref={input} type="file" multiple hidden accept={accept} onChange={(event) => void openFromDisk(event.target.files).finally(() => (event.target.value = ""))} />
            <select aria-label="Language" value={locale} onChange={(event) => setLocale(event.target.value)}>
              {EDITOR_LOCALES.map((code) => (
                <option key={code} value={code}>
                  {code}
                </option>
              ))}
            </select>
          </div>
        </header>
        {bannerShown ? (
          <div className="pg-banner" role="note">
            <span>
              Use the <strong>MR Office desktop app</strong> for the full service: Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt and saving PDF edits.
            </span>
            <a href={DESKTOP_DOWNLOAD_URL} target="_blank" rel="noreferrer">
              Download MR Office (free)
            </a>
            <a className="pg-banner-web" href={WEB_APP_URL} target="_blank" rel="noreferrer">
              mroffice.ai
            </a>
            <button
              type="button"
              aria-label="Hide this message"
              onClick={() => {
                setBannerShown(false);
                try {
                  localStorage.setItem(BANNER_KEY, "hidden");
                } catch {
                  // storage unavailable
                }
              }}
            >
              ×
            </button>
          </div>
        ) : null}
        {tabs.length > 0 ? (
          <nav className="pg-tabs" role="tablist">
            {tabs.map((tab) => (
              <div key={tab.file.id} className="pg-tab" data-active={tab.file.id === active || undefined}>
                <button type="button" role="tab" aria-selected={tab.file.id === active} onClick={() => setActive(tab.file.id)}>
                  {tab.file.name}
                  {tab.dirty ? " •" : ""}
                </button>
                <button type="button" aria-label={`Close ${tab.file.name}`} onClick={() => close(tab.file.id)}>
                  ×
                </button>
              </div>
            ))}
          </nav>
        ) : null}
        {notice ? (
          <div className="pg-notice" role="status">
            {notice}
            {noticeLink ? (
              <>
                <a className="pg-alt pg-alt-primary" href={noticeLink.action.href} target="_blank" rel="noreferrer">
                  {noticeLink.action.label}
                </a>
                {noticeLink.secondary ? (
                  <a className="pg-alt-secondary" href={noticeLink.secondary.href} target="_blank" rel="noreferrer">
                    {noticeLink.secondary.label}
                  </a>
                ) : null}
              </>
            ) : null}
            {alternative ? (
              <button type="button" className="pg-alt" onClick={() => { const next = alternative; setAlternative(null); openFile(next.file, next.editor); }}>
                Open in the {EDITOR_LOADERS[alternative.editor].label} editor instead
              </button>
            ) : null}
            <button type="button" aria-label="Dismiss" onClick={() => { setNotice(null); setAlternative(null); setNoticeLink(null); }}>
              ×
            </button>
          </div>
        ) : null}
        <main
          className="pg-stage"
          onDragOver={(event) => event.preventDefault()}
          onDrop={(event) => {
            event.preventDefault();
            void openFromDisk(event.dataTransfer.files);
          }}
        >
          {tabs.length === 0 ? (
            <div className="pg-empty">
              <h1>Open a file, or start a new one</h1>
              <p>Everything stays in this browser tab. Save downloads the file back to your computer.</p>
              <p className="pg-kinds">
                {Object.values(EDITOR_LOADERS)
                  .map((entry) => `${entry.label}: ${entry.extensions.slice(0, 8).join(" ")}${entry.extensions.length > 8 ? " …" : ""}`)
                  .join("  ·  ")}
              </p>
            </div>
          ) : null}
          {tabs.map((tab) => {
            const Editor = editorComponent(tab.editor);
            return (
              <section key={tab.file.id} className="pg-editor" hidden={tab.file.id !== active}>
                <Suspense fallback={<div className="pg-loading">Loading the {EDITOR_LOADERS[tab.editor].label} editor…</div>}>
                  <Editor
                    file={tab.file}
                    registerSave={(saver) => {
                      savers.current.set(tab.file.id, saver);
                      return () => {
                        if (savers.current.get(tab.file.id) === saver) savers.current.delete(tab.file.id);
                      };
                    }}
                    onSaveRequest={() => void save(tab.file.id)}
                    onDirty={() => setTabs((current) => current.map((entry) => (entry.file.id === tab.file.id && !entry.dirty ? { ...entry, dirty: true } : entry)))}
                  />
                </Suspense>
              </section>
            );
          })}
        </main>
        <footer className="pg-foot">
          <a href={WEB_APP_URL}>MR Office</a> open-source editors · <a href={DESKTOP_DOWNLOAD_URL}>Get the free desktop app</a> · AGPL-3.0 with the terms in NOTICE.md
        </footer>
      </div>
    </EditorHostProvider>
  );
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The desktop app's parts fetched after installing (apps/desktop
 * src/classic/ondemand-assets.ts): the code pages' Python ("pyodide") and the
 * ink readers ("ink-models"). The app fetches them in the background after
 * its first launch; a feature that needs one before then asks for it here and
 * waits, while the download's progress shows on the transfer line
 * (components/ui/TransferProgressLine, lib/transfer-store) and in the Dock.
 * Outside the desktop app, and in a package that still carries them, every
 * call answers at once.
 */
import { IS_DESKTOP_CLASSIC } from "./desktop-classic";
import { beginTransfer } from "./transfer-store";

export type DesktopOndemandGroup = "pyodide" | "ink-models";

export interface DesktopOndemandStatus {
  name: string;
  ready: boolean;
  bytes: number;
  have: number;
  downloading: "foreground" | "background" | null;
  error: string | null;
}

interface OndemandBridge {
  status: () => Promise<DesktopOndemandStatus[] | null>;
  ensure: (group: string) => Promise<{ ok: boolean; error?: string }>;
  onChange: (listener: () => void) => () => void;
}

export function desktopOndemand(): OndemandBridge | null {
  if (!IS_DESKTOP_CLASSIC || typeof window === "undefined") return null;
  const bridge = (window as unknown as { mraiDesktop?: { classic?: boolean; ondemand?: OndemandBridge } }).mraiDesktop;
  return bridge?.classic && bridge.ondemand?.ensure ? bridge.ondemand : null;
}

const pending = new Map<DesktopOndemandGroup, Promise<void>>();

/**
 * Resolves once the group is on this computer (at once when it already is,
 * or outside the desktop app); rejects with the reason when it cannot be
 * fetched now (offline before it ever was).
 */
export function ensureDesktopOndemand(group: DesktopOndemandGroup): Promise<void> {
  const bridge = desktopOndemand();
  if (!bridge) return Promise.resolve();
  const running = pending.get(group);
  if (running) return running;
  const task = (async () => {
    const statuses = await bridge.status().catch(() => null);
    const status = statuses?.find((item) => item.name === group);
    if (!status || status.ready) return;
    // Its progress on the transfer line while someone waits for it.
    const transfer = beginTransfer({ direction: "download", name: ondemandName(group), totalBytes: status.bytes });
    transfer.update(status.have);
    const stop = bridge.onChange(() => {
      void bridge.status().then((now) => {
        const current = now?.find((item) => item.name === group);
        if (current) transfer.update(current.have);
      }, () => undefined);
    });
    try {
      const outcome = await bridge.ensure(group);
      if (!outcome?.ok) throw new Error(outcome?.error || "download failed");
    } finally {
      stop();
      transfer.end();
    }
  })().finally(() => pending.delete(group));
  pending.set(group, task);
  return task;
}

/** What the transfer line calls it, in the page's language. */
function ondemandName(group: DesktopOndemandGroup): string {
  const zh = typeof document !== "undefined" && /^zh/i.test(document.documentElement.lang || navigator.language || "");
  if (group === "pyodide") return zh ? "Python 运行环境" : "Python";
  return zh ? "手写识别" : "Handwriting recognition";
}

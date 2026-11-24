/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The parts of the app fetched after installing (scripts/classic-ondemand.mjs):
 * the code pages' Python (Pyodide) and the ink readers' models. The package
 * lists them in mrnote/ondemand-assets.json (path, size, SHA-256 per file);
 * they are kept at <data>/ondemand/<group>/<path>, where the web server's
 * Python sandbox, the window checker and the API read them (services.ts).
 *
 * - After a launch the app fetches what is missing in the background: once it
 *   has run a while and the computer is online, one file at a time, at a
 *   modest rate, from our download host (the faster of its two addresses).
 * - A page that needs a group before then asks for it (`ensure`): the
 *   download goes on at full speed and its progress is told to the windows.
 * - Every file resumes where it stopped (HTTP Range) and is kept only with the
 *   size and SHA-256 the package names (libreoffice-installer `download`);
 *   what fails is tried again later, and on the next launch.
 * - Files an earlier version of the app fetched and this one does not name
 *   are removed once its groups are complete.
 *
 * No Electron here, so tests run it with plain Node.
 */

import fs from "node:fs";
import path from "node:path";

import { download, type DownloadOptions, InstallError } from "./libreoffice-installer";

export type OndemandGroupName = "pyodide" | "ink-models";

export interface OndemandFile {
  readonly path: string;
  readonly size: number;
  readonly sha256: string;
}

export interface OndemandManifest {
  readonly version: number;
  readonly hosts: readonly string[];
  readonly groups: Readonly<Record<string, { readonly bytes: number; readonly files: readonly OndemandFile[] }>>;
}

export interface OndemandGroupStatus {
  readonly name: string;
  readonly ready: boolean;
  readonly bytes: number;
  readonly have: number;
  /** A download of this group is running, in the foreground (someone waits) or not. */
  readonly downloading: "foreground" | "background" | null;
  readonly error: string | null;
}

const SAFE_PATH = /^(?!.*(?:^|\/)\.\.?(?:\/|$))[A-Za-z0-9._+@=-]+(?:\/[A-Za-z0-9._+@=-]+)*$/;

export function readManifest(file: string): OndemandManifest | null {
  let raw: unknown;
  try {
    raw = JSON.parse(fs.readFileSync(file, "utf8"));
  } catch {
    return null;
  }
  const manifest = raw as OndemandManifest;
  if (manifest?.version !== 1 || !Array.isArray(manifest.hosts) || typeof manifest.groups !== "object") return null;
  for (const host of manifest.hosts) if (typeof host !== "string" || !host.startsWith("https://") || !host.endsWith("/")) return null;
  for (const group of Object.values(manifest.groups)) {
    if (!Array.isArray(group?.files)) return null;
    for (const entry of group.files) {
      if (typeof entry?.path !== "string" || !SAFE_PATH.test(entry.path)) return null;
      if (typeof entry.size !== "number" || entry.size < 0 || typeof entry.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(entry.sha256)) return null;
    }
  }
  return manifest;
}

/** Bytes per second for a background download; a foreground one is not limited. */
export const BACKGROUND_RATE = 2 * 1024 * 1024;

/** A fetch whose body is read no faster than `rate()` bytes per second (0: no limit), checked as it goes. */
export function throttledFetch(rate: () => number, base: typeof fetch = fetch): typeof fetch {
  return (async (input: Parameters<typeof fetch>[0], init?: Parameters<typeof fetch>[1]) => {
    const response = await base(input, init);
    if (!response.body) return response;
    const reader = response.body.getReader();
    let windowStart = Date.now();
    let windowBytes = 0;
    const body = new ReadableStream<Uint8Array>({
      async pull(controller) {
        const { done, value } = await reader.read();
        if (done) {
          controller.close();
          return;
        }
        controller.enqueue(value);
        const limit = rate();
        if (limit > 0) {
          windowBytes += value.length;
          const due = windowStart + (windowBytes / limit) * 1000;
          const wait = due - Date.now();
          if (wait > 0) await new Promise((resolve) => setTimeout(resolve, Math.min(wait, 2000)));
          if (Date.now() - windowStart > 10_000) {
            windowStart = Date.now();
            windowBytes = 0;
          }
        }
      },
      cancel(reason) {
        return reader.cancel(reason);
      },
    });
    return new Response(body, { status: response.status, statusText: response.statusText, headers: response.headers });
  }) as typeof fetch;
}

export interface OndemandOptions {
  readonly manifestFile: string;
  readonly root: string;
  readonly fetch?: typeof fetch;
  readonly log?: (line: string) => void;
  /** Told whenever a group's state changes or a download moves on. */
  readonly onChange?: () => void;
}

export class OndemandAssets {
  readonly manifest: OndemandManifest | null;
  private readonly root: string;
  private readonly baseFetch: typeof fetch;
  private readonly log: (line: string) => void;
  private readonly onChange: () => void;
  private hosts: string[] | null = null;
  private readonly running = new Map<string, Promise<void>>();
  private readonly foreground = new Set<string>();
  private readonly progress = new Map<string, number>();
  private readonly errors = new Map<string, string>();
  private readonly readyCache = new Map<string, boolean>();

  constructor(options: OndemandOptions) {
    this.manifest = readManifest(options.manifestFile);
    this.root = options.root;
    this.baseFetch = options.fetch ?? fetch;
    this.log = options.log ?? (() => undefined);
    this.onChange = options.onChange ?? (() => undefined);
  }

  /** Where a group's files are kept (they may not all be there yet). */
  dir(group: string): string {
    return path.join(this.root, group);
  }

  has(group: string): boolean {
    return Boolean(this.manifest?.groups[group]);
  }

  groups(): string[] {
    return Object.keys(this.manifest?.groups ?? {});
  }

  private fileAt(group: string, file: OndemandFile): string {
    return path.join(this.dir(group), ...file.path.split("/"));
  }

  private fileHere(group: string, file: OndemandFile): boolean {
    try {
      return fs.statSync(this.fileAt(group, file)).size === file.size;
    } catch {
      return false;
    }
  }

  /** Every file at its listed size (each was checked by SHA-256 when it came). */
  ready(group: string): boolean {
    const listed = this.manifest?.groups[group];
    if (!listed) return true;
    if (this.readyCache.get(group)) return true;
    const ok = listed.files.every((file) => this.fileHere(group, file));
    if (ok) this.readyCache.set(group, true);
    return ok;
  }

  status(group: string): OndemandGroupStatus {
    const listed = this.manifest?.groups[group];
    const bytes = listed?.bytes ?? 0;
    const ready = this.ready(group);
    let have = ready ? bytes : 0;
    if (!ready && listed) {
      for (const file of listed.files) if (this.fileHere(group, file)) have += file.size;
      have += this.progress.get(group) ?? 0;
    }
    return {
      name: group,
      ready,
      bytes,
      have: Math.min(have, bytes),
      downloading: this.running.has(group) ? (this.foreground.has(group) ? "foreground" : "background") : null,
      error: this.errors.get(group) ?? null,
    };
  }

  /** The download host's two addresses, the one that answers sooner first. */
  private async orderedHosts(): Promise<string[]> {
    if (this.hosts) return this.hosts;
    const hosts = [...(this.manifest?.hosts ?? [])];
    const probe = this.manifest ? Object.values(this.manifest.groups)[0]?.files[0] : undefined;
    if (hosts.length < 2 || !probe) return (this.hosts = hosts);
    const timed = await Promise.all(
      hosts.map(async (host, index) => {
        const started = Date.now();
        try {
          const response = await this.baseFetch(`${host}${probe.sha256}`, { method: "HEAD", signal: AbortSignal.timeout(5000) });
          return { host, ms: response.ok ? Date.now() - started : Infinity, index };
        } catch {
          return { host, ms: Infinity, index };
        }
      }),
    );
    timed.sort((a, b) => a.ms - b.ms || a.index - b.index);
    this.log(`on-demand hosts: ${timed.map((item) => `${item.host} ${Number.isFinite(item.ms) ? `${item.ms} ms` : "no answer"}`).join(", ")}`);
    // Only remembered when one answered: offline now is not offline later.
    if (timed.some((item) => Number.isFinite(item.ms))) this.hosts = timed.map((item) => item.host);
    return timed.map((item) => item.host);
  }

  /**
   * Fetch a group's missing files. `foreground`: someone waits for it (full
   * speed; a background download of it already running speeds up). Resolves
   * when the group is ready; rejects when a file cannot be had now.
   */
  ensure(group: string, foreground: boolean): Promise<void> {
    if (!this.has(group) || this.ready(group)) return Promise.resolve();
    if (foreground && !this.foreground.has(group)) {
      this.foreground.add(group);
      this.onChange();
    }
    const running = this.running.get(group);
    if (running) return running;
    const task = this.fetchGroup(group).finally(() => {
      this.running.delete(group);
      this.foreground.delete(group);
      this.progress.delete(group);
      this.onChange();
    });
    this.running.set(group, task);
    this.onChange();
    return task;
  }

  private async fetchGroup(group: string): Promise<void> {
    const listed = this.manifest!.groups[group];
    const parts = path.join(this.root, ".parts");
    const rate = () => (this.foreground.has(group) ? 0 : BACKGROUND_RATE);
    const doFetch = throttledFetch(rate, this.baseFetch);
    this.errors.delete(group);
    try {
      const hosts = await this.orderedHosts();
      for (const file of listed.files) {
        if (this.fileHere(group, file)) continue;
        const options: DownloadOptions = {
          fetch: doFetch,
          log: this.log,
          onProgress: (_phase, received) => {
            this.progress.set(group, received);
            this.onChange();
          },
        };
        const done = await download(
          { file: file.sha256, size: file.size, sha256: file.sha256, urls: hosts.map((host) => `${host}${file.sha256}`) },
          { downloads: parts },
          options,
        );
        const at = this.fileAt(group, file);
        fs.mkdirSync(path.dirname(at), { recursive: true });
        fs.renameSync(done, at);
        this.progress.delete(group);
        this.onChange();
      }
      this.readyCache.delete(group);
      if (!this.ready(group)) throw new InstallError("download", `${group} is not complete`);
      this.log(`on-demand ${group}: ready (${listed.files.length} files)`);
      this.prune();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.errors.set(group, message);
      this.log(`on-demand ${group}: ${message}`);
      throw error;
    }
  }

  /** Once every group is ready: remove what an earlier version fetched and this one does not list. */
  private prune(): void {
    if (!this.manifest || !this.groups().every((group) => this.ready(group))) return;
    const wanted = new Set<string>();
    for (const [group, listed] of Object.entries(this.manifest.groups)) {
      for (const file of listed.files) wanted.add(path.resolve(this.fileAt(group, file)));
    }
    const walk = (dir: string): void => {
      let entries: fs.Dirent[];
      try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
      } catch {
        return;
      }
      for (const entry of entries) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          walk(full);
          try {
            if (!fs.readdirSync(full).length) fs.rmdirSync(full);
          } catch {
            // in use or gone
          }
        } else if (!wanted.has(path.resolve(full))) {
          fs.rmSync(full, { force: true });
        }
      }
    };
    walk(this.root);
  }

  /** Every group, in the background, one after another (Python first: the code pages use it most). */
  async prefetch(): Promise<boolean> {
    const order = this.groups().sort((a, b) => (a === "pyodide" ? -1 : b === "pyodide" ? 1 : a.localeCompare(b)));
    let ok = true;
    for (const group of order) {
      try {
        await this.ensure(group, false);
      } catch {
        ok = false;
      }
    }
    return ok;
  }
}

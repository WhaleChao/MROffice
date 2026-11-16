/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * MR Note Classic's own LibreOffice (docs/desktop-classic.md): fetched from
 * The Document Foundation's official download site the first time the person
 * asks for it, kept in the app's data folder, never installed into the
 * system and never needing an administrator.
 *
 * - One pinned release per platform (PACKAGES): the file's size and SHA-256,
 *   as TDF publishes them beside the file on download.documentfoundation.org
 *   and on its permanent archive downloadarchive.documentfoundation.org (both
 *   checked to agree when the pin was made). Whatever mirror the download site
 *   redirects to, nothing that does not match is ever opened.
 * - macOS: the official .dmg is attached read-only and unbrowsed
 *   (`hdiutil attach -nobrowse -readonly`), LibreOffice.app is copied out with
 *   `ditto` (its signature intact) and the image is detached. A file the app
 *   downloads itself carries no quarantine; any that came along is removed
 *   from the copy, so a headless soffice never waits on a Gatekeeper prompt.
 * - Windows: the official .msi is unpacked by an administrative install
 *   (`msiexec /a … /qn TARGETDIR=…`): files only, no registry, no elevation.
 *   The Visual C++ runtime the package carries as merge modules is copied
 *   beside soffice.exe, which is where an app-local runtime is looked for.
 * - Before it is used, the copy converts a small text file headless; only
 *   then is it moved into place (one rename) and its record written.
 *
 * The local API finds it by the path of its soffice (MRAI_DESKTOP_LIBREOFFICE,
 * apps/api app/desktop/libreoffice.py), looked at on every conversion: it is
 * used as soon as it is in place and dropped as soon as it is removed.
 *
 * No Electron here: scripts/classic-libreoffice-check.mjs runs the same
 * steps with plain Node (and the Windows CI job does).
 */

import { type ChildProcess, spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { once } from "node:events";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

export const LIBREOFFICE_VERSION = "26.2.6";
/** The build of that release (the archive names its folder by it). */
export const LIBREOFFICE_BUILD = "26.2.6.3";

export interface LibreOfficePackage {
  readonly key: "mac-arm64" | "mac-x64" | "win-x64";
  readonly version: string;
  readonly kind: "dmg" | "msi";
  readonly file: string;
  readonly size: number;
  readonly sha256: string;
  /** Tried in order: the download site (a mirror near the person), then TDF's own archive. */
  readonly urls: readonly string[];
  /** About what the unpacked copy takes on disk. */
  readonly unpackedBytes: number;
}

function pinned(
  key: LibreOfficePackage["key"],
  dir: string,
  name: (version: string) => string,
  kind: LibreOfficePackage["kind"],
  size: number,
  sha256: string,
  unpackedBytes: number,
): LibreOfficePackage {
  return {
    key,
    version: LIBREOFFICE_VERSION,
    kind,
    file: name(LIBREOFFICE_VERSION),
    size,
    sha256,
    urls: [
      `https://download.documentfoundation.org/libreoffice/stable/${LIBREOFFICE_VERSION}/${dir}/${name(LIBREOFFICE_VERSION)}`,
      `https://downloadarchive.documentfoundation.org/libreoffice/old/${LIBREOFFICE_BUILD}/${dir}/${name(LIBREOFFICE_BUILD)}`,
    ],
    unpackedBytes,
  };
}

/** LibreOffice 26.2.6 (build 26.2.6.3), TDF's published SHA-256 of each file. */
export const PACKAGES: Readonly<Record<LibreOfficePackage["key"], LibreOfficePackage>> = {
  "mac-arm64": pinned(
    "mac-arm64",
    "mac/aarch64",
    (v) => `LibreOffice_${v}_MacOS_aarch64.dmg`,
    "dmg",
    297_798_926,
    "94bb3248df074c225490a8a6d1d9dc87c7d6783dbb7a8e9f0d0c3d94348552af",
    950_000_000,
  ),
  "mac-x64": pinned(
    "mac-x64",
    "mac/x86_64",
    (v) => `LibreOffice_${v}_MacOS_x86-64.dmg`,
    "dmg",
    308_308_875,
    "135b8a95b8133396d54bf8e726dbc0066145efa0d785963fc3d4592acbfcfe5b",
    950_000_000,
  ),
  "win-x64": pinned(
    "win-x64",
    "win/x86_64",
    (v) => `LibreOffice_${v}_Win_x86-64.msi`,
    "msi",
    373_252_096,
    "f9877032fd908beb9c0ddf06df4af5c2e85f419c42e14876c4cce5aae5fb2660",
    1_400_000_000,
  ),
};

/**
 * The package for this computer. A Mac with Apple silicon gets the arm64 one
 * even when the app itself runs under Rosetta; Windows on Arm runs the x64
 * build the app ships for (emulated), so it gets the x64 one too.
 */
export function packageFor(platform: NodeJS.Platform, arch: string, arm64Translation = false): LibreOfficePackage | null {
  if (platform === "darwin") return arch === "arm64" || arm64Translation ? PACKAGES["mac-arm64"] : PACKAGES["mac-x64"];
  if (platform === "win32" && (arch === "x64" || arch === "arm64")) return PACKAGES["win-x64"];
  return null;
}

export interface RuntimePaths {
  /** The folder all of it lives in. */
  readonly root: string;
  /** The unpacked LibreOffice (LibreOffice.app on macOS, a folder on Windows). */
  readonly install: string;
  /** Its soffice: what the API is told to run. */
  readonly binary: string;
  /** What is installed (libreoffice.json), written last. */
  readonly record: string;
  /** Downloads in progress (kept across a failure, to resume). */
  readonly downloads: string;
}

export function runtimePaths(root: string, platform: NodeJS.Platform = process.platform): RuntimePaths {
  const install = path.join(root, platform === "darwin" ? "LibreOffice.app" : "LibreOffice");
  return {
    root,
    install,
    binary:
      platform === "darwin"
        ? path.join(install, "Contents", "MacOS", "soffice")
        : platform === "win32"
          ? path.join(install, "program", "soffice.exe")
          : path.join(install, "program", "soffice"),
    record: path.join(root, "libreoffice.json"),
    downloads: path.join(root, "downloads"),
  };
}

export interface InstalledRecord {
  readonly version: string;
  readonly build: string;
  readonly package: string;
  readonly sha256: string;
  readonly installedAt: string;
  readonly bytes: number;
}

/** What is installed, when its soffice is really there. */
export function readRecord(paths: RuntimePaths): InstalledRecord | null {
  try {
    if (!fs.statSync(paths.binary).isFile()) return null;
    const record = JSON.parse(fs.readFileSync(paths.record, "utf8")) as InstalledRecord;
    return typeof record?.version === "string" ? record : null;
  } catch {
    return null;
  }
}

/** Free bytes for this user where `dir` (or its nearest existing parent) is, or null when unknown. */
export function freeBytes(dir: string): number | null {
  let probe = dir;
  for (;;) {
    try {
      const stats = fs.statfsSync(probe);
      return Number(stats.bavail) * Number(stats.bsize);
    } catch {
      const parent = path.dirname(probe);
      if (parent === probe) return null;
      probe = parent;
    }
  }
}

const MARGIN_BYTES = 300_000_000;

/** What the install still needs on disk: the rest of the download, the unpacked copy, a margin. */
export function bytesNeeded(pkg: LibreOfficePackage, downloadedBytes: number): number {
  return Math.max(0, pkg.size - downloadedBytes) + pkg.unpackedBytes + MARGIN_BYTES;
}

export class InstallError extends Error {
  constructor(
    readonly code: "no_space" | "download" | "checksum" | "unpack" | "check" | "cancelled" | "unsupported" | "busy",
    message: string,
    readonly details: Record<string, unknown> = {},
  ) {
    super(message);
    this.name = "InstallError";
  }
}

export type Phase = "space" | "download" | "verify" | "unpack" | "check" | "done";

export interface InstallOptions {
  readonly fetch?: typeof fetch;
  readonly signal?: AbortSignal;
  readonly onProgress?: (phase: Phase, received: number, total: number) => void;
  /** Tests: the steps that run a program. */
  readonly run?: Runner;
  readonly platform?: NodeJS.Platform;
  readonly log?: (line: string) => void;
}

// --- download ---------------------------------------------------------------------

function partFile(pkg: LibreOfficePackage, paths: RuntimePaths): string {
  return path.join(paths.downloads, `${pkg.file}.part`);
}

export function downloadedBytes(pkg: LibreOfficePackage, paths: RuntimePaths): number {
  try {
    return Math.min(fs.statSync(partFile(pkg, paths)).size, pkg.size);
  } catch {
    return 0;
  }
}

async function hashOf(file: string, signal?: AbortSignal): Promise<ReturnType<typeof createHash>> {
  const hash = createHash("sha256");
  const stream = fs.createReadStream(file, { highWaterMark: 4 << 20 });
  for await (const chunk of stream) {
    if (signal?.aborted) {
      stream.destroy();
      throw new InstallError("cancelled", "cancelled");
    }
    hash.update(chunk as Buffer);
  }
  return hash;
}

/**
 * The package, downloaded (resuming a partial one), its size and SHA-256
 * checked. Each URL is tried in turn; a file that does not match is deleted.
 */
export async function download(pkg: LibreOfficePackage, paths: RuntimePaths, options: InstallOptions = {}): Promise<string> {
  const doFetch = options.fetch ?? fetch;
  const signal = options.signal;
  const part = partFile(pkg, paths);
  fs.mkdirSync(paths.downloads, { recursive: true });
  let lastError: unknown = null;
  for (const url of pkg.urls) {
    try {
      let have = downloadedBytes(pkg, paths);
      if (have > pkg.size) {
        fs.rmSync(part, { force: true });
        have = 0;
      }
      let hash = have > 0 ? await hashOf(part, signal) : createHash("sha256");
      if (have < pkg.size) {
        options.log?.(`downloading ${url}${have ? ` from byte ${have}` : ""}`);
        const response = await doFetch(url, {
          signal,
          redirect: "follow",
          headers: have > 0 ? { range: `bytes=${have}-` } : {},
        });
        if (response.status === 200 && have > 0) {
          // The server sends the whole file, not the rest: it starts over.
          have = 0;
          hash = createHash("sha256");
        }
        if (!(response.status === 200 || response.status === 206) || !response.body) {
          throw new InstallError("download", `HTTP ${response.status} for ${url}`, { status: response.status, url });
        }
        const out = fs.createWriteStream(part, { flags: have > 0 ? "a" : "w" });
        let received = have;
        let shown = 0;
        try {
          const reader = response.body.getReader();
          for (;;) {
            const { done, value } = await reader.read();
            if (done) break;
            received += value.length;
            if (received > pkg.size) throw new InstallError("download", `the download is larger than ${pkg.size} bytes`, { url });
            hash.update(value);
            if (!out.write(value)) await once(out, "drain");
            if (Date.now() - shown > 200) {
              shown = Date.now();
              options.onProgress?.("download", received, pkg.size);
            }
          }
        } finally {
          out.end();
          await once(out, "close").catch(() => undefined);
        }
        options.onProgress?.("download", received, pkg.size);
        if (received !== pkg.size) {
          // Kept: the next try resumes from here.
          throw new InstallError("download", `the download stopped at ${received} of ${pkg.size} bytes`, { url, received });
        }
      }
      options.onProgress?.("verify", pkg.size, pkg.size);
      const digest = hash.digest("hex");
      if (digest !== pkg.sha256) {
        fs.rmSync(part, { force: true });
        throw new InstallError("checksum", `SHA-256 ${digest} is not the published ${pkg.sha256}`, { url, digest });
      }
      const done = path.join(paths.downloads, pkg.file);
      fs.renameSync(part, done);
      return done;
    } catch (error) {
      if (signal?.aborted) throw new InstallError("cancelled", "cancelled");
      lastError = error;
      options.log?.(`download from ${url} failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  throw lastError instanceof InstallError
    ? lastError
    : new InstallError("download", lastError instanceof Error ? lastError.message : String(lastError));
}

// --- programs ---------------------------------------------------------------------

export interface RunOptions {
  readonly signal?: AbortSignal;
  readonly timeoutMs?: number;
  readonly cwd?: string;
  readonly env?: NodeJS.ProcessEnv;
  /** Windows: the arguments are one command line as given (msiexec parses its own). */
  readonly verbatim?: boolean;
}

export type Runner = (command: string, args: readonly string[], options?: RunOptions) => Promise<{ code: number; output: string }>;

export const runProgram: Runner = (command, args, options = {}) =>
  new Promise((resolve, reject) => {
    let child: ChildProcess;
    try {
      child = spawn(command, [...args], {
        cwd: options.cwd,
        env: options.env,
        stdio: ["ignore", "pipe", "pipe"],
        windowsHide: true,
        windowsVerbatimArguments: options.verbatim === true,
        detached: process.platform !== "win32",
      });
    } catch (error) {
      reject(error);
      return;
    }
    let output = "";
    const keep = (chunk: Buffer) => {
      if (output.length < 64_000) output += chunk.toString("utf8");
    };
    child.stdout?.on("data", keep);
    child.stderr?.on("data", keep);
    const stop = () => {
      try {
        if (process.platform === "win32") spawn("taskkill", ["/pid", String(child.pid), "/t", "/f"], { windowsHide: true });
        else process.kill(-(child.pid as number), "SIGKILL");
      } catch {
        child.kill("SIGKILL");
      }
    };
    const timer = options.timeoutMs ? setTimeout(stop, options.timeoutMs) : null;
    const onAbort = () => stop();
    options.signal?.addEventListener("abort", onAbort, { once: true });
    child.on("error", (error) => {
      if (timer) clearTimeout(timer);
      options.signal?.removeEventListener("abort", onAbort);
      reject(error);
    });
    child.on("close", (code, signal) => {
      if (timer) clearTimeout(timer);
      options.signal?.removeEventListener("abort", onAbort);
      resolve({ code: code ?? (signal ? 128 : 1), output });
    });
  });

async function mustRun(run: Runner, code: InstallError["code"], command: string, args: readonly string[], options?: RunOptions): Promise<string> {
  const result = await run(command, args, options);
  if (options?.signal?.aborted) throw new InstallError("cancelled", "cancelled");
  if (result.code !== 0) {
    throw new InstallError(code, `${path.basename(command)} exited ${result.code}: ${result.output.trim().slice(-600)}`, { exitCode: result.code });
  }
  return result.output;
}

// --- unpack -----------------------------------------------------------------------

/** msiexec's own command line for an administrative install (files only). */
export function msiexecCommandLine(msi: string, target: string, log: string): string {
  const quote = (value: string) => `"${value.replace(/"/g, "")}"`;
  return `/a ${quote(msi)} /qn TARGETDIR=${quote(target)} /L*v ${quote(log)}`;
}

/** The folder holding soffice[.exe] under `root` (an admin install nests it), or null. */
export function findProgramDir(root: string, executable = "soffice.exe", depth = 4): string | null {
  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(root, { withFileTypes: true });
  } catch {
    return null;
  }
  if (path.basename(root).toLowerCase() === "program" && entries.some((entry) => entry.isFile() && entry.name.toLowerCase() === executable)) {
    return root;
  }
  if (depth <= 0) return null;
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const found = findProgramDir(path.join(root, entry.name), executable, depth - 1);
    if (found) return found;
  }
  return null;
}

const VC_RUNTIME = /^(msvcp140.*|vcruntime140.*|concrt140|vccorlib140|vcomp140)\.dll$/i;

/**
 * The Visual C++ runtime DLLs an admin install leaves where the merge
 * modules would have put them in System32, copied beside soffice.exe (an
 * app-local runtime), unless one of that name is already there.
 */
export function placeVcRuntime(extracted: string, programDir: string): string[] {
  const copied: string[] = [];
  const walk = (dir: string, depth: number) => {
    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        if (depth > 0 && path.resolve(full) !== path.resolve(programDir)) walk(full, depth - 1);
      } else if (VC_RUNTIME.test(entry.name)) {
        const target = path.join(programDir, entry.name);
        if (!fs.existsSync(target)) {
          fs.copyFileSync(full, target);
          copied.push(entry.name);
        }
      }
    }
  };
  walk(extracted, 5);
  return copied;
}

function directoryBytes(dir: string): number {
  let total = 0;
  const walk = (current: string) => {
    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (entry.isFile()) {
        try {
          total += fs.statSync(full).size;
        } catch {
          // gone
        }
      }
    }
  };
  walk(dir);
  return total;
}

async function unpackDmg(dmg: string, staging: string, paths: RuntimePaths, options: InstallOptions): Promise<string> {
  const run = options.run ?? runProgram;
  const mount = path.join(paths.root, `mount-${process.pid}-${Date.now()}`);
  fs.mkdirSync(mount, { recursive: true });
  await mustRun(
    run,
    "unpack",
    "/usr/bin/hdiutil",
    ["attach", "-nobrowse", "-readonly", "-noautoopen", "-noverify", "-mountpoint", mount, dmg],
    { signal: options.signal, timeoutMs: 180_000 },
  );
  try {
    const source = path.join(mount, "LibreOffice.app");
    if (!fs.existsSync(path.join(source, "Contents", "MacOS", "soffice"))) {
      throw new InstallError("unpack", "the disk image holds no LibreOffice.app");
    }
    const app = path.join(staging, "LibreOffice.app");
    await mustRun(run, "unpack", "/usr/bin/ditto", [source, app], { signal: options.signal, timeoutMs: 600_000 });
    // Downloaded by the app itself, the image had no quarantine; whatever
    // came along is dropped so a headless soffice never waits on a prompt.
    await run("/usr/bin/xattr", ["-dr", "com.apple.quarantine", app], { timeoutMs: 120_000 });
    return app;
  } finally {
    const detached = await run("/usr/bin/hdiutil", ["detach", mount, "-quiet"], { timeoutMs: 60_000 }).catch(() => ({ code: 1, output: "" }));
    if (detached.code !== 0) await run("/usr/bin/hdiutil", ["detach", mount, "-force", "-quiet"], { timeoutMs: 60_000 }).catch(() => undefined);
    fs.rmSync(mount, { recursive: true, force: true });
  }
}

async function unpackMsi(msi: string, staging: string, options: InstallOptions): Promise<string> {
  const run = options.run ?? runProgram;
  const extracted = path.join(staging, "admin");
  fs.mkdirSync(extracted, { recursive: true });
  const log = path.join(staging, "msiexec.log");
  const msiexec = path.join(process.env.SystemRoot || process.env.windir || "C:\\Windows", "System32", "msiexec.exe");
  const result = await run(msiexec, [msiexecCommandLine(msi, extracted, log)], {
    signal: options.signal,
    timeoutMs: 900_000,
    verbatim: true,
  });
  if (options.signal?.aborted) throw new InstallError("cancelled", "cancelled");
  // 3010: done, a restart is advised (never for an admin install, but not a failure).
  if (result.code !== 0 && result.code !== 3010) {
    let tail = "";
    try {
      tail = fs.readFileSync(log, "utf16le").slice(-1500);
    } catch {
      // no log
    }
    throw new InstallError("unpack", `msiexec /a exited ${result.code}${tail ? `: ${tail}` : ""}`, { exitCode: result.code });
  }
  const program = findProgramDir(extracted);
  if (!program) throw new InstallError("unpack", "the package unpacked without program\\soffice.exe");
  placeVcRuntime(extracted, program);
  const install = path.join(staging, "LibreOffice");
  fs.renameSync(path.dirname(program), install);
  fs.rmSync(extracted, { recursive: true, force: true });
  // An admin image keeps a copy of the package's database at its root.
  for (const name of fs.readdirSync(install)) {
    if (name.toLowerCase().endsWith(".msi")) fs.rmSync(path.join(install, name), { force: true });
  }
  return install;
}

/** The copy converts a small text file headless (its own throwaway profile). */
export async function smokeTest(binary: string, options: { run?: Runner; signal?: AbortSignal; timeoutMs?: number } = {}): Promise<void> {
  const run = options.run ?? runProgram;
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), "mrai-lo-check-"));
  try {
    const input = path.join(scratch, "check.txt");
    fs.writeFileSync(input, "MR Note Classic LibreOffice check\n");
    const out = path.join(scratch, "out");
    fs.mkdirSync(out);
    const profile = pathToFileURL(path.join(scratch, "profile")).href;
    const result = await run(
      binary,
      [`-env:UserInstallation=${profile}`, "--headless", "--invisible", "--norestore", "--nologo", "--nodefault", "--convert-to", "odt", "--outdir", out, input],
      { signal: options.signal, timeoutMs: options.timeoutMs ?? 240_000, cwd: scratch },
    );
    if (options.signal?.aborted) throw new InstallError("cancelled", "cancelled");
    const written = path.join(out, "check.odt");
    if (!fs.existsSync(written) || fs.statSync(written).size === 0) {
      throw new InstallError("check", `LibreOffice did not convert a test file (exit ${result.code}): ${result.output.trim().slice(-600)}`, {
        exitCode: result.code,
      });
    }
  } finally {
    fs.rmSync(scratch, { recursive: true, force: true });
  }
}

/** Removes `target` by moving it aside first (one rename), so a half-deleted copy is never in place. */
export function removeTree(target: string, root: string): void {
  if (!fs.existsSync(target)) return;
  const aside = path.join(root, `.removing-${path.basename(target)}-${Date.now()}`);
  fs.renameSync(target, aside);
  fs.rmSync(aside, { recursive: true, force: true });
}

/** Leftovers of an install or removal that did not finish. */
export function cleanLeftovers(paths: RuntimePaths): void {
  let names: string[] = [];
  try {
    names = fs.readdirSync(paths.root);
  } catch {
    return;
  }
  for (const name of names) {
    if (name.startsWith(".removing-") || name.startsWith("staging-") || name.startsWith("mount-")) {
      const full = path.join(paths.root, name);
      if (name.startsWith("mount-")) {
        // Never delete through a still-attached image.
        try {
          if (fs.readdirSync(full).length) continue;
        } catch {
          continue;
        }
      }
      fs.rmSync(full, { recursive: true, force: true });
    }
  }
}

/**
 * Space, download, check, unpack, try, move into place, record. Resolves to
 * the record; throws InstallError (code `cancelled` when the signal fired).
 */
export async function install(pkg: LibreOfficePackage, paths: RuntimePaths, options: InstallOptions = {}): Promise<InstalledRecord> {
  const platform = options.platform ?? process.platform;
  fs.mkdirSync(paths.root, { recursive: true });
  cleanLeftovers(paths);
  options.onProgress?.("space", 0, 0);
  const have = downloadedBytes(pkg, paths);
  const done = path.join(paths.downloads, pkg.file);
  const alreadyDownloaded = fs.existsSync(done) && fs.statSync(done).size === pkg.size;
  const need = alreadyDownloaded ? pkg.unpackedBytes + MARGIN_BYTES : bytesNeeded(pkg, have);
  const free = freeBytes(paths.root);
  if (free !== null && free < need) {
    throw new InstallError("no_space", `needs ${need} bytes free, ${free} are`, { need, free });
  }
  let file = done;
  if (alreadyDownloaded) {
    options.onProgress?.("verify", pkg.size, pkg.size);
    const digest = (await hashOf(done, options.signal)).digest("hex");
    if (digest !== pkg.sha256) {
      fs.rmSync(done, { force: true });
      file = await download(pkg, paths, options);
    }
  } else {
    file = await download(pkg, paths, options);
  }
  options.onProgress?.("unpack", 0, 0);
  const staging = path.join(paths.root, `staging-${Date.now()}`);
  fs.mkdirSync(staging, { recursive: true });
  try {
    const unpacked =
      platform === "darwin"
        ? await unpackDmg(file, staging, paths, options)
        : platform === "win32"
          ? await unpackMsi(file, staging, options)
          : (() => {
              throw new InstallError("unsupported", `no LibreOffice package for ${platform}`);
            })();
    const stagedBinary = path.join(unpacked, path.relative(paths.install, paths.binary));
    if (!fs.existsSync(stagedBinary)) throw new InstallError("unpack", `no ${path.relative(paths.install, paths.binary)} in the unpacked copy`);
    options.onProgress?.("check", 0, 0);
    await smokeTest(stagedBinary, { run: options.run, signal: options.signal });
    // Into place: the old copy (another version, or a broken one) aside first.
    try {
      fs.rmSync(paths.record, { force: true });
      removeTree(paths.install, paths.root);
    } catch (error) {
      throw new InstallError("busy", `the LibreOffice in use could not be replaced: ${error instanceof Error ? error.message : String(error)}`);
    }
    fs.renameSync(unpacked, paths.install);
    const record: InstalledRecord = {
      version: pkg.version,
      build: LIBREOFFICE_BUILD,
      package: pkg.file,
      sha256: pkg.sha256,
      installedAt: new Date().toISOString(),
      bytes: directoryBytes(paths.install),
    };
    fs.writeFileSync(paths.record, `${JSON.stringify(record, null, 2)}\n`);
    options.onProgress?.("done", 1, 1);
    // The package is not kept: the copy is what is used (and disks are small).
    fs.rmSync(file, { force: true });
    return record;
  } finally {
    fs.rmSync(staging, { recursive: true, force: true });
  }
}

/** The app's copy and everything downloaded for it. */
export function uninstall(paths: RuntimePaths): void {
  fs.rmSync(paths.record, { force: true });
  removeTree(paths.install, paths.root);
  fs.rmSync(paths.downloads, { recursive: true, force: true });
  cleanLeftovers(paths);
}

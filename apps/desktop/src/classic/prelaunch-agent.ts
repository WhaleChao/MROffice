/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The parts of ./prelaunch.ts that need no Electron (its tests run on Node):
 * the login launch's argument and the macOS LaunchAgent that opens the app at
 * login.
 */

export const BACKGROUND_ARG = "--mrai-background";

/** Whether this launch is the login's (no window until the app is opened). */
export function launchedInBackground(argv: readonly string[] = process.argv): boolean {
  return argv.includes(BACKGROUND_ARG);
}

/** Whether the app starts at login: on unless the person turned it off
 * (``choice``: the menu's checkbox as last set; undefined, never set). */
export function prelaunchWanted(choice: boolean | undefined): boolean {
  return choice !== false;
}

export function agentLabel(appId: string): string {
  return `${appId}.prelaunch`;
}

function xml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function launchAgentPlist(appId: string, bundle: string): string {
  const args = ["/usr/bin/open", "-g", "-j", "-a", bundle, "--args", BACKGROUND_ARG];
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${xml(agentLabel(appId))}</string>
  <key>ProgramArguments</key>
  <array>
${args.map((arg) => `    <string>${xml(arg)}</string>`).join("\n")}
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>LimitLoadToSessionType</key>
  <string>Aqua</string>
</dict>
</plist>
`;
}

/** Restarts after a service stopped while nothing of the app was on screen,
 * at most this many within ``BACKGROUND_RESTART_WINDOW_MS``: a service that
 * keeps dying is not restarted for ever behind the person's back. */
export const BACKGROUND_RESTARTS_MAX = 3;
export const BACKGROUND_RESTART_WINDOW_MS = 30 * 60_000;

/** Whether a service that stopped while the app showed nothing (started in
 * the background at login, or every window closed) is restarted quietly now,
 * and the restart times to keep (``history``: the earlier ones, ms). */
export function backgroundRestartAllowed(history: readonly number[], now: number): { allowed: boolean; history: number[] } {
  const recent = history.filter((at) => Number.isFinite(at) && at <= now && now - at < BACKGROUND_RESTART_WINDOW_MS);
  if (recent.length >= BACKGROUND_RESTARTS_MAX) return { allowed: false, history: recent };
  return { allowed: true, history: [...recent, now] };
}

/** The arguments a quiet restart relaunches with: the app's own switches
 * (never a document or link the first launch was given: it was opened then)
 * and the login's, so it starts in the background again. ``script``: an
 * unpacked run's (``electron <script>``) first argument, kept as it is. */
export function backgroundRestartArgs(argv: readonly string[], script = false): string[] {
  const kept = script && argv[1] ? [argv[1]] : [];
  const own = argv.slice(script ? 2 : 1).filter((arg) => arg.startsWith("--") && arg !== BACKGROUND_ARG);
  return [...kept, ...own, BACKGROUND_ARG];
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which bundled services this app itself asked to stop.
 *
 * A service that exits after it was asked to is not a crash. On Windows a
 * quit asks the API to stop by itself first (main.ts will-quit,
 * stopServiceGracefully: Windows has no SIGTERM) and only marks the app as
 * quitting after that; the API's exit then reached the crash watch, which
 * counted a service crash and opened "a background part stopped" with
 * Restart/Quit on an ordinary quit. Production desktop usage, 11-22/11-23:
 * service crashes on Windows only (Office 0.1.6 and 0.1.7, Classic 0.1.11,
 * up to three in one day for one person), none on macOS, whose quit stops
 * the services after it marks itself quitting.
 */
const asked = new WeakSet<object>();

/** Note that this app asked `service` to stop (a request, a signal or a kill). */
export function markStopAsked(service: object): void {
  asked.add(service);
}

/** True when the service's exit was asked for by this app. */
export function stopWasAsked(service: object): boolean {
  return asked.has(service);
}

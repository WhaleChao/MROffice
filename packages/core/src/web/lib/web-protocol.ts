/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether a page left open across a release still works with the server.
 *
 * A new build alone is no reason to interrupt anyone: the page keeps working
 * and loads the new build at its next page change (WebBuildFreshnessGuard).
 * Only a change that breaks open pages asks them to reload now: raise
 * WEB_CLIENT_PROTOCOL when a build's pages need what an older server does not
 * have, and WEB_MIN_CLIENT_PROTOCOL (never above WEB_CLIENT_PROTOCOL) when
 * this build's server or its collaborators refuse what an older page sends --
 * a save, room or editor schema it no longer accepts. /web-build.json
 * publishes both; a page whose protocol is below the server's minimum shows
 * the reload notice. A save refused on an old build shows it as well, without
 * waiting for anyone to raise a number (lib/web-build.ts).
 */
export const WEB_CLIENT_PROTOCOL = 1;
export const WEB_MIN_CLIENT_PROTOCOL = 1;

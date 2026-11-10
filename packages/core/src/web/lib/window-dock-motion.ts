/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Minimizing as on a Mac's Dock: a window is drawn into its place in the
 * minimized tray, and comes back out of that place when restored.
 *
 * Minimize: just before the window leaves the desk (Window.tsx), an inert
 * picture of it (lib/inert-ghost) is noted with its place on screen. The
 * tray (MinimizedTray), drawing the window's new tile, flies the picture
 * into the tile and shows the tile as it lands. Restore: pressing a tile
 * notes its place; the window, mounting again, grows out of it into its own
 * place (the live window itself, not a picture).
 *
 * Nothing waits on the motion: window state changes exactly as before, a
 * picture nobody claims soon is dropped, and with reduced motion (system or
 * Settings) or in a test DOM nothing is noted, so nothing plays.
 */
import { exitMotionAllowed } from "./inert-ghost";
import { surfacePictureOf } from "./desk-motion";

/** Length of either flight. */
export const DOCK_FLIGHT_MS = 460;
/** A noted place older than this belongs to some other gesture. */
const CLAIM_WINDOW_MS = 3000;
const GHOST_ATTRIBUTE = "data-window-dock-ghost";
/** Elements the picture copies with their contents (lib/desk-motion). */
const PICTURE_BUDGET = 6000;

interface Departure {
  picture: HTMLElement;
  /** Puts the picture's scrollers where the window's were, once shown. */
  settle: () => void;
  /** On screen (canvas zoom included). */
  rect: DOMRect;
  /** The window's own CSS size, which the picture keeps. */
  width: number;
  height: number;
  at: number;
}

const departures = new Map<string, Departure>();
const returns = new Map<string, { rect: DOMRect; at: number }>();

const fresh = (at: number) => performance.now() - at < CLAIM_WINDOW_MS;

function motionAllowed(): boolean {
  return typeof window !== "undefined" && exitMotionAllowed(window);
}

/** How long a picture waits for its tile before it simply fades in place. */
const UNCLAIMED_FADE_AFTER_MS = 160;

/**
 * Take the leaving window's picture and place, for its tray tile to claim.
 * `fadeIfUnclaimed`: the window is leaving now, so if no tray on screen
 * claims the picture at once (the sidebar hidden) it fades where it stood.
 */
export function noteDockDeparture(
  windowId: string,
  shell: HTMLElement | null,
  { fadeIfUnclaimed = false }: { fadeIfUnclaimed?: boolean } = {},
): void {
  if (!shell || !motionAllowed()) return;
  const rect = shell.getBoundingClientRect();
  if (rect.width < 24 || rect.height < 24) return;
  // The picture keeps what is on screen (a copied frame or player, which
  // would load its document again, keeps only its box); the rest of a long
  // page is an empty box of its size (lib/desk-motion).
  const { element: picture, settle } = surfacePictureOf(shell, GHOST_ATTRIBUTE, {
    stripAttributes: ["data-window-id"],
    budget: PICTURE_BUDGET,
  });
  const departure: Departure = {
    picture,
    settle,
    rect,
    width: shell.offsetWidth || rect.width,
    height: shell.offsetHeight || rect.height,
    at: performance.now(),
  };
  departures.set(windowId, departure);
  if (fadeIfUnclaimed) {
    window.setTimeout(() => {
      if (departures.get(windowId) !== departure) return;
      departures.delete(windowId);
      const host = flightHost(departure);
      try {
        const fade = host.animate(
          [
            { transform: "scale(1)", opacity: 1 },
            { transform: "translateY(18px) scale(0.92)", opacity: 0 },
          ],
          { duration: 200, easing: "cubic-bezier(0.4, 0, 0.2, 1)", fill: "forwards" },
        );
        fade.onfinish = () => host.remove();
      } catch {
        host.remove();
      }
      window.setTimeout(() => host.remove(), 440);
    }, UNCLAIMED_FADE_AFTER_MS);
  }
}

/** The picture, on screen exactly where the window stood. */
function flightHost(departure: Departure): HTMLElement {
  const from = departure.rect;
  const host = document.createElement("div");
  host.className = "wm-dock-flight";
  host.setAttribute(GHOST_ATTRIBUTE, "");
  host.setAttribute("aria-hidden", "true");
  Object.assign(host.style, {
    left: `${from.left}px`,
    top: `${from.top}px`,
    width: `${from.width}px`,
    height: `${from.height}px`,
  });
  const picture = departure.picture;
  Object.assign(picture.style, {
    position: "absolute",
    left: "0",
    top: "0",
    margin: "0",
    width: `${departure.width}px`,
    height: `${departure.height}px`,
    transformOrigin: "0 0",
    transform: `scale(${from.width / departure.width}, ${from.height / departure.height})`,
    animation: "none",
    transition: "none",
  });
  host.append(picture);
  document.body.append(host);
  departure.settle();
  return host;
}

/**
 * Fly a waiting picture of `windowId` into its tile. Calls `landed` when the
 * tile may show (at once when nothing flies); returns whether one flies.
 */
export function flyIntoDockTile(
  windowId: string,
  tile: HTMLElement,
  landed: () => void,
): boolean {
  const departure = departures.get(windowId);
  const to = tile.getBoundingClientRect();
  // A tray that is not on screen (the sidebar's other layout) leaves the
  // picture for the one that is.
  if (!departure || to.width === 0 || to.height === 0) return false;
  departures.delete(windowId);
  if (!fresh(departure.at)) return false;

  const from = departure.rect;
  const host = flightHost(departure);

  const dx = to.left + to.width / 2 - (from.left + from.width / 2);
  const dy = to.top + to.height / 2 - (from.top + from.height / 2);
  const sx = to.width / from.width;
  const sy = to.height / from.height;
  let done = false;
  const finish = () => {
    if (done) return;
    done = true;
    host.remove();
    landed();
  };
  try {
    // Narrows toward the tray first, then shrinks into the tile: the Genie's
    // pull without bending the picture.
    const flight = host.animate(
      [
        { transform: "translate(0px, 0px) scale(1, 1)", opacity: 1 },
        {
          offset: 0.45,
          transform: `translate(${dx * 0.4}px, ${dy * 0.2}px) scale(${Math.max(sx, 0.42)}, ${Math.max(sy, 0.8)})`,
          opacity: 1,
        },
        { transform: `translate(${dx}px, ${dy}px) scale(${sx}, ${sy})`, opacity: 0.25 },
      ],
      { duration: DOCK_FLIGHT_MS, easing: "cubic-bezier(0.55, 0, 0.3, 1)", fill: "forwards" },
    );
    flight.onfinish = finish;
    flight.oncancel = finish;
  } catch {
    finish();
    return false;
  }
  // However the animation ends, the picture never outlives the flight.
  window.setTimeout(finish, DOCK_FLIGHT_MS + 240);
  return true;
}

/** A tile is pressed to restore its window: note where the window comes from. */
export function noteDockReturn(windowId: string, tile: HTMLElement | null): void {
  if (!tile || !motionAllowed()) return;
  const rect = tile.getBoundingClientRect();
  if (rect.width === 0 || rect.height === 0) return;
  returns.set(windowId, { rect, at: performance.now() });
}

/** The restored window, just mounted, grows out of its tile into its place. */
export function growFromDockTile(windowId: string, shell: HTMLElement | null): void {
  const origin = returns.get(windowId);
  returns.delete(windowId);
  if (!origin || !shell || !fresh(origin.at) || !motionAllowed()) return;
  const to = shell.getBoundingClientRect();
  if (to.width === 0 || to.height === 0 || !shell.offsetWidth) return;
  // The shell moves in the canvas's own (zoomed) space.
  const zoom = to.width / shell.offsetWidth || 1;
  const from = origin.rect;
  const dx = (from.left + from.width / 2 - (to.left + to.width / 2)) / zoom;
  const dy = (from.top + from.height / 2 - (to.top + to.height / 2)) / zoom;
  const sx = from.width / to.width;
  const sy = from.height / to.height;
  try {
    shell.animate(
      [
        { transform: `translate(${dx}px, ${dy}px) scale(${sx}, ${sy})`, opacity: 0.3 },
        {
          offset: 0.55,
          transform: `translate(${dx * 0.2}px, ${dy * 0.08}px) scale(${Math.max(sx, 0.62)}, ${Math.max(sy, 0.88)})`,
          opacity: 1,
        },
        { transform: "translate(0px, 0px) scale(1, 1)", opacity: 1 },
      ],
      { duration: DOCK_FLIGHT_MS, easing: "cubic-bezier(0.2, 0.75, 0.2, 1)" },
    );
  } catch {
    // An engine without Web Animations shows the window in place.
  }
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Rescue for `Alt-<letter>` shortcuts on macOS.
 *
 * On a Mac, Option+letter TYPES a character: Option+F is "ƒ", Option+Shift+F is
 * "Ï". CodeMirror knows this and deliberately refuses to fall back to the
 * physical key for Alt combinations there — see the guard in its own keymap:
 *
 *     // Alt-combinations on macOS tend to be typed characters
 *     !(browser.mac && event.altKey && !(event.ctrlKey || event.metaKey))
 *
 * The consequence is that EVERY `Alt-<letter>` binding is dead on macOS, no
 * matter how it is written: the event's key is never the letter the binding
 * names. Format Document (Shift+Alt+F), Organize Imports (Shift+Alt+O) and
 * Toggle Word Wrap (Alt+Z) all did nothing, while the menus and the shortcut
 * reference went on advertising them.
 *
 * `event.code` is the physical key, unaffected by what the OS decided to type,
 * so that is what this matches on. The table is DERIVED from the real key
 * bindings rather than written out again, so a chord cannot be rescued into
 * doing something different from what its binding says.
 *
 * Nothing changes off macOS: there CodeMirror's own fallback already works, and
 * running this too would execute the command twice.
 */

import { isMacPlatform } from "@mrnote/core/web/lib/mac-platform.ts";
import { EditorView, type KeyBinding } from "@codemirror/view";
import type { Extension } from "@codemirror/state";

/** `Alt-z` / `Shift-Alt-f` → the physical key and whether Shift is required. */
export interface PhysicalAltChord {
  code: string;
  shift: boolean;
  binding: KeyBinding;
}

const ALT_LETTER = /^(Shift-)?Alt-([A-Za-z])$/;

/** Whether this browser is the one with the Option-types-a-character problem. */
export { isMacPlatform };

/** The Alt+letter chords among `bindings`, keyed by physical key. */
export function physicalAltChords(
  bindings: readonly KeyBinding[],
): PhysicalAltChord[] {
  const chords: PhysicalAltChord[] = [];
  for (const binding of bindings) {
    const match = binding.key ? ALT_LETTER.exec(binding.key) : null;
    if (!match || !binding.run) continue;
    chords.push({
      code: `Key${match[2]!.toUpperCase()}`,
      shift: Boolean(match[1]),
      binding,
    });
  }
  return chords;
}

/** The parts of a keyboard event this matcher reads. */
export interface PhysicalChordEvent {
  code: string;
  altKey: boolean;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  defaultPrevented: boolean;
  preventDefault: () => void;
}

/**
 * Run the chord whose PHYSICAL key matches this event, if any. Returns whether
 * the event was consumed. Pure enough to test without a live editor, which is
 * the point: the bug this exists for is invisible to a test that fabricates
 * key events.
 */
export function runPhysicalAltChord<V>(
  chords: readonly PhysicalAltChord[],
  event: PhysicalChordEvent,
  view: V,
): boolean {
  // Only the combination CodeMirror gives up on: Alt without Ctrl/Meta.
  if (!event.altKey || event.ctrlKey || event.metaKey) return false;
  // Something already claimed this key; do not run a command twice.
  if (event.defaultPrevented) return false;
  for (const chord of chords) {
    if (chord.code !== event.code) continue;
    if (chord.shift !== event.shiftKey) continue;
    const ran = (chord.binding.run as ((view: V) => boolean) | undefined)?.(
      view,
    );
    // A command that declines (returns false) leaves the key alone, exactly as
    // it would through the keymap.
    if (ran || chord.binding.preventDefault) {
      event.preventDefault();
      return true;
    }
    return false;
  }
  return false;
}

/**
 * An extension that makes the `Alt-<letter>` bindings in `bindings` reachable on
 * macOS. Returns nothing at all on other platforms.
 */
export function macOptionChordFallback(
  bindings: readonly KeyBinding[],
  isMac: boolean = isMacPlatform(),
): Extension {
  if (!isMac) return [];
  const chords = physicalAltChords(bindings);
  if (chords.length === 0) return [];
  return EditorView.domEventHandlers({
    keydown: (event, view) => runPhysicalAltChord(chords, event, view),
  });
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spreadsheet's command registry — slice 1: declarative keybindings.
 *
 * Commands live today as `useCallback`s inside SpreadsheetWindow and their
 * shortcuts as a long if/else chain inside two keydown handlers. That shape
 * has no seam for a keybinding overview, per-command telemetry, or a future
 * scripting facade — every consumer would have to re-read the chain. This
 * module is the seam: commands are IDs, bindings are DATA, and dispatch is a
 * lookup. The design (a registry of command ids consulted by an input layer)
 * is the one Univer's OSS command service uses (Apache-2.0, dream-num/univer);
 * the code is not derived from theirs — it is a table and two functions
 * written against this component's existing handlers.
 *
 * Deliberately NOT in this slice: undo journalling (the sheet snapshot undo
 * already owns that), command payloads, or migrating every shortcut. Handlers
 * are re-registered per render by the host — a registry entry is a live ref
 * into component state, not a global.
 *
 * `stage` preserves the host's dispatch ORDER exactly: the keydown chain
 * consults "chrome" bindings at the top of the handler (where undo/redo have
 * always run, before every other branch) and "selection" bindings down in the
 * default branch (where the format cluster has always run). Collapsing the two
 * into one early consult would MOVE the format cluster ahead of branches that
 * today outrank it, which is a behaviour change this slice refuses to make.
 * A binding whose key a branch between the two would otherwise swallow (F2,
 * PageDown, Alt+Shift+Arrow, Ctrl+Backspace, Shift+Space, the Shift+F-keys)
 * sits in "chrome"; every other Ctrl/Cmd binding sits in "selection".
 *
 * Slice 2 — Excel's grid shortcuts (Microsoft Support, "Keyboard shortcuts in
 * Excel", the Windows, macOS and Excel for the web tables). Each one runs a
 * command the ribbon or a menu already owns; the host's handler goes through
 * the same control authority the button asks, so a protected or read-only
 * sheet refuses the key exactly as it greys the button. A handler that returns
 * `false` did not run: the key is left to the rest of the chain (and the
 * browser), so preventDefault only follows a shortcut that acted.
 */

export type SpreadsheetCommandId =
  | "undo"
  | "redo"
  | "copySelection"
  | "cutSelection"
  // Dispatched from the input's native paste event, not a keybinding: the
  // browser owns Ctrl/Cmd+V delivery, the registry owns what happens next.
  // Payload: { event: ClipboardEvent-like, row, col }.
  | "pasteIntoCell"
  // Scripted point-write (facade.setCell): the same guarded single value-commit
  // path typing uses — protection notice, undo step, autosave — never a direct
  // model write. Payload: { row, col, value }.
  | "setCellValue"
  // Scripted bulk write (facade.setRange): the shared paste-grid commit path —
  // one undo step, protection guard, autosave. Payload: { row, col, values }.
  | "setRangeValues"
  | "fillDown"
  | "fillRight"
  | "fillLeft"
  | "toggleBold"
  | "toggleItalic"
  | "toggleUnderline"
  | "openFormatCells"
  // Excel's Insert / Delete shortcuts: whole rows or columns selected are
  // inserted or deleted at once, any other selection opens the dialog.
  | "insertCells"
  | "deleteCells"
  | SpreadsheetShortcutCommandId;

/**
 * Slice 2: Excel's grid shortcuts. Each id names a command the ribbon, a menu
 * or the tab bar already runs; none has a second implementation here.
 */
export type SpreadsheetShortcutCommandId =
  | "createTable"
  | "insertLink"
  | "toggleAutoFilter"
  | "toggleStrikethrough"
  | "hideRows"
  | "unhideRows"
  | "hideColumns"
  | "unhideColumns"
  | SpreadsheetShortcutNumberFormatCommand
  | "outlineBorder"
  | "removeBorders"
  | "selectEntireRow"
  | "selectEntireColumn"
  | "selectCurrentRegion"
  | "copyFormulaFromAbove"
  | "copyValueFromAbove"
  | "editNote"
  | "newThreadedComment"
  | "insertWorksheet"
  | "nextSheet"
  | "previousSheet"
  | "groupSelection"
  | "ungroupSelection"
  | "openNameManager"
  | "openInsertFunction"
  | "openPasteSpecial"
  | "scrollToActiveCell"
  | "print"
  // F7: Review > Spelling. In Edit mode it checks the entry being edited, so
  // it is consulted before the Edit-mode branches.
  | "checkSpelling";

/** The Ctrl+Shift+~ / ! / @ / # / $ / % / ^ number formats. */
export type SpreadsheetShortcutNumberFormatCommand =
  | "numberFormatGeneral"
  | "numberFormatNumber"
  | "numberFormatTime"
  | "numberFormatDate"
  | "numberFormatCurrency"
  | "numberFormatPercent"
  | "numberFormatScientific";

export interface SpreadsheetKeybinding {
  command: SpreadsheetCommandId;
  /** Where in the host's keydown chain this binding is consulted. */
  stage: "chrome" | "selection";
  /** Matched against event.key, lowercased. Mutually exclusive with `code`. */
  key?: string;
  /** Matched against event.code, for layout-independent digits/symbols. */
  code?: string;
  /** Cmd (mac) or Ctrl: true = required, false = both must be absent. */
  mod: boolean;
  /** Tri-state: true = required, false = must be absent, "any" = ignored. */
  shift?: boolean | "any";
  /**
   * Same tri-state; undeclared means Alt/Option must be absent. A binding that
   * requires Alt never matches AltGr (Windows reports it as Ctrl+Alt): AltGr+V
   * types "@" on a Hungarian keyboard, it is not Ctrl+Alt+V.
   */
  alt?: boolean | "any";
  /**
   * Consulted only while the active cell is NOT in Edit mode: there Space
   * types a space, Ctrl+Backspace deletes a word and PageDown belongs to the
   * entry, as in Excel.
   */
  outsideEdit?: true;
  /**
   * The handler decides preventDefault itself (copy/cut, whose native
   * clipboard behaviour must survive when nothing was snapshotted). The
   * dispatcher must NOT call preventDefault for these.
   */
  managesDefault?: true;
}

/**
 * The bindings the host consults. One flat table so "what does Ctrl+D do"
 * has exactly one answer; the host owns nothing but the handlers.
 */
export const SPREADSHEET_KEYBINDINGS: readonly SpreadsheetKeybinding[] = [
  // Chrome stage — always outranked everything, keeps doing so.
  { command: "undo", stage: "chrome", key: "z", mod: true, shift: false },
  { command: "redo", stage: "chrome", key: "z", mod: true, shift: true },
  { command: "redo", stage: "chrome", key: "y", mod: true, shift: "any" },
  // Keys an Edit-mode, navigation or F2 branch would take before the default
  // branch. Shift+F2 edits the note (plain F2 still enters Edit mode);
  // Ctrl+Shift+F2 starts a threaded comment.
  { command: "editNote", stage: "chrome", key: "f2", mod: false, shift: true },
  {
    command: "newThreadedComment",
    stage: "chrome",
    key: "f2",
    mod: true,
    shift: true,
  },
  {
    command: "openInsertFunction",
    stage: "chrome",
    key: "f3",
    mod: false,
    shift: true,
  },
  { command: "openNameManager", stage: "chrome", key: "f3", mod: true, shift: false },
  { command: "checkSpelling", stage: "chrome", key: "f7", mod: false, shift: false },
  // Excel inserts the new worksheet before the active one.
  {
    command: "insertWorksheet",
    stage: "chrome",
    key: "f11",
    mod: false,
    shift: true,
  },
  // Ctrl+PageDown / Ctrl+PageUp (desktop and Mac); Excel for the web adds Alt
  // because the browser keeps Ctrl+PageDown for its own tabs.
  {
    command: "nextSheet",
    stage: "chrome",
    key: "pagedown",
    mod: true,
    shift: false,
    alt: "any",
    outsideEdit: true,
  },
  {
    command: "previousSheet",
    stage: "chrome",
    key: "pageup",
    mod: true,
    shift: false,
    alt: "any",
    outsideEdit: true,
  },
  {
    command: "groupSelection",
    stage: "chrome",
    key: "arrowright",
    mod: false,
    shift: true,
    alt: true,
    outsideEdit: true,
  },
  {
    command: "ungroupSelection",
    stage: "chrome",
    key: "arrowleft",
    mod: false,
    shift: true,
    alt: true,
    outsideEdit: true,
  },
  // Excel for Mac lists the same key as Ctrl+Delete (its Backspace).
  {
    command: "scrollToActiveCell",
    stage: "chrome",
    key: "backspace",
    mod: true,
    shift: false,
    outsideEdit: true,
  },
  // Shift+Space is printable: it must outrank the locked-cell entry notice.
  {
    command: "selectEntireRow",
    stage: "chrome",
    key: " ",
    mod: false,
    shift: true,
    outsideEdit: true,
  },
  // Selection stage — clipboard first (they outranked the cluster and still
  // do; the matcher returns the FIRST match in table order).
  {
    command: "copySelection",
    stage: "selection",
    key: "c",
    mod: true,
    shift: "any",
    alt: "any",
    managesDefault: true,
  },
  {
    command: "cutSelection",
    stage: "selection",
    key: "x",
    mod: true,
    shift: "any",
    alt: "any",
    managesDefault: true,
  },
  // The format cluster in the keydown default branch.
  { command: "fillDown", stage: "selection", key: "d", mod: true, shift: "any" },
  // Cmd+R is the browser reload; the host preventDefaults on dispatch so the
  // in-app fill wins. Excel has no Fill Left key (the command stays in the
  // Fill menu): Ctrl+L is Create Table there, see below.
  { command: "fillRight", stage: "selection", key: "r", mod: true, shift: "any" },
  { command: "toggleBold", stage: "selection", key: "b", mod: true, shift: false },
  { command: "toggleItalic", stage: "selection", key: "i", mod: true, shift: false },
  {
    command: "toggleUnderline",
    stage: "selection",
    key: "u",
    mod: true,
    shift: false,
  },
  // Matched by code so the main-row "1" resolves regardless of layout; the
  // numpad 1 is "Numpad1" and intentionally not bound (Excel parity).
  {
    command: "openFormatCells",
    stage: "selection",
    code: "Digit1",
    mod: true,
    shift: false,
  },
  // Excel: Ctrl+Shift+Plus (Ctrl+Shift+= on the main row, or Ctrl with the
  // keypad plus) opens Insert, Ctrl+Minus opens Delete; Excel for Mac also
  // takes Cmd+Hyphen. Matched by code so the shifted "=" still resolves, and
  // the host preventDefaults so the browser does not zoom instead.
  { command: "insertCells", stage: "selection", code: "Equal", mod: true, shift: true },
  { command: "insertCells", stage: "selection", code: "NumpadAdd", mod: true, shift: "any" },
  { command: "deleteCells", stage: "selection", code: "Minus", mod: true, shift: false },
  { command: "deleteCells", stage: "selection", code: "NumpadSubtract", mod: true, shift: false },
  // Ctrl+L is Create Table in Excel for Windows and the one Excel for the web
  // documents, because the browser keeps Ctrl+T for a new tab. Ctrl+T stays
  // bound for a host that lets the page see it.
  { command: "createTable", stage: "selection", key: "l", mod: true, shift: false },
  { command: "createTable", stage: "selection", key: "t", mod: true, shift: false },
  { command: "toggleAutoFilter", stage: "selection", key: "l", mod: true, shift: true },
  { command: "insertLink", stage: "selection", key: "k", mod: true, shift: false },
  // Ctrl+P / Cmd+P is File > Print in Excel (Windows, Mac and the web), so the
  // grid prints the workbook rather than letting the browser print the page.
  { command: "print", stage: "selection", key: "p", mod: true, shift: false },
  // The digit row by code, so Ctrl+Shift+4 is "$" on any layout.
  {
    command: "toggleStrikethrough",
    stage: "selection",
    code: "Digit5",
    mod: true,
    shift: false,
  },
  { command: "hideRows", stage: "selection", code: "Digit9", mod: true, shift: false },
  { command: "unhideRows", stage: "selection", code: "Digit9", mod: true, shift: true },
  { command: "hideColumns", stage: "selection", code: "Digit0", mod: true, shift: false },
  { command: "unhideColumns", stage: "selection", code: "Digit0", mod: true, shift: true },
  // Ctrl+Shift+~ ! @ # $ % ^ on a US keyboard.
  ...(
    [
      ["numberFormatGeneral", "Backquote"],
      ["numberFormatNumber", "Digit1"],
      ["numberFormatTime", "Digit2"],
      ["numberFormatDate", "Digit3"],
      ["numberFormatCurrency", "Digit4"],
      ["numberFormatPercent", "Digit5"],
      ["numberFormatScientific", "Digit6"],
    ] as const
  ).map(
    ([command, code]): SpreadsheetKeybinding => ({
      command,
      stage: "selection",
      code,
      mod: true,
      shift: true,
    }),
  ),
  // Ctrl+Shift+& and Ctrl+Shift+_ .
  { command: "outlineBorder", stage: "selection", code: "Digit7", mod: true, shift: true },
  { command: "removeBorders", stage: "selection", code: "Minus", mod: true, shift: true },
  // Ctrl+Shift+* on the main row, or Ctrl with the keypad *.
  {
    command: "selectCurrentRegion",
    stage: "selection",
    code: "Digit8",
    mod: true,
    shift: true,
  },
  {
    command: "selectCurrentRegion",
    stage: "selection",
    code: "NumpadMultiply",
    mod: true,
    shift: "any",
  },
  {
    command: "selectEntireColumn",
    stage: "selection",
    key: " ",
    mod: true,
    shift: false,
    outsideEdit: true,
  },
  // Ctrl+' and Ctrl+Shift+" .
  {
    command: "copyFormulaFromAbove",
    stage: "selection",
    code: "Quote",
    mod: true,
    shift: false,
    outsideEdit: true,
  },
  {
    command: "copyValueFromAbove",
    stage: "selection",
    code: "Quote",
    mod: true,
    shift: true,
    outsideEdit: true,
  },
  // Ctrl+Alt+V (Excel for Mac also Cmd+Option+V). By code: Option+V types
  // another character on a Mac.
  {
    command: "openPasteSpecial",
    stage: "selection",
    code: "KeyV",
    mod: true,
    shift: false,
    alt: true,
  },
];

/**
 * The number format each Ctrl+Shift+digit applies (Microsoft Support,
 * "Keyboard shortcuts in Excel"): General; Number with two decimals, the
 * thousands separator and a minus sign; Time h:mm AM/PM; Date d-mmm-yy;
 * Currency with two decimals and negatives in parentheses; Percentage with no
 * decimals; Scientific with two decimals. These are the SpreadsheetML
 * built-ins 0, 4, 18, 15, 7, 9 and 11. Currency is the one the region
 * changes: zh-CN Excel's built-in 7 is the yuan with a leading minus sign.
 * The symbol is spelled bare, as the Home number gallery and Format Cells
 * write it, so the dialog reopens the cell on Currency.
 */
export function shortcutNumberFormatCode(
  command: SpreadsheetShortcutNumberFormatCommand,
  locale: string,
): string {
  switch (command) {
    case "numberFormatGeneral":
      return "";
    case "numberFormatNumber":
      return "#,##0.00";
    case "numberFormatTime":
      return "h:mm AM/PM";
    case "numberFormatDate":
      return "d-mmm-yy";
    case "numberFormatCurrency":
      return locale.toLowerCase().startsWith("zh")
        ? "¥#,##0.00;¥-#,##0.00"
        : "$#,##0.00_);($#,##0.00)";
    case "numberFormatPercent":
      return "0%";
    case "numberFormatScientific":
      return "0.00E+00";
  }
}

/** The subset of KeyboardEvent the matcher reads — keeps it testable bare. */
export interface KeybindingEventLike {
  key: string;
  code: string;
  metaKey: boolean;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
  /** Present on real events; reports AltGr as "AltGraph". */
  getModifierState?(key: "AltGraph"): boolean;
}

/** What the host knows about the active cell when it asks. */
export interface KeybindingHostState {
  /**
   * Whether the active cell is in Edit mode. `outsideEdit` bindings match only
   * when this is known to be false.
   */
  cellEditing?: boolean;
}

/**
 * The command a key event resolves to at `stage`, or null. Modifier semantics
 * are EXACT: an undeclared alt must be absent, shift follows its tri-state,
 * and mod means meta-or-ctrl exactly as the existing handlers read it.
 */
export function matchSpreadsheetKeybinding(
  event: KeybindingEventLike,
  stage: SpreadsheetKeybinding["stage"],
  bindings: readonly SpreadsheetKeybinding[] = SPREADSHEET_KEYBINDINGS,
  host: KeybindingHostState = {},
): SpreadsheetKeybinding | null {
  const mod = event.metaKey || event.ctrlKey;
  const altGraph =
    event.ctrlKey &&
    event.altKey &&
    !event.metaKey &&
    event.getModifierState?.("AltGraph") === true;
  for (const binding of bindings) {
    if (binding.stage !== stage) continue;
    if (binding.mod !== mod) continue;
    if (binding.outsideEdit && host.cellEditing !== false) continue;
    if (binding.alt === true && altGraph) continue;
    if (binding.alt !== "any" && event.altKey !== (binding.alt ?? false))
      continue;
    if (binding.shift !== "any" && event.shiftKey !== (binding.shift ?? false))
      continue;
    if (binding.code !== undefined) {
      if (event.code !== binding.code) continue;
    } else if (binding.key !== event.key.toLowerCase()) {
      continue;
    }
    return binding;
  }
  return null;
}

/**
 * A handler. Returning `false` means it did not run (the control is not
 * offered here, nothing to act on); anything else counts as ran.
 */
export type SpreadsheetCommandHandler = (payload?: unknown) => void | boolean;

/**
 * A per-window registry: ids to live handlers. The host re-registers every
 * render so a handler always closes over current state; dispatch returns
 * whether anything ran, which is what the caller's preventDefault keys off.
 */
export interface SpreadsheetCommandRegistry {
  register(id: SpreadsheetCommandId, run: SpreadsheetCommandHandler): void;
  /** `payload` is opaque to the registry; a keyboard-dispatched command gets
   *  the originating event (copy/cut read caret state off it). */
  dispatch(id: SpreadsheetCommandId, payload?: unknown): boolean;
  registered(): readonly SpreadsheetCommandId[];
}

export function createSpreadsheetCommandRegistry(): SpreadsheetCommandRegistry {
  const handlers = new Map<SpreadsheetCommandId, SpreadsheetCommandHandler>();
  return {
    register(id, run) {
      handlers.set(id, run);
    },
    dispatch(id, payload) {
      const run = handlers.get(id);
      if (!run) return false;
      return run(payload) !== false;
    },
    registered() {
      return [...handlers.keys()];
    },
  };
}

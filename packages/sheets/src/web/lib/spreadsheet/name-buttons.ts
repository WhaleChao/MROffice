/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Formulas > Defined Names: Define Name and the Name Manager's New / Edit /
 * Delete as pure workbook transactions. Moved out of SpreadsheetWindow so the
 * ribbon and the spreadsheet editor commands run the same code; behaviour is
 * unchanged. The window keeps its dialog authority, notices and focus.
 *
 * Workbook names are mirrored into every sheet's `namedRanges`; a worksheet
 * name lives on its own sheet only (scopeSheetId).
 */

import type { CellMap } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  addNamedRange,
  isCellRefString,
  isRangeRefString,
  isValidName,
  isValidRef,
  namedRangeIdentity,
  normalizeNamedRange,
  normalizeNamedRanges,
  qualifyWorkbookNamedRangeRef,
  removeNamedRange,
  renameDefinedNameReferences,
  type NamedRange,
  type NamedRangeScope,
} from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { rewriteSheetFormulaReferences } from "./sheet-formula-references";
import { Parser, stripFormulaParameterPrefix, tokenize, type Node as FormulaNode } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { nameRefShownFrom, typedNameRefHasRelativeAxis } from "@mrnote/core/web/lib/spreadsheet/name-relative-refs.ts";

/** The parts of a sheet the name buttons read and write. */
export interface NameButtonSheet {
  id: string;
  name: string;
  cells: CellMap;
  namedRanges?: NamedRange[];
}

type NamesMutation<S> = (current: readonly NamedRange[], candidate: S) => readonly NamedRange[];

export function namedRangeValuesEqual(left: NamedRange, right: NamedRange): boolean {
  return (
    left.name === right.name &&
    left.ref === right.ref &&
    (left.scope ?? "workbook") === (right.scope ?? "workbook") &&
    (left.scopeSheetId ?? "") === (right.scopeSheetId ?? "") &&
    (left.nativeBlocker ?? "") === (right.nativeBlocker ?? "") &&
    (left.nativeSourceName ?? "") === (right.nativeSourceName ?? "") &&
    (left.nativeSourceScopeSheetId ?? "") ===
      (right.nativeSourceScopeSheetId ?? "") &&
    (left.hidden === true) === (right.hidden === true) &&
    (left.comment ?? "") === (right.comment ?? "")
  );
}

/**
 * Build the workbook Name Manager catalog from the sheet-local projection.
 * Workbook names are mirrored into every sheet; they are emitted once only
 * when every mirror agrees. A divergent mirror is retained read-only instead
 * of granting a mutation authority from an arbitrary copy.
 */
export function workbookNamedRangeCatalog(
  // Each sheet's names are normalized here, so a stored sheet's raw
  // `namedRanges` is as good as a typed one.
  workbook: readonly { readonly id: string; readonly namedRanges?: unknown }[],
): NamedRange[] {
  const catalog = new Map<string, NamedRange>();
  for (const candidate of workbook) {
    for (const entry of normalizeNamedRanges(candidate.namedRanges)) {
      if (
        entry.scope === "worksheet" &&
        entry.scopeSheetId !== candidate.id
      ) {
        continue;
      }
      const identity = namedRangeIdentity(entry);
      const previous = catalog.get(identity);
      if (!previous) {
        catalog.set(identity, entry);
        continue;
      }
      if (!namedRangeValuesEqual(previous, entry)) {
        catalog.set(identity, {
          ...previous,
          nativeBlocker: "ambiguous-identity",
        });
      }
    }
  }
  return [...catalog.values()];
}

/**
 * Whether renaming `source` (a workbook name) to `nextName` would hand its
 * unqualified references on some sheet to that sheet's own local name of
 * the same spelling: the formula text would read the local name after the
 * rewrite, and silently change value (B509). A worksheet name takes its own
 * sheet's references first either way, so only a workbook source is at risk.
 */
/** The names a formula's LET variables and LAMBDA parameters bind, lower-cased. */
export function formulaLocalNames(formula: string): Set<string> {
  const names = new Set<string>();
  if (!/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(formula)) return names;
  const tokens = tokenize(formula.replace(/^=/, ""));
  if (!tokens) return names;
  let root: FormulaNode;
  try {
    root = new Parser(tokens).parse();
  } catch {
    return names;
  }
  const walk = (node: FormulaNode | undefined): void => {
    if (!node) return;
    if (node.kind === "call") {
      const fn = node.name.toUpperCase().replace(/^_XLFN\./, "");
      // LET(name1, value1, ..., calculation); LAMBDA(param1, ..., calculation).
      const bound = fn === "LET" ? node.args.filter((_, i) => i % 2 === 0 && i < node.args.length - 1)
        : fn === "LAMBDA" ? node.args.slice(0, -1) : [];
      for (const arg of bound) if (arg.kind === "name") names.add(stripFormulaParameterPrefix(arg.name).toLowerCase());
    }
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) for (const item of value) (Array.isArray(item) ? item : [item]).forEach((n) => n && typeof n === "object" && "kind" in n && walk(n as FormulaNode));
      else if (value && typeof value === "object" && "kind" in value) walk(value as FormulaNode);
    }
  };
  walk(root);
  return names;
}

/**
 * Whether renaming `source` to `nextName` would hand one of its references to
 * a LET variable or LAMBDA parameter of that spelling in the same formula:
 * the local binding wins inside its formula, so the formula would silently
 * read the local value instead of the name (live audit S-164). Such a rename
 * is refused, as a capture by a sheet's local name is (B509).
 */
export function renameCapturedByFormulaLocals<S extends NameButtonSheet>(
  workbook: readonly S[],
  source: NamedRange,
  nextName: string,
): boolean {
  if (source.name.toLowerCase() === nextName.toLowerCase()) return false;
  const next = nextName.toLowerCase();
  const probe = "\u0000";
  const captures = (text: unknown): boolean =>
    typeof text === "string" &&
    text.startsWith("=") &&
    renameDefinedNameReferences(text, source.name, probe, undefined, true) !== text &&
    formulaLocalNames(text).has(next);
  return workbook.some((sheet) =>
    Object.values(sheet.cells).some((cell) => captures(cell.value) || captures(cell.formula)) ||
    normalizeNamedRanges(sheet.namedRanges).some((entry) => captures(entry.ref.startsWith("=") ? entry.ref : `=${entry.ref}`)),
  );
}

export function renameCapturesReferences<S extends NameButtonSheet>(
  workbook: readonly S[],
  source: NamedRange,
  nextName: string,
): boolean {
  if ((source.scope ?? "workbook") !== "workbook") return false;
  if (source.name.toLowerCase() === nextName.toLowerCase()) return false;
  const probe = "\u0000";
  return workbook.some((sheet) => {
    const shadowed = normalizeNamedRanges(sheet.namedRanges).some(
      (entry) =>
        !entry.nativeBlocker &&
        entry.scope === "worksheet" &&
        entry.scopeSheetId === sheet.id &&
        entry.name.toLowerCase() === nextName.toLowerCase(),
    );
    if (!shadowed) return false;
    return Object.values(sheet.cells).some((cell) => {
      for (const text of [cell.value, cell.formula]) {
        if (typeof text !== "string" || !text.startsWith("=")) continue;
        if (renameDefinedNameReferences(text, source.name, probe, undefined, true) !== text) {
          return true;
        }
      }
      return false;
    });
  });
}

export function rewriteDefinedNameInSheet<S extends NameButtonSheet>(
  sheet: S,
  source: NamedRange,
  nextName: string,
  sourceSheetName?: string,
): S {
  if (source.name.toLowerCase() === nextName.toLowerCase()) return sheet;
  const sourceScope = source.scope ?? "workbook";
  const owningLocalSheet =
    sourceScope === "worksheet" && source.scopeSheetId === sheet.id;
  const localShadow = normalizeNamedRanges(sheet.namedRanges).some(
    (entry) =>
      !entry.nativeBlocker &&
      entry.scope === "worksheet" &&
      entry.scopeSheetId === sheet.id &&
      entry.name.toLowerCase() === source.name.toLowerCase(),
  );
  const rewriteUnqualified =
    sourceScope === "worksheet" ? owningLocalSheet : !localShadow;
  const qualifierSheetName =
    sourceScope === "worksheet" ? sourceSheetName : undefined;
  let cellsChanged = false;
  const cells: CellMap = {};
  for (const [key, cell] of Object.entries(sheet.cells)) {
    let nextCell = cell;
    if (typeof cell.value === "string" && cell.value.startsWith("=")) {
      const value = renameDefinedNameReferences(
        cell.value,
        source.name,
        nextName,
        qualifierSheetName,
        rewriteUnqualified,
      );
      if (value !== cell.value) nextCell = { ...nextCell, value };
    }
    if (typeof cell.formula === "string" && cell.formula.startsWith("=")) {
      const formula = renameDefinedNameReferences(
        cell.formula,
        source.name,
        nextName,
        qualifierSheetName,
        rewriteUnqualified,
      );
      if (formula !== cell.formula) nextCell = { ...nextCell, formula };
    }
    if (nextCell !== cell) cellsChanged = true;
    cells[key] = nextCell;
  }
  let next = cellsChanged ? { ...sheet, cells } : sheet;
  if (sheet.namedRanges?.length) {
    let namesChanged = false;
    const namedRanges = sheet.namedRanges.map((entry) => {
      if (entry.nativeBlocker) return entry;
      const entryRewriteUnqualified =
        sourceScope === "worksheet"
          ? entry.scope === "worksheet" &&
            entry.scopeSheetId === source.scopeSheetId
          : entry.scope !== "worksheet" || !localShadow;
      const ref = renameDefinedNameReferences(
        entry.ref,
        source.name,
        nextName,
        qualifierSheetName,
        entryRewriteUnqualified,
      );
      if (ref === entry.ref) return entry;
      namesChanged = true;
      return { ...entry, ref };
    });
    if (namesChanged) next = { ...next, namedRanges };
  }
  // Every other formula holder of the sheet: conditional-format rules,
  // validations, table recipes, sparklines and charts (live audit B132).
  return rewriteSheetFormulaReferences(
    next,
    (formula) =>
      renameDefinedNameReferences(
        formula,
        source.name,
        nextName,
        qualifierSheetName,
        rewriteUnqualified,
      ),
    { skipNamedRanges: true },
  );
}

/** Every sheet with its own names run through `mutate`; an empty list is dropped. */
export function namesMutatedSheets<S extends Pick<NameButtonSheet, "namedRanges">>(
  workbook: readonly S[],
  mutate: NamesMutation<S>,
): S[] {
  return workbook.map((candidate) => {
    const nextNames = mutate(normalizeNamedRanges(candidate.namedRanges), candidate);
    const next: S = { ...candidate };
    if (nextNames.length > 0) next.namedRanges = [...nextNames];
    else delete next.namedRanges;
    return next;
  });
}

export interface DefineNamePlan {
  name: string;
  /** The reference as stored: a cell or range absolute and sheet-qualified. */
  ref: string;
  scope: NamedRangeScope;
  scopeSheetId?: string;
}

/**
 * Define Name's OK: what the name will store, or null when the name or the
 * "Refers to" text is not valid. An empty `explicitRef` names the selection;
 * a leading "=" (Excel's refers-to form) is stripped, so a named formula or
 * LAMBDA can be defined too.
 */
export function defineNamePlan(
  activeSheet: Pick<NameButtonSheet, "id" | "name">,
  rawName: string,
  explicitRef: string | undefined,
  selection: { r1: number; c1: number; r2: number; c2: number },
  explicitScope?: NamedRangeScope,
  /**
   * The active cell a typed relative reference is relative to (Excel's New
   * Name); without it a typed relative reference is refused.
   */
  activeCell?: { row: number; col: number },
): DefineNamePlan | null {
  const name = rawName.trim();
  if (!isValidName(name)) return null;
  const typed = (explicitRef ?? "").trim();
  let ref: string;
  if (typed) {
    ref = typed.startsWith("=") ? typed.slice(1).trim() : typed;
  } else {
    const a1 = `${columnLetters(selection.c1)}${selection.r1 + 1}`;
    const single = selection.r1 === selection.r2 && selection.c1 === selection.c2;
    ref = single ? a1 : `${a1}:${columnLetters(selection.c2)}${selection.r2 + 1}`;
  }
  if (!isValidRef(ref)) return null;
  const scope = explicitScope ?? "workbook";
  // A cell or range is stored sheet-qualified at either scope, as Excel
  // stores it (a selection absolute, a typed relative axis as its offset from
  // the active cell); a worksheet formula with no relative reference keeps
  // its text.
  const storedRef =
    scope === "workbook" ||
    isCellRefString(ref) ||
    isRangeRefString(ref) ||
    (typed && typedNameRefHasRelativeAxis(ref))
      ? qualifyWorkbookNamedRangeRef(ref, activeSheet.name, {
          authored: Boolean(typed),
          ...(typed && activeCell ? { base: activeCell } : {}),
        })
      : ref;
  if (!storedRef) return null;
  return {
    name,
    ref: storedRef,
    scope,
    ...(scope === "worksheet" ? { scopeSheetId: activeSheet.id } : {}),
  };
}

/** The names mutation Define Name's OK runs (throws on a read-only name). */
export function definedNameMutation<S extends Pick<NameButtonSheet, "id">>(
  plan: DefineNamePlan,
): NamesMutation<S> {
  const context =
    plan.scope === "worksheet"
      ? ({ scope: "worksheet", scopeSheetId: plan.scopeSheetId } as const)
      : ({ scope: "workbook" } as const);
  return (current, candidate) => {
    if (context.scope === "worksheet" && candidate.id !== context.scopeSheetId) {
      return current;
    }
    return addNamedRange(current, plan.name, plan.ref, context);
  };
}

/** The names mutation that removes one defined name (Define Name > Delete). */
export function deletedNameMutation<S extends Pick<NameButtonSheet, "id">>(
  entry: NamedRange,
): NamesMutation<S> {
  const context =
    entry.scope === "worksheet"
      ? ({ scope: "worksheet", scopeSheetId: entry.scopeSheetId } as const)
      : ({ scope: "workbook" } as const);
  return (current, candidate) => {
    if (context.scope === "worksheet" && candidate.id !== context.scopeSheetId) {
      return current;
    }
    return removeNamedRange(current, entry.name, context);
  };
}

/** A Name Manager write: New / Edit's OK, or Delete. */
export type NameManagerRequest =
  | { kind: "delete"; source: NamedRange | null }
  | {
      kind: "apply";
      action: "create" | "edit";
      source: NamedRange | null;
      name: string;
      ref: string;
      scope: NamedRangeScope;
      /**
       * The New / Edit Name box's Comment (XLS-A066); undefined keeps the
       * source's, an empty one clears it.
       */
      comment?: string;
      /**
       * The active cell: a typed relative reference is relative to it, and
       * Edit showed the source's relative reference from it (an unchanged
       * Refers to keeps the stored text).
       */
      base?: { row: number; col: number };
    };

/**
 * The workbook after a Name Manager write, made from `activeSheet` (where a
 * bare cell reference of a workbook name points). "protected" when a sheet
 * the name lives on is protected; null when the write is refused (stale or
 * read-only source, invalid name or reference, duplicate, nothing changed).
 */
export function nameManagerSheets<S extends NameButtonSheet>(
  workbook: readonly S[],
  activeSheet: S,
  request: NameManagerRequest,
  isProtected: (sheet: S) => boolean,
): S[] | "protected" | null {
  const catalog = workbookNamedRangeCatalog(workbook);
  const source = request.source;
  const currentSource = source
    ? catalog.find(
        (entry) => namedRangeIdentity(entry) === namedRangeIdentity(source),
      ) ?? null
    : null;
  if (
    (source &&
      (!currentSource ||
        !namedRangeValuesEqual(currentSource, source) ||
        currentSource.nativeBlocker)) ||
    (!source && request.kind === "delete")
  ) {
    return null;
  }
  const sourceScope =
    source?.scope ?? (request.kind === "apply" ? request.scope : "workbook");
  const affectedLocalSheetId =
    source?.scopeSheetId ??
    (request.kind === "apply" && request.scope === "worksheet"
      ? activeSheet.id
      : undefined);
  const affectedSheetIds = new Set(
    sourceScope === "worksheet" && affectedLocalSheetId
      ? [affectedLocalSheetId]
      : workbook.map((candidate) => candidate.id),
  );
  if (
    workbook.some(
      (candidate) => affectedSheetIds.has(candidate.id) && isProtected(candidate),
    )
  ) {
    return "protected";
  }
  if (request.kind === "delete") {
    if (!source) return null;
    const sourceIdentity = namedRangeIdentity(source);
    return workbook.map((candidate) => {
      const names = normalizeNamedRanges(candidate.namedRanges).filter(
        (entry) => namedRangeIdentity(entry) !== sourceIdentity,
      );
      const next: S = { ...candidate };
      if (names.length) next.namedRanges = names;
      else delete next.namedRanges;
      return next;
    });
  }

  const name = request.name.trim();
  let ref = request.ref.trim();
  if (ref.startsWith("=")) ref = ref.slice(1).trim();
  if (!isValidName(name) || !isValidRef(ref)) return null;
  if (request.action === "edit" && !source) return null;
  const scope = source?.scope ?? request.scope;
  if (source && scope !== request.scope) return null;
  const scopeSheetId =
    scope === "worksheet" ? source?.scopeSheetId ?? activeSheet.id : undefined;
  const shownSourceRef = source
    ? nameRefShownFrom(source.ref, request.base).replace(/^=\s*/, "").trim()
    : null;
  if (shownSourceRef !== null && shownSourceRef === ref) {
    // Refers to as Edit showed it: the stored reference stays as it is.
    ref = source!.ref;
  } else if (
    isCellRefString(ref) ||
    isRangeRefString(ref) ||
    typedNameRefHasRelativeAxis(ref)
  ) {
    // Stored the way Excel stores a name's reference: sheet-qualified -- the
    // active sheet for a workbook name, the owning sheet for a worksheet name
    // (where a bare ref already resolved) -- a relative axis as its offset
    // from the active cell.
    const qualifierSheetName =
      scope === "worksheet"
        ? workbook.find((candidate) => candidate.id === scopeSheetId)?.name
        : activeSheet.name;
    const qualified = qualifierSheetName
      ? qualifyWorkbookNamedRangeRef(ref, qualifierSheetName, {
          authored: true,
          ...(request.base ? { base: request.base } : {}),
        })
      : null;
    if (!qualified) return null;
    ref = qualified;
  }
  // A rename that would make a formula read another name: refused whole,
  // the model untouched (B509).
  if (source && (renameCapturesReferences(workbook, source, name) || renameCapturedByFormulaLocals(workbook, source, name))) return null;
  if (source && source.name.toLowerCase() !== name.toLowerCase()) {
    const sourceSheetName =
      source.scope === "worksheet"
        ? workbook.find((candidate) => candidate.id === source.scopeSheetId)?.name
        : undefined;
    ref = renameDefinedNameReferences(ref, source.name, name, sourceSheetName, true);
  }
  const target = normalizeNamedRange({
    ...(source ?? {}),
    name,
    ref,
    scope,
    ...(scopeSheetId ? { scopeSheetId } : {}),
    comment: request.comment ?? source?.comment,
  });
  if (!target) return null;
  const sourceIdentity = source ? namedRangeIdentity(source) : null;
  const targetIdentity = namedRangeIdentity(target);
  if (
    catalog.some(
      (entry) =>
        namedRangeIdentity(entry) === targetIdentity &&
        namedRangeIdentity(entry) !== sourceIdentity,
    )
  ) {
    return null;
  }
  if (source && namedRangeValuesEqual(source, target)) return null;
  const sourceSheetName =
    source?.scope === "worksheet"
      ? workbook.find((candidate) => candidate.id === source.scopeSheetId)?.name
      : undefined;
  return workbook.map((candidate) => {
    const rewritten = source
      ? rewriteDefinedNameInSheet(candidate, source, target.name, sourceSheetName)
      : candidate;
    const names = normalizeNamedRanges(rewritten.namedRanges);
    let nextNames = names;
    if (sourceIdentity) {
      nextNames = names.map((entry) =>
        namedRangeIdentity(entry) === sourceIdentity ? target : entry,
      );
    } else if (
      scope === "workbook" ||
      (scope === "worksheet" && candidate.id === scopeSheetId)
    ) {
      nextNames = [...names, target];
    }
    const next: S = { ...rewritten };
    if (nextNames.length) next.namedRanges = nextNames;
    else delete next.namedRanges;
    return next;
  });
}

function columnLetters(index: number): string {
  let value = index + 1;
  let text = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    text = String.fromCharCode(65 + remainder) + text;
    value = Math.floor((value - 1) / 26);
  }
  return text;
}

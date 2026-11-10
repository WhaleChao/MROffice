/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The `tools` engine's wire: the small stateless editor algorithms that run on
 * the server (`engine-host/engines/tools.ts`). Types and plain constants only.
 *
 * - `shapes.merge`: Merge Shapes (union / combine / fragment / intersect /
 *   subtract) of two deck shapes, `lib/presentation-shape-merge.ts`.
 * - `document.compare`: Compare Documents' word-level redline of two
 *   documents, `lib/document/compare.ts`; with `combine`, Combine Documents'
 *   merge of two copies' revisions, `lib/document/combine.ts`.
 * - `document.draftDiff`: the page-draft review's alignment of a page and its
 *   rewrite, `lib/page-draft-diff-engine.ts`.
 * - `sheet.flashFill`: Flash Fill's string-program synthesis,
 *   `lib/spreadsheet/flash-fill.ts`.
 * - `document.answerCommand`: an AI answer and the member's instruction read
 *   into a native document command,
 *   `components/console/editor/native-document/documentAnswerCommand.ts`.
 */

import type {
  EditableAnswerCommandContext,
  EditableAnswerCommandPayload,
  NativeDocumentCommand,
} from "../../components/console/editor/native-document/documentCommands";
import type { DocumentCombineFormatSource, DocumentCombineOutcome } from "@/lib/document/combine";
import type { DocumentCompareOutcome, PMNodeJSON } from "@/lib/document/compare";
import type { DraftRewriteDiff, DraftTextBlock } from "@/lib/page-draft-diff";
import type { MergeableShape, MergedShapeResult } from "../presentation-shapes";

export type ToolsMergeShapesOperation = "union" | "combine" | "fragment" | "intersect" | "subtract";

export interface ToolsShapesMergeRequest {
  op: ToolsMergeShapesOperation;
  /** The shape further back; the results take its paint. Geometry fields only
   * (`mergeableShapeFields`). */
  first: MergeableShape;
  second: MergeableShape;
}

export interface ToolsShapesMergeResult {
  /** In order (Fragment: A minus B, A and B, B minus A); empty when nothing is left. */
  results: MergedShapeResult[];
}

export interface ToolsDocumentCompareRequest {
  /** The older version (the base) and the live document, as editor JSON. */
  base: PMNodeJSON;
  current: PMNodeJSON;
  /** The redline marks' time: the asking browser's clock (ms since epoch). */
  timestamp: number;
  /**
   * Combine instead of Compare (Word's Combine): `base` is the original and
   * `current` the revised copy; both may carry tracked changes.
   */
  combine?: ToolsDocumentCombineOptions;
  /**
   * The edit an author made with Track Changes on, recorded as Word records
   * it (the Copilot's own way, `documentCopilotTracked`): the words that
   * changed as insertions and deletions, formatting as formatting revisions,
   * table rows as row revisions, and a change read no other way as its old
   * blocks struck and its new ones inserted. `author` names the revisions.
   * It takes precedence over `combine`, which an engine without it reads.
   */
  track?: { author: string | null };
}

/** `document.compare` with `track`: the current document with the edit as
 * tracked revisions. `recorded` is false only when nothing changed; `whole`
 * when the blocks were struck and inserted whole (`reason`: why). */
export interface ToolsDocumentTrackedOutcome {
  outcome: "tracked";
  doc: PMNodeJSON;
  recorded: boolean;
  whole: boolean;
  reason: string | null;
}

/** Combine's choices (Word's Combine dialog and its formatting prompt). */
export interface ToolsDocumentCombineOptions {
  /** "Label unmarked changes with" (null: anonymous). At most 256 characters. */
  author: string | null;
  /** Whose formatting to keep where the copies differ; null asks first. */
  formatFrom: DocumentCombineFormatSource | null;
}

export const TOOLS_COMBINE_MAX_AUTHOR_LENGTH = 256;

export interface ToolsDraftDiffRequest {
  /** The page as it reads now and the draft, cut into blocks (`lib/page-draft-diff.ts`). */
  before: DraftTextBlock[];
  after: DraftTextBlock[];
}

/** Flash Fill's bounds: the examples and targets of one block of data. */
export const TOOLS_FLASH_FILL_MAX_EXAMPLES = 20_000;
export const TOOLS_FLASH_FILL_MAX_TARGETS = 200_000;
/** The cells to the left of the column (a row's inputs). */
export const TOOLS_FLASH_FILL_MAX_INPUTS = 16_384;

export interface ToolsFlashFillRequest {
  /** The column's filled cells: each one's inputs (the cells to its left, as
   * shown) and its own text. */
  examples: { inputs: string[]; output: string }[];
  /** Each empty target row's inputs. */
  targets: string[][];
}

export interface ToolsAnswerCommandRequest {
  /** The answer (its markdown, the action, the instruction it answers). */
  payload: EditableAnswerCommandPayload;
  /** The live edit context at apply time (page type, selection, block ids). */
  context: EditableAnswerCommandContext | null;
}

export interface ToolsEngineMethods {
  "shapes.merge": { input: ToolsShapesMergeRequest; output: ToolsShapesMergeResult };
  /** The redline and its change count, or why the two were not compared. */
  "document.compare": {
    input: ToolsDocumentCompareRequest;
    /** Compare's outcome; with `combine` in the request, Combine's. With
     * `track` (the server's Agent proposals ask for it), a
     * `ToolsDocumentTrackedOutcome` instead. */
    output: DocumentCompareOutcome | DocumentCombineOutcome;
  };
  /** Null when there is nothing to compare against or it is over the review's budget. */
  "document.draftDiff": { input: ToolsDraftDiffRequest; output: { diff: DraftRewriteDiff | null } };
  /** Null: no pattern; else one value per target (null: that row stays blank). */
  "sheet.flashFill": { input: ToolsFlashFillRequest; output: { resolved: (string | null)[] | null } };
  /** The command the browser validates and applies. */
  "document.answerCommand": { input: ToolsAnswerCommandRequest; output: { command: NativeDocumentCommand } };
}

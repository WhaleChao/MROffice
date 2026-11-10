"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Flag-gated ({@link ENABLE_CODE_DEBUG_V1}) debugging panel for the code editor.
 *
 * This component is mounted by {@link CodeWindow} ONLY when the flag is ON, so the
 * OFF path renders nothing (zero DOM / state / listeners — byte-for-byte today).
 *
 * It is a SELF-CONTAINED breakpoint + debug-session UI: the user manages line
 * breakpoints (add the editor's current line, or any line; remove via chip), runs
 * "Debug", and — when execution pauses at a breakpoint — sees the current line, a
 * best-effort snapshot of in-scope locals, and Continue / Step controls. JS runs
 * through {@link runJavaScriptDebug} (live pause on the isolated worker channel);
 * TypeScript is type-stripped first via the existing ts-runner transpile; Python
 * runs through {@link runPythonDebug} (record-and-replay on the kernel — same
 * Continue/Step surface, and breakpoints work inside function bodies).
 *
 * The panel NEVER edits the editor document or touches the normal Run path. An
 * optional host callback only requests selection/scroll navigation for current
 * lines, pauses, breakpoints, and stack frames.
 */

import {
  useCallback,
  useDeferredValue,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
} from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import {
  MAX_LOGPOINT_MESSAGE_CHARS,
  appendLogpointLine,
  formatLogpoint,
  isLogpointMessage,
  type LogpointLine,
} from "../../../lib/code-runner/debug-logpoint";
import {
  addDebugWatch,
  isWatchableName,
  MAX_DEBUG_WATCHES,
  MAX_DEBUG_WATCH_NAME_CHARS,
  removeDebugWatch,
  renameDebugWatch,
  resolveDebugWatches,
  type ResolvedDebugWatch,
} from "../../../lib/code-runner/debug-watch";
import { Loader2 } from "lucide-react";
import {
  ArrowDownToLine,
  ArrowUpFromLine,
  Bug,
  ChevronRight,
  Circle,
  Diamond,
  Download,
  RotateCcw,
  Search,
  SkipForward,
  Upload,
  X,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import {
  officeDialogClassName,
  officeDialogParts,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

import { downloadTextFile } from "@mrnote/core/web/lib/download-file.ts";
import {
  runJavaScriptDebug,
  type RunResult,
  type WorkerDebugPausedMessage,
} from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import { boundExecutionResult } from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";
import {
  appendDebugSnapshotEvaluation,
  MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS,
  normalizeDebugSnapshotExpression,
  type DebugSnapshotEvaluationEntry,
  type DebugSnapshotEvaluationReceipt,
} from "@mrnote/core/web/lib/code-runner/debug-snapshot-evaluate.ts";
import { isValidBreakpointCondition } from "@mrnote/core/web/lib/code-runner/js-debug.ts";
import { unbindableBreakpointLines } from "../../../lib/code-runner/debug-breakpoint-binding";
import {
  isValidPythonBreakpointCondition,
  parseHitCondition,
  runPythonDebug,
  type PythonDebugPausedMessage,
} from "../../../lib/code-runner/python-debug";
import { transpileTypeScript } from "../../../lib/code-runner/ts-runner";
import {
  CODE_PAGE_DEBUG_TIMELINE_PAUSE_LIMIT,
  CODE_PAGE_DEBUG_TIMELINE_STACK_LIMIT,
  CODE_DEBUG_CONFIGURATION_BACKUP_FILENAME,
  EMPTY_CODE_DEBUG_CONFIGURATION,
  MAX_CODE_DEBUG_CONFIGURATION_BACKUP_BYTES,
  MAX_CODE_PAGE_BREAKPOINT_LINE,
  MAX_CODE_PAGE_BREAKPOINT_EXPRESSION_CHARS,
  createCodeDebugConfigurationBackup,
  normalizeCodeDebugConfiguration,
  parseCodeDebugConfigurationBackup,
  type CodeDebugConfiguration,
  type PersistedCodeDebugPause,
  type PersistedCodeDebugTimelineEntry,
} from "../../../lib/code-runner/code-page-debug-persistence";
import { CodeDebugTimelinePanel } from "./CodeRunTimelinePanel";
import AnsiText from "./code-output/AnsiText";

import styles from "./CodeSidePanels.module.css";

const CODE_DEBUG_BREAKPOINT_INITIAL_LIMIT = 40;

export interface CodeDebugPanelProps {
  /**
   * The HOST is already executing something (a file run, or a console
   * submission). The panel's own `running` only covers a debug session, so
   * without this the Debug button would start a second execution in the same
   * kernel while the first is still in it.
   */
  hostBusy?: boolean;
  /**
   * Atomically acquire the host page's execution slot. A null result means a
   * duplicate view won the race and Debug must remain fail-closed; the returned
   * callback releases exactly this session's claim.
   */
  acquireExecution?: () => (() => void) | null;
  /**
   * Stable identity of the page/editor surface that owns this panel. CodeWindow
   * keeps the panel mounted across ribbon navigation and page replacement, so
   * source text alone is not a sufficient owner (two pages may be identical).
   */
  viewStateKey?: string;
  /** Current editor source (read-only here). */
  content: string;
  /**
   * Optional transient selected-test source. The authored `content` remains
   * the line/configuration authority; only this exact source is dispatched.
   */
  testExecution?: CodeDebugTestExecution;
  /** Current language id ("javascript"/"typescript"/"python" are debuggable). */
  language: string;
  /**
   * The editor's current 1-based caret line, if the host can supply it, so
   * "Add breakpoint at current line" targets the right line. Optional: when
   * absent the user adds a breakpoint by typing a line number.
   */
  currentLine?: number;
  /** Kernel namespace for Python debug runs (the page id). */
  pythonNamespace?: string;
  /** Page-scoped JavaScript iframe owner shared with normal Run and Console. */
  sandboxOwner?: string;
  /** Pre-supplied stdin for Python input() calls, same as a normal run. */
  stdin?: string;
  /**
   * Requests host/editor navigation. Passive pauses use `focusEditor=false`;
   * explicit clicks on a breakpoint, paused line, or stack frame use true.
   */
  onNavigateToLine?: (line: number, focusEditor: boolean) => void;
  /** Mirrors the active paused statement into editor chrome; null clears it. */
  onPausedLineChange?: (line: number | null) => void;
  /**
   * Optional controlled debugger configuration. CodeWindow supplies this so
   * authored breakpoints and watches survive page save/reopen; standalone
   * panel tests may omit it.
   */
  configuration?: CodeDebugConfiguration;
  onConfigurationChange?: (configuration: CodeDebugConfiguration) => void;
  /** Debug may inspect source, but page-owned breakpoints/watches stay inert. */
  configurationReadOnly?: boolean;
  /** Monotonic host request for a real debugger action. */
  commandRequest?: CodeDebugCommandRequest;
  /** Consumes one exact host request, including an owner-mismatched request. */
  onCommandRequestHandled?: (requestId: number) => void;
  /** Reports live debugger availability/session state to host chrome. */
  onStateChange?: (state: CodeDebugState) => void;
  /** Durable structural history; raw runtime values never enter these entries. */
  persistedTimeline?: readonly PersistedCodeDebugTimelineEntry[];
  currentSourceFingerprint?: string | null;
  selectedTimelineIndex?: number | null;
  onTimelineSelect?: (index: number) => void;
  onTimelineDebugAgain?: (index: number) => void;
  onTimelineClear?: () => void;
  /** Publishes only a normally settled, current, non-test session summary. */
  onSessionSettled?: (settlement: CodeDebugSessionSettlement) => void;
}

function boundedDebugFailure(error: unknown, fallback: string): RunResult {
  let message = fallback;
  try {
    if (error instanceof Error && error.message.trim()) {
      message = error.message;
    } else if (typeof error === "string" && error.trim()) {
      message = error;
    }
  } catch {
    // Unknown/proxied values remain the localized fallback.
  }
  return boundExecutionResult({ logs: [], error: message });
}

export interface CodeDebugTestExecution {
  readonly content: string;
  readonly identity: string;
  readonly testName: string;
  readonly stopLine: number;
}

export type CodeDebugCommand =
  | "start"
  | "restart"
  | "stop"
  | "addBreakpointCurrent"
  | "runToCursor"
  | "continue"
  | "step"
  | "stepOver"
  | "stepInto"
  | "stepOut";

export interface CodeDebugCommandRequest {
  requestId: number;
  command: CodeDebugCommand;
  /** Exact page/editor surface that was visible when the command was invoked. */
  viewStateKey?: string;
}

export interface CodeDebugState {
  debuggable: boolean;
  running: boolean;
  paused: boolean;
  pausedLine: number | null;
  currentLine: number | null;
  breakpointCount: number;
  supportsDepthStepping: boolean;
}

export interface CodeDebugSessionSettlement {
  readonly content: string;
  readonly language: string;
  readonly viewStateKey: string | undefined;
  readonly completedAt: number;
  readonly outcome: "success" | "error";
  readonly pauses: readonly PersistedCodeDebugPause[];
}

/** Languages this panel can debug (JS + type-stripped TS + recorded Python). */
export function isDebuggableLanguage(language: string): boolean {
  return (
    language === "javascript" ||
    language === "typescript" ||
    language === "python"
  );
}

/**
 * The js and python controllers share this driving surface. The depth-based
 * step variants + frame inspection are Python-only (record-and-replay) — the JS
 * controller has none, so they are optional and feature-detected at call time.
 */
interface DebugController {
  readonly result: Promise<RunResult>;
  continue(): void;
  step(): void;
  evaluate?(expression: string): Promise<DebugSnapshotEvaluationReceipt>;
  stepOver?(): void;
  stepInto?(): void;
  stepOut?(): void;
  frameLocals?(frameIndex: number): Record<string, string> | null;
}

interface ActiveDebugSession {
  readonly id: number;
  readonly content: string;
  readonly language: string;
  readonly viewStateKey: string | undefined;
  readonly executionIdentity: string | undefined;
  readonly ownerGeneration: number;
  readonly abort: AbortController;
  readonly releaseExecution: (() => void) | null;
  readonly pauses: PersistedCodeDebugPause[];
}

async function readDebugConfigurationBackup(file: File): Promise<string> {
  if (typeof file.text === "function") return file.text();
  if (typeof FileReader === "undefined") {
    throw new Error("debug_configuration_backup_read_unavailable");
  }
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () =>
      typeof reader.result === "string"
        ? resolve(reader.result)
        : reject(new Error("debug_configuration_backup_not_text"));
    reader.onerror = () =>
      reject(
        reader.error ?? new Error("debug_configuration_backup_read_failed"),
      );
    reader.readAsText(file, "utf-8");
  });
}

/** Pause messages from either engine (python adds stack/kind/etc. extras). */
type DebugPausedMessage = WorkerDebugPausedMessage | PythonDebugPausedMessage;

function pauseIsUncaughtException(pause: DebugPausedMessage): boolean {
  return (
    ("reason" in pause && pause.reason === "exception") ||
    ("kind" in pause && pause.kind === "exception")
  );
}

function structuralPauseReceipt(
  pause: DebugPausedMessage,
): PersistedCodeDebugPause {
  const kind = pauseIsUncaughtException(pause)
    ? "exception"
    : "kind" in pause && pause.kind === "step"
      ? "step"
      : "breakpoint";
  const stack =
    "stack" in pause
      ? pause.stack.slice(0, CODE_PAGE_DEBUG_TIMELINE_STACK_LIMIT).map((frame) => ({
          name: frame.name,
          line: frame.line,
        }))
      : [];
  return { line: pause.line, kind, stack };
}

interface DebugWatchPanelProps {
  readonly watches: readonly ResolvedDebugWatch[];
  readonly valuesAvailable: boolean;
  readonly draft: string;
  readonly editingWatchId: string | null;
  readonly configurationReadOnly: boolean;
  readonly capacityReached: boolean;
  readonly commitDisabled: boolean;
  readonly validationMessage: string | null;
  readonly labels: {
    readonly title: string;
    readonly limit: string;
    readonly placeholder: string;
    readonly add: string;
    readonly submit: string;
    readonly cancel: string;
    readonly outOfScope: string;
    readonly valueUnavailable: string;
    readonly remove: (name: string) => string;
  };
  readonly onDraftChange: (value: string) => void;
  readonly onCommit: () => void;
  readonly onCancel: () => void;
  readonly onEdit: (watch: ResolvedDebugWatch) => void;
  readonly onRemove: (watch: ResolvedDebugWatch) => void;
}

function DebugWatchPanel({
  watches,
  valuesAvailable,
  draft,
  editingWatchId,
  configurationReadOnly,
  capacityReached,
  commitDisabled,
  validationMessage,
  labels,
  onDraftChange,
  onCommit,
  onCancel,
  onEdit,
  onRemove,
}: DebugWatchPanelProps) {
  return (
    <section
      aria-label={labels.title}
      data-testid="code-debug-watch-panel"
      className={officeDialogParts.group}
    >
      <div className={officeDialogParts.row}>
        <span className={officeDialogParts.sectionTitle}>{labels.title}</span>
        {capacityReached ? (
          <span
            role="status"
            data-testid="code-debug-watch-limit"
            className={officeDialogParts.message}
            data-tone="warning"
          >
            {labels.limit}
          </span>
        ) : null}
        <input
          data-testid="code-debug-watch-input"
          value={draft}
          disabled={configurationReadOnly || capacityReached}
          onChange={(event) =>
            onDraftChange(
              event.target.value.slice(0, MAX_DEBUG_WATCH_NAME_CHARS),
            )
          }
          maxLength={MAX_DEBUG_WATCH_NAME_CHARS}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              onCommit();
            } else if (event.key === "Escape" && editingWatchId) {
              event.preventDefault();
              onCancel();
            }
          }}
          placeholder={labels.placeholder}
          aria-label={labels.placeholder}
          aria-invalid={validationMessage ? "true" : undefined}
          aria-describedby={
            validationMessage ? "code-debug-watch-validation" : undefined
          }
          className={`${officeDialogParts.mono} ${styles.watchInput}`}
        />
        <button
          type="button"
          data-testid="code-debug-watch-add"
          onClick={onCommit}
          disabled={commitDisabled}
          title={capacityReached ? labels.limit : undefined}
          className={officeDialogParts.compactButton}
        >
          {editingWatchId ? labels.submit : labels.add}
        </button>
        {editingWatchId ? (
          <button
            type="button"
            data-testid="code-debug-watch-edit-cancel"
            onClick={onCancel}
            className={officeDialogParts.compactButton}
          >
            {labels.cancel}
          </button>
        ) : null}
        {validationMessage ? (
          <span
            id="code-debug-watch-validation"
            role="status"
            data-testid="code-debug-watch-validation"
            className={officeDialogParts.message}
            data-tone="warning"
          >
            {validationMessage}
          </span>
        ) : null}
      </div>
      {watches.length > 0 ? (
        <div className={`${officeDialogParts.well} ${styles.tableWell}`}>
          <table
            data-testid="code-debug-watches"
            className={`${officeDialogParts.table} ${officeDialogParts.mono}`}
          >
            <tbody>
              {watches.map((watch) => {
                const outOfScope = valuesAvailable && watch.outOfScope;
                return (
                  <tr key={watch.id} data-testid="code-debug-watch-row">
                    <td className={styles.muted}>
                      <button
                        type="button"
                        data-testid="code-debug-watch-edit"
                        disabled={configurationReadOnly}
                        onClick={() => onEdit(watch)}
                        aria-label={`${labels.submit}: ${watch.name}`}
                        className={styles.cellButton}
                      >
                        {watch.name}
                      </button>
                    </td>
                    <td
                      data-testid="code-debug-watch-value"
                      data-out-of-scope={outOfScope ? "true" : "false"}
                      className={
                        outOfScope ? `${styles.muted} ${styles.italic}` : undefined
                      }
                    >
                      {!valuesAvailable
                        ? labels.valueUnavailable
                        : outOfScope
                          ? labels.outOfScope
                          : (watch.value ?? "")}
                    </td>
                    <td className={styles.cellAction}>
                      <button
                        type="button"
                        data-testid="code-debug-watch-remove"
                        disabled={configurationReadOnly}
                        onClick={() => onRemove(watch)}
                        aria-label={labels.remove(watch.name)}
                        className={officeDialogParts.iconButton}
                      >
                        &times;
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : null}
    </section>
  );
}

export default function CodeDebugPanel({
  content,
  testExecution,
  language,
  viewStateKey,
  currentLine,
  pythonNamespace,
  sandboxOwner,
  stdin,
  onNavigateToLine,
  onPausedLineChange,
  configuration,
  onConfigurationChange,
  configurationReadOnly = false,
  commandRequest,
  onCommandRequestHandled,
  onStateChange,
  persistedTimeline = [],
  currentSourceFingerprint = null,
  selectedTimelineIndex = null,
  onTimelineSelect,
  onTimelineDebugAgain,
  onTimelineClear,
  onSessionSettled,
  hostBusy = false,
  acquireExecution,
}: CodeDebugPanelProps) {
  const t = useTranslations("console-notebooks");
  const [localConfiguration, setLocalConfiguration] =
    useState<CodeDebugConfiguration>(EMPTY_CODE_DEBUG_CONFIGURATION);
  const [lineInput, setLineInput] = useState("");
  const [breakpointQuery, setBreakpointQuery] = useState("");
  const deferredBreakpointQuery = useDeferredValue(breakpointQuery);
  const [breakpointVisibleLimit, setBreakpointVisibleLimit] = useState(
    CODE_DEBUG_BREAKPOINT_INITIAL_LIMIT,
  );
  const [running, setRunning] = useState(false);
  const [activePaused, setActivePaused] =
    useState<DebugPausedMessage | null>(null);
  const [pauseHistory, setPauseHistory] =
    useState<readonly DebugPausedMessage[]>([]);
  const [inspectedPauseIndex, setInspectedPauseIndex] = useState<number | null>(
    null,
  );
  const paused =
    inspectedPauseIndex === null
      ? activePaused
      : (pauseHistory[inspectedPauseIndex] ?? null);
  const pauseIsLive = inspectedPauseIndex === null && activePaused !== null;
  // Which call-stack frame's locals to show (0 = innermost, the default). Reset
  // to the innermost frame on every fresh pause. Python-only inspection.
  const [selectedFrame, setSelectedFrame] = useState(0);
  // Locals for the selected OUTER frame, resolved from the recording at click
  // time (in the handler, never during render — reading the controller ref
  // during render is disallowed). null = the frame's locals were never captured.
  const [selectedFrameLocals, setSelectedFrameLocals] = useState<
    Record<string, string> | null
  >(null);
  const [result, setResult] = useState<RunResult | null>(null);
  const controllerRef = useRef<DebugController | null>(null);
  const abortRef = useRef<AbortController | null>(null);
  const continueRef = useRef<(() => void) | null>(null);
  const sessionSequenceRef = useRef(0);
  const ownerGenerationRef = useRef(0);
  const activeSessionRef = useRef<ActiveDebugSession | null>(null);
  const lastInputRef = useRef({
    content,
    language,
    viewStateKey,
    executionContent: testExecution?.content,
    executionIdentity: testExecution?.identity,
  });
  const configurationImportInputRef = useRef<HTMLInputElement | null>(null);
  const configurationImportRequestIdRef = useRef(0);
  const [configurationTransferError, setConfigurationTransferError] = useState<
    string | null
  >(null);
  const [pendingConfigurationImport, setPendingConfigurationImport] = useState<{
    configuration: CodeDebugConfiguration;
    ownerGeneration: number;
    requestId: number;
    baseline: string;
  } | null>(null);
  const [watchDraft, setWatchDraft] = useState("");
  const [editingWatchId, setEditingWatchId] = useState<string | null>(null);
  const watchIdRef = useRef(0);
  const [evaluationDraft, setEvaluationDraft] = useState("");
  const [evaluationHistory, setEvaluationHistory] = useState<
    readonly DebugSnapshotEvaluationEntry[]
  >([]);
  const [evaluationBusy, setEvaluationBusy] = useState(false);
  const [timelineDetailFocus, setTimelineDetailFocus] = useState<{
    owner: string;
    key: string;
  } | null>(null);
  const evaluationOwnerGenerationRef = useRef(0);
  const evaluationEntryIdRef = useRef(0);
  // False after unmount, so the controller-result microtask (which the unmount
  // abort itself resolves) does not setState on a dead component.
  const aliveRef = useRef(true);

  const debuggable = isDebuggableLanguage(language);
  const debugActionBlocked = running || hostBusy || !debuggable;
  const breakpointAuthoringBlocked = configurationReadOnly || !debuggable;
  const isConditionValid =
    language === "python"
      ? isValidPythonBreakpointCondition
      : isValidBreakpointCondition;
  const effectiveConfiguration = configuration ?? localConfiguration;
  const pauseOnUncaughtExceptions =
    effectiveConfiguration.pauseOnUncaughtExceptions === true;
  const breakpoints = useMemo(
    () => effectiveConfiguration.breakpoints.map((breakpoint) => breakpoint.line),
    [effectiveConfiguration],
  );
  const typedBreakpointLine = /^\d+$/.test(lineInput)
    ? Number(lineInput)
    : null;
  const typedBreakpointValid =
    typedBreakpointLine !== null &&
    Number.isSafeInteger(typedBreakpointLine) &&
    typedBreakpointLine >= 1 &&
    typedBreakpointLine <= MAX_CODE_PAGE_BREAKPOINT_LINE;
  const typedBreakpointDuplicate =
    typedBreakpointLine !== null &&
    typedBreakpointValid &&
    breakpoints.includes(typedBreakpointLine);
  const typedBreakpointMessage =
    lineInput.length > 0 && !typedBreakpointValid
      ? t("code.debugBreakpointInvalid", {
          max: MAX_CODE_PAGE_BREAKPOINT_LINE,
        })
      : typedBreakpointDuplicate
        ? t("code.debugBreakpointDuplicate")
        : null;
  const typedBreakpointDisabled =
    breakpointAuthoringBlocked ||
    !typedBreakpointValid ||
    typedBreakpointDuplicate;
  const conditions = useMemo(
    () =>
      Object.fromEntries(
        effectiveConfiguration.breakpoints.flatMap((breakpoint) =>
          breakpoint.condition
            ? [[breakpoint.line, breakpoint.condition] as const]
            : [],
        ),
      ),
    [effectiveConfiguration],
  );
  const hitConditions = useMemo(
    () =>
      Object.fromEntries(
        effectiveConfiguration.breakpoints.flatMap((breakpoint) =>
          breakpoint.hitCondition
            ? [[breakpoint.line, breakpoint.hitCondition] as const]
            : [],
        ),
      ),
    [effectiveConfiguration],
  );
  const logMessages = useMemo(
    () =>
      Object.fromEntries(
        effectiveConfiguration.breakpoints.flatMap((breakpoint) =>
          breakpoint.logMessage
            ? [[breakpoint.line, breakpoint.logMessage] as const]
            : [],
        ),
      ),
    [effectiveConfiguration],
  );
  const normalizedBreakpointQuery = deferredBreakpointQuery
    .trim()
    .toLocaleLowerCase();
  const filteredBreakpoints = useMemo(
    () =>
      normalizedBreakpointQuery.length === 0
        ? breakpoints
        : breakpoints.filter((line) =>
            `${line}\n${conditions[line] ?? ""}\n${logMessages[line] ?? ""}\n${
              hitConditions[line] ?? ""
            }`
              .toLocaleLowerCase()
              .includes(normalizedBreakpointQuery),
          ),
    [
      breakpoints,
      conditions,
      hitConditions,
      logMessages,
      normalizedBreakpointQuery,
    ],
  );
  const visibleBreakpoints = filteredBreakpoints.slice(
    0,
    breakpointVisibleLimit,
  );
  const hiddenBreakpointCount = Math.max(
    0,
    filteredBreakpoints.length - visibleBreakpoints.length,
  );

  const updateConfiguration = useCallback(
    (
      update: (current: CodeDebugConfiguration) => CodeDebugConfiguration,
    ) => {
      if (configurationReadOnly) return;
      const next = normalizeCodeDebugConfiguration(
        update(configuration ?? localConfiguration),
      );
      if (configuration === undefined) setLocalConfiguration(next);
      onConfigurationChange?.(next);
    },
    [
      configuration,
      configurationReadOnly,
      localConfiguration,
      onConfigurationChange,
    ],
  );

  const setPauseOnUncaughtExceptions = useCallback(
    (enabled: boolean) => {
      updateConfiguration((current) => ({
        ...current,
        pauseOnUncaughtExceptions: enabled,
      }));
    },
    [updateConfiguration],
  );

  const handleDownloadConfiguration = useCallback(() => {
    setConfigurationTransferError(null);
    const backup = createCodeDebugConfigurationBackup(effectiveConfiguration);
    try {
      if (
        !backup.ok ||
        !downloadTextFile(
          CODE_DEBUG_CONFIGURATION_BACKUP_FILENAME,
          backup.ok ? backup.json : "",
          "application/json;charset=utf-8",
        )
      ) {
        setConfigurationTransferError(t("toolbar.exportDownloadUnavailable"));
      }
    } catch {
      setConfigurationTransferError(t("toolbar.exportDownloadUnavailable"));
    }
  }, [effectiveConfiguration, t]);

  const handleImportConfiguration = useCallback(
    async (event: ChangeEvent<HTMLInputElement>) => {
      const file = event.currentTarget.files?.[0];
      // Selecting the same corrected file must retrigger after a failed read.
      event.currentTarget.value = "";
      if (!file || breakpointAuthoringBlocked) return;

      const requestId = ++configurationImportRequestIdRef.current;
      const ownerGeneration = ownerGenerationRef.current;
      setConfigurationTransferError(null);
      setPendingConfigurationImport(null);
      if (file.size > MAX_CODE_DEBUG_CONFIGURATION_BACKUP_BYTES) {
        setConfigurationTransferError(t("editor.importFailed"));
        return;
      }
      try {
        const raw = await readDebugConfigurationBackup(file);
        if (
          !aliveRef.current ||
          configurationImportRequestIdRef.current !== requestId ||
          ownerGenerationRef.current !== ownerGeneration
        ) {
          return;
        }
        const parsed = parseCodeDebugConfigurationBackup(raw);
        if (!parsed.ok) {
          setConfigurationTransferError(t("editor.importFailed"));
          return;
        }
        setPendingConfigurationImport({
          configuration: parsed.configuration,
          ownerGeneration,
          requestId,
          baseline: JSON.stringify(effectiveConfiguration),
        });
      } catch {
        if (
          aliveRef.current &&
          configurationImportRequestIdRef.current === requestId &&
          ownerGenerationRef.current === ownerGeneration
        ) {
          setConfigurationTransferError(t("editor.importFailed"));
        }
      }
    },
    [breakpointAuthoringBlocked, effectiveConfiguration, t],
  );

  const confirmConfigurationImport = useCallback(() => {
    const pending = pendingConfigurationImport;
    if (!pending || breakpointAuthoringBlocked) return;
    if (
      pending.ownerGeneration !== ownerGenerationRef.current ||
      pending.requestId !== configurationImportRequestIdRef.current ||
      pending.baseline !== JSON.stringify(effectiveConfiguration)
    ) {
      setPendingConfigurationImport(null);
      setConfigurationTransferError(t("code.debugImportStale"));
      return;
    }
    configurationImportRequestIdRef.current += 1;
    updateConfiguration(() => pending.configuration);
    setPendingConfigurationImport(null);
    setConfigurationTransferError(null);
  }, [
    breakpointAuthoringBlocked,
    effectiveConfiguration,
    pendingConfigurationImport,
    t,
    updateConfiguration,
  ]);

  // Drop any in-flight debug session if the panel unmounts (window closed).
  // React Strict Mode deliberately runs setup → cleanup → setup in development;
  // reset the liveness guard during setup so the second, real mount can still
  // publish the final debug result after Continue.
  useLayoutEffect(() => {
    aliveRef.current = true;
    return () => {
      aliveRef.current = false;
      ownerGenerationRef.current += 1;
      evaluationOwnerGenerationRef.current += 1;
      configurationImportRequestIdRef.current += 1;
      const active = activeSessionRef.current;
      activeSessionRef.current = null;
      abortRef.current?.abort();
      abortRef.current = null;
      controllerRef.current = null;
      continueRef.current = null;
      active?.releaseExecution?.();
    };
  }, []);

  const addBreakpoint = useCallback(
    (line: number) => {
      if (!debuggable) return;
      if (
        !Number.isFinite(line) ||
        !Number.isInteger(line) ||
        line < 1 ||
        line > MAX_CODE_PAGE_BREAKPOINT_LINE
      ) {
        return;
      }
      updateConfiguration((current) => ({
        ...current,
        breakpoints: current.breakpoints.some(
          (breakpoint) => breakpoint.line === line,
        )
          ? current.breakpoints
          : [...current.breakpoints, { line }],
      }));
    },
    [debuggable, updateConfiguration],
  );

  const removeBreakpoint = useCallback(
    (line: number) => {
      updateConfiguration((current) => ({
        ...current,
        breakpoints: current.breakpoints.filter(
          (breakpoint) => breakpoint.line !== line,
        ),
      }));
    },
    [updateConfiguration],
  );

  const setCondition = useCallback(
    (line: number, value: string) => {
      updateConfiguration((current) => ({
        ...current,
        breakpoints: current.breakpoints.map((breakpoint) =>
          breakpoint.line === line
            ? {
                ...breakpoint,
                condition: value.trim() === "" ? undefined : value,
              }
            : breakpoint,
        ),
      }));
    },
    [updateConfiguration],
  );

  const setLogMessage = useCallback(
    (line: number, value: string) => {
      updateConfiguration((current) => ({
        ...current,
        breakpoints: current.breakpoints.map((breakpoint) =>
          breakpoint.line === line
            ? {
                ...breakpoint,
                logMessage: value.trim() === "" ? undefined : value,
              }
            : breakpoint,
        ),
      }));
    },
    [updateConfiguration],
  );

  const setHitCondition = useCallback(
    (line: number, value: string) => {
      updateConfiguration((current) => ({
        ...current,
        breakpoints: current.breakpoints.map((breakpoint) =>
          breakpoint.line === line
            ? {
                ...breakpoint,
                hitCondition: value.trim() === "" ? undefined : value,
              }
            : breakpoint,
        ),
      }));
    },
    [updateConfiguration],
  );

  // Every fresh pause shows the innermost frame's locals until the user clicks
  // an outer frame.
  // Lines a logpoint reported during the current run. Cleared when a new run
  // starts, like the paused marker and the result.
  const [logLines, setLogLines] = useState<readonly LogpointLine[]>([]);

  const retireSnapshotEvaluation = useCallback((clearHistory: boolean) => {
    evaluationOwnerGenerationRef.current += 1;
    setEvaluationBusy(false);
    if (clearHistory) {
      setEvaluationDraft("");
      setEvaluationHistory((current) => (current.length === 0 ? current : []));
    }
  }, []);

  // Logpoint lines are read from a ref so this callback stays stable: the
  // controller is handed `onPause` once, at the start of a run, and would
  // otherwise keep calling the version that existed then.
  //
  // Published from an EFFECT, not during render — writing a ref while rendering
  // is a side effect in render, and a pause can only arrive from an async
  // worker message, which is always after the commit.
  // Same reason as `logMessagesRef`: the pause handler must not close over a
  // controller that did not exist when it was created.
  const logMessagesRef = useRef<Record<number, string>>({});
  useEffect(() => {
    logMessagesRef.current = logMessages;
  }, [logMessages]);

  const handlePause = useCallback(
    (p: DebugPausedMessage) => {
      // A new pause is a new expression owner. A result from the previous
      // frame must never hydrate this frame even when both report the same line.
      retireSnapshotEvaluation(false);
      // A logpoint REPORTS instead of stopping. The line still carries a real
      // breakpoint — that is how the run captures locals there — so the stop
      // arrives and is turned into a line of output and released immediately.
      // An uncaught exception is terminal authority, not a logpoint hit, even
      // when its last statement happens to share a configured logpoint line.
      const isUncaughtException = pauseIsUncaughtException(p);
      const message = logMessagesRef.current[p.line];
      if (!isUncaughtException && isLogpointMessage(message)) {
        setLogLines((current) =>
          appendLogpointLine(current, {
            line: p.line,
            text: formatLogpoint(message, p.locals ?? null),
          }),
        );
        continueRef.current?.();
        return;
      }
      const active = activeSessionRef.current;
      if (
        active &&
        active.pauses.length < CODE_PAGE_DEBUG_TIMELINE_PAUSE_LIMIT
      ) {
        active.pauses.push(structuralPauseReceipt(p));
      }
      setPauseHistory((current) => {
        const candidate = [...current, p];
        return candidate.length <= CODE_PAGE_DEBUG_TIMELINE_PAUSE_LIMIT
          ? candidate
          : candidate.slice(candidate.length - CODE_PAGE_DEBUG_TIMELINE_PAUSE_LIMIT);
      });
      setInspectedPauseIndex(null);
      setSelectedFrame(0);
      setSelectedFrameLocals(null);
      setActivePaused(p);
      onPausedLineChange?.(p.line);
      onNavigateToLine?.(p.line, false);
    },
    [onNavigateToLine, onPausedLineChange, retireSnapshotEvaluation],
  );

  const clearPausedLine = useCallback(() => {
    setActivePaused(null);
    setInspectedPauseIndex(null);
    onPausedLineChange?.(null);
  }, [onPausedLineChange]);

  // Stop and Restart are hard session boundaries. Retire the identity before
  // aborting so a worker's final pause/result cannot repopulate a session the
  // user has already left (or overwrite the replacement session).
  const retireActiveSession = useCallback(() => {
    const active = activeSessionRef.current;
    activeSessionRef.current = null;
    const abort = active?.abort ?? abortRef.current;
    abortRef.current = null;
    controllerRef.current = null;
    continueRef.current = null;
    abort?.abort();
    active?.releaseExecution?.();
    retireSnapshotEvaluation(true);
    setRunning(false);
    clearPausedLine();
    setPauseHistory((current) => (current.length === 0 ? current : []));
  }, [clearPausedLine, retireSnapshotEvaluation]);

  // A debugger controller always executes the exact page/source/language
  // snapshot captured when Debug was pressed. Page identity is independent of
  // bytes: A -> B must retire A even when both pages contain identical source.
  // Advance the owner generation before aborting so late worker messages cannot
  // attach old output, pauses or logpoints to the replacement surface.
  useLayoutEffect(() => {
    const previous = lastInputRef.current;
    lastInputRef.current = {
      content,
      language,
      viewStateKey,
      executionContent: testExecution?.content,
      executionIdentity: testExecution?.identity,
    };
    if (
      previous.content === content &&
      previous.language === language &&
      previous.viewStateKey === viewStateKey &&
      previous.executionContent === testExecution?.content &&
      previous.executionIdentity === testExecution?.identity
    ) {
      return;
    }

    ownerGenerationRef.current += 1;
    configurationImportRequestIdRef.current += 1;
    // Deliberate owner-boundary retirement: source/page replacement must clear
    // the committed debugger before the replacement surface can act.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    retireActiveSession();
    setConfigurationTransferError(null);
    setPendingConfigurationImport(null);
    setResult(null);
    setLogLines((current) => (current.length === 0 ? current : []));
    setSelectedFrame(0);
    setSelectedFrameLocals(null);
    setWatchDraft("");
    setEditingWatchId(null);
  }, [content, language, retireActiveSession, testExecution, viewStateKey]);

  const handleAddCurrent = useCallback(() => {
    if (typeof currentLine === "number") addBreakpoint(currentLine);
  }, [addBreakpoint, currentLine]);

  // `handleDebug` is declared below; the ref keeps this callback from having to
  // be defined after it, and is published from an effect rather than render.
  const handleDebugRef = useRef<
    ((options?: { alsoStopAt?: number; restart?: boolean }) => void) | null
  >(null);

  // Run to Cursor (VS Code parity): debug up to where the caret is, without
  // first placing a breakpoint there and then having to remember to remove it.
  const handleRunToCursor = useCallback(() => {
    if (typeof currentLine !== "number") return;
    handleDebugRef.current?.({ alsoStopAt: currentLine });
  }, [currentLine]);

  const handleRestart = useCallback(() => {
    handleDebugRef.current?.({ restart: true });
  }, []);

  const handleAddTyped = useCallback(() => {
    if (
      typedBreakpointLine === null ||
      !typedBreakpointValid ||
      typedBreakpointDuplicate
    ) {
      return;
    }
    addBreakpoint(typedBreakpointLine);
    setLineInput("");
  }, [
    addBreakpoint,
    typedBreakpointDuplicate,
    typedBreakpointLine,
    typedBreakpointValid,
  ]);

  const handleDebug = useCallback((options?: {
    alsoStopAt?: number;
    restart?: boolean;
  }) => {
    // A disabled button is not the only way in (the host also drives this via
    // its command bridge), so refuse here too rather than trusting the UI.
    if (hostBusy || !debuggable) return;
    if (activeSessionRef.current) {
      if (!options?.restart) return;
      retireActiveSession();
    }
    // Only valid conditions ship; broken ones stay UI-flagged.
    const validConditions: Record<number, string> = {};
    for (const [line, expr] of Object.entries(conditions)) {
      if (isConditionValid(expr)) {
        validConditions[Number(line)] = expr;
      }
    }

    // Run to Cursor is a ONE-SHOT stop: the line is added to what this run
    // reports, never to the saved configuration, so it leaves no breakpoint
    // behind in the list, on the gutter, or in the page's stored settings.
    const oneShot = options?.alsoStopAt ?? testExecution?.stopLine;
    const runBreakpoints =
      typeof oneShot === "number" &&
      Number.isInteger(oneShot) &&
      oneShot >= 1 &&
      !breakpoints.includes(oneShot)
        ? [...breakpoints, oneShot]
        : breakpoints;

    // Transpile TS → JS before opening a session. A compile error has no live
    // controller to invalidate and surfaces exactly like a normal TS run.
    const executionContent = testExecution?.content ?? content;
    let debugContent = executionContent;
    if (language === "typescript") {
      const transpiled = transpileTypeScript(executionContent);
      if (transpiled.error || transpiled.code === undefined) {
        setResult({
          logs: [],
          error: transpiled.error ?? "TypeScript compile error",
        });
        return;
      }
      debugContent = transpiled.code;
    }

    // The rendered hostBusy state can be one paint behind a sibling view. The
    // page registry is the dispatch-time authority and closes that same-frame
    // race before a worker/kernel is started or the current result is cleared.
    const releaseExecution = acquireExecution?.() ?? null;
    if (acquireExecution && !releaseExecution) return;

    setResult(null);
    setLogLines([]);
    setPauseHistory([]);
    setInspectedPauseIndex(null);
    retireSnapshotEvaluation(true);
    clearPausedLine();

    const abort = new AbortController();
    const sessionId = sessionSequenceRef.current + 1;
    const ownerGeneration = ownerGenerationRef.current;
    sessionSequenceRef.current = sessionId;
    activeSessionRef.current = {
      id: sessionId,
      content,
      language,
      viewStateKey,
      executionIdentity: testExecution?.identity,
      ownerGeneration,
      abort,
      releaseExecution,
      pauses: [],
    };
    const handleSessionPause = (pause: DebugPausedMessage) => {
      const active = activeSessionRef.current;
      if (
        aliveRef.current &&
        active?.id === sessionId &&
        active.ownerGeneration === ownerGeneration &&
        ownerGenerationRef.current === ownerGeneration
      ) {
        handlePause(pause);
      }
    };

    let controller: DebugController;
    try {
      if (language === "python") {
        // Only parseable hit-counts ship; unparseable ones stay UI-flagged.
        const validHitConditions: Record<number, string> = {};
        for (const [line, expr] of Object.entries(hitConditions)) {
          if (parseHitCondition(expr)) validHitConditions[Number(line)] = expr;
        }
        // Recorded run + interactive replay on the page's live kernel (same
        // namespace/stdin as a normal run; breakpoints work inside functions).
        controller = runPythonDebug(executionContent, {
          breakpoints: runBreakpoints,
          conditions: validConditions,
          hitConditions: validHitConditions,
          namespace: pythonNamespace,
          stdin,
          signal: abort.signal,
          pauseOnUncaughtExceptions,
          onPause: handleSessionPause,
        });
      } else {
        controller = runJavaScriptDebug(debugContent, {
          breakpoints: runBreakpoints,
          conditions: validConditions,
          signal: abort.signal,
          pauseOnUncaughtExceptions,
          sandboxOwner,
          onPause: handleSessionPause,
        });
      }
    } catch (error) {
      if (activeSessionRef.current?.id === sessionId) {
        activeSessionRef.current = null;
      }
      releaseExecution?.();
      setResult(
        boundedDebugFailure(error, t("document.capability.unavailable")),
      );
      clearPausedLine();
      return;
    }
    setRunning(true);
    abortRef.current = abort;
    controllerRef.current = controller;
    // The pause handler releases a logpoint through this rather than through
    // `controllerRef`, which it cannot name (it is declared after it).
    continueRef.current = () => controller.continue();
    const settleSession = (res: RunResult, trustedReceipt: boolean) => {
      const active = activeSessionRef.current;
      if (
        active?.id !== sessionId ||
        active.ownerGeneration !== ownerGeneration ||
        ownerGenerationRef.current !== ownerGeneration
      ) {
        return;
      }
      const boundedResult = boundExecutionResult(res);
      // The console is scoped to this one debug session. Expressions and values
      // may contain secrets, so release them as soon as the session settles.
      retireSnapshotEvaluation(true);
      if (
        trustedReceipt &&
        active.executionIdentity === undefined &&
        onSessionSettled
      ) {
        try {
          onSessionSettled({
            content: active.content,
            language: active.language,
            viewStateKey: active.viewStateKey,
            completedAt: Date.now(),
            outcome: boundedResult.error === undefined ? "success" : "error",
            pauses: [...active.pauses],
          });
        } catch {
          // Host persistence failure must not resurrect or wedge a settled run.
        }
      }
      activeSessionRef.current = null;
      active.releaseExecution?.();
      if (abortRef.current === abort) abortRef.current = null;
      if (controllerRef.current === controller) controllerRef.current = null;
      continueRef.current = null;
      if (!aliveRef.current) return;
      setResult(boundedResult);
      setRunning(false);
      clearPausedLine();
    };
    void controller.result.then(
      (settled) => settleSession(settled, true),
      (error: unknown) =>
        settleSession(
          boundedDebugFailure(
            error,
            t("document.capability.unavailable"),
          ),
          false,
        ),
    );
  }, [
    acquireExecution,
    breakpoints,
    clearPausedLine,
    hostBusy,
    conditions,
    content,
    debuggable,
    handlePause,
    hitConditions,
    isConditionValid,
    language,
    pauseOnUncaughtExceptions,
    onSessionSettled,
    pythonNamespace,
    retireSnapshotEvaluation,
    retireActiveSession,
    sandboxOwner,
    stdin,
    t,
    testExecution,
    viewStateKey,
  ]);

  // Published in layout so a host command consumed later in this same commit
  // cannot reach the previous page's restart/run-to-cursor closure.
  useLayoutEffect(() => {
    handleDebugRef.current = handleDebug;
  }, [handleDebug]);

  const handleContinue = useCallback(() => {
    if (!pauseIsLive) return;
    retireSnapshotEvaluation(false);
    clearPausedLine();
    controllerRef.current?.continue();
  }, [clearPausedLine, pauseIsLive, retireSnapshotEvaluation]);

  const handleStep = useCallback(() => {
    if (!pauseIsLive) return;
    retireSnapshotEvaluation(false);
    clearPausedLine();
    controllerRef.current?.step();
  }, [clearPausedLine, pauseIsLive, retireSnapshotEvaluation]);

  const handleStepOver = useCallback(() => {
    if (!pauseIsLive) return;
    retireSnapshotEvaluation(false);
    clearPausedLine();
    controllerRef.current?.stepOver?.();
  }, [clearPausedLine, pauseIsLive, retireSnapshotEvaluation]);

  const handleStepInto = useCallback(() => {
    if (!pauseIsLive) return;
    retireSnapshotEvaluation(false);
    clearPausedLine();
    controllerRef.current?.stepInto?.();
  }, [clearPausedLine, pauseIsLive, retireSnapshotEvaluation]);

  const handleStepOut = useCallback(() => {
    if (!pauseIsLive) return;
    retireSnapshotEvaluation(false);
    clearPausedLine();
    controllerRef.current?.stepOut?.();
  }, [clearPausedLine, pauseIsLive, retireSnapshotEvaluation]);

  // Click a call-stack frame to inspect its locals (resolved from the
  // recording). Innermost (0) shows the event's own locals; outer frames fall
  // back to the last event paused in that frame, or null if never captured.
  const handleSelectFrame = useCallback(
    (frameIndex: number) => {
      setSelectedFrame(frameIndex);
      // Resolve the frame's locals NOW (in the handler), not during render.
      setSelectedFrameLocals(
        frameIndex > 0 && pauseIsLive
          ? controllerRef.current?.frameLocals?.(frameIndex) ?? null
          : null,
      );
      const frame =
        paused && "stack" in paused ? paused.stack[frameIndex] : undefined;
      if (frame) onNavigateToLine?.(frame.line, true);
    },
    [onNavigateToLine, pauseIsLive, paused],
  );

  const handleStop = useCallback(() => {
    retireActiveSession();
  }, [retireActiveSession]);

  // CodeWindow's Debug ribbon drives the exact same handlers as this panel's
  // controls. A request id is consumed once, so StrictMode/effect replays never
  // start a second debug worker or advance the controller twice.
  const lastCommandRequestRef = useRef<number | undefined>(undefined);
  const executeHostCommand = useCallback(
    (request: CodeDebugCommandRequest) => {
      if (request.requestId === lastCommandRequestRef.current) return;
      lastCommandRequestRef.current = request.requestId;
      // Acknowledge before owner rejection or dispatch. The host owns a
      // separate monotonic sequence, so clearing this exact payload prevents a
      // consumed A command from replaying if the panel later renders page B.
      onCommandRequestHandled?.(request.requestId);
      if (
        viewStateKey !== undefined &&
        request.viewStateKey !== viewStateKey
      ) {
        return;
      }
      switch (request.command) {
        case "start":
          handleDebug();
          break;
        case "restart":
          handleRestart();
          break;
        case "stop":
          handleStop();
          break;
        case "addBreakpointCurrent":
          handleAddCurrent();
          break;
        case "runToCursor":
          handleRunToCursor();
          break;
        case "continue":
          handleContinue();
          break;
        case "step":
          handleStep();
          break;
        case "stepOver":
          handleStepOver();
          break;
        case "stepInto":
          handleStepInto();
          break;
        case "stepOut":
          handleStepOut();
          break;
      }
    },
    [
      handleAddCurrent,
      handleContinue,
      handleDebug,
      handleRestart,
      handleRunToCursor,
      handleStep,
      handleStepInto,
      handleStepOut,
      handleStepOver,
      handleStop,
      onCommandRequestHandled,
      viewStateKey,
    ],
  );
  useLayoutEffect(() => {
    if (!commandRequest) return;
    // Execute against the committed owner immediately. Deferring to a
    // microtask let a page replacement commit before an A-scoped command ran;
    // the explicit viewStateKey check above is the final authority boundary.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    executeHostCommand(commandRequest);
  }, [commandRequest, executeHostCommand]);

  useEffect(() => {
    onStateChange?.({
      debuggable,
      running,
      paused: pauseIsLive,
      pausedLine: pauseIsLive ? activePaused?.line ?? null : null,
      currentLine: currentLine ?? null,
      breakpointCount: breakpoints.length,
      supportsDepthStepping: language === "python",
    });
  }, [
    breakpoints.length,
    currentLine,
    debuggable,
    language,
    onStateChange,
    activePaused,
    pauseIsLive,
    running,
  ]);

  // Python recording extras (absent on JS pauses): user-frame stack, a raised
  // condition, and the one-time "recording capped" note.
  const pausedStack =
    paused && "stack" in paused && paused.stack.length > 1 ? paused.stack : null;
  // Locals to display: the paused event's own (innermost / JS) by default, or
  // the selected outer frame's recording-resolved snapshot. `null` means the
  // outer frame's locals were never captured — render the "no locals" note.
  const isPython = language === "python";

  /**
   * Breakpoints this source cannot pause at.
   *
   * The JavaScript debugger works by inserting an awaited hook at safe
   * statement boundaries, so it cannot pause inside a NON-async function body:
   * `await` is not legal there. A breakpoint on such a line is not an error —
   * it simply never fires — and saying nothing would leave the user pressing
   * Debug and watching their program run straight past the line they marked.
   *
   * VS Code shows exactly this state as an unverified (hollow) breakpoint; so
   * does this panel. Python is excluded: its debugger traces real frames and
   * can stop anywhere.
   */
  const unbindableBreakpoints = useMemo(
    () => new Set(unbindableBreakpointLines(content, breakpoints, language)),
    [breakpoints, content, language],
  );
  // Watched names are authored debugger configuration: they outlive any one
  // pause and the page itself, while their resolved values remain session-only.
  const watches = useMemo(
    () => effectiveConfiguration.watches ?? [],
    [effectiveConfiguration.watches],
  );
  const watchCapacityReached =
    editingWatchId === null && watches.length >= MAX_DEBUG_WATCHES;
  const watchDraftName = watchDraft.trim();
  const watchDraftValid = isWatchableName(watchDraftName);
  const watchDraftDuplicate =
    watchDraftValid &&
    watches.some(
      (watch) =>
        watch.id !== editingWatchId && watch.name === watchDraftName,
    );
  const watchDraftUnchanged =
    editingWatchId !== null &&
    watches.some(
      (watch) =>
        watch.id === editingWatchId && watch.name === watchDraftName,
    );
  const watchValidationMessage =
    watchDraft.length > 0 && !watchDraftValid
      ? t("code.debugWatchInvalid")
      : watchDraftDuplicate
        ? t("code.debugWatchDuplicate")
        : null;
  const watchCommitDisabled =
    configurationReadOnly ||
    watchCapacityReached ||
    !watchDraftValid ||
    watchDraftDuplicate ||
    watchDraftUnchanged;
  const commitWatch = useCallback(() => {
    const name = watchDraft.trim();
    let committed = false;
    updateConfiguration((current) => {
      const currentWatches = current.watches ?? [];
      const next = editingWatchId
        ? renameDebugWatch(currentWatches, editingWatchId, name)
        : addDebugWatch(currentWatches, name, () => {
            let candidate: string;
            do {
              watchIdRef.current += 1;
              candidate = `watch-${watchIdRef.current}`;
            } while (currentWatches.some((watch) => watch.id === candidate));
            return candidate;
          });
      if (next === currentWatches) return current;
      committed = true;
      return { ...current, watches: next };
    });
    // Invalid and duplicate drafts stay visible so the user can fix them.
    if (committed) {
      setWatchDraft("");
      setEditingWatchId(null);
    }
  }, [editingWatchId, updateConfiguration, watchDraft]);

  const cancelWatchEdit = useCallback(() => {
    setEditingWatchId(null);
    setWatchDraft("");
  }, []);

  useEffect(() => {
    if (
      editingWatchId !== null &&
      !watches.some((watch) => watch.id === editingWatchId)
    ) {
      // External configuration replacement invalidated the edited watch.
      // eslint-disable-next-line react-hooks/set-state-in-effect
      cancelWatchEdit();
    }
  }, [cancelWatchEdit, editingWatchId, watches]);

  // Memoized because the watch list resolves against it: a fresh object every
  // render would re-resolve every watch on every render.
  const displayedLocals = useMemo(
    () =>
      isPython && pausedStack && selectedFrame > 0
        ? selectedFrameLocals
        : (paused?.locals ?? {}),
    [isPython, pausedStack, selectedFrame, selectedFrameLocals, paused],
  );
  // Resolved against the frame the user is LOOKING at, so stepping out changes
  // the answers the way the locals table above it does.
  const resolvedWatches = useMemo(
    () => resolveDebugWatches(watches, paused ? displayedLocals : null),
    [watches, paused, displayedLocals],
  );
  const localsEntries = displayedLocals ? Object.entries(displayedLocals) : [];
  const pausedConditionError =
    paused && "conditionError" in paused && paused.conditionError
      ? paused.conditionError
      : null;
  const pausedConditionNote =
    paused && "conditionNote" in paused && paused.conditionNote
      ? paused.conditionNote
      : null;
  const pausedTruncated = paused && "truncated" in paused && paused.truncated === true;
  const pausedIsUncaughtException =
    paused !== null &&
    (("reason" in paused && paused.reason === "exception") ||
      ("kind" in paused && paused.kind === "exception"));
  const pausedException =
    pausedIsUncaughtException &&
    paused &&
    "exception" in paused &&
    typeof paused.exception === "string"
      ? paused.exception
      : null;
  const snapshotEvaluationAvailable =
    (language === "javascript" || language === "typescript") && pauseIsLive;
  const handleEvaluateSnapshot = useCallback(async () => {
    const expression = normalizeDebugSnapshotExpression(evaluationDraft);
    const controller = controllerRef.current;
    const active = activeSessionRef.current;
    const livePause = activePaused;
    if (
      !expression ||
      evaluationBusy ||
      !snapshotEvaluationAvailable ||
      !controller?.evaluate ||
      !active ||
      !livePause
    ) {
      return;
    }

    evaluationOwnerGenerationRef.current += 1;
    const requestGeneration = evaluationOwnerGenerationRef.current;
    const sessionId = active.id;
    const ownerGeneration = active.ownerGeneration;
    const pauseStep = livePause.step;
    evaluationEntryIdRef.current += 1;
    const entryId = evaluationEntryIdRef.current;
    setEvaluationDraft("");
    setEvaluationBusy(true);

    let receipt: DebugSnapshotEvaluationReceipt;
    try {
      receipt = await controller.evaluate(expression);
    } catch {
      // The production controller always resolves a bounded receipt, but a
      // transport failure must not leave the form permanently busy.
      receipt = {
        status: "error",
        value: t("document.capability.unavailable"),
      };
    }
    const current = activeSessionRef.current;
    if (
      !aliveRef.current ||
      evaluationOwnerGenerationRef.current !== requestGeneration ||
      current?.id !== sessionId ||
      current.ownerGeneration !== ownerGeneration ||
      ownerGenerationRef.current !== ownerGeneration ||
      controllerRef.current !== controller ||
      activePaused?.step !== pauseStep
    ) {
      return;
    }
    setEvaluationBusy(false);
    setEvaluationHistory((history) =>
      appendDebugSnapshotEvaluation(history, {
        id: entryId,
        expression,
        status: receipt.status,
        value: receipt.value,
      }),
    );
  }, [
    activePaused,
    evaluationBusy,
    evaluationDraft,
    snapshotEvaluationAvailable,
    t,
  ]);
  const handleInspectPause = useCallback(
    (index: number) => {
      const snapshot = pauseHistory[index];
      if (!snapshot) return;
      setInspectedPauseIndex(snapshot === activePaused ? null : index);
      setSelectedFrame(0);
      setSelectedFrameLocals(null);
      onNavigateToLine?.(snapshot.line, true);
    },
    [activePaused, onNavigateToLine, pauseHistory],
  );
  const selectedPersistedEntry =
    selectedTimelineIndex === null
      ? null
      : (persistedTimeline[selectedTimelineIndex] ?? null);
  const selectedPersistedEntryIsCurrent = Boolean(
    selectedPersistedEntry &&
      currentSourceFingerprint !== null &&
      selectedPersistedEntry.language === language &&
      selectedPersistedEntry.sourceFingerprint === currentSourceFingerprint,
  );
  const selectedPersistedEntryOwner =
    selectedPersistedEntry && selectedTimelineIndex !== null
      ? `${selectedTimelineIndex}:${selectedPersistedEntry.completedAt}:${selectedPersistedEntry.language}:${selectedPersistedEntry.sourceFingerprint}`
      : null;
  const selectedPersistedEntryNavigationKeys = selectedPersistedEntry
    ? selectedPersistedEntry.pauses.flatMap((pause, pauseIndex) => [
        `pause:${pauseIndex}`,
        ...pause.stack.map(
          (_frame, frameIndex) => `frame:${pauseIndex}:${frameIndex}`,
        ),
      ])
    : [];
  const selectedPersistedEntryRovingKey =
    selectedPersistedEntryIsCurrent &&
    selectedPersistedEntryOwner !== null &&
    timelineDetailFocus?.owner === selectedPersistedEntryOwner &&
    selectedPersistedEntryNavigationKeys.includes(timelineDetailFocus.key)
      ? timelineDetailFocus.key
      : (selectedPersistedEntryNavigationKeys[0] ?? null);

  // A bottom panel in the dialog chrome (code tone, the task panes' glass
  // card): chrome buttons (Debug is the primary), fields, checks, segmented
  // pauses, messages, wells and tables, around the Code-owned debugger.
  return (
    <div
      data-testid="code-debug-panel"
      data-breakpoints={breakpoints.length}
      className={`${officeDialogClassName({ tone: "code", pane: true })} ${styles.bottomPane} ${styles.debugPane}`}
    >
      <div className={officeDialogParts.row}>
        <button
          type="button"
          onClick={() => handleDebug()}
          disabled={debugActionBlocked}
          data-testid="code-debug-run"
          aria-keyshortcuts="F5"
          title={
            !debuggable
              ? t("code.debugUnsupported")
              : hostBusy
                ? t("code.debugBlockedByRun")
                : t("code.debugRun")
          }
          className={`${officeDialogParts.primaryButton} ${officeDialogParts.compactButton}`}
        >
          {running ? (
            <Loader2 size={13} className="animate-spin" />
          ) : (
            <Bug size={13} />
          )}
          {t("code.debugRun")}
        </button>
        {running ? (
          <>
            <button
              type="button"
              onClick={handleRestart}
              data-testid="code-debug-restart"
              aria-keyshortcuts="Control+Shift+F5 Meta+Shift+F5"
              title={t("code.debugRestart")}
              className={officeDialogParts.compactButton}
            >
              <RotateCcw size={12} aria-hidden="true" />
              {t("code.debugRestart")}
            </button>
            <button
              type="button"
              onClick={handleStop}
              data-testid="code-debug-stop"
              aria-keyshortcuts="Shift+F5"
              title={t("code.debugStop")}
              className={officeDialogParts.compactButton}
            >
              {t("code.debugStop")}
            </button>
          </>
        ) : null}
        {testExecution ? (
          <span
            role="status"
            data-testid="code-debug-test-target"
            title={testExecution.testName}
            className={`${officeDialogParts.hint} ${styles.ellipsis} ${styles.debugTarget}`}
          >
            {t("code.testExplorer.debuggingTest", {
              name: testExecution.testName,
            })}
          </span>
        ) : null}
        {typeof currentLine === "number" ? (
          <button
            type="button"
            onClick={handleAddCurrent}
            disabled={breakpointAuthoringBlocked}
            data-testid="code-debug-add-current"
            data-line={currentLine}
            title={t("code.debugAddBreakpointCurrent", { line: currentLine })}
            className={officeDialogParts.compactButton}
          >
            {t("code.debugAddBreakpointCurrent", { line: currentLine })}
          </button>
        ) : null}
        {typeof currentLine === "number" ? (
          <button
            type="button"
            onClick={handleRunToCursor}
            disabled={debugActionBlocked}
            data-testid="code-debug-run-to-cursor"
            data-line={currentLine}
            title={t("code.debugRunToCursor", { line: currentLine })}
            className={officeDialogParts.compactButton}
          >
            {t("code.debugRunToCursorShort")}
          </button>
        ) : null}
        <span className={`${officeDialogParts.row} ${styles.nowrap}`}>
          <input
            type="number"
            min={1}
            max={MAX_CODE_PAGE_BREAKPOINT_LINE}
            value={lineInput}
            disabled={breakpointAuthoringBlocked}
            onChange={(event) => {
              const bounded = event.target.value.slice(
                0,
                String(MAX_CODE_PAGE_BREAKPOINT_LINE).length,
              );
              setLineInput(
                Number(bounded) > MAX_CODE_PAGE_BREAKPOINT_LINE
                  ? String(MAX_CODE_PAGE_BREAKPOINT_LINE)
                  : bounded,
              );
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter") handleAddTyped();
            }}
            placeholder={t("code.debugLinePlaceholder")}
            data-testid="code-debug-line-input"
            aria-label={t("code.debugLinePlaceholder")}
            aria-invalid={typedBreakpointMessage ? "true" : undefined}
            aria-describedby={
              typedBreakpointMessage ? "code-debug-line-status" : undefined
            }
            className={styles.lineInput}
          />
          <button
            type="button"
            onClick={handleAddTyped}
            disabled={typedBreakpointDisabled}
            data-testid="code-debug-add-line"
            title={typedBreakpointMessage ?? t("code.debugAddBreakpoint")}
            className={officeDialogParts.compactButton}
          >
            {t("code.debugAddBreakpoint")}
          </button>
          {typedBreakpointMessage ? (
            <span
              id="code-debug-line-status"
              role="status"
              data-testid="code-debug-line-status"
              className={officeDialogParts.message}
              data-tone="warning"
            >
              {typedBreakpointMessage}
            </span>
          ) : null}
        </span>
        <label
          title={t("code.debugPauseOnUncaughtHint")}
          className={officeDialogParts.check}
        >
          <input
            type="checkbox"
            checked={pauseOnUncaughtExceptions}
            disabled={running || breakpointAuthoringBlocked}
            onChange={(event) =>
              setPauseOnUncaughtExceptions(event.currentTarget.checked)
            }
            data-testid="code-debug-pause-on-uncaught"
          />
          {t("code.debugPauseOnUncaught")}
        </label>
        <span className={`${officeDialogParts.row} ${styles.nowrap} ${styles.end}`}>
          <button
            type="button"
            onClick={handleDownloadConfiguration}
            data-testid="code-debug-config-download"
            title={t("code.download")}
            aria-label={t("code.download")}
            className={officeDialogParts.compactButton}
          >
            <Download size={12} aria-hidden="true" />
            {t("code.download")}
          </button>
          <button
            type="button"
            onClick={() => configurationImportInputRef.current?.click()}
            disabled={breakpointAuthoringBlocked}
            data-testid="code-debug-config-import"
            title={t("editor.import")}
            aria-label={t("editor.import")}
            className={officeDialogParts.compactButton}
          >
            <Upload size={12} aria-hidden="true" />
            {t("editor.import")}
          </button>
          <input
            ref={configurationImportInputRef}
            type="file"
            disabled={breakpointAuthoringBlocked}
            accept="application/json,.json"
            onChange={handleImportConfiguration}
            data-testid="code-debug-config-file"
            tabIndex={-1}
            aria-hidden="true"
            hidden
          />
        </span>
      </div>

      {configurationTransferError ? (
        <p
          role="alert"
          data-testid="code-debug-config-transfer-error"
          className={officeDialogParts.error}
        >
          {configurationTransferError}
        </p>
      ) : null}

      {pendingConfigurationImport && !breakpointAuthoringBlocked ? (
        <div
          role="status"
          data-testid="code-debug-config-import-confirmation"
          className={officeDialogParts.row}
        >
          <span className={officeDialogParts.message} data-tone="warning">
            {t("code.debugImportConfirm", {
              breakpoints:
                pendingConfigurationImport.configuration.breakpoints.length,
              watches:
                pendingConfigurationImport.configuration.watches?.length ?? 0,
            })}
          </span>
          <button
            type="button"
            onClick={confirmConfigurationImport}
            data-testid="code-debug-config-import-confirm"
            className={`${officeDialogParts.primaryButton} ${officeDialogParts.compactButton}`}
          >
            {t("code.debugImportApply")}
          </button>
          <button
            type="button"
            onClick={() => setPendingConfigurationImport(null)}
            data-testid="code-debug-config-import-cancel"
            className={officeDialogParts.compactButton}
          >
            {t("common.cancel")}
          </button>
        </div>
      ) : null}

      {persistedTimeline.length > 0 &&
      onTimelineSelect &&
      onTimelineDebugAgain &&
      onTimelineClear ? (
        <div className={styles.bleed}>
          <CodeDebugTimelinePanel
            entries={persistedTimeline}
            currentSourceFingerprint={currentSourceFingerprint}
            currentLanguage={language}
            selectedIndex={selectedTimelineIndex}
            debugLabel={t("code.debugRun")}
            unavailableLabel={t("document.capability.unavailable")}
            clearLabel={t("code.console.clear")}
            confirmClearLabel={t("code.confirmClearExecutionHistory")}
            cancelLabel={t("common.cancel")}
            entryLabel={(entry) =>
              entry.pauses[0]
                ? t("code.debugPausedAtLine", { line: entry.pauses[0].line })
                : t("code.debugRun")
            }
            onSelect={onTimelineSelect}
            onDebugAgain={onTimelineDebugAgain}
            onClear={onTimelineClear}
            actionsBlocked={running || hostBusy}
          />
          {selectedPersistedEntry ? (
            <div
              data-testid="code-debug-timeline-details"
              data-state={selectedPersistedEntryIsCurrent ? "current" : "stale"}
              onKeyDown={(event) => {
                if (
                  !selectedPersistedEntryIsCurrent ||
                  selectedPersistedEntryOwner === null ||
                  ![
                    "ArrowDown",
                    "ArrowLeft",
                    "ArrowRight",
                    "ArrowUp",
                    "End",
                    "Home",
                  ].includes(event.key)
                ) {
                  return;
                }
                const controls = Array.from(
                  event.currentTarget.querySelectorAll<HTMLButtonElement>(
                    "button[data-debug-history-nav-key]:not(:disabled)",
                  ),
                );
                if (controls.length === 0) return;
                const currentIndex = controls.indexOf(
                  event.target as HTMLButtonElement,
                );
                const fallbackIndex = controls.findIndex(
                  (control) =>
                    control.dataset.debugHistoryNavKey ===
                    selectedPersistedEntryRovingKey,
                );
                const originIndex =
                  currentIndex >= 0 ? currentIndex : fallbackIndex;
                let nextIndex: number;
                if (event.key === "Home") {
                  nextIndex = 0;
                } else if (event.key === "End") {
                  nextIndex = controls.length - 1;
                } else if (
                  event.key === "ArrowDown" ||
                  event.key === "ArrowRight"
                ) {
                  nextIndex = (Math.max(originIndex, 0) + 1) % controls.length;
                } else {
                  nextIndex =
                    (Math.max(originIndex, 0) - 1 + controls.length) %
                    controls.length;
                }
                const nextControl = controls[nextIndex];
                const key = nextControl?.dataset.debugHistoryNavKey;
                const line = Number(nextControl?.dataset.debugHistoryLine);
                if (
                  !nextControl ||
                  !key ||
                  !Number.isSafeInteger(line) ||
                  line < 1
                ) {
                  return;
                }
                event.preventDefault();
                event.stopPropagation();
                setTimelineDetailFocus({
                  owner: selectedPersistedEntryOwner,
                  key,
                });
                nextControl.focus();
                onNavigateToLine?.(line, true);
              }}
              className={`${officeDialogParts.row} ${styles.timelineDetails}`}
            >
              {selectedPersistedEntry.pauses.length > 0 ? (
                selectedPersistedEntry.pauses.map((pause, pauseIndex) => (
                  <span
                    key={`${pause.line}:${pause.kind}:${pauseIndex}`}
                    className={`${officeDialogParts.row} ${styles.nowrap}`}
                  >
                    <button
                      type="button"
                      disabled={!selectedPersistedEntryIsCurrent}
                      tabIndex={
                        selectedPersistedEntryIsCurrent &&
                        selectedPersistedEntryRovingKey === `pause:${pauseIndex}`
                          ? 0
                          : -1
                      }
                      data-debug-history-nav-key={`pause:${pauseIndex}`}
                      data-debug-history-line={pause.line}
                      onFocus={() => {
                        if (selectedPersistedEntryOwner === null) return;
                        setTimelineDetailFocus({
                          owner: selectedPersistedEntryOwner,
                          key: `pause:${pauseIndex}`,
                        });
                      }}
                      onClick={() => onNavigateToLine?.(pause.line, true)}
                      data-testid={`code-debug-timeline-pause-${pauseIndex}`}
                      title={
                        selectedPersistedEntryIsCurrent
                          ? t("code.debugPausedAtLine", { line: pause.line })
                          : t("document.capability.unavailable")
                      }
                      className={officeDialogParts.compactButton}
                    >
                      {t("code.debugBreakpointLabel", { line: pause.line })}
                    </button>
                    {pause.stack.map((frame, frameIndex) => (
                      <button
                        key={`${frame.name}:${frame.line}:${frameIndex}`}
                        type="button"
                        disabled={!selectedPersistedEntryIsCurrent}
                        tabIndex={
                          selectedPersistedEntryIsCurrent &&
                          selectedPersistedEntryRovingKey ===
                            `frame:${pauseIndex}:${frameIndex}`
                            ? 0
                            : -1
                        }
                        data-debug-history-nav-key={`frame:${pauseIndex}:${frameIndex}`}
                        data-debug-history-line={frame.line}
                        onFocus={() => {
                          if (selectedPersistedEntryOwner === null) return;
                          setTimelineDetailFocus({
                            owner: selectedPersistedEntryOwner,
                            key: `frame:${pauseIndex}:${frameIndex}`,
                          });
                        }}
                        onClick={() => onNavigateToLine?.(frame.line, true)}
                        data-testid={`code-debug-timeline-frame-${pauseIndex}-${frameIndex}`}
                        title={
                          selectedPersistedEntryIsCurrent
                            ? t("code.debugInspectFrame", { name: frame.name })
                            : t("document.capability.unavailable")
                        }
                        className={`${officeDialogParts.mono} ${styles.frameButton}`}
                      >
                        {`${frame.name}:${frame.line}`}
                      </button>
                    ))}
                  </span>
                ))
              ) : (
                <span role="status" className={officeDialogParts.message}>
                  {selectedPersistedEntryIsCurrent
                    ? t("code.noOutput")
                    : t("document.capability.unavailable")}
                </span>
              )}
            </div>
          ) : null}
        </div>
      ) : null}

      {/* Breakpoint chips */}
      {breakpoints.length > 0 ? (
        <div data-testid="code-debug-breakpoints" className={styles.chips}>
          {breakpoints.length > CODE_DEBUG_BREAKPOINT_INITIAL_LIMIT ||
          breakpointQuery.length > 0 ? (
            <div
              className={`${officeDialogParts.row} ${styles.nowrap} ${styles.fullRow}`}
            >
              <div
                className={`${officeDialogParts.inputGroup} ${styles.breakpointSearch}`}
                role="search"
              >
                <Search size={14} aria-hidden="true" />
                <input
                  type="search"
                  value={breakpointQuery}
                  maxLength={128}
                  onChange={(event) => {
                    setBreakpointQuery(event.currentTarget.value);
                    setBreakpointVisibleLimit(CODE_DEBUG_BREAKPOINT_INITIAL_LIMIT);
                  }}
                  onKeyDown={(event) => {
                    if (event.key !== "Escape" || breakpointQuery.length === 0)
                      return;
                    event.preventDefault();
                    setBreakpointQuery("");
                    setBreakpointVisibleLimit(CODE_DEBUG_BREAKPOINT_INITIAL_LIMIT);
                  }}
                  placeholder={t("code.debugBreakpointsSearch")}
                  aria-label={t("code.debugBreakpointsSearch")}
                  data-testid="code-debug-breakpoints-search"
                />
              </div>
              <span
                data-testid="code-debug-breakpoints-count"
                aria-live="polite"
                className={styles.count}
              >
                {visibleBreakpoints.length}/{filteredBreakpoints.length}
              </span>
            </div>
          ) : null}
          {visibleBreakpoints.map((line) => {
            // A conditional/hit-count breakpoint gets a distinct diamond vs a
            // plain circle, mirroring VSCode's different breakpoint marker.
            const isConditional =
              !!conditions[line]?.trim() || (isPython && !!hitConditions[line]?.trim());
            const unbound = unbindableBreakpoints.has(line);
            const conditionInvalid = Boolean(
              conditions[line] && !isConditionValid(conditions[line]),
            );
            const hitInvalid = Boolean(
              hitConditions[line] && !parseHitCondition(hitConditions[line]),
            );
            return (
            <span
              key={line}
              data-testid={`code-debug-breakpoint-${line}`}
              data-conditional={isConditional ? "true" : undefined}
              data-bound={unbound ? "false" : "true"}
              title={unbound ? t("code.debugBreakpointUnbound") : undefined}
              // An unverified breakpoint is drawn the way VS Code draws it:
              // present but greyed out, so it never reads as armed.
              className={styles.breakpointChip}
            >
              <span
                aria-hidden="true"
                data-testid={`code-debug-glyph-${line}`}
                data-shape={isConditional ? "conditional" : "plain"}
                title={
                  unbound
                    ? t("code.debugBreakpointUnbound")
                    : isConditional
                      ? t("code.debugConditionHint")
                      : t("code.debugBreakpointLabel", { line })
                }
                className={styles.breakpointGlyph}
              >
                {isConditional ? (
                  <Diamond
                    size={9}
                    fill={unbound ? "none" : "currentColor"}
                    aria-hidden="true"
                  />
                ) : (
                  <Circle
                    size={9}
                    fill={unbound ? "none" : "currentColor"}
                    aria-hidden="true"
                  />
                )}
              </span>
              {onNavigateToLine ? (
                <button
                  type="button"
                  onClick={() => onNavigateToLine(line, true)}
                  data-testid={`code-debug-breakpoint-line-${line}`}
                  title={t("code.debugBreakpointLabel", { line })}
                  className={styles.chipLine}
                >
                  {t("code.debugBreakpointLabel", { line })}
                </button>
              ) : (
                t("code.debugBreakpointLabel", { line })
              )}
              <input
                value={conditions[line] ?? ""}
                disabled={breakpointAuthoringBlocked}
                onChange={(event) => setCondition(line, event.target.value)}
                maxLength={MAX_CODE_PAGE_BREAKPOINT_EXPRESSION_CHARS}
                placeholder={t("code.debugConditionPlaceholder")}
                title={t("code.debugConditionHint")}
                data-testid={`code-debug-condition-${line}`}
                spellCheck={false}
                aria-invalid={conditionInvalid ? "true" : undefined}
                className={`${officeDialogParts.mono} ${styles.chipInput} ${styles.conditionInput}`}
              />
              <input
                value={logMessages[line] ?? ""}
                disabled={breakpointAuthoringBlocked}
                onChange={(event) => setLogMessage(line, event.target.value)}
                maxLength={MAX_LOGPOINT_MESSAGE_CHARS}
                placeholder={t("code.debugLogpointPlaceholder")}
                title={t("code.debugLogpointHint", { name: "{name}" })}
                data-testid={`code-debug-logpoint-${line}`}
                spellCheck={false}
                className={`${officeDialogParts.mono} ${styles.chipInput} ${styles.logpointInput}`}
              />
              {isPython ? (
                <input
                  value={hitConditions[line] ?? ""}
                  disabled={breakpointAuthoringBlocked}
                  onChange={(event) => setHitCondition(line, event.target.value)}
                  maxLength={MAX_CODE_PAGE_BREAKPOINT_EXPRESSION_CHARS}
                  placeholder={t("code.debugHitPlaceholder")}
                  title={t("code.debugHitHint")}
                  data-testid={`code-debug-hit-${line}`}
                  spellCheck={false}
                  aria-invalid={hitInvalid ? "true" : undefined}
                  className={`${officeDialogParts.mono} ${styles.chipInput} ${styles.hitInput}`}
                />
              ) : null}
              <button
                type="button"
                onClick={() => removeBreakpoint(line)}
                disabled={configurationReadOnly}
                aria-label={t("code.debugRemoveBreakpoint", { line })}
                title={t("code.debugRemoveBreakpoint", { line })}
                className={`${officeDialogParts.iconButton} ${styles.chipRemove}`}
              >
                <X size={11} />
              </button>
            </span>
            );
          })}
          {filteredBreakpoints.length === 0 ? (
            <p
              data-testid="code-debug-breakpoints-search-empty"
              role="status"
              className={`${officeDialogParts.message} ${styles.fullRow}`}
            >
              {t("code.findInFiles.empty")}
            </p>
          ) : null}
          {hiddenBreakpointCount > 0 ? (
            <div className={`${officeDialogParts.row} ${styles.fullRow}`}>
              <button
                type="button"
                data-testid="code-debug-breakpoints-show-more"
                onClick={() =>
                  setBreakpointVisibleLimit((current) =>
                    Math.min(
                      filteredBreakpoints.length,
                      current + CODE_DEBUG_BREAKPOINT_INITIAL_LIMIT,
                    ),
                  )
                }
                className={officeDialogParts.compactButton}
              >
                {t("code.debugBreakpointsShowMore")}
              </button>
              <span role="status" className={officeDialogParts.hint}>
                {t("code.debugBreakpointsHidden", {
                  count: hiddenBreakpointCount,
                })}
              </span>
            </div>
          ) : null}
          {unbindableBreakpoints.size > 0 ? (
            // A hollow dot is easy to miss, and "why did it not stop?" is the
            // most expensive question a debugger can leave unanswered.
            <p
              data-testid="code-debug-unbound-note"
              className={`${officeDialogParts.message} ${styles.fullRow}`}
            >
              {t("code.debugBreakpointUnboundNote")}
            </p>
          ) : null}
        </div>
      ) : (
        <p className={officeDialogParts.message}>
          {t("code.debugNoBreakpoints")}
        </p>
      )}

      {pauseHistory.length > 0 ? (
        <section
          aria-label={t("code.debugRun")}
          data-testid="code-debug-live-timeline"
          className={`${officeDialogParts.tabs} ${styles.pauseStrip}`}
        >
          {pauseHistory.map((snapshot, index) => {
            const live = snapshot === activePaused;
            const selected = paused === snapshot;
            return (
              <button
                key={`${snapshot.step}:${snapshot.line}:${index}`}
                type="button"
                onClick={() => handleInspectPause(index)}
                aria-pressed={selected}
                data-testid={`code-debug-live-pause-${index}`}
                data-state={live ? "live" : "historical"}
                title={t("code.debugPausedAtLine", { line: snapshot.line })}
                className={officeDialogParts.tab}
              >
                {t("code.debugBreakpointLabel", { line: snapshot.line })}
              </button>
            );
          })}
        </section>
      ) : null}

      <DebugWatchPanel
        watches={resolvedWatches}
        valuesAvailable={paused !== null}
        draft={watchDraft}
        editingWatchId={editingWatchId}
        configurationReadOnly={configurationReadOnly}
        capacityReached={watchCapacityReached}
        commitDisabled={watchCommitDisabled}
        validationMessage={watchValidationMessage}
        labels={{
          title: t("code.debugWatchTitle"),
          limit: t("code.debugWatchLimit", { max: MAX_DEBUG_WATCHES }),
          placeholder: t("code.debugWatchPlaceholder"),
          add: t("code.debugWatchAdd"),
          submit: t("code.rename.submit"),
          cancel: t("code.rename.cancel"),
          outOfScope: t("code.debugWatchOutOfScope"),
          valueUnavailable: t("code.debugWatchNotPaused"),
          remove: (name) => t("code.debugWatchRemove", { name }),
        }}
        onDraftChange={setWatchDraft}
        onCommit={commitWatch}
        onCancel={cancelWatchEdit}
        onEdit={(watch) => {
          setEditingWatchId(watch.id);
          setWatchDraft(watch.name);
        }}
        onRemove={(watch) => {
          if (editingWatchId === watch.id) cancelWatchEdit();
          updateConfiguration((current) => {
            const currentWatches = current.watches ?? [];
            const next = removeDebugWatch(currentWatches, watch.id);
            return next === currentWatches
              ? current
              : { ...current, watches: next };
          });
        }}
      />

      {/* Paused panel: current line + locals snapshot + Continue / Step */}
      {paused ? (
        <div data-testid="code-debug-paused" className={styles.pausedBox}>
          <div className={officeDialogParts.row}>
            {onNavigateToLine ? (
              <button
                type="button"
                onClick={() => onNavigateToLine(paused.line, true)}
                data-testid="code-debug-paused-line"
                title={t("code.debugPausedAtLine", { line: paused.line })}
                className={styles.pausedLine}
              >
                {t("code.debugPausedAtLine", { line: paused.line })}
              </button>
            ) : (
              <strong
                data-testid="code-debug-paused-line"
                className={styles.pausedTitle}
              >
                {t("code.debugPausedAtLine", { line: paused.line })}
              </strong>
            )}
            <button
              type="button"
              onClick={handleContinue}
              disabled={!pauseIsLive}
              data-testid="code-debug-continue"
              title={t("code.debugContinue")}
              className={officeDialogParts.compactButton}
            >
              <ChevronRight size={12} />
              {t("code.debugContinue")}
            </button>
            {isPython ? (
              <>
                <button
                  type="button"
                  onClick={handleStepOver}
                  disabled={!pauseIsLive}
                  data-testid="code-debug-step-over"
                  title={t("code.debugStepOver")}
                  className={officeDialogParts.compactButton}
                >
                  <SkipForward size={12} />
                  {t("code.debugStepOver")}
                </button>
                <button
                  type="button"
                  onClick={handleStepInto}
                  disabled={!pauseIsLive}
                  data-testid="code-debug-step-into"
                  title={t("code.debugStepInto")}
                  className={officeDialogParts.compactButton}
                >
                  <ArrowDownToLine size={12} />
                  {t("code.debugStepInto")}
                </button>
                <button
                  type="button"
                  onClick={handleStepOut}
                  disabled={!pauseIsLive}
                  data-testid="code-debug-step-out"
                  title={t("code.debugStepOut")}
                  className={officeDialogParts.compactButton}
                >
                  <ArrowUpFromLine size={12} />
                  {t("code.debugStepOut")}
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={handleStep}
                disabled={!pauseIsLive}
                data-testid="code-debug-step"
                title={t("code.debugStep")}
                className={officeDialogParts.compactButton}
              >
                <SkipForward size={12} />
                {t("code.debugStep")}
              </button>
            )}
          </div>
          {pausedIsUncaughtException ? (
            <p
              role="alert"
              data-testid="code-debug-uncaught-exception"
              className={`${officeDialogParts.error} ${styles.preWrap}`}
            >
              {t("code.debugPausedOnUncaught", {
                message: pausedException ?? t("code.noOutput"),
              })}
            </p>
          ) : null}
          {isPython ? (
            // Python cannot truly block mid-run here, so a debug run executes
            // ONCE to completion under a tracer and this panel replays what it
            // recorded. Stepping therefore feels identical to the JavaScript
            // debugger while meaning something different: by the time you read
            // these locals, the rest of the program has already run. Someone
            // who sets a breakpoint to catch state BEFORE a destructive line
            // deserves to know that, rather than discovering it afterwards.
            <p
              data-testid="code-debug-replay-note"
              className={officeDialogParts.message}
              data-tone="warning"
            >
              {t("code.debugReplayNote")}
            </p>
          ) : null}
          {pausedStack ? (
            <div
              data-testid="code-debug-stack"
              onKeyDown={(event) => {
                if (event.altKey || event.ctrlKey || event.metaKey) return;
                const target = event.target;
                if (!(target instanceof HTMLButtonElement)) return;
                const controls = Array.from(
                  event.currentTarget.querySelectorAll<HTMLButtonElement>(
                    "button[data-debug-frame-index]",
                  ),
                );
                const currentIndex = controls.indexOf(target);
                if (currentIndex < 0 || controls.length === 0) return;
                let nextIndex: number;
                if (event.key === "ArrowDown" || event.key === "ArrowRight") {
                  nextIndex = (currentIndex + 1) % controls.length;
                } else if (
                  event.key === "ArrowUp" ||
                  event.key === "ArrowLeft"
                ) {
                  nextIndex =
                    (currentIndex - 1 + controls.length) % controls.length;
                } else if (event.key === "Home") {
                  nextIndex = 0;
                } else if (event.key === "End") {
                  nextIndex = controls.length - 1;
                } else {
                  return;
                }
                const next = controls[nextIndex];
                const frameIndex = Number(next?.dataset.debugFrameIndex);
                if (!next || !Number.isInteger(frameIndex)) return;
                event.preventDefault();
                event.stopPropagation();
                handleSelectFrame(frameIndex);
                next.focus();
              }}
              className={`${officeDialogParts.mono} ${styles.stack}`}
            >
              {pausedStack.map((frame, frameIndex) => (
                <span
                  key={`${frame.name}:${frame.line}:${frameIndex}`}
                  className={styles.stackFrame}
                >
                  {frameIndex > 0 ? (
                    <span className={styles.muted}>←</span>
                  ) : null}
                  {/* Click a frame to inspect its locals (Python record-and-replay). */}
                  <button
                    type="button"
                    onClick={() => handleSelectFrame(frameIndex)}
                    data-testid={`code-debug-stack-frame-${frameIndex}`}
                    data-debug-frame-index={frameIndex}
                    aria-pressed={selectedFrame === frameIndex}
                    tabIndex={selectedFrame === frameIndex ? 0 : -1}
                    title={t("code.debugInspectFrame", { name: frame.name })}
                    className={styles.frameButton}
                  >
                    {`${frame.name}:${frame.line}`}
                  </button>
                </span>
              ))}
            </div>
          ) : null}
          {pausedConditionError ? (
            <p
              data-testid="code-debug-condition-error"
              className={officeDialogParts.error}
            >
              {t("code.debugConditionError", { message: pausedConditionError })}
            </p>
          ) : null}
          {pausedConditionNote ? (
            <p
              data-testid="code-debug-condition-note"
              className={officeDialogParts.message}
              data-tone="warning"
            >
              {t("code.debugConditionNote", { message: pausedConditionNote })}
            </p>
          ) : null}
          {pausedTruncated ? (
            <p
              data-testid="code-debug-truncated"
              className={officeDialogParts.message}
              data-tone="warning"
            >
              {t("code.debugTruncatedNote")}
            </p>
          ) : null}
          {localsEntries.length > 0 ? (
            <div className={`${officeDialogParts.well} ${styles.tableWell}`}>
              <table
                data-testid="code-debug-locals"
                className={`${officeDialogParts.table} ${officeDialogParts.mono}`}
              >
                <thead>
                  <tr>
                    <th scope="col">{t("code.debugLocalsName")}</th>
                    <th scope="col">{t("code.debugLocalsValue")}</th>
                  </tr>
                </thead>
                <tbody>
                  {localsEntries.map(([name, value]) => (
                    <tr key={name}>
                      <td className={styles.muted}>{name}</td>
                      <td>{value}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <span className={officeDialogParts.message}>
              {t("code.debugNoLocals")}
            </span>
          )}

          {snapshotEvaluationAvailable ? (
            <section
              data-testid="code-debug-console"
              aria-label={t("code.console.title")}
              className={`${officeDialogParts.group} ${styles.debugConsole}`}
            >
              <div className={officeDialogParts.header}>
                <strong className={officeDialogParts.sectionTitle}>
                  {t("code.console.title")}
                </strong>
                <button
                  type="button"
                  data-testid="code-debug-console-clear"
                  onClick={() =>
                    setEvaluationHistory((current) =>
                      current.length === 0 ? current : [],
                    )
                  }
                  disabled={evaluationHistory.length === 0}
                  className={officeDialogParts.compactButton}
                >
                  {t("code.console.clear")}
                </button>
              </div>
              <form
                onSubmit={(event) => {
                  event.preventDefault();
                  void handleEvaluateSnapshot();
                }}
                className={`${officeDialogParts.row} ${styles.nowrap}`}
              >
                <input
                  data-testid="code-debug-console-input"
                  value={evaluationDraft}
                  onChange={(event) =>
                    setEvaluationDraft(
                      event.target.value.slice(
                        0,
                        MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS,
                      ),
                    )
                  }
                  maxLength={MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS}
                  placeholder={t("code.console.placeholder")}
                  aria-label={t("code.console.placeholder")}
                  disabled={evaluationBusy}
                  className={`${officeDialogParts.mono} ${styles.grow}`}
                />
                <button
                  type="submit"
                  data-testid="code-debug-console-submit"
                  disabled={
                    evaluationBusy ||
                    normalizeDebugSnapshotExpression(evaluationDraft) === null
                  }
                  className={`${officeDialogParts.primaryButton} ${officeDialogParts.compactButton}`}
                >
                  {evaluationBusy
                    ? t("code.console.busy")
                    : t("code.console.title")}
                </button>
              </form>
              <div
                data-testid="code-debug-console-history"
                aria-live="polite"
                className={`${officeDialogParts.well} ${officeDialogParts.mono} ${styles.codeWell}`}
              >
                {evaluationHistory.length === 0 ? (
                  <span className={styles.muted}>
                    {t("code.console.empty")}
                  </span>
                ) : (
                  evaluationHistory.map((entry) => (
                    <div
                      key={entry.id}
                      data-testid="code-debug-console-entry"
                      data-status={entry.status}
                      role={entry.status === "error" ? "alert" : undefined}
                      className={styles.evaluationEntry}
                    >
                      <div className={styles.muted}>
                        {`> ${entry.expression}`}
                      </div>
                      <div
                        className={
                          entry.status === "error"
                            ? `${styles.outputLine} ${styles.failed}`
                            : styles.outputLine
                        }
                      >
                        {entry.value}
                      </div>
                    </div>
                  ))
                )}
              </div>
            </section>
          ) : null}
        </div>
      ) : null}

      {/* What the logpoints reported during this run. */}
      {logLines.length > 0 ? (
        <div
          data-testid="code-debug-logpoint-output"
          className={`${officeDialogParts.well} ${officeDialogParts.mono} ${styles.codeWell}`}
        >
          {logLines.map((entry, index) => (
            <div
              key={`${entry.line}-${index}`}
              data-testid="code-debug-logpoint-line"
              data-line={entry.line}
            >
              <span className={`${styles.muted} ${styles.logpointLabel}`}>
                {t("code.debugBreakpointLabel", { line: entry.line })}
              </span>
              {entry.text}
            </div>
          ))}
        </div>
      ) : null}

      {/* Final result of a finished debug run (logs / error / value). */}
      {result ? (
        <div
          data-testid="code-debug-result"
          className={`${officeDialogParts.well} ${officeDialogParts.mono} ${styles.codeWell} ${styles.resultWell}`}
        >
          {result.error ? (
            <pre className={`${styles.outputLine} ${styles.failed}`}>
              <AnsiText text={result.error} />
            </pre>
          ) : null}
          {result.logs.map((line, index) => {
            const stream =
              result.logStreams?.[index] === "stderr" ? "stderr" : "stdout";
            return (
            <pre
              key={index}
              data-testid={`code-debug-result-${stream}`}
              data-stream={stream}
              className={
                stream === "stderr"
                  ? `${styles.outputLine} ${styles.failed}`
                  : styles.outputLine
              }
            >
              <AnsiText text={line} />
            </pre>
            );
          })}
          {result.result !== undefined ? (
            <pre className={`${styles.outputLine} ${styles.accent}`}>
              {"⇒ "}
              <AnsiText text={result.result} />
            </pre>
          ) : null}
          {/* A raising condition that produced no pauses still surfaces its
              one-time note here so the user knows a breakpoint was skipped. */}
          {!paused &&
          typeof (result as { debugConditionNote?: string }).debugConditionNote ===
            "string" ? (
            <pre
              data-testid="code-debug-result-condition-note"
              className={`${styles.outputLine} ${styles.warning}`}
            >
              {t("code.debugConditionNote", {
                message: (result as { debugConditionNote?: string })
                  .debugConditionNote as string,
              })}
            </pre>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}

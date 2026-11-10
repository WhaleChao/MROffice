/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { BackgroundCalculationClient, type BackgroundCalculationClientOptions } from "@mrnote/core/web/lib/spreadsheet/background-calculation-client.ts";
import { BackgroundCalculationHost, type BackgroundCalculationHostBinding } from "@mrnote/core/web/lib/spreadsheet/background-calculation-host.ts";
import { BackgroundCalculationCancelledError, sameCalculationToken, type BackgroundCalculationResult, type CalculationRuleRequest } from "@mrnote/core/web/lib/spreadsheet/background-calculation-protocol.ts";
import type { WorksheetCalculationToken, WorksheetFormulaCoordinate } from "@mrnote/core/web/lib/spreadsheet/calculation-mode.ts";
import type { EngineValue, WorkbookFormulaResultSeed } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { settleCalculationPivots, type CalculationPivots } from "@mrnote/core/web/lib/spreadsheet/calculation-pivots.ts";
import type { CalculationExternalDatum } from "@mrnote/core/web/lib/spreadsheet/external-data.ts";
import { type ExcelDateSystemArg } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

export interface RunSpreadsheetCalculationOptions {
  /** Exact inputs captured by the host for this operation, never recovered from an older render. */
  readonly bindings: readonly BackgroundCalculationHostBinding[];
  /** Created by the existing beginWorksheetCalculation authority. */
  readonly token: WorksheetCalculationToken;
  /** Workbook scope: every requested formula. Worksheet scope: target sheet only. */
  readonly coordinates: readonly WorksheetFormulaCoordinate[];
  readonly formulaResultSeed?: WorkbookFormulaResultSeed;
  /**
   * The rule formulas on screen (conditional formats, validation, link
   * locations), answered in the same pass as the cells (`ruleValues`), so
   * the grid's formats come from the calculation its values do (live audit
   * B470), as the automatic calculation's do.
   */
  readonly rules?: readonly CalculationRuleRequest[];
  /** What WEBSERVICE / IMPORTRANGE read, as the browser holds it
   * (external-data); the answer names the calls read without it. */
  readonly externalData?: readonly CalculationExternalDatum[];
  readonly workbookName?: string;
  /** The folder of the workbook's location (CELL("filename"), X-194). */
  readonly workbookFolder?: string;
  /** The workbook reads serials in the 1904 date system. */
  readonly date1904?: ExcelDateSystemArg;
  /** Session lifecycle only; never cloned or sent through the Worker protocol. */
  readonly signal?: AbortSignal;
  /** Live authority fields; its targetSheetId may remain null for the mounted host. */
  readonly getAuthorityToken: () => WorksheetCalculationToken | null;
  /**
   * The PivotTables as the grid shows them now (the session starts holding
   * them), and how to read them from an answer: the calculation settles them
   * like an automatic one (at most two replacements; `pivotStale` after).
   */
  readonly pivots?: CalculationPivots | null;
  readonly derivePivots?: (results: ReadonlyMap<string, EngineValue>) => CalculationPivots | null;
  /** Injectable transport/scheduler settings, with normal production defaults. */
  readonly clientOptions?: Pick<BackgroundCalculationClientOptions, "createWorker" | "workerTimeoutMs">;
}

/**
 * F9 / Shift+F9 preparation only. The host first cancels its automatic job and
 * advances existing calculation authority, then passes the captured bindings.
 * It alone publishes the resulting ledger. No model/authority writes occur here.
 */
export async function runSpreadsheetCalculation(options: RunSpreadsheetCalculationOptions): Promise<BackgroundCalculationResult & { pivotStale?: boolean }> {
  const token = { ...options.token };
  const getCurrentToken = (): WorksheetCalculationToken | null => {
    if (options.signal?.aborted) return null;
    const live = options.getAuthorityToken();
    if (!live || options.signal?.aborted) return null;
    // targetSheetId selects this calculation's scope; it is not a separate
    // workbook revision. Never replace any live version field with the request.
    const scoped = { ...live, targetSheetId: token.targetSheetId };
    return sameCalculationToken(scoped, token) ? scoped : null;
  };
  if (!getCurrentToken()) throw new BackgroundCalculationCancelledError();
  const host = new BackgroundCalculationHost();
  let client: BackgroundCalculationClient | null = null;
  const abort = () => { client?.cancel(); client?.dispose(); };
  options.signal?.addEventListener("abort", abort, { once: true });
  try {
    if (!getCurrentToken()) throw new BackgroundCalculationCancelledError();
    const prepared = host.reset(token, options.bindings, { name: options.workbookName, folder: options.workbookFolder }, options.date1904);
    if (!prepared.ok) throw new Error(`Explicit calculation input unavailable: ${prepared.reason}`);
    if (!getCurrentToken()) throw new BackgroundCalculationCancelledError();
    client = new BackgroundCalculationClient({ ...options.clientOptions, getCurrentToken });
    client.reset(options.pivots ? { ...prepared.snapshot, pivots: options.pivots } : prepared.snapshot);
    if (!getCurrentToken()) throw new BackgroundCalculationCancelledError();
    const request = { token, coordinates: options.coordinates, formulaResultSeed: options.formulaResultSeed,
      ...(options.rules?.length ? { rules: options.rules } : {}),
      ...(options.externalData ? { externalData: options.externalData } : {}) };
    let result: BackgroundCalculationResult & { pivotStale?: boolean } = await client.calculate(request);
    if (options.derivePivots) {
      const settled = await settleCalculationPivots(client, request, result, options.derivePivots, () => Boolean(getCurrentToken()));
      if (!settled) throw new BackgroundCalculationCancelledError();
      result = settled.pivotStale ? { ...settled.answer, pivotStale: true } : settled.answer;
    }
    if (!getCurrentToken() || !sameCalculationToken(result.token, token)) throw new BackgroundCalculationCancelledError();
    return result;
  } finally { options.signal?.removeEventListener("abort", abort); client?.dispose(); host.clear(); }
}

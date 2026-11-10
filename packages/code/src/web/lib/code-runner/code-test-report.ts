/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  CODE_TEST_RUN_TIMELINE_KEY,
  CODE_TEST_RUN_TIMELINE_VERSION,
  restoreCodeTestRunTimeline,
  type PersistedCodeTestRunTimelineEntry,
} from "./code-test-explorer";

export const CODE_TEST_JUNIT_XML_UTF8_BYTE_LIMIT = 2_000_000;

export interface CodeTestJUnitReport {
  readonly filename: string;
  readonly mimeType: "application/junit+xml;charset=utf-8";
  readonly content: string;
}

function xmlSafe(value: string): string {
  return value
    .replace(
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/g,
      "\ufffd",
    )
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function seconds(milliseconds: number): string {
  return (milliseconds / 1_000).toFixed(3);
}

function reportFilename(completedAt: string, runId: string): string {
  const timestamp = completedAt
    .replace(/\.\d{3}Z$/, "Z")
    .replace(/[^0-9A-Za-z]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 32);
  const safeRunId = runId
    .replace(/[^0-9A-Za-z_-]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 64);
  return `code-tests-${timestamp || "run"}-${safeRunId || "results"}.junit.xml`;
}

/**
 * Serialize one normalized Test Explorer history entry as a bounded JUnit XML
 * report. The report is historical evidence of that run, written from the
 * run's own receipt (its sources' paths and languages, its results and
 * times), never execution authority: it never becomes a rerun. Editing the
 * source after the run no longer refuses the report ("history unavailable"
 * for every run older than the last keystroke); a malformed entry still
 * fails closed.
 */
export async function createCodeTestJUnitReport(
  entryValue: unknown,
): Promise<CodeTestJUnitReport | null> {
  try {
    const entries = restoreCodeTestRunTimeline({
      [CODE_TEST_RUN_TIMELINE_KEY]: {
        version: CODE_TEST_RUN_TIMELINE_VERSION,
        entries: [entryValue],
      },
    });
    const entry: PersistedCodeTestRunTimelineEntry | undefined = entries[0];
    if (entries.length !== 1 || !entry) return null;

    const { receipt } = entry;
    const suites = receipt.sources.map((source) => {
      const results = receipt.results.filter(
        (result) => result.pageId === source.pageId,
      );
      if (results.length === 0) return null;
      const failures = results.filter((result) => result.status === "failed");
      const durationMs = results.reduce(
        (total, result) => total + result.durationMs,
        0,
      );
      const cases = results
        .map((result) => {
          const failure =
            result.status === "failed"
              ? `\n      <failure message="${xmlSafe(result.message ?? "Test failed")}" type="test-failure">${xmlSafe(result.message ?? "Test failed")}</failure>`
              : "";
          return [
            `    <testcase name="${xmlSafe(result.name)}" classname="${xmlSafe(result.path)}" file="${xmlSafe(result.path)}" line="${result.line}" time="${seconds(result.durationMs)}">`,
            `      <properties><property name="language" value="${xmlSafe(result.language)}"/><property name="testId" value="${xmlSafe(result.testId)}"/></properties>${failure}`,
            "    </testcase>",
          ].join("\n");
        })
        .join("\n");
      return [
        `  <testsuite name="${xmlSafe(source.path)}" package="${xmlSafe(source.language)}" tests="${results.length}" failures="${failures.length}" errors="0" skipped="0" time="${seconds(durationMs)}" timestamp="${xmlSafe(receipt.completedAt)}">`,
        cases,
        "  </testsuite>",
      ].join("\n");
    });
    if (suites.some((suite) => suite === null)) return null;

    const failures = receipt.results.filter(
      (result) => result.status === "failed",
    ).length;
    const durationMs = receipt.results.reduce(
      (total, result) => total + result.durationMs,
      0,
    );
    const content = [
      '<?xml version="1.0" encoding="UTF-8"?>',
      `<testsuites name="MRNote Code Tests" id="${xmlSafe(receipt.runId)}" tests="${receipt.results.length}" failures="${failures}" errors="0" skipped="0" time="${seconds(durationMs)}" timestamp="${xmlSafe(receipt.completedAt)}">`,
      ...(suites as string[]),
      "</testsuites>",
      "",
    ].join("\n");
    if (
      typeof TextEncoder === "undefined" ||
      new TextEncoder().encode(content).byteLength >
        CODE_TEST_JUNIT_XML_UTF8_BYTE_LIMIT
    ) {
      return null;
    }
    return {
      filename: reportFilename(receipt.completedAt, receipt.runId),
      mimeType: "application/junit+xml;charset=utf-8",
      content,
    };
  } catch {
    return null;
  }
}

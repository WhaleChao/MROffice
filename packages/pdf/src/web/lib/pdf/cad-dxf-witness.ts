/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PdfCadDxfWitness {
  readonly pageCount: number;
  readonly entityCount: number;
  readonly version: "AC1027";
  readonly units: "millimeters";
}

interface DxfPair {
  readonly code: string;
  readonly value: string;
}

function parseAsciiPairs(bytes: Uint8Array): readonly DxfPair[] | null {
  if (bytes.byteLength < 20 || bytes.some((value) => value > 0x7f)) return null;
  const source = new TextDecoder("ascii", { fatal: true }).decode(bytes);
  if (
    !source.startsWith("0\r\nSECTION\r\n") ||
    !source.endsWith("0\r\nEOF\r\n")
  ) {
    return null;
  }
  const lines = source.slice(0, -2).split("\r\n");
  if (lines.length % 2 !== 0) return null;
  const pairs: DxfPair[] = [];
  for (let index = 0; index < lines.length; index += 2) {
    const code = lines[index];
    const value = lines[index + 1];
    if (!/^-?\d+$/.test(code) || value.length > 512) return null;
    pairs.push(Object.freeze({ code, value }));
  }
  return Object.freeze(pairs);
}

function exactHeaderValue(
  pairs: readonly DxfPair[],
  variable: string,
  valueCode: string,
): string | null {
  const matches: string[] = [];
  for (let index = 0; index + 1 < pairs.length; index += 1) {
    if (
      pairs[index].code === "9" &&
      pairs[index].value === variable &&
      pairs[index + 1].code === valueCode
    ) {
      matches.push(pairs[index + 1].value);
    }
  }
  return matches.length === 1 ? matches[0] : null;
}

function entitySectionBounds(
  pairs: readonly DxfPair[],
): Readonly<{ start: number; end: number }> | null {
  let match: { start: number; end: number } | null = null;
  for (let index = 0; index + 1 < pairs.length; index += 1) {
    if (
      pairs[index].code !== "0" ||
      pairs[index].value !== "SECTION" ||
      pairs[index + 1].code !== "2" ||
      pairs[index + 1].value !== "ENTITIES"
    ) {
      continue;
    }
    const start = index + 2;
    const end = pairs.findIndex(
      (pair, pairIndex) =>
        pairIndex >= start && pair.code === "0" && pair.value === "ENDSEC",
    );
    if (end < start || match !== null) return null;
    match = { start, end };
  }
  return match ? Object.freeze(match) : null;
}

export function verifyPdfCadDxfOutput(
  bytes: Uint8Array,
  expected: Readonly<{ pageCount: number; entityCount: number }>,
): PdfCadDxfWitness | null {
  if (
    !Number.isSafeInteger(expected.pageCount) ||
    expected.pageCount < 1 ||
    expected.pageCount > 200 ||
    !Number.isSafeInteger(expected.entityCount) ||
    expected.entityCount < 1 ||
    expected.entityCount > 100_000
  ) {
    return null;
  }
  const pairs = parseAsciiPairs(bytes);
  if (
    !pairs ||
    exactHeaderValue(pairs, "$ACADVER", "1") !== "AC1027" ||
    exactHeaderValue(pairs, "$INSUNITS", "70") !== "4" ||
    exactHeaderValue(pairs, "$MEASUREMENT", "70") !== "1"
  ) {
    return null;
  }

  const expectedLayers = new Set(
    Array.from(
      { length: expected.pageCount },
      (_, index) => `PDF_PAGE_${String(index + 1).padStart(4, "0")}`,
    ),
  );
  const declaredLayers = new Set<string>();
  for (let index = 0; index + 1 < pairs.length; index += 1) {
    if (pairs[index].code !== "0" || pairs[index].value !== "LAYER") continue;
    if (pairs[index + 1].code !== "2") return null;
    const name = pairs[index + 1].value;
    if (!expectedLayers.has(name) || declaredLayers.has(name)) return null;
    declaredLayers.add(name);
  }
  if (declaredLayers.size !== expectedLayers.size) return null;

  const bounds = entitySectionBounds(pairs);
  if (!bounds) return null;
  let entityCount = 0;
  for (let index = bounds.start; index < bounds.end; ) {
    const first = pairs[index];
    if (
      first.code !== "0" ||
      (first.value !== "LINE" && first.value !== "LWPOLYLINE")
    ) {
      return null;
    }
    const next = pairs.findIndex(
      (pair, pairIndex) => pairIndex > index && pair.code === "0",
    );
    const end = next < 0 || next > bounds.end ? bounds.end : next;
    const layers = pairs
      .slice(index + 1, end)
      .filter((pair) => pair.code === "8")
      .map((pair) => pair.value);
    if (
      layers.length !== 1 ||
      !expectedLayers.has(layers[0]) ||
      !declaredLayers.has(layers[0])
    ) {
      return null;
    }
    entityCount += 1;
    index = end;
  }
  if (entityCount !== expected.entityCount) return null;
  return Object.freeze({
    pageCount: expected.pageCount,
    entityCount,
    version: "AC1027",
    units: "millimeters",
  });
}

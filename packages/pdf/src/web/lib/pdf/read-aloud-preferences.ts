/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizePdfLanguageTag } from "./document-properties";

export const PDF_READ_ALOUD_PREFERENCE_SCHEMA_VERSION = 1 as const;
export const PDF_READ_ALOUD_PREFERENCE_STORAGE_KEY =
  "mrnote:pdf:read-aloud-preferences:v1";

const PDF_READ_ALOUD_RATE_MIN = 0.5;
const PDF_READ_ALOUD_RATE_MAX = 2;
const PDF_READ_ALOUD_VOICE_URI_MAX_LENGTH = 256;
const PREFERENCE_KEYS = [
  "schemaVersion",
  "lang",
  "rate",
  "voiceURI",
  "followDocumentLanguage",
] as const;

export interface PdfReadAloudPreferences {
  readonly lang: string;
  readonly rate: number;
  readonly voiceURI: string | null;
  readonly followDocumentLanguage: boolean;
}

export interface PdfReadAloudPreferenceStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

interface PersistedPdfReadAloudPreferences extends PdfReadAloudPreferences {
  readonly schemaVersion: typeof PDF_READ_ALOUD_PREFERENCE_SCHEMA_VERSION;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function hasExactKeys(value: Record<string, unknown>): boolean {
  const actual = Object.keys(value).sort();
  const expected = [...PREFERENCE_KEYS].sort();
  return (
    actual.length === expected.length &&
    actual.every((key, index) => key === expected[index])
  );
}

function normalizeRate(value: unknown): number | null {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < PDF_READ_ALOUD_RATE_MIN ||
    value > PDF_READ_ALOUD_RATE_MAX
  ) {
    return null;
  }
  const tenths = Math.round(value * 10);
  return Math.abs(value * 10 - tenths) < Number.EPSILON * 10
    ? tenths / 10
    : null;
}

function normalizeVoiceUri(value: unknown): string | null | undefined {
  if (value === null) return null;
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > PDF_READ_ALOUD_VOICE_URI_MAX_LENGTH ||
    value.trim() !== value ||
    /\p{Cc}/u.test(value)
  ) {
    return undefined;
  }
  return value;
}

export function createDefaultPdfReadAloudPreferences(
  locale: string,
): PdfReadAloudPreferences {
  const canonicalLocale = normalizePdfLanguageTag(locale);
  const localeLanguage = canonicalLocale?.toLowerCase();
  return Object.freeze({
    lang:
      localeLanguage === "zh"
        ? "zh-CN"
        : localeLanguage === "en"
          ? "en-US"
          : canonicalLocale ||
            (locale.toLowerCase().startsWith("zh") ? "zh-CN" : "en-US"),
    rate: 1,
    voiceURI: null,
    followDocumentLanguage: true,
  });
}

export function parsePdfReadAloudPreferences(
  value: unknown,
): PdfReadAloudPreferences | null {
  if (!isPlainObject(value) || !hasExactKeys(value)) return null;
  if (value.schemaVersion !== PDF_READ_ALOUD_PREFERENCE_SCHEMA_VERSION) {
    return null;
  }
  const lang = normalizePdfLanguageTag(value.lang);
  const rate = normalizeRate(value.rate);
  const voiceURI = normalizeVoiceUri(value.voiceURI);
  if (
    !lang ||
    rate === null ||
    voiceURI === undefined ||
    typeof value.followDocumentLanguage !== "boolean"
  ) {
    return null;
  }
  return Object.freeze({
    lang,
    rate,
    voiceURI,
    followDocumentLanguage: value.followDocumentLanguage,
  });
}

export function loadPdfReadAloudPreferences(
  storage: PdfReadAloudPreferenceStorage,
): PdfReadAloudPreferences | null {
  try {
    const raw = storage.getItem(PDF_READ_ALOUD_PREFERENCE_STORAGE_KEY);
    if (!raw) return null;
    return parsePdfReadAloudPreferences(JSON.parse(raw) as unknown);
  } catch {
    return null;
  }
}

export function savePdfReadAloudPreferences(
  storage: PdfReadAloudPreferenceStorage,
  preferences: PdfReadAloudPreferences,
): boolean {
  const parsed = parsePdfReadAloudPreferences({
    schemaVersion: PDF_READ_ALOUD_PREFERENCE_SCHEMA_VERSION,
    ...preferences,
  });
  if (!parsed) return false;
  const persisted: PersistedPdfReadAloudPreferences = {
    schemaVersion: PDF_READ_ALOUD_PREFERENCE_SCHEMA_VERSION,
    ...parsed,
  };
  try {
    storage.setItem(
      PDF_READ_ALOUD_PREFERENCE_STORAGE_KEY,
      JSON.stringify(persisted),
    );
    return true;
  } catch {
    return false;
  }
}

export function resolvePdfReadAloudLanguage(
  preferences: PdfReadAloudPreferences,
  documentLanguage: unknown,
): string {
  if (preferences.followDocumentLanguage) {
    const canonicalDocumentLanguage = normalizePdfLanguageTag(documentLanguage);
    if (canonicalDocumentLanguage) return canonicalDocumentLanguage;
  }
  return preferences.lang;
}

export function pdfSpeechVoiceMatchesLanguage(
  voiceLanguage: unknown,
  requestedLanguage: unknown,
): boolean {
  const voice = normalizePdfLanguageTag(voiceLanguage);
  const requested = normalizePdfLanguageTag(requestedLanguage);
  if (!voice || !requested) return false;
  if (voice.toLowerCase() === requested.toLowerCase()) return true;
  return (
    voice.split("-", 1)[0].toLowerCase() ===
    requested.split("-", 1)[0].toLowerCase()
  );
}

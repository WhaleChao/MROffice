"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Check } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { useMemo, useState, type RefObject } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { normalizePdfLanguageTag } from "../../../../lib/pdf/document-properties";
import {
  pdfSpeechVoiceMatchesLanguage,
  type PdfReadAloudPreferences,
} from "../../../../lib/pdf/read-aloud-preferences";

export interface PdfSpeechVoiceOption {
  readonly name: string;
  readonly lang: string;
  readonly voiceURI: string;
  readonly localService: boolean;
  readonly default: boolean;
}

export type PdfReadAloudSettings = PdfReadAloudPreferences;

interface PdfReadAloudSettingsDialogProps {
  readonly open: boolean;
  readonly locale: string;
  readonly value: PdfReadAloudSettings;
  readonly documentLanguage: string;
  readonly voices: readonly PdfSpeechVoiceOption[];
  readonly returnFocusRef?: RefObject<HTMLElement | null>;
  readonly onApply: (value: PdfReadAloudSettings) => void;
  readonly onCancel: () => void;
}

function normalizedLanguage(value: string): string {
  return normalizePdfLanguageTag(value) || "en-US";
}

function PdfReadAloudSettingsDialogSession({
  locale,
  value,
  documentLanguage,
  voices,
  returnFocusRef,
  onApply,
  onCancel,
}: Omit<PdfReadAloudSettingsDialogProps, "open">): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const [draft, setDraft] = useState<PdfReadAloudSettings>(() => ({
    lang: normalizedLanguage(value.lang),
    rate: Math.min(2, Math.max(0.5, value.rate)),
    voiceURI: voices.some((voice) => voice.voiceURI === value.voiceURI)
      ? value.voiceURI
      : null,
    followDocumentLanguage: value.followDocumentLanguage,
  }));
  const canonicalDocumentLanguage = normalizePdfLanguageTag(documentLanguage);
  const activeLanguage =
    draft.followDocumentLanguage && canonicalDocumentLanguage
      ? canonicalDocumentLanguage
      : draft.lang;
  const languages = useMemo(() => {
    const values = new Map<string, string>();
    const currentLanguage = normalizedLanguage(value.lang);
    values.set(currentLanguage.toLowerCase(), currentLanguage);
    if (canonicalDocumentLanguage) {
      values.set(
        canonicalDocumentLanguage.toLowerCase(),
        canonicalDocumentLanguage,
      );
    }
    for (const voice of voices) {
      const language = normalizedLanguage(voice.lang);
      values.set(language.toLowerCase(), language);
    }
    return [...values.values()].sort((left, right) =>
      left.localeCompare(right, locale),
    );
  }, [canonicalDocumentLanguage, locale, value.lang, voices]);
  const availableVoices = voices.filter((voice) =>
    pdfSpeechVoiceMatchesLanguage(voice.lang, activeLanguage),
  );
  const selectedVoice = availableVoices.some(
    (voice) => voice.voiceURI === draft.voiceURI,
  )
    ? (draft.voiceURI ?? "")
    : "";

  return (
    <OfficeDialogShell
      open
      title={t("pdf.readAloud.title")}
      description={t("pdf.readAloud.description")}
      defaultActionLabel={t("pdf.readAloud.apply")}
      cancelActionLabel={t("pdf.readAloud.cancel")}
      initialFocus="cancel"
      returnFocusRef={returnFocusRef}
      onDefaultAction={() => onApply(Object.freeze({ ...draft }))}
      onCancel={onCancel}
      testIdPrefix="pdf-read-aloud-settings"
      className="sm:max-w-[520px]"
    >
      <div className="grid gap-4">
        <label className={`${parts.check} cursor-pointer min-h-11`}>
          <span className="relative size-11 shrink-0">
            <input
              type="checkbox"
              checked={draft.followDocumentLanguage}
              className={`${parts.markInput} absolute inset-0 size-11 cursor-pointer`}
              data-testid="pdf-read-aloud-follow-document-language"
              aria-describedby="pdf-read-aloud-document-language-hint"
              onChange={(event) => {
                const followDocumentLanguage = event.currentTarget.checked;
                setDraft((current) => ({
                  ...current,
                  followDocumentLanguage,
                  voiceURI: null,
                }));
              }}
            />
            <span
              aria-hidden
              className={`${parts.mark} pointer-events-none absolute left-3.5 top-3.5 inline-flex size-4 items-center justify-center`}
              data-testid="pdf-read-aloud-follow-document-language-visual"
            >
              <Check
                size={12}
                strokeWidth={3}
                className={
                  draft.followDocumentLanguage ? "opacity-100" : "opacity-0"
                }
              />
            </span>
          </span>
          <span className="grid min-w-0 gap-0.5">
            <span className="font-medium">
              {t("pdf.readAloud.followDocument")}
            </span>
            <span
              id="pdf-read-aloud-document-language-hint"
              className={parts.hint}
            >
              {canonicalDocumentLanguage
                ? t("pdf.readAloud.documentLanguage", {
                    language: canonicalDocumentLanguage,
                  })
                : t("pdf.readAloud.documentLanguageMissing")}
            </span>
          </span>
        </label>

        <label className={parts.field}>
          {t("pdf.readAloud.defaultLanguage")}
          <select
            value={draft.lang}
            data-testid="pdf-read-aloud-language"
            onChange={(event) => {
              const lang = event.currentTarget.value;
              setDraft((current) => ({
                ...current,
                lang,
                voiceURI: null,
              }));
            }}
          >
            {languages.map((language) => (
              <option key={language} value={language}>
                {language}
              </option>
            ))}
          </select>
        </label>

        <label className={parts.field}>
          {t("pdf.readAloud.systemVoice")}
          <select
            value={selectedVoice}
            data-testid="pdf-read-aloud-voice"
            onChange={(event) => {
              const voiceURI = event.currentTarget.value || null;
              setDraft((current) => ({
                ...current,
                voiceURI,
              }));
            }}
          >
            <option value="">{t("pdf.readAloud.systemDefaultVoice")}</option>
            {availableVoices.map((voice) => (
              <option key={voice.voiceURI} value={voice.voiceURI}>
                {voice.name}
                {voice.localService ? t("pdf.readAloud.localVoiceSuffix") : ""}
              </option>
            ))}
          </select>
        </label>

        <label className={parts.field}>
          <span className="flex items-center justify-between gap-3">
            <span>{t("pdf.readAloud.speed")}</span>
            <output
              htmlFor="pdf-read-aloud-rate"
              className="min-w-12 text-end tabular-nums"
              data-testid="pdf-read-aloud-rate-value"
            >
              {draft.rate.toFixed(1)}×
            </output>
          </span>
          <input
            id="pdf-read-aloud-rate"
            type="range"
            min={0.5}
            max={2}
            step={0.1}
            value={draft.rate}
            className="cursor-pointer h-11 min-h-11"
            data-testid="pdf-read-aloud-rate"
            onChange={(event) => {
              const rate = Number(event.currentTarget.value);
              setDraft((current) => ({
                ...current,
                rate,
              }));
            }}
          />
        </label>

        <p className={`${parts.hint} m-0`}>
          {voices.length === 0
            ? t("pdf.readAloud.voicesPending")
            : t("pdf.readAloud.voicesHint", { language: activeLanguage })}
        </p>
      </div>
    </OfficeDialogShell>
  );
}

export default function PdfReadAloudSettingsDialog(
  props: PdfReadAloudSettingsDialogProps,
): React.JSX.Element | null {
  if (!props.open) return null;
  return (
    <PdfReadAloudSettingsDialogSession
      key={`${props.value.lang}:${props.value.rate}:${props.value.voiceURI ?? "default"}`}
      locale={props.locale}
      value={props.value}
      documentLanguage={props.documentLanguage}
      voices={props.voices}
      returnFocusRef={props.returnFocusRef}
      onApply={props.onApply}
      onCancel={props.onCancel}
    />
  );
}

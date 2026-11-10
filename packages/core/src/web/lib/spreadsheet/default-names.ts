/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { contentLocaleOf, pageInterfaceLocale, type ContentLocale } from "../i18n/content-locale";

export interface SpreadsheetDefaultNames {
  /** A new worksheet's name before its number ("Sheet" -> "Sheet1"). */
  readonly sheet: string;
  /** A new table's name before its number. */
  readonly table: string;
  /** A table column's name before its number, where it has no header text. */
  readonly column: string;
}

export const SPREADSHEET_DEFAULT_NAMES: Readonly<Record<ContentLocale, SpreadsheetDefaultNames>> = {
  en: { sheet: "Sheet", table: "Table", column: "Column" },
  zh: { sheet: "Sheet", table: "表", column: "列" },
  "zh-TW": { sheet: "工作表", table: "表格", column: "欄" },
  ja: { sheet: "Sheet", table: "テーブル", column: "列" },
  ko: { sheet: "Sheet", table: "표", column: "열" },
  de: { sheet: "Tabelle", table: "Tabelle", column: "Spalte" },
  fr: { sheet: "Feuil", table: "Tableau", column: "Colonne" },
  es: { sheet: "Hoja", table: "Tabla", column: "Columna" },
  it: { sheet: "Foglio", table: "Tabella", column: "Colonna" },
  nl: { sheet: "Blad", table: "Tabel", column: "Kolom" },
  "pt-BR": { sheet: "Planilha", table: "Tabela", column: "Coluna" },
  da: { sheet: "Ark", table: "Tabel", column: "Kolonne" },
  no: { sheet: "Ark", table: "Tabell", column: "Kolonne" },
  sv: { sheet: "Blad", table: "Tabell", column: "Kolumn" },
  fi: { sheet: "Taul", table: "Taulukko", column: "Sarake" },
  id: { sheet: "Lembar", table: "Tabel", column: "Kolom" },
  th: { sheet: "Sheet", table: "Table", column: "คอลัมน์" },
  vi: { sheet: "Sheet", table: "Table", column: "Column" },
};

/**
 * The default names for `locale`, or -- with none given -- the language of the
 * page this runs in (English where there is none: the server, a test).
 */
export function spreadsheetDefaultNames(locale?: string | null): SpreadsheetDefaultNames {
  return SPREADSHEET_DEFAULT_NAMES[contentLocaleOf(locale ?? pageInterfaceLocale())];
}

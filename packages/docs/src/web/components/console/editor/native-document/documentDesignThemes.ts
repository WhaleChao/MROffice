/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/core";
import { closeHistory } from "@tiptap/pm/history";

import {
  DEFAULT_DOCUMENT_DESIGN,
  documentDesignEquals,
  documentDesignWithCustomFonts,
  documentDesignWithScheme,
  documentDesignWithTheme,
  documentDesignWithThemeFile,
  normalizeDocumentDesign,
  type DocumentDesign,
} from "@mrnote/core/web/lib/document-design.ts";
import {
  OFFICE_COLOR_SCHEMES,
  OFFICE_EFFECT_SCHEMES,
  OFFICE_FONT_SCHEMES,
} from "@mrnote/core/web/lib/office-theme-catalog.ts";
import { documentParagraphSpacingPreset } from "@mrnote/core/web/lib/document-paragraph-spacing.ts";
import type { OfficeThemeFileTheme } from "@mrnote/core/web/lib/office-theme-file.ts";
import type { SpreadsheetNativeThemeFonts } from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";
import {
  documentDesignWithCustomStyleSet,
  type DocumentCustomStyleSet,
} from "@mrnote/core/web/lib/document-custom-style-sets.ts";
import {
  WORD_DEFAULT_STYLE_SET_ID,
  isWordStyleSetId,
  loadWordStyleSets,
  wordStyleSetsAvailable,
} from "@mrnote/core/web/lib/office-word-style-sets.ts";

import { restyleDocumentDesign } from "@mrnote/core/web/components/console/editor/native-document/documentDesignRestyle.ts";

/** The document's design, or Word's new document's for one that has none. */
export function currentDocumentDesign(editor: Editor): DocumentDesign {
  return normalizeDocumentDesign(editor.state.doc.attrs.documentDesign) ?? DEFAULT_DOCUMENT_DESIGN;
}

/**
 * One Design command: `change` maps the current design to the next (null
 * refuses). A document with no design takes one even when the choice is
 * the default's, the way Word's document always has one. False when nothing
 * changed or the editor cannot write.
 */
export function applyDocumentDesignChange(
  editor: Editor,
  change: (design: DocumentDesign) => DocumentDesign | null,
): boolean {
  if (editor.isDestroyed || !editor.isEditable) return false;
  // Word's style sets (served Office data) decide every style's formatting:
  // until they are here nothing is written, and they are asked for now.
  if (!wordStyleSetsAvailable()) {
    void loadWordStyleSets();
    return false;
  }
  const state = editor.state;
  const stored = normalizeDocumentDesign(state.doc.attrs.documentDesign);
  const next = change(stored ?? DEFAULT_DOCUMENT_DESIGN);
  if (!next) return false;
  if (stored && documentDesignEquals(stored, next)) return false;
  const tr = restyleDocumentDesign(closeHistory(state.tr), stored, next);
  if (!tr) return false;
  if (editor.isDestroyed || !editor.isEditable || editor.state !== state) return false;
  editor.view.dispatch(tr);
  return true;
}

const COLOR_IDS = new Set(OFFICE_COLOR_SCHEMES.map((scheme) => scheme.id));
const FONT_IDS = new Set(OFFICE_FONT_SCHEMES.map((scheme) => scheme.id));
const EFFECT_IDS = new Set(OFFICE_EFFECT_SCHEMES.map((scheme) => scheme.id));

/** Design > Themes: an Office theme's colours, fonts and effects together. */
export function applyDocumentTheme(editor: Editor, themeId: unknown): boolean {
  if (typeof themeId !== "string") return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithTheme(design, themeId));
}

/**
 * Themes > Browse for Themes: the theme an Office theme file or a themed
 * Office document carries (lib/office-theme-file.ts), applied whole.
 */
export function applyDocumentThemeFile(editor: Editor, file: OfficeThemeFileTheme | null): boolean {
  if (!file) return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithThemeFile(design, file));
}

/** Themes > Reset to the Theme from Template: Word's template theme is the Office theme. */
export function resetDocumentTheme(editor: Editor): boolean {
  return applyDocumentTheme(editor, "office");
}

/** Design > Colors: one colour scheme; the theme's fonts and effects stay. */
export function applyDocumentColors(editor: Editor, colorsId: unknown): boolean {
  if (typeof colorsId !== "string" || !COLOR_IDS.has(colorsId)) return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithScheme(design, "colors", colorsId));
}

/** Design > Fonts: one font scheme (heading and body fonts). */
export function applyDocumentFonts(editor: Editor, fontsId: unknown): boolean {
  if (typeof fontsId !== "string" || !FONT_IDS.has(fontsId)) return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithScheme(design, "fonts", fontsId));
}

/** Customize Fonts...: the document theme's own named font scheme. */
export function applyDocumentCustomFonts(editor: Editor, fonts: SpreadsheetNativeThemeFonts | null): boolean {
  if (!fonts) return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithCustomFonts(design, fonts));
}

/** The theme's effects (shapes and SmartArt draw with them). */
export function applyDocumentEffects(editor: Editor, effectsId: unknown): boolean {
  if (typeof effectsId !== "string" || !EFFECT_IDS.has(effectsId)) return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithScheme(design, "effects", effectsId));
}

/** The style-set gallery: Word's style set replaces the defaults and Quick Styles. */
export function applyDocumentStyleSet(editor: Editor, styleSetId: unknown): boolean {
  if (!isWordStyleSetId(styleSetId)) return false;
  return applyDocumentDesignChange(editor, (design) => ({ ...design, styleSet: styleSetId }));
}

/** A Custom set (Save as a New Style Set): its style set and Paragraph Spacing. */
export function applyDocumentCustomStyleSet(editor: Editor, entry: DocumentCustomStyleSet | null): boolean {
  if (!entry || !isWordStyleSetId(entry.styleSet)) return false;
  return applyDocumentDesignChange(editor, (design) => documentDesignWithCustomStyleSet(design, entry));
}

/** Style sets > Reset to the Default Style Set: Word's Normal template's styles. */
export function resetDocumentStyleSet(editor: Editor): boolean {
  return applyDocumentStyleSet(editor, WORD_DEFAULT_STYLE_SET_ID);
}

/** Design > Paragraph Spacing; "default" returns to the style set's own spacing. */
export function applyDocumentParagraphSpacing(editor: Editor, presetId: unknown): boolean {
  const preset = documentParagraphSpacingPreset(presetId);
  if (!preset) return false;
  return applyDocumentDesignChange(editor, (design) => ({
    ...design,
    paragraphSpacing: preset.id === "default" ? null : preset.id,
  }));
}

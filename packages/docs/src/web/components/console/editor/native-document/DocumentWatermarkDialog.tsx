"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pxToPt } from "@mrnote/core/web/lib/office-units.ts";
import { useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as parts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  OFFICE_FONT_FAMILY_VALUES,
  officePrimaryFontFamily,
} from "@mrnote/core/web/components/office-ribbon/officeTypographyCatalog.ts";

import { DocumentDialogIcon } from "@mrnote/core/web/components/console/editor/native-document/DocumentDialogIcon.tsx";
import { documentValueInputClassName } from "@mrnote/core/web/components/console/editor/documentValueInput.ts";
import styles from "@mrnote/core/web/components/console/editor/native-document/DocumentLayoutDialogs.module.css";
import {
  DEFAULT_WATERMARK_COLOR,
  DEFAULT_WATERMARK_OPACITY,
  MAX_WATERMARK_LENGTH,
  PICTURE_WATERMARK_SCALE_PRESETS,
  WATERMARK_FONT_SIZE_PRESETS,
  WATERMARK_PRESETS,
  normalizePictureWatermark,
  normalizeWatermark,
  normalizeWatermarkFontFamily,
  normalizeWatermarkFontSizePt,
  pictureWatermarkSizePt,
  type DocumentWatermark,
  type PictureWatermarkConfig,
  type WatermarkConfig,
} from "@mrnote/core/web/components/console/editor/native-document/watermark.ts";

/*
 * Word's Printed Watermark dialog (Design > Watermark > Custom Watermark…;
 * "Insert a watermark", Microsoft Support):
 * - No watermark / Picture watermark / Text watermark.
 * - Picture: Select Picture…, Scale (Auto or a percentage of the picture's
 *   size), Washout (on by default).
 * - Text: Language (which ready-made texts the Text box lists), Text, Font,
 *   Size (Auto or points), Color, Semitransparent (on by default: the text is
 *   drawn at half opacity), Layout Diagonal or Horizontal.
 * - Apply shows the result and keeps the dialog open; OK applies and closes;
 *   Cancel closes without changing anything further.
 * Word writes the text stamp as a header-anchored VML WordArt shape and the
 * picture as a header-anchored VML picture (see docx-export.ts).
 */

export interface DocumentWatermarkDialogProps {
  readonly initial: DocumentWatermark | null;
  /** The page's text area, for the picture's Auto scale. */
  readonly textArea: { widthPt: number; heightPt: number };
  /** Pictures already in the document the user can pick. */
  readonly documentPictures: readonly string[];
  readonly canEdit?: boolean;
  /** Draft text seeded by the caller (kept across reopen). */
  readonly onApply: (next: DocumentWatermark | null) => boolean;
  readonly onClose: () => void;
  readonly returnFocusTarget?: () => HTMLElement | null;
}

type Kind = "none" | "picture" | "text";

/** Word's ready-made watermark texts per Language. */
const WATERMARK_LANGUAGE_TEXTS = {
  en: [
    "ASAP",
    "CONFIDENTIAL",
    "COPY",
    "DO NOT COPY",
    "DRAFT",
    "ORIGINAL",
    "PERSONAL",
    "SAMPLE",
    "TOP SECRET",
    "URGENT",
  ],
  zh: ["保密", "传阅", "草稿", "副本", "机密", "紧急", "绝密", "内部", "严禁复制", "样本", "原件"],
  ja: ["社外秘", "複製禁止", "下書き", "見本", "至急", "極秘", "回覧"],
} as const;
type WatermarkLanguage = keyof typeof WATERMARK_LANGUAGE_TEXTS;

const RASTER_DATA_URL_RE = /^data:image\/(?:png|jpeg|gif);base64,/;

/** Read a picture into a raster data URL plus its natural pixel size. */
async function loadPicture(
  source: string,
): Promise<{ src: string; width: number; height: number }> {
  let dataUrl = source;
  if (!source.startsWith("data:")) {
    const response = await fetch(source);
    if (!response.ok) throw new Error("fetch_failed");
    const blob = await response.blob();
    dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(new Error("read_failed"));
      reader.readAsDataURL(blob);
    });
  }
  const image = await new Promise<HTMLImageElement>((resolve, reject) => {
    const element = new Image();
    element.onload = () => resolve(element);
    element.onerror = () => reject(new Error("decode_failed"));
    element.src = dataUrl;
  });
  const width = image.naturalWidth;
  const height = image.naturalHeight;
  if (!(width > 0 && height > 0)) throw new Error("empty_picture");
  if (!RASTER_DATA_URL_RE.test(dataUrl)) {
    // Word stores a raster; anything else (SVG, WebP) is drawn once to PNG.
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("canvas_unavailable");
    context.drawImage(image, 0, 0);
    dataUrl = canvas.toDataURL("image/png");
  }
  return { src: dataUrl, width, height };
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error("read_failed"));
    reader.readAsDataURL(file);
  });
}

export function DocumentWatermarkDialog({
  initial,
  textArea,
  documentPictures,
  canEdit = true,
  onApply,
  onClose,
  returnFocusTarget,
}: DocumentWatermarkDialogProps) {
  const t = useTranslations("console-notebooks");
  const initialText: WatermarkConfig | null =
    initial && !("kind" in initial) ? initial : null;
  const initialPicture: PictureWatermarkConfig | null =
    initial && "kind" in initial ? initial : null;
  const textRef = useRef<HTMLInputElement | null>(null);
  const colorButtonRef = useRef<HTMLButtonElement | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);
  const [kind, setKind] = useState<Kind>(
    initialPicture ? "picture" : initialText ? "text" : "text",
  );
  // Text watermark drafts.
  const [language, setLanguage] = useState<WatermarkLanguage>("en");
  const [text, setText] = useState(initialText?.text ?? "");
  const [fontFamily, setFontFamily] = useState<string | null>(
    initialText?.fontFamily ?? null,
  );
  const [sizeDraft, setSizeDraft] = useState(
    initialText?.fontSizePt === undefined ? "" : String(initialText.fontSizePt),
  );
  const [color, setColor] = useState(initialText?.color ?? DEFAULT_WATERMARK_COLOR);
  const [paletteOpen, setPaletteOpen] = useState(false);
  const [semitransparent, setSemitransparent] = useState(
    initialText ? initialText.opacity < 1 : true,
  );
  const [diagonal, setDiagonal] = useState(initialText?.diagonal ?? true);
  // Picture watermark drafts.
  const [picture, setPicture] = useState<{
    src: string;
    width: number;
    height: number;
  } | null>(null);
  const [scale, setScale] = useState<string>("auto");
  const [washout, setWashout] = useState(initialPicture?.washout ?? true);
  const [pictureUrl, setPictureUrl] = useState("");
  const [pictureError, setPictureError] = useState<string | null>(null);
  const [pictureBusy, setPictureBusy] = useState(false);

  // An existing picture watermark: learn its natural size, and show Scale as
  // Auto when its drawn size is the Auto size, else as a percentage.
  useEffect(() => {
    if (!initialPicture) return;
    let live = true;
    loadPicture(initialPicture.src)
      .then((loaded) => {
        if (!live) return;
        setPicture(loaded);
        const auto = pictureWatermarkSizePt(loaded.width, loaded.height, "auto", textArea);
        if (auto && Math.abs(auto.widthPt - initialPicture.widthPt) < 0.5) {
          setScale("auto");
        } else {
          setScale(String(Math.round((initialPicture.widthPt / (pxToPt(loaded.width))) * 100)));
        }
      })
      .catch(() => {
        if (live) setPictureError(t("toolbar.watermarkPictureUnreadable"));
      });
    return () => {
      live = false;
    };
    // Loaded once for the watermark the dialog opened on.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const fontOptions = useMemo(() => {
    const names = new Set<string>();
    for (const stack of OFFICE_FONT_FAMILY_VALUES) {
      const name = normalizeWatermarkFontFamily(officePrimaryFontFamily(stack));
      if (name) names.add(name);
    }
    if (fontFamily) names.add(fontFamily);
    return [...names];
  }, [fontFamily]);

  const candidate = useMemo((): DocumentWatermark | null | "invalid" => {
    if (kind === "none") return null;
    if (kind === "picture") {
      if (!picture) return "invalid";
      const scaleValue = scale === "auto" ? "auto" : Number(scale);
      if (scaleValue !== "auto" && !(scaleValue > 0 && scaleValue <= 5000)) {
        return "invalid";
      }
      const size = pictureWatermarkSizePt(picture.width, picture.height, scaleValue, textArea);
      if (!size) return "invalid";
      return (
        normalizePictureWatermark({
          kind: "picture",
          src: picture.src,
          widthPt: size.widthPt,
          heightPt: size.heightPt,
          washout,
        }) ?? "invalid"
      );
    }
    const size = sizeDraft.trim() ? normalizeWatermarkFontSizePt(Number(sizeDraft)) : undefined;
    if (sizeDraft.trim() && size === undefined) return "invalid";
    const opacity = semitransparent
      ? initialText && initialText.opacity < 1
        ? initialText.opacity
        : initialText
          ? 0.5
          : DEFAULT_WATERMARK_OPACITY
      : 1;
    const next = normalizeWatermark({
      text,
      color,
      opacity,
      diagonal,
      ...(fontFamily ? { fontFamily } : {}),
      ...(size === undefined ? {} : { fontSizePt: size }),
    });
    return next ?? "invalid";
  }, [
    color,
    diagonal,
    fontFamily,
    initialText,
    kind,
    picture,
    scale,
    semitransparent,
    sizeDraft,
    text,
    textArea,
    washout,
  ]);

  const choosePicture = async (source: string | File) => {
    setPictureBusy(true);
    setPictureError(null);
    try {
      const raw = typeof source === "string" ? source.trim() : await readFileAsDataUrl(source);
      if (typeof source === "string" && !/^(data:image\/|https?:\/\/|\/)/i.test(raw)) {
        throw new Error("unsupported_source");
      }
      setPicture(await loadPicture(raw));
      setKind("picture");
    } catch {
      setPictureError(t("toolbar.watermarkPictureUnreadable"));
    } finally {
      setPictureBusy(false);
    }
  };

  const apply = (): boolean => {
    if (!canEdit || candidate === "invalid") return false;
    return onApply(candidate);
  };

  const presetTexts = WATERMARK_LANGUAGE_TEXTS[language];

  return (
    <OfficeDialogShell
      open
      testIdPrefix="document-toolbar-watermark"
      title={t("toolbar.watermarkDialogTitle")}
      description=""
      defaultActionLabel={t("toolbar.dialogOk")}
      defaultActionTestId="document-toolbar-watermark-apply"
      defaultActionDisabled={!canEdit || candidate === "invalid"}
      cancelActionLabel={t("toolbar.pageSetupCancel")}
      cancelActionTestId="document-toolbar-watermark-cancel"
      onDefaultAction={() => {
        if (apply()) onClose();
      }}
      onCancel={onClose}
      initialFocusRef={textRef}
      returnFocusTarget={returnFocusTarget}
      className="sm:max-w-[560px]"
      footerStart={
        <button
          type="button"
          className="office-dialog__button"
          data-testid="document-toolbar-watermark-apply-keep"
          disabled={!canEdit || candidate === "invalid"}
          onClick={() => {
            apply();
          }}
        >
          {t("toolbar.apply")}
        </button>
      }
    >
      <div data-testid="document-toolbar-watermark-form" className="office-dialog__body">
        <label className="office-dialog__check">
          <input
            type="radio"
            name="document-watermark-kind"
            checked={kind === "none"}
            disabled={!canEdit}
            data-testid="document-toolbar-watermark-kind-none"
            onChange={() => setKind("none")}
          />
          <DocumentDialogIcon icon="status.blocked" />
          {t("toolbar.watermarkNone")}
        </label>

        <label className="office-dialog__check">
          <input
            type="radio"
            name="document-watermark-kind"
            checked={kind === "picture"}
            disabled={!canEdit}
            data-testid="document-toolbar-watermark-kind-picture"
            onChange={() => setKind("picture")}
          />
          <DocumentDialogIcon icon="image.insert" />
          {t("toolbar.watermarkPicture")}
        </label>
        <fieldset className="office-dialog__group" disabled={!canEdit || kind !== "picture"}>
          <div className={styles.valueRow}>
            <button
              type="button"
              className="office-dialog__button"
              data-testid="document-toolbar-watermark-picture-select"
              onClick={() => fileRef.current?.click()}
            >
              <DocumentDialogIcon icon="image.upload" />
              {t("toolbar.watermarkSelectPicture")}
            </button>
            <input
              ref={fileRef}
              type="file"
              accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml"
              hidden
              data-testid="document-toolbar-watermark-picture-file"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0];
                if (file) void choosePicture(file);
                event.currentTarget.value = "";
              }}
            />
            {picture ? (
              // eslint-disable-next-line @next/next/no-img-element
              <img
                src={picture.src}
                alt=""
                className={styles.pictureThumb}
                data-testid="document-toolbar-watermark-picture-preview"
              />
            ) : null}
          </div>
          <div className={styles.valueRow}>
            <input
              type="text"
              className="note-editor-office-toolbar__input"
              placeholder={t("toolbar.watermarkPictureUrl")}
              aria-label={t("toolbar.watermarkPictureUrl")}
              data-testid="document-toolbar-watermark-picture-url"
              value={pictureUrl}
              onChange={(event) => setPictureUrl(event.currentTarget.value)}
            />
            <button
              type="button"
              className="office-dialog__button"
              data-testid="document-toolbar-watermark-picture-url-use"
              disabled={!pictureUrl.trim() || pictureBusy}
              onClick={() => void choosePicture(pictureUrl)}
            >
              {t("toolbar.watermarkPictureUse")}
            </button>
          </div>
          {documentPictures.length > 0 ? (
            <div className={styles.valueRow} role="group" aria-label={t("toolbar.watermarkDocumentPictures")}>
              {documentPictures.slice(0, 6).map((src, index) => (
                <button
                  key={`${index}-${src.slice(0, 32)}`}
                  type="button"
                  className="office-dialog__button"
                  data-testid={`document-toolbar-watermark-picture-doc-${index}`}
                  onClick={() => void choosePicture(src)}
                >
                  {/* eslint-disable-next-line @next/next/no-img-element */}
                  <img src={src} alt="" className={styles.pictureThumb} />
                </button>
              ))}
            </div>
          ) : null}
          <div className={styles.twoColumns}>
            <label className="office-dialog__field">
              <span className="office-dialog__label">{t("toolbar.watermarkScale")}</span>
              <select
                data-testid="document-toolbar-watermark-picture-scale"
                value={scale}
                onChange={(event) => setScale(event.currentTarget.value)}
              >
                <option value="auto">{t("toolbar.watermarkSizeAuto")}</option>
                {PICTURE_WATERMARK_SCALE_PRESETS.map((value) => (
                  <option key={value} value={String(value)}>
                    {value}%
                  </option>
                ))}
                {scale !== "auto" &&
                !PICTURE_WATERMARK_SCALE_PRESETS.some((value) => String(value) === scale) ? (
                  <option value={scale}>{scale}%</option>
                ) : null}
              </select>
            </label>
            <div className="office-dialog__field">
              <label className="office-dialog__check">
                <input
                  type="checkbox"
                  checked={washout}
                  data-testid="document-toolbar-watermark-picture-washout"
                  onChange={(event) => setWashout(event.currentTarget.checked)}
                />
                {t("toolbar.watermarkWashout")}
              </label>
            </div>
          </div>
          {pictureError ? (
            <p className="office-dialog__message" role="alert">
              {pictureError}
            </p>
          ) : null}
        </fieldset>

        <label className="office-dialog__check">
          <input
            type="radio"
            name="document-watermark-kind"
            checked={kind === "text"}
            disabled={!canEdit}
            data-testid="document-toolbar-watermark-kind-text"
            onChange={() => setKind("text")}
          />
          <DocumentDialogIcon icon="text.type" />
          {t("toolbar.watermarkText")}
        </label>
        <fieldset className="office-dialog__group" disabled={!canEdit || kind !== "text"}>
          <div className={styles.twoColumns}>
            <label className="office-dialog__field">
              <span className="office-dialog__label">{t("toolbar.watermarkLanguage")}</span>
              <select
                data-testid="document-toolbar-watermark-language"
                value={language}
                onChange={(event) =>
                  setLanguage(event.currentTarget.value as WatermarkLanguage)
                }
              >
                <option value="en">English</option>
                <option value="zh">中文</option>
                <option value="ja">日本語</option>
              </select>
            </label>
            <label className="office-dialog__field">
              <span className="office-dialog__label">{t("toolbar.watermarkTextLabel")}</span>
              <input
                ref={textRef}
                type="text"
                list="document-watermark-texts"
                className="note-editor-office-toolbar__input"
                data-testid="document-toolbar-watermark-input"
                value={text}
                maxLength={MAX_WATERMARK_LENGTH}
                placeholder={t("toolbar.watermarkPlaceholder")}
                onChange={(event) => {
                  setText(event.currentTarget.value);
                  setKind("text");
                }}
              />
              <datalist id="document-watermark-texts">
                {presetTexts.map((value) => (
                  <option key={value} value={value} />
                ))}
              </datalist>
            </label>
            <label className="office-dialog__field">
              <span className="office-dialog__label">{t("toolbar.watermarkFont")}</span>
              <select
                data-testid="document-toolbar-watermark-font"
                value={fontFamily ?? ""}
                onChange={(event) => setFontFamily(event.currentTarget.value || null)}
              >
                <option value="">{t("toolbar.watermarkFontDefault")}</option>
                {fontOptions.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
            </label>
            <label className="office-dialog__field">
              <span className="office-dialog__label">{t("toolbar.watermarkSize")}</span>
              <input
                type="text"
                inputMode="decimal"
                list="document-watermark-sizes"
                className={documentValueInputClassName("note-editor-office-toolbar__input")}
                data-testid="document-toolbar-watermark-size"
                placeholder={t("toolbar.watermarkSizeAuto")}
                value={sizeDraft}
                onChange={(event) => setSizeDraft(event.currentTarget.value)}
              />
              <datalist id="document-watermark-sizes">
                {WATERMARK_FONT_SIZE_PRESETS.map((value) => (
                  <option key={value} value={String(value)} />
                ))}
              </datalist>
            </label>
            <div className="office-dialog__field" style={{ position: "relative" }}>
              <span className="office-dialog__label">{t("toolbar.watermarkColor")}</span>
              <button
                ref={colorButtonRef}
                type="button"
                className="office-dialog__button"
                aria-haspopup="dialog"
                aria-expanded={paletteOpen}
                data-testid="document-toolbar-watermark-color"
                data-color={color}
                onClick={() => setPaletteOpen((open) => !open)}
              >
                <span className={styles.swatch} style={{ background: color }} aria-hidden />
              </button>
              {paletteOpen ? (
                <OfficeColorPalettePopover
                  anchorRef={colorButtonRef}
                  zIndex={1000}
                  value={color}
                  labels={{
                    themeColors: t("toolbar.themeColors"),
                    standardColors: t("toolbar.standardColors"),
                    recentColors: t("spreadsheet.recentColors"),
                    moreColors: t("toolbar.watermarkColor"),
                  }}
                  onPick={(next) => {
                    if (next) setColor(next.toLowerCase());
                    setPaletteOpen(false);
                  }}
                  onDismiss={() => setPaletteOpen(false)}
                  testIdPrefix="document-toolbar-watermark-color"
                />
              ) : null}
            </div>
            <div className="office-dialog__field">
              <label className="office-dialog__check">
                <input
                  type="checkbox"
                  checked={semitransparent}
                  data-testid="document-toolbar-watermark-semitransparent"
                  onChange={(event) => setSemitransparent(event.currentTarget.checked)}
                />
                {t("toolbar.watermarkSemitransparent")}
              </label>
            </div>
          </div>
          <div className={styles.choiceRow} role="radiogroup" aria-label={t("toolbar.watermarkLayout")}>
            <span className="office-dialog__label">{t("toolbar.watermarkLayout")}</span>
            {([true, false] as const).map((value) => (
              <label className="office-dialog__check" key={String(value)}>
                <input
                  type="radio"
                  name="document-watermark-layout"
                  checked={diagonal === value}
                  data-testid={`document-toolbar-watermark-layout-${value ? "diagonal" : "horizontal"}`}
                  onChange={() => setDiagonal(value)}
                />
                <DocumentDialogIcon icon={value ? "object.rotate" : "nav.move.horizontal"} />
                {t(value ? "toolbar.watermarkDiagonal" : "toolbar.watermarkHorizontal")}
              </label>
            ))}
          </div>
          <div className={styles.valueRow}>
            {WATERMARK_PRESETS.map((preset) => (
              <button
                key={preset}
                type="button"
                className={parts.compactButton}
                data-testid={`document-toolbar-watermark-preset-${preset.toLowerCase()}`}
                onClick={() => {
                  setText(preset);
                  setKind("text");
                }}
              >
                {preset}
              </button>
            ))}
          </div>
        </fieldset>
      </div>
    </OfficeDialogShell>
  );
}

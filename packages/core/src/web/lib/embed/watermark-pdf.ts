/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An embed session's watermark, drawn over the editor and stamped into the
 * PDFs that leave the frame (lib/embed/session-access: `watermark`).
 *
 * The server stamps every PDF the page export hands a session whose
 * watermark is `inExports` (app/middleware/embed_export_watermark.py); what
 * the frame makes itself (a PDF an editor draws in the browser and saves or
 * prints) is stamped here, the same way: the text tiled across every page,
 * rotated, translucent, in the session's colour and size. The text is drawn
 * into a picture first (a canvas), so any script a person's name or email is
 * written in prints as it reads, without a font for it in the PDF.
 *
 * A blob the page export returned is marked (`markServerStamped`) so it is
 * not stamped twice.
 */

import type { EmbedWatermark } from "./session-access";

const stamped = new WeakSet<Blob>();

/** A PDF the server already stamped (lib/page-export marks them). */
export function markServerStamped(blob: Blob): Blob {
  stamped.add(blob);
  return blob;
}

export function isStamped(blob: Blob): boolean {
  return stamped.has(blob);
}

const WATERMARK_FONT = `system-ui, -apple-system, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif`;

/** The widest line's width at ``fontSize``: measured in the watermark's font
 * where a canvas can, else a letter's usual width (a CJK one the font size). */
function watermarkTextWidth(lines: readonly string[], fontSize: number): number {
  try {
    const context = typeof document === "undefined" ? null : document.createElement("canvas").getContext("2d");
    if (context) {
      context.font = `${fontSize}px ${WATERMARK_FONT}`;
      const measured = Math.max(0, ...lines.map((line) => context.measureText(line).width));
      if (Number.isFinite(measured) && measured > 0) return measured;
    }
  } catch {
    // estimated below
  }
  return Math.max(0, ...lines.map((line) => [...line].reduce((width, letter) => width + (letter.codePointAt(0)! >= 0x2e80 ? 1 : 0.6) * fontSize, 0)));
}

/**
 * The tile's size in PDF points: it holds the text turned by the angle and a
 * gap of four letters around it (the tile is repeated, and what runs past its
 * edge is cut off) -- an email and a date on the old fixed 12-letter tile
 * were cut and ran into the next one's. The server's stamp
 * (app/services/embed/access.py) steps along the turned text the same way.
 */
export function watermarkTileSize(watermark: EmbedWatermark): { width: number; height: number } {
  const size = watermark.fontSize;
  const lines = watermarkLines(watermark);
  const textWidth = watermarkTextWidth(lines, size);
  const textHeight = size * 1.25 * Math.max(1, lines.length);
  const turn = (watermark.angle * Math.PI) / 180;
  const cos = Math.abs(Math.cos(turn));
  const sin = Math.abs(Math.sin(turn));
  return {
    width: Math.ceil(Math.max(textWidth * cos + textHeight * sin + size * 4, size * 12, 220)),
    height: Math.ceil(Math.max(textWidth * sin + textHeight * cos + size * 4, size * 6, 140)),
  };
}

export function watermarkLines(watermark: EmbedWatermark): string[] {
  return watermark.text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 2);
}

/**
 * One tile of the watermark as an SVG (the overlay's background and the
 * picture stamped into PDFs): the text centred, rotated by the angle.
 */
export function watermarkTileSvg(watermark: EmbedWatermark): string {
  const { width, height } = watermarkTileSize(watermark);
  const lines = watermarkLines(watermark);
  const escape = (value: string) =>
    value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const size = watermark.fontSize;
  const first = height / 2 - ((lines.length - 1) * size * 1.25) / 2;
  const text = lines
    .map(
      (line, index) =>
        `<text x="${width / 2}" y="${first + index * size * 1.25}" text-anchor="middle" dominant-baseline="middle">${escape(line)}</text>`,
    )
    .join("");
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
    `<g fill="${watermark.color}" fill-opacity="${watermark.opacity}" font-family="${WATERMARK_FONT.replace(/"/g, "'")}" font-size="${size}" transform="rotate(${watermark.angle} ${width / 2} ${height / 2})">${text}</g>` +
    `</svg>`
  );
}

async function tilePng(watermark: EmbedWatermark, scale: number): Promise<Uint8Array> {
  const { width, height } = watermarkTileSize(watermark);
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(width * scale);
  canvas.height = Math.ceil(height * scale);
  const context = canvas.getContext("2d");
  if (!context) throw new Error("no canvas");
  const lines = watermarkLines(watermark);
  context.scale(scale, scale);
  context.translate(width / 2, height / 2);
  context.rotate((watermark.angle * Math.PI) / 180);
  context.globalAlpha = watermark.opacity;
  context.fillStyle = watermark.color;
  context.font = `${watermark.fontSize}px ${WATERMARK_FONT}`;
  context.textAlign = "center";
  context.textBaseline = "middle";
  const first = -((lines.length - 1) * watermark.fontSize * 1.25) / 2;
  lines.forEach((line, index) => context.fillText(line, 0, first + index * watermark.fontSize * 1.25));
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!blob) throw new Error("no tile");
  return new Uint8Array(await blob.arrayBuffer());
}

/** `pdf` with the watermark over every page; throws rather than hand out an
 * unstamped copy. */
export async function stampPdfBlob(pdf: Blob, watermark: EmbedWatermark): Promise<Blob> {
  if (isStamped(pdf)) return pdf;
  const { PDFDocument } = await import("pdf-lib");
  const document = await PDFDocument.load(new Uint8Array(await pdf.arrayBuffer()), { ignoreEncryption: false });
  const tile = await document.embedPng(await tilePng(watermark, 3));
  const { width: tileWidth, height: tileHeight } = watermarkTileSize(watermark);
  for (const page of document.getPages()) {
    const { width, height } = page.getSize();
    for (let row = 0, y = 0; y < height; row += 1, y += tileHeight) {
      for (let x = row % 2 ? -tileWidth / 2 : 0; x < width; x += tileWidth) {
        page.drawImage(tile, { x, y, width: tileWidth, height: tileHeight });
      }
    }
  }
  const bytes = await document.save();
  const out = new Blob([bytes as BlobPart], { type: "application/pdf" });
  return markServerStamped(out);
}

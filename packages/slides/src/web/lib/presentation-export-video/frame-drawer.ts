/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const XMLNS = "http://www.w3.org/2000/svg";














export class DefaultStylesKeptIntact<K, V extends Map<unknown, unknown>> extends Map<K, V> {
  override get(key: K): V | undefined {
    const styles = super.get(key);
    return styles ? (new Map(styles) as V) : undefined;
  }
}















export function standardStyleProperties(ownerDocument: Document = document): string[] {
  const computed = (ownerDocument.defaultView ?? window).getComputedStyle(ownerDocument.documentElement);
  const names: string[] = [];
  for (let index = 0; index < computed.length; index += 1) {
    const name = computed.item(index);
    if (name && !name.startsWith("--")) names.push(name);
  }
  return names;
}

/** Make a screenshot context's default-style cache hand out copies (see {@link DefaultStylesKeptIntact}). */
export function keepDefaultStylesIntact(context: { defaultComputedStyles?: Map<string, Map<string, string>> }): void {
  const current = context.defaultComputedStyles;
  if (current instanceof DefaultStylesKeptIntact) return;
  context.defaultComputedStyles = new DefaultStylesKeptIntact(current ?? []);
}

/** What `modern-screenshot` resets after each frame of a reused context. */
type ScreenshotContextState = {
  svgStyleElement?: HTMLStyleElement;
  svgDefsElement?: SVGDefsElement;
  svgStyles: Map<string, string[]>;
  /** The characters each family drew (`frame-fonts`): this frame's only. */
  fontFamilies?: Map<string, Set<string>>;
  ownerDocument?: Document;
};

/** A frame's markup as the body of its `data:` URL (see `draw`). */
export function frameDataUrlBody(markup: string): string {
  return markup.replace(/[\u0000-\u0008\v\f\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]|[%#\t\n\r]/gu, (character) => {
    switch (character) {
      case "%":
        return "%25";
      case "#":
        return "%23";
      case "\t":
        return "%09";
      case "\n":
        return "%0A";
      case "\r":
        return "%0D";
      default:
        return "";
    }
  });
}

export type FrameDrawer = {
  /** The frame as it stands now, on the one canvas. */
  draw(): Promise<HTMLCanvasElement>;
  /** Let go of the canvas and the picture element. */
  dispose(): void;
};

export function createFrameDrawer({
  width,
  height,
  background,
  toSvg,
  context,
  ownerDocument = document,
}: {
  width: number;
  height: number;
  /** The colour under the frame, or null for a transparent one. */
  background: string | null;
  /** The show at this instant as one SVG (`domToForeignObjectSvg`). */
  toSvg: () => Promise<Element>;
  /** The screenshot context the SVG is made with, reset after each frame. */
  context?: ScreenshotContextState;
  ownerDocument?: Document;
}): FrameDrawer {
  const canvas = ownerDocument.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context2d = canvas.getContext("2d");
  const picture = ownerDocument.createElement("img");
  picture.decoding = "sync";
  const serializer = new XMLSerializer();
  const resetContext = () => {
    if (!context) return;
    const doc = context.ownerDocument ?? ownerDocument;
    const style = doc.createElement("style");
    style.appendChild(
      doc.createTextNode("\n.______background-clip--text {\n  background-clip: text;\n  -webkit-background-clip: text;\n}\n"),
    );
    context.svgStyleElement = style;
    context.svgDefsElement = doc.createElementNS(XMLNS, "defs") as SVGDefsElement;
    context.svgStyles.clear();
    // The next frame names its own characters, so it carries only its own
    // faces (`frame-fonts`), not every face the recording has met.
    context.fontFamilies?.clear();
  };
  return {
    async draw() {
      const svg = await toSvg();
      resetContext();
      // A data: URL, as `modern-screenshot` makes: a picture of an SVG with a
      // foreignObject loaded from a blob: URL taints the canvas, and a
      // tainted canvas cannot become a video frame. In one pass: characters
      // XML cannot hold are dropped (as `modern-screenshot` does: they would
      // make the whole picture unreadable), and only what a URL cannot carry
      // as it is is escaped -- `%` and `#`, and the tabs and line breaks a URL
      // parser strips -- not every character `encodeURIComponent` escapes (a
      // copy 1.4 times the frame's markup, 30 times a second; the picture is
      // the same pixel for pixel, four decks).
      picture.src = `data:image/svg+xml;charset=utf-8,${frameDataUrlBody(serializer.serializeToString(svg))}`;
      try {
        // A frame the browser cannot decode keeps the one before it.
        const decoded = await picture.decode().then(
          () => true,
          () => false,
        );
        if (decoded && context2d) {
          context2d.clearRect(0, 0, width, height);
          if (background) {
            context2d.fillStyle = background;
            context2d.fillRect(0, 0, width, height);
          }
          context2d.drawImage(picture, 0, 0, width, height);
        }
      } finally {
        // The frame's picture goes with it: the next frame's replaces it.
        picture.removeAttribute("src");
      }
      return canvas;
    },
    dispose() {
      picture.removeAttribute("src");
      canvas.width = 0;
      canvas.height = 0;
    },
  };
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Served faces' glyph outlines for WordArt text outlines: the face's bytes
 * read once (the same URL the page's @font-face draws from, so the browser's
 * cache answers), each glyph's outline merged where its contours overlap
 * (glyph-outline.ts). Runs in the outline worker; where there is none (tests,
 * a browser without module workers) on the page, as the same code.
 */
import {
  readCffSfntProgramAsTrueType,
  readSfntProgram,
  type SfntProgram,
} from "../pdf/engine/native/sfnt/sfnt-reader";
import { glyphOutline, type PathUnion } from "./glyph-outline";

export interface FaceOutlines {
  readonly unitsPerEm: number;
  /** [glyph id, path data (font units, y up), merged, bbox, horizontal
   * advance, vertical origin's height above the baseline] per glyph asked
   * for: an upright glyph in vertical text stands with its advance centred
   * on the line and its vertical origin at the top of its box (vmtx's top
   * side bearing over the ink, else the face's ascent). */
  readonly glyphs: readonly (readonly [number, string, boolean, readonly [number, number, number, number] | null, number, number])[];
  /** Glyph ids the face could not give an outline for. */
  readonly failed: readonly number[];
}

type Bytes = (url: string) => Promise<ArrayBuffer>;
type PathKitLoader = () => Promise<PathUnion>;
/** No PathKit here: overlapping contours are stroked as they stand. */
export const NO_PATH_UNION: PathKitLoader = () => Promise.reject(new Error("no PathKit here"));

/** PathKit's loader (pathkit-wasm's PathKitInit). */
export type PathKitInit = (options: { wasmBinary: Uint8Array }) => Promise<{
  FromSVGString(d: string): { setFillType(type: unknown): void; simplify(): unknown; toSVGString(): string; delete(): void } | null;
  FillType: { WINDING: unknown };
}>;

/**
 * PathKit (pathkit-wasm, BSD-3-Clause: Skia's PathOps), loaded the first
 * time a glyph's contours overlap: its loader from `loader` (the worker reads
 * the package's script as a file, so no bundle carries its Node branch), its
 * wasm from `wasmBytes`.
 */
export function pathKitUnion(
  loader: () => Promise<PathKitInit>, wasmBytes: () => Promise<ArrayBuffer | Uint8Array>,
): PathKitLoader {
  let loading: Promise<PathUnion> | null = null;
  return () => {
    loading ??= (async () => {
      const init = await loader();
      const kit = await init({ wasmBinary: new Uint8Array(await wasmBytes()) });
      return {
        union(d: string): string | null {
          const path = kit.FromSVGString(d);
          if (!path) return null;
          try {
            path.setFillType(kit.FillType.WINDING);
            if (!path.simplify()) return null;
            return path.toSVGString();
          } catch {
            return null;
          } finally {
            path.delete();
          }
        },
      };
    })();
    loading.catch(() => {
      loading = null;
    });
    return loading;
  };
}

export class OutlineEngine {
  private readonly programs = new Map<string, Promise<SfntProgram>>();
  private union: PathUnion | null = null;

  constructor(private readonly bytes: Bytes, private readonly loadUnion: PathKitLoader) {}

  private program(url: string): Promise<SfntProgram> {
    let known = this.programs.get(url);
    if (!known) {
      known = this.bytes(url).then((buffer) => {
        let data = new Uint8Array(buffer);
        let signature = String.fromCharCode(...data.subarray(0, 4));
        if (signature === "ttcf") {
          // A collection: its first face, as @font-face draws one. Its table
          // offsets count from the file's start, so its table directory is
          // put at the start of a copy.
          const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
          const at = view.getUint32(12);
          const length = 12 + 16 * view.getUint16(at + 4);
          const copy = data.slice();
          copy.set(data.subarray(at, at + length), 0);
          data = copy;
          signature = String.fromCharCode(...data.subarray(0, 4));
        }
        return signature === "OTTO" ? readCffSfntProgramAsTrueType(data) : readSfntProgram(data);
      });
      this.programs.set(url, known);
      known.catch(() => this.programs.delete(url));
    }
    return known;
  }

  /** Forget the faces read (a document closed): their programs are the
   * engine's only large memory. */
  release(): void {
    this.programs.clear();
  }

  /** Each path (SVG path data, non-zero) with its overlaps merged; a path
   * PathKit declines, or every path where it cannot load, as it stands. */
  async unionPaths(paths: readonly string[]): Promise<string[]> {
    try {
      this.union ??= await this.loadUnion();
    } catch {
      return [...paths];
    }
    const union = this.union;
    return paths.map((d) => (d ? union.union(d) ?? d : d));
  }

  async outlines(url: string, glyphIds: readonly number[]): Promise<FaceOutlines> {
    const program = await this.program(url);
    const glyphs: [number, string, boolean, readonly [number, number, number, number] | null, number, number][] = [];
    const topBearing = topSideBearings(program);
    const failed: number[] = [];
    for (const glyphId of glyphIds) {
      const outline = glyphOutline(program, glyphId, this.union);
      if (!outline) {
        failed.push(glyphId);
        continue;
      }
      let { d, merged } = outline;
      if (outline.overlaps && !merged) {
        // The first overlapping glyph loads PathKit; a union it declines
        // leaves the contours as they stand.
        this.union ??= await this.loadUnion().catch(() => null);
        const union = this.union?.union(d) ?? null;
        if (union !== null) {
          d = union;
          merged = true;
        }
      }
      const tsb = topBearing?.(glyphId);
      const vOrigin = tsb !== undefined && outline.bbox ? tsb + outline.bbox[3] : program.ascent;
      glyphs.push([glyphId, d, merged, outline.bbox, program.advances[glyphId] ?? 0, vOrigin]);
    }
    return { unitsPerEm: program.unitsPerEm, glyphs, failed };
  }
}

/** Each glyph's top side bearing (`vmtx`), or null when the face has no
 * vertical metrics. */
function topSideBearings(program: SfntProgram): ((glyphId: number) => number) | null {
  const vhea = program.tables.get("vhea");
  const vmtx = program.tables.get("vmtx");
  if (!vhea || !vmtx || vhea.length < 36) return null;
  const metrics = new DataView(vhea.buffer, vhea.byteOffset, vhea.byteLength).getUint16(34);
  const view = new DataView(vmtx.buffer, vmtx.byteOffset, vmtx.byteLength);
  return (glyphId) => {
    const at = glyphId < metrics ? glyphId * 4 + 2 : metrics * 4 + (glyphId - metrics) * 2;
    return at + 2 <= view.byteLength ? view.getInt16(at) : 0;
  };
}

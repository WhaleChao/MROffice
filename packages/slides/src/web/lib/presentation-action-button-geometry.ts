/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ACTION_BUTTON_PATHS } from "./presentation-action-button-paths";
import type { NativeDrawingGeometry } from "@mrnote/core/web/lib/native-drawing-geometry.ts";

/** Default button icons retain their aspect ratio when their frame is resized. */
export function presentationActionButtonGeometry(kind: string, width: number, height: number): NativeDrawingGeometry | null {
  const paths = Object.hasOwn(ACTION_BUTTON_PATHS, kind) ? ACTION_BUTTON_PATHS[kind] : undefined;
  if (!paths || [width, height].some(value => !Number.isFinite(value) || value <= 0 || value > 100_000)) return null;
  const scale = Math.min(width, height) / 100;
  const point = (value: number, size: number) => value === 0 ? 0 : value === 100 ? size : size / 2 + (value - 50) * scale;
  const number = (value: number) => String(Number(value.toFixed(7)));
  return {
    schemaVersion: "drawingml-paths-v1", viewBox: [width, height],
    paths: paths.map(path => {
      const tokens = path.d.match(/[MLAZ]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)/g)!;
      const commands: string[] = [];
      for (let i = 0; i < tokens.length;) {
        const command = tokens[i++];
        if (command === "Z") { commands.push(command); continue; }
        if (command === "A") {
          const rx = Number(tokens[i++]) * scale, ry = Number(tokens[i++]) * scale;
          const rotation = tokens[i++], large = tokens[i++], sweep = tokens[i++];
          const x = point(Number(tokens[i++]), width), y = point(Number(tokens[i++]), height);
          commands.push(`A${number(rx)} ${number(ry)} ${rotation} ${large} ${sweep} ${number(x)} ${number(y)}`);
        } else {
          commands.push(`${command}${number(point(Number(tokens[i++]), width))} ${number(point(Number(tokens[i++]), height))}`);
        }
      }
      return { ...path, d: commands.join(" ") };
    }),
  };
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One sentence for every reason 转换为可编辑 brought nothing back.
 *
 * Two shapes of refusal reach the same three hosts and the same key family:
 * the PICTURE refusals R15 added (`no-page` / `unavailable` / `incomplete`,
 * `SMARTART-IMPORT.md` §5.57), which are plain strings, and the CONVERSION
 * refusals one layer below them (`over-capacity` / `no-scene`, §5.61), which
 * carry two numbers. Spelling the key and choosing the values is the whole of
 * the difference, so it is done once here rather than three times in three
 * host components.
 *
 * The host still calls its own `t`: the message catalogue is typed per surface
 * and this module has no business importing a translator.
 */

import type { SmartArtConvertRefusal } from "./native-convert-model";
import type { SmartArtConversionRefusal } from "./slot-media-client";

export type SmartArtRefusal = SmartArtConversionRefusal | SmartArtConvertRefusal;

export interface SmartArtRefusalMessage {
  /** The flat key under `toolbar.smartartContext.convertRefused`. */
  readonly key: string;
  /**
   * The numbers the capacity sentence interpolates. Always present, because a
   * message that names no placeholder simply ignores them, and a message that
   * names one must never be handed nothing.
   */
  readonly values: { readonly limit: number; readonly count: number };
}

export function smartArtRefusalMessage(
  refusal: SmartArtRefusal,
): SmartArtRefusalMessage {
  const reason = typeof refusal === "string" ? refusal : refusal.reason;
  return {
    key: `toolbar.smartartContext.convertRefused.${reason}`,
    values: { limit: 0, count: 0 },
  };
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The editors' interface text, by locale and namespace (generated).
export const MESSAGE_LOADERS: Record<string, () => Promise<Record<string, Record<string, unknown>>>> = {
  "da": async () => ({ "console-notebooks": (await import("./da/console-notebooks.json")).default, "common": (await import("./da/common.json")).default }),
  "de": async () => ({ "console-notebooks": (await import("./de/console-notebooks.json")).default, "common": (await import("./de/common.json")).default }),
  "en": async () => ({ "console-notebooks": (await import("./en/console-notebooks.json")).default, "common": (await import("./en/common.json")).default }),
  "en-GB": async () => ({ "console-notebooks": (await import("./en-GB/console-notebooks.json")).default, "common": (await import("./en-GB/common.json")).default }),
  "es": async () => ({ "console-notebooks": (await import("./es/console-notebooks.json")).default, "common": (await import("./es/common.json")).default }),
  "es-419": async () => ({ "console-notebooks": (await import("./es-419/console-notebooks.json")).default, "common": (await import("./es-419/common.json")).default }),
  "fi": async () => ({ "console-notebooks": (await import("./fi/console-notebooks.json")).default, "common": (await import("./fi/common.json")).default }),
  "fr": async () => ({ "console-notebooks": (await import("./fr/console-notebooks.json")).default, "common": (await import("./fr/common.json")).default }),
  "id": async () => ({ "console-notebooks": (await import("./id/console-notebooks.json")).default, "common": (await import("./id/common.json")).default }),
  "it": async () => ({ "console-notebooks": (await import("./it/console-notebooks.json")).default, "common": (await import("./it/common.json")).default }),
  "ja": async () => ({ "console-notebooks": (await import("./ja/console-notebooks.json")).default, "common": (await import("./ja/common.json")).default }),
  "ko": async () => ({ "console-notebooks": (await import("./ko/console-notebooks.json")).default, "common": (await import("./ko/common.json")).default }),
  "nl": async () => ({ "console-notebooks": (await import("./nl/console-notebooks.json")).default, "common": (await import("./nl/common.json")).default }),
  "no": async () => ({ "console-notebooks": (await import("./no/console-notebooks.json")).default, "common": (await import("./no/common.json")).default }),
  "pt-BR": async () => ({ "console-notebooks": (await import("./pt-BR/console-notebooks.json")).default, "common": (await import("./pt-BR/common.json")).default }),
  "sv": async () => ({ "console-notebooks": (await import("./sv/console-notebooks.json")).default, "common": (await import("./sv/common.json")).default }),
  "th": async () => ({ "console-notebooks": (await import("./th/console-notebooks.json")).default, "common": (await import("./th/common.json")).default }),
  "vi": async () => ({ "console-notebooks": (await import("./vi/console-notebooks.json")).default, "common": (await import("./vi/common.json")).default }),
  "zh": async () => ({ "console-notebooks": (await import("./zh/console-notebooks.json")).default, "common": (await import("./zh/common.json")).default }),
  "zh-TW": async () => ({ "console-notebooks": (await import("./zh-TW/console-notebooks.json")).default, "common": (await import("./zh-TW/common.json")).default }),
};
export const MESSAGE_NAMESPACES = ["common","console-notebooks"];

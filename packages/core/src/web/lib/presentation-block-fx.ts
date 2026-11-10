/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-block present-mode entrance animations + block hyperlinks
 * (PowerPoint-parity subset).
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * controls, and the present-mode block renderer so the allowed animation set and
 * the href scheme whitelist live in exactly one place. Keep this module
 * dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * imported from the zod schema.
 *
 * Mirrors the conventions of `presentation-transitions.ts` (slide-level) and
 * `presentation-shapes.ts`: a frozen allowed list, a type guard, a normalizer
 * that collapses unknown/stale/hand-edited values to a safe default, and a
 * class-name resolver that returns `null` for the no-op default.
 */

/**
 * The supported block animations. `none` is the implicit default.
 *
 * Categories (PowerPoint parity):
 *  - ENTRANCE (original set, order-locked for back-compat): `none`, `fade`,
 *    `fly-in`, `zoom` — play once as the block first appears.
 *  - EXIT (additive): `fade-out`, `fly-out` — play the block off-stage. With
 *    `animation-fill-mode: both` the keyframe's final (hidden / translated-away)
 *    frame holds, so the block stays gone after the effect.
 *  - EMPHASIS (additive): `pulse`, `spin`, `bounce`, `grow-shrink` — draw the
 *    eye and RETURN to the block's resting state (each keyframe ends at the
 *    neutral opacity/transform) so the block looks unchanged once it settles.
 *
 * Every non-`none` value has a matching
 * `.presentation-present__block--anim-<value>` rule + `@keyframes` in
 * `styles/native-page-window.css`; the timeline and single-block controls share
 * this one vocabulary, so adding a value here without its CSS would render as a
 * no-op. Keep the two in lock-step.
 */
export const BLOCK_ANIMATIONS = [
  "none",
  "fade",
  "fly-in",
  "zoom",
  "fade-out",
  "fly-out",
  "pulse",
  "spin",
  "bounce",
  "grow-shrink",
] as const;

export type BlockAnimation = (typeof BLOCK_ANIMATIONS)[number];

/** Provider-free PowerPoint click-navigation targets we can execute natively. */
export const PRESENTATION_RELATIVE_CLICK_ACTION_TARGETS = [
  "previous",
  "next",
  "first",
  "last",
  "last-viewed",
  // Action Settings > Hyperlink to: End Show (PPT-B012).
  "end-show",
] as const;

export type PresentationRelativeClickActionTarget =
  (typeof PRESENTATION_RELATIVE_CLICK_ACTION_TARGETS)[number];

export interface PresentationRelativeClickAction {
  readonly kind: "relative-slide";
  readonly target: PresentationRelativeClickActionTarget;
}

export const PRESENTATION_MEDIA_CLICK_ACTION_KINDS = ["video", "audio"] as const;

export type PresentationMediaClickActionKind =
  (typeof PRESENTATION_MEDIA_CLICK_ACTION_KINDS)[number];

/** Stable same-slide media action used by Movie/Sound Action Buttons. */
export interface PresentationMediaClickAction {
  readonly kind: "play-media";
  readonly mediaKind: PresentationMediaClickActionKind;
  readonly targetBlockId: string;
}

export type PresentationClickAction =
  | PresentationRelativeClickAction
  | PresentationMediaClickAction;

const PRESENTATION_RELATIVE_CLICK_ACTION_TARGET_SET: ReadonlySet<string> =
  new Set(PRESENTATION_RELATIVE_CLICK_ACTION_TARGETS);

export const PRESENTATION_RELATIVE_CLICK_ACTION_URIS = Object.freeze({
  previous: "ppaction://hlinkshowjump?jump=previousslide",
  next: "ppaction://hlinkshowjump?jump=nextslide",
  first: "ppaction://hlinkshowjump?jump=firstslide",
  last: "ppaction://hlinkshowjump?jump=lastslide",
  "last-viewed": "ppaction://hlinkshowjump?jump=lastslideviewed",
  "end-show": "ppaction://hlinkshowjump?jump=endshow",
} satisfies Record<PresentationRelativeClickActionTarget, string>);

const PRESENTATION_RELATIVE_CLICK_ACTION_BY_URI = new Map(
  Object.entries(PRESENTATION_RELATIVE_CLICK_ACTION_URIS).map(
    ([target, uri]) => [uri, target as PresentationRelativeClickActionTarget],
  ),
);

/** Exact, strict persisted action parser; extra keys and stale targets refuse. */
export function normalizePresentationRelativeClickAction(
  value: unknown,
): PresentationRelativeClickAction | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    keys.length !== 2 ||
    !keys.includes("kind") ||
    !keys.includes("target") ||
    record.kind !== "relative-slide" ||
    typeof record.target !== "string" ||
    !PRESENTATION_RELATIVE_CLICK_ACTION_TARGET_SET.has(record.target)
  ) {
    return null;
  }
  return {
    kind: "relative-slide",
    target: record.target as PresentationRelativeClickActionTarget,
  };
}

const PRESENTATION_MEDIA_CLICK_ACTION_KIND_SET: ReadonlySet<string> = new Set(
  PRESENTATION_MEDIA_CLICK_ACTION_KINDS,
);

export function normalizePresentationMediaClickAction(
  value: unknown,
): PresentationMediaClickAction | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    Object.keys(record).length !== 3 ||
    record.kind !== "play-media" ||
    typeof record.mediaKind !== "string" ||
    !PRESENTATION_MEDIA_CLICK_ACTION_KIND_SET.has(record.mediaKind) ||
    typeof record.targetBlockId !== "string" ||
    !/^[A-Za-z0-9_-]{1,200}$/.test(record.targetBlockId)
  ) {
    return null;
  }
  return {
    kind: "play-media",
    mediaKind: record.mediaKind as PresentationMediaClickActionKind,
    targetBlockId: record.targetBlockId,
  };
}

export function normalizePresentationClickAction(
  value: unknown,
): PresentationClickAction | null {
  return (
    normalizePresentationRelativeClickAction(value) ??
    normalizePresentationMediaClickAction(value)
  );
}

export function presentationRelativeClickActionUri(
  value: unknown,
): string | null {
  const action = normalizePresentationRelativeClickAction(value);
  return action ? PRESENTATION_RELATIVE_CLICK_ACTION_URIS[action.target] : null;
}

export function presentationRelativeClickActionFromUri(
  value: unknown,
): PresentationRelativeClickAction | null {
  if (typeof value !== "string") return null;
  const target = PRESENTATION_RELATIVE_CLICK_ACTION_BY_URI.get(value);
  return target ? { kind: "relative-slide", target } : null;
}

/** Exact twelve authorable Action Button DrawingML geometries. */
export const PRESENTATION_ACTION_BUTTON_SHAPE_KINDS = [
  "action-button-back-previous",
  "action-button-forward-next",
  "action-button-beginning",
  "action-button-end",
  "action-button-home",
  "action-button-information",
  "action-button-return",
  "action-button-movie",
  "action-button-document",
  "action-button-sound",
  "action-button-help",
  "action-button-custom",
] as const;

export type PresentationActionButtonShapeKind =
  (typeof PRESENTATION_ACTION_BUTTON_SHAPE_KINDS)[number];

/** Imported button geometry can have no hlinkClick; never invent a jump. */
export function presentationHasAbsentNativeAction(block: Record<string, unknown>): boolean {
  const source = block.source;
  return Boolean(source && typeof source === "object" && !Array.isArray(source) &&
    (source as Record<string, unknown>).nativeActionDisposition === "absent" &&
    block.href === undefined && block.clickAction === undefined);
}

const PRESENTATION_ACTION_BUTTON_SHAPE_KIND_SET: ReadonlySet<string> = new Set(
  PRESENTATION_ACTION_BUTTON_SHAPE_KINDS,
);

/** Six provider-free buttons whose DrawingML geometry has a fixed native jump. */
export const PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS = Object.freeze({
  "action-button-back-previous": {
    kind: "relative-slide",
    target: "previous",
  },
  "action-button-forward-next": { kind: "relative-slide", target: "next" },
  "action-button-beginning": { kind: "relative-slide", target: "first" },
  "action-button-end": { kind: "relative-slide", target: "last" },
  "action-button-home": { kind: "relative-slide", target: "first" },
  "action-button-return": {
    kind: "relative-slide",
    target: "last-viewed",
  },
} satisfies Record<string, PresentationRelativeClickAction>);

export function isPresentationActionButtonShapeKind(
  value: unknown,
): value is PresentationActionButtonShapeKind {
  return (
    typeof value === "string" &&
    PRESENTATION_ACTION_BUTTON_SHAPE_KIND_SET.has(value)
  );
}

export function defaultPresentationClickActionForShape(
  value: unknown,
): PresentationRelativeClickAction | null {
  if (
    !isPresentationActionButtonShapeKind(value) ||
    !(value in PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS)
  ) {
    return null;
  }
  return {
    ...PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS[
      value as keyof typeof PRESENTATION_ACTION_BUTTON_DEFAULT_ACTIONS
    ],
  };
}

/** The animation applied when a block has no explicit (or invalid) value. */
export const DEFAULT_BLOCK_ANIMATION: BlockAnimation = "none";

const BLOCK_ANIMATION_SET: ReadonlySet<string> = new Set(BLOCK_ANIMATIONS);

/**
 * Type guard: is `value` exactly one of the allowed animation identifiers?
 * Exact, case-sensitive — use {@link normalizeBlockAnimation} for lenient input.
 */
export function isBlockAnimation(value: unknown): value is BlockAnimation {
  return typeof value === "string" && BLOCK_ANIMATION_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed animation. Unknown,
 * empty, or non-string values collapse to {@link DEFAULT_BLOCK_ANIMATION} so a
 * stale or hand-edited deck can never inject an unsupported animation. The
 * comparison is case-insensitive and trims surrounding whitespace; everything
 * else (including any other casing/spacing variation) falls back to the default.
 */
export function normalizeBlockAnimation(value: unknown): BlockAnimation {
  if (typeof value !== "string") return DEFAULT_BLOCK_ANIMATION;
  const candidate = value.trim().toLowerCase();
  return isBlockAnimation(candidate) ? candidate : DEFAULT_BLOCK_ANIMATION;
}

/**
 * The CSS class that plays an animation's keyframes when a block's slide becomes
 * active in present mode. Returns `null` for `none` (and any value that
 * normalizes to it) so the caller can skip adding an animation class entirely.
 * The class names are defined in `styles/native-page-window.css`
 * (`.presentation-present__block--anim-*`).
 */
export function blockAnimationClassName(value: unknown): string | null {
  const animation = normalizeBlockAnimation(value);
  if (animation === "none") return null;
  return `presentation-present__block--anim-${animation}`;
}

/**
 * The `@keyframes` NAME a non-`none` block animation references, or `null` for
 * `none`. Mirrors the `.presentation-present__block--anim-* { animation: <name> }`
 * rules in `styles/native-page-window.css` (every rule leads with this name).
 * Used by the present-mode renderer to COMPOSE the entrance/emphasis/exit
 * keyframe with a motion-path keyframe inline (a single CSS `animation` shorthand
 * can list both), so a step carrying both effects plays both without one
 * overwriting the other. Kept in lock-step with the CSS (the CSS-presence test
 * asserts each class's `animation:` leads with exactly this name).
 */
export function blockAnimationKeyframeName(value: unknown): string | null {
  const animation = normalizeBlockAnimation(value);
  if (animation === "none") return null;
  return `presentation-present-block-${animation}`;
}

/**
 * The URL schemes a block hyperlink may use. Deliberately tiny: anything that
 * can execute script (`javascript:`) or smuggle inline payloads (`data:`,
 * `vbscript:`, `file:`, etc.) is rejected. `mailto:` is allowed for contact
 * links; `http`/`https` for the web.
 */
export const ALLOWED_HREF_SCHEMES = ["http:", "https:", "mailto:"] as const;

const ALLOWED_HREF_SCHEME_SET: ReadonlySet<string> = new Set(
  ALLOWED_HREF_SCHEMES,
);

/**
 * Leading scheme of a URL, e.g. `javascript:` or `https:`. A scheme is
 * `ALPHA *( ALPHA / DIGIT / "+" / "-" / "." ) ":"` per RFC 3986; we capture the
 * scheme (including the colon) and lower-case it for comparison. Anchored at the
 * start so leading whitespace must already be trimmed by the caller.
 */
const URL_SCHEME_RE = /^([a-zA-Z][a-zA-Z0-9+.-]*:)/;

/**
 * Matches any ASCII control character (U+0000–U+001F or U+007F). Browsers strip
 * embedded control chars (newline/tab/etc.) when parsing a URL, which would let
 * `java\tscript:alert(1)` slip past a naive scheme check; rejecting them up
 * front closes that bypass.
 */
const CONTROL_CHAR_RE = /[\u0000-\u001f\u007f]/;

/**
 * The pseudo-scheme for an INTERNAL slide link: `slide://<slideId>`. Clicking a
 * block carrying one jumps to that slide in present mode instead of navigating
 * away. This is a separate, additive target kind to the external
 * {@link ALLOWED_HREF_SCHEMES} web/mail whitelist; it never resolves to a URL,
 * so it can never be turned into a clickable `href` (the renderer wires it to an
 * in-app jump handler, never an `<a href>`).
 */
export const SLIDE_LINK_SCHEME = "slide:";

const SLIDE_LINK_PREFIX = "slide://";

/**
 * The id-shape a `slide://` target may carry. Deliberately strict — the same
 * conservative slug class the rest of the deck uses for ids (alphanumerics plus
 * `-`/`_`, bounded length). Anchored end-to-end so no stray scheme, slash,
 * whitespace, control char, or path/query can ride along inside the target. This
 * is the structural floor; {@link sanitizeBlockHref} additionally validates the
 * id against the live deck when a known-id set is supplied.
 */
const SLIDE_LINK_ID_RE = /^[A-Za-z0-9_-]{1,200}$/;

/**
 * Parse a raw href into a `slide://` internal target. Returns the bare slide id
 * when `value` is a well-formed `slide://<id>` whose id matches the strict id
 * shape (and, when `knownSlideIds` is supplied, references a slide that EXISTS
 * in the deck); otherwise `null`. Never throws. The match is exact on the
 * `slide://` prefix (case-sensitive scheme to stay byte-stable with what we
 * persist) AFTER trimming + control-char rejection, so an obfuscated or
 * partial scheme can't slip through.
 */
export function parseSlideLinkTarget(
  value: unknown,
  knownSlideIds?: ReadonlySet<string> | readonly string[],
): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (CONTROL_CHAR_RE.test(trimmed)) return null;
  if (!trimmed.startsWith(SLIDE_LINK_PREFIX)) return null;
  const id = trimmed.slice(SLIDE_LINK_PREFIX.length);
  if (!SLIDE_LINK_ID_RE.test(id)) return null;
  if (knownSlideIds) {
    const known =
      knownSlideIds instanceof Set ? knownSlideIds : new Set(knownSlideIds);
    if (!known.has(id)) return null;
  }
  return id;
}

/**
 * Build the canonical `slide://<id>` href for an internal slide link from a bare
 * slide id, or `null` when the id is not a valid slide-id shape. Used by the
 * editor's slide picker so the persisted value always round-trips through
 * {@link parseSlideLinkTarget}.
 */
export function buildSlideLinkHref(slideId: unknown): string | null {
  if (typeof slideId !== "string") return null;
  const trimmed = slideId.trim();
  if (!SLIDE_LINK_ID_RE.test(trimmed)) return null;
  return `${SLIDE_LINK_PREFIX}${trimmed}`;
}

/**
 * Is `value` an internal `slide://` link (well-formed scheme + id shape, and —
 * when `knownSlideIds` is given — an existing slide)? Thin boolean wrapper over
 * {@link parseSlideLinkTarget}.
 */
export function isSlideLinkHref(
  value: unknown,
  knownSlideIds?: ReadonlySet<string> | readonly string[],
): boolean {
  return parseSlideLinkTarget(value, knownSlideIds) !== null;
}

/**
 * Validate and normalize a block hyperlink. Returns the trimmed href when it is
 * safe to render, or `null` when it is missing, malformed, or uses a disallowed
 * / dangerous scheme (e.g. `javascript:`, `data:`, `vbscript:`, `file:`).
 *
 * Two accepted target kinds:
 *  - An EXTERNAL web/mail URL against the {@link ALLOWED_HREF_SCHEMES} whitelist
 *    (`http`/`https`/`mailto`) — rendered as a real `<a href>`.
 *  - An INTERNAL `slide://<slideId>` link (see {@link parseSlideLinkTarget}) —
 *    rendered as an in-app jump, never a navigable URL. When `knownSlideIds` is
 *    supplied the slide id MUST reference an existing slide; otherwise only the
 *    strict scheme + id shape is enforced (so a render edge with no deck context
 *    still rejects garbage but tolerates a not-yet-resolved id).
 *
 * Rules (external):
 *  - Non-string / empty / whitespace-only → `null`.
 *  - An explicit scheme must be one of the whitelist (case-insensitive scheme).
 *    The scheme is lower-cased; the rest of the URL is preserved verbatim so a
 *    case-sensitive path/query/mailto address is not mangled.
 *  - A scheme-less value (no leading `scheme:`) is treated as a relative /
 *    protocol-relative reference and rejected — block links must be explicit,
 *    absolute web/mail URLs so present mode never resolves against an unexpected
 *    base. (`//evil.example` is scheme-less here and therefore rejected.)
 *  - Control characters (incl. embedded newlines/tabs that browsers strip when
 *    parsing a URL, a classic `java\nscript:` bypass) cause a rejection.
 *  - An http(s) URL must parse and name a host; a `mailto:` must name an
 *    address (`https://` or `mailto:` alone is rejected).
 */
export function sanitizeBlockHref(
  value: unknown,
  knownSlideIds?: ReadonlySet<string> | readonly string[],
): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (CONTROL_CHAR_RE.test(trimmed)) return null;
  // Internal slide link: a well-formed `slide://<id>` (and an existing slide
  // when a deck id set is supplied) round-trips to its canonical form. Checked
  // before the scheme whitelist because `slide:` is deliberately NOT in it.
  const slideTarget = parseSlideLinkTarget(trimmed, knownSlideIds);
  if (slideTarget !== null) return `${SLIDE_LINK_PREFIX}${slideTarget}`;
  const match = URL_SCHEME_RE.exec(trimmed);
  if (!match) {
    // No explicit scheme: relative or protocol-relative — not allowed.
    return null;
  }
  const scheme = match[1].toLowerCase();
  if (!ALLOWED_HREF_SCHEME_SET.has(scheme)) return null;
  // Preserve the original (case-sensitive) remainder, only normalizing the
  // scheme casing so persisted hrefs stay stable.
  const href = scheme + trimmed.slice(match[1].length);
  // A web address must name a host and a mail link an address: `https://`
  // alone is no destination at all (PPT-B021).
  if (scheme === "mailto:") return href.length > scheme.length ? href : null;
  try {
    return new URL(href).hostname ? href : null;
  } catch {
    // A link a deck states that no URL parser reads ("http:// https:/x...", a
    // typo PowerPoint keeps and opens as written) stays as the package says:
    // dropping it made every save of the deck plan the link's removal, which
    
    // whitelisted above; only a bare scheme names nothing.
    return href.slice(scheme.length).replace(/^\/\//, "").trim() ? href : null;
  }
}

/**
 * Convenience: is `value` an acceptable block hyperlink (external URL or
 * internal `slide://` link)? Thin wrapper over {@link sanitizeBlockHref} for
 * places that only need a boolean.
 */
export function isSafeBlockHref(
  value: unknown,
  knownSlideIds?: ReadonlySet<string> | readonly string[],
): boolean {
  return sanitizeBlockHref(value, knownSlideIds) !== null;
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isDoNotTrackEnabled, sendErrorBeacon } from "./errors/report-beacon";
import { inkFieldSwitchQuery } from "./ink-field-switches";

/** Dispatched on `window` when the browser withheld a pen's contacts (a `deaf` stall). */
export const INK_SCRIBBLE_STALL_EVENT = "mrai:ink-scribble-stall";

/** A pen contact this long without an input event of any kind is a stall. */
export const INK_INPUT_STALL_MS = 2000;
/**
 * An input event handed to the page this long after it happened: the page's
 * own main thread was busy for that long (its timers could not run either,
 * so a silence they were to measure is only seen here, afterwards).
 */
export const INK_INPUT_BUSY_MS = 800;
/**
 * A pen's hover that pauses this long and resumes with nothing heard in
 * between was a contact the page never got (a tap is shorter than a word,
 * a pen put down longer than this has left).
 */
export const INK_INPUT_DEAF_GAP_MIN_MS = 60;
export const INK_INPUT_DEAF_GAP_MAX_MS = 3000;
/** This many unheard contacts within {@link INK_INPUT_DEAF_WINDOW_MS}: the page is deaf. */
export const INK_INPUT_DEAF_CONTACTS = 3;
export const INK_INPUT_DEAF_WINDOW_MS = 12000;
/** The hover resumes this close (px) to where it paused: the pen came down where it was. */
const DEAF_NEAR_PX = 160;
/** Work of the page's own logged in the trail when it took at least this long. */
const MARK_MIN_MS = 24;
const TRAIL_DEFAULT = 48;















const HEARTBEAT_MS = 250;
const HELD_MIN_MS = 400;
const SPLIT_MIN_MS = 100;
const SPLIT_EVENTS = new Set(["pointerdown", "pointerup", "click", "gesturestart", "gestureend"]);
const MAX_REPORTS_DEFAULT = 4;

/**
 * Field diagnosis switches, read once from the address the tab was opened
 * with (lib/ink-field-switches keeps them for the tab): `inkdiag=1` reports every deaf
 * episode instead of one per page; `inkpd=0` stops the ink surfaces cancelling
 * a pen's pointerdown (an A/B test of WebKit dropping Pencil contacts);
 * `inkscrib=off` stops them cancelling a Pencil's touchstart (lib/ink-touch-shield).
 */
const fieldSwitches: {
  readonly diag: boolean;
  readonly penPreventDefault: boolean;
  /** `inkscrib=off`: the ink surfaces stop cancelling a Pencil's touchstart (Scribble comparison). */
  readonly scribble: string | null;
  /** `inkshield=off`: no hand guard or finger pan, the ink surfaces only cancel touchstarts (Excalidraw's arrangement, a field comparison). */
  readonly shield: boolean;
} = (() => {
  try {
    const query = inkFieldSwitchQuery();
    return {
      diag: query.get("inkdiag") === "1",
      penPreventDefault: query.get("inkpd") !== "0",
      scribble: query.get("inkscrib"),
      shield: query.get("inkshield") !== "off",
    };
  } catch {
    return { diag: false, penPreventDefault: true, scribble: null, shield: true };
  }
})();

/** The `inkscrib` field switch (`off` lets Scribble see the ink surfaces again). */
/** The `inkshield` field switch: false runs the ink surfaces without the hand guard and finger pan. */
export function inkShieldEnabled(): boolean {
  return fieldSwitches.shield;
}

export function inkScribbleCancelSwitch(): string | null {
  return fieldSwitches.scribble;
}

/** Under the field diagnosis the trail is longer, more reports go, and every
 * written character's trail is sent once the pen rests. */
const TRAIL = fieldSwitches.diag ? 400 : TRAIL_DEFAULT;
const MAX_REPORTS = fieldSwitches.diag ? 80 : MAX_REPORTS_DEFAULT;
const DIAG_DUMP_IDLE_MS = 1800;

/** False under the `inkpd=0` field test: a pen's pointerdown is not cancelled. */
export function inkCancelsPenPointerDown(): boolean {
  return fieldSwitches.penPreventDefault;
}

interface TraceEntry {
  at: number;
  code: string;
}

export interface InkInputStall {
  /**
   * `silent`: nothing reached the page with a pen down; `busy`: input reached
   * it late; `deaf`: the pen's hover kept arriving but its contacts never did.
   */
  kind: "silent" | "busy" | "deaf" | "trace";
  /** How long the window heard nothing with a pen down, ms (`deaf`: the last unheard contact's length). */
  silentMs: number;
  /** The most a timer due in the silence ran late, ms (main thread busy). */
  lateMs: number;
  /** How the silence ended: the next event's type, or the page went away. */
  endedBy: string;
  /** `deaf`: how many contacts went unheard within the window. */
  contacts?: number;
  /** `deaf`: what lay under the pen and could take its contact (see the shield). */
  diag?: string;
  trail: readonly TraceEntry[];
}

type StallSink = (stall: InkInputStall, state: string) => void;

const EVENT_CODES: Record<string, string> = {
  pointerdown: "pd",
  pointermove: "pm",
  pointerup: "pu",
  pointercancel: "pc",
  touchstart: "ts",
  touchmove: "tm",
  touchend: "te",
  touchcancel: "tc",
  click: "ck",
  gesturestart: "gs",
  gestureend: "ge",
};

function eventCode(event: Event): string {
  const code = EVENT_CODES[event.type] ?? event.type.slice(0, 3);
  if ("pointerType" in event) {
    const pointer = event as PointerEvent;
    const kind = pointer.pointerType ? pointer.pointerType.charAt(0) : "?";
    return `${code}${kind}${pointer.buttons ? "" : "0"}`;
  }
  const touches = (event as TouchEvent).changedTouches;
  if (touches) {
    let kinds = "";
    for (let index = 0; index < touches.length; index += 1) {
      kinds += (touches[index] as Touch & { touchType?: string }).touchType === "stylus" ? "s" : "f";
    }
    return `${code}${kinds}${(event as TouchEvent).touches?.length ?? 0}`;
  }
  return code;
}

function stylusDown(event: Event): boolean {
  const touches = (event as TouchEvent).touches;
  if (!touches) return false;
  for (let index = 0; index < touches.length; index += 1) {
    if ((touches[index] as Touch & { touchType?: string }).touchType === "stylus") return true;
  }
  return false;
}

/** The recorder; one per window while an ink surface holds the touch shield. */
export class InkInputTrace {
  private readonly trail: TraceEntry[] = [];
  private busyReportedAt = Number.NEGATIVE_INFINITY;
  private lastEvent: Event | null = null;
  private penDown = false;
  private lastHeard = Number.NEGATIVE_INFINITY;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private timerDue = 0;
  private lateMs = 0;
  private stalled = false;
  private reports = 0;
  // The pen's hover, for a page gone deaf to its contact (see the top).
  private hoverAt = Number.NEGATIVE_INFINITY;
  private hoverPoint: { x: number; y: number } | null = null;
  private heardSinceHover = false;
  private penContactsHeard = 0;
  private unheard: number[] = [];
  private deafReported = false;

  constructor(
    private readonly sink: StallSink,
    private readonly describeState: () => string = () => "",
    private readonly now: () => number = () => performance.now(),
    /** What the page is at a point, for a deaf report (under the pen). */
    private readonly describePoint: (x: number, y: number) => string = () => "",
  ) {
    if (fieldSwitches.diag) this.watchFrames();
    this.beat();
  }

  private heartbeat: ReturnType<typeof setTimeout> | null = null;
  private disposed = false;

  /** The heartbeat (see {@link HEARTBEAT_MS}); a hidden page's timers are throttled, so it says nothing then. */
  private beat(): void {
    if (this.disposed || typeof setTimeout !== "function") return;
    const due = this.now() + HEARTBEAT_MS;
    this.heartbeat = setTimeout(() => {
      this.heartbeat = null;
      const late = this.now() - due;
      const hidden = typeof document !== "undefined" && document.visibilityState === "hidden";
      if (late >= HELD_MIN_MS && !hidden) this.push(due, `held${Math.round(late)}`);
      this.beat();
    }, HEARTBEAT_MS);
  }

  private splitPending = false;

  /** After a discrete input: script until the next frame, and that frame's cost (see {@link SPLIT_MIN_MS}). */
  private splitAfter(handedAt: number): void {
    if (this.splitPending || this.disposed || typeof requestAnimationFrame !== "function") return;
    this.splitPending = true;
    requestAnimationFrame(() => {
      const frame = this.now();
      setTimeout(() => {
        this.splitPending = false;
        if (this.disposed) return;
        const script = frame - handedAt;
        const render = this.now() - frame;
        if (script + render >= SPLIT_MIN_MS) this.push(handedAt, `js${Math.round(script)}rf${Math.round(render)}`);
      }, 0);
    });
  }

  private frameHandle = 0;
  /**
   * Field diagnosis: every frame the main thread held longer than two
   * frames, entered where it began (`busy<ms>`), so a contact the page never
   * heard can be set against what the page was doing at that moment.
   */
  private watchFrames(): void {
    if (typeof requestAnimationFrame !== "function") return;
    let last = this.now();
    const tick = () => {
      const now = this.now();
      const gap = now - last;
      if (gap > 34) this.push(last, `busy${Math.round(gap)}`);
      last = now;
      // What the frame cost after its callbacks (style, layout, paint, and
      // any task queued ahead): `F<ms>`, entered where the frame began.
      setTimeout(() => {
        const cost = this.now() - now;
        if (cost >= 8) this.push(now, `F${Math.round(cost)}`);
      }, 0);
      this.frameHandle = requestAnimationFrame(tick);
    };
    this.frameHandle = requestAnimationFrame(tick);
  }

  /** Every input event the window hears, before anything can swallow it. */
  heard(event: Event): void {
    const now = this.now();
    const at = event.timeStamp || now;
    const code = eventCode(event);
    if (this.stalled) this.endStall(at, code);
    else if (this.penDown && at - this.lastHeard >= INK_INPUT_STALL_MS) {
      // A silence the timers could not see: the main thread was held through it.
      this.stalled = true;
      this.endStall(at, code);
    }
    const lag = now - at;
    // Field diagnosis: how long a pen's contact waited before the page saw
    // it (the main thread was busy for that long when it arrived). The
    // touchstart's own answer to WebKit cannot come sooner than this.
    if (fieldSwitches.diag && lag >= 2 && ((event.type === "pointerdown" && (event as PointerEvent).pointerType === "pen") || (event.type === "touchstart" && code.includes("s")))) {
      this.push(now, `H:${code}+${Math.round(lag)}`);
    }
    if (lag >= INK_INPUT_BUSY_MS && now - this.busyReportedAt > INK_INPUT_STALL_MS) {
      this.busyReportedAt = now;
      this.push(at, code);
      this.report({ kind: "busy", silentMs: Math.round(lag), lateMs: Math.round(lag), endedBy: code, trail: this.trail.slice() });
    }
    this.lastHeard = at;
    this.lastEvent = event;
    // A move the pointer coalesced into the last entry adds nothing to tell.
    const last = this.trail[this.trail.length - 1];
    if (last && last.code === code && (code.startsWith("pm") || code.startsWith("tm"))) last.at = at;
    else this.push(at, code);
    const pointer = event as PointerEvent;
    const pen = pointer.pointerType === "pen";
    if (event.type === "pointermove" && pen && pointer.buttons === 0) {
      this.hoverResumed(at, pointer.clientX, pointer.clientY);
      this.hoverAt = at;
      this.hoverPoint = { x: pointer.clientX, y: pointer.clientY };
      this.heardSinceHover = false;
    } else {
      this.heardSinceHover = true;
      if ((event.type === "pointerdown" && pen) || (event.type === "touchstart" && code.includes("s"))) {
        // The page hears this pen's contact: whatever gaps came before were its lifts.
        this.penContactsHeard += 1;
        this.unheard = [];
      }
    }
    if (pen || code.includes("s")) this.scheduleDump();
    if (SPLIT_EVENTS.has(event.type)) this.splitAfter(now);
    if (event.type === "pointerdown" && pen) {
      this.penDown = true;
      this.arm();
    } else if (
      ((event.type === "pointerup" || event.type === "pointercancel") && pen) ||
      (event.type === "pointermove" && pen && (event as PointerEvent).buttons === 0) ||
      ((event.type === "touchend" || event.type === "touchcancel") && code.includes("s") && !stylusDown(event))
    ) {
      this.penDown = false;
      this.disarm();
    }
  }

  /** The page's own work that took `ms` (see {@link inkTraceWork}). */
  work(code: string, ms: number): void {
    this.push(this.now(), `${code}${Math.round(ms)}`);
  }

  /** A surface's own decision, entered in the trail (field diagnosis). */
  note(code: string): void {
    this.push(this.now(), code);
  }

  private dumpTimer: ReturnType<typeof setTimeout> | null = null;

  /** Field diagnosis: once the pen rests, the trail since the last one goes. */
  private scheduleDump(): void {
    if (!fieldSwitches.diag) return;
    if (this.dumpTimer !== null) clearTimeout(this.dumpTimer);
    this.dumpTimer = setTimeout(() => {
      this.dumpTimer = null;
      if (this.trail.length === 0) return;
      this.report({ kind: "trace", silentMs: 0, lateMs: 0, endedBy: "rest", trail: this.trail.slice() });
      this.trail.length = 0;
    }, DIAG_DUMP_IDLE_MS);
  }

  /** The trail so far, relative to its last entry (for the black box). */
  recent(count: number): string {
    const tail = this.trail.slice(-count);
    const end = tail.length > 0 ? tail[tail.length - 1].at : 0;
    return tail.map((entry) => `${Math.round(entry.at - end)} ${entry.code}`).join("\n");
  }

  /** The touch shield took this event from the page. */
  swallowed(event: Event): void {
    const last = this.trail[this.trail.length - 1];
    if (last && this.lastEvent === event && !last.code.endsWith("!")) last.code += "!";
  }

  /** The page is going away or hidden: a stall still running ends here. */
  leaving(reason: string): void {
    if (this.stalled) this.endStall(this.now(), reason);
    this.penDown = false;
    this.disarm();
  }

  dispose(): void {
    this.disposed = true;
    if (this.heartbeat !== null) clearTimeout(this.heartbeat);
    this.heartbeat = null;
    this.disarm();
    if (this.frameHandle && typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.frameHandle);
    this.frameHandle = 0;
  }

  private push(at: number, code: string): void {
    this.trail.push({ at, code });
    if (this.trail.length > TRAIL) this.trail.shift();
  }

  /**
   * The pen's hover is back. Paused for the length of a contact, resumed
   * close by, nothing heard meanwhile, on a page this pen has written on:
   * a contact the page never got. Enough of them within the window is a
   * deaf page, reported once.
   */
  private hoverResumed(at: number, x: number, y: number): void {
    if (this.heardSinceHover || !this.hoverPoint || this.penContactsHeard === 0) return;
    const gap = at - this.hoverAt;
    // The diagnosis counts a dot's short contact too.
    if (fieldSwitches.diag && gap >= 12 && gap < INK_INPUT_DEAF_GAP_MIN_MS && !this.heardSinceHover) this.push(at, `gap${Math.round(gap)}`);
    if (!(gap >= INK_INPUT_DEAF_GAP_MIN_MS && gap <= INK_INPUT_DEAF_GAP_MAX_MS)) return;
    if (Math.hypot(x - this.hoverPoint.x, y - this.hoverPoint.y) > DEAF_NEAR_PX) return;
    this.unheard = this.unheard.filter((when) => at - when <= INK_INPUT_DEAF_WINDOW_MS);
    this.unheard.push(at);
    this.push(at, `deaf${Math.round(gap)}`);
    if (this.unheard.length < INK_INPUT_DEAF_CONTACTS || this.deafReported) return;
    // Under the field diagnosis every episode is reported (3 unheard contacts
    // each), up to the page's report limit.
    if (fieldSwitches.diag) this.unheard = [];
    else this.deafReported = true;
    this.report({
      kind: "deaf",
      silentMs: Math.round(gap),
      lateMs: 0,
      endedBy: "pmp0",
      contacts: this.unheard.length,
      diag: this.describePoint(x, y),
      trail: this.trail.slice(),
    });
  }

  private arm(): void {
    this.disarm();
    this.lateMs = 0;
    this.schedule(INK_INPUT_STALL_MS / 4);
  }

  private schedule(delay: number): void {
    this.timerDue = this.now() + delay;
    this.timer = setTimeout(() => this.tick(), delay);
  }

  private tick(): void {
    this.timer = null;
    const now = this.now();
    this.lateMs = Math.max(this.lateMs, now - this.timerDue);
    if (!this.penDown) return;
    if (!this.stalled && now - this.lastHeard >= INK_INPUT_STALL_MS) {
      this.stalled = true;
      this.push(now, "stall");
    }
    // Keep watching while the pen is down: the silence's length and how late
    // the timers ran are what the report is for.
    this.schedule(INK_INPUT_STALL_MS / 4);
  }

  private disarm(): void {
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
  }

  private endStall(at: number, endedBy: string): void {
    this.stalled = false;
    const stall: InkInputStall = {
      kind: "silent",
      silentMs: Math.round(at - this.lastHeard),
      lateMs: Math.round(this.lateMs),
      endedBy,
      trail: this.trail.slice(),
    };
    this.lateMs = 0;
    this.report(stall);
  }

  private report(stall: InkInputStall): void {
    // The browser withholding the pen's contacts (iPadOS Scribble's gesture
    // fault): the page can tell the writer how to stop it (ScribbleStallNotice).
    if (stall.kind === "deaf" && typeof window !== "undefined") {
      window.dispatchEvent(new CustomEvent(INK_SCRIBBLE_STALL_EVENT));
    }
    if (this.reports >= MAX_REPORTS) return;
    this.reports += 1;
    this.sink(stall, this.describeState());
  }
}

/** The last whole lines of `text` that fit in `limit` characters. */
function tailWithin(text: string, limit: number): string {
  if (text.length <= limit) return text;
  const cut = text.slice(text.length - Math.max(0, limit));
  const newline = cut.indexOf("\n");
  return newline >= 0 ? cut.slice(newline + 1) : cut;
}

/** The stall as a log line: the summary, then the trail relative to its end. */
export function formatInkInputStall(stall: InkInputStall, state: string): { message: string; stack: string } {
  const end = stall.trail.length > 0 ? stall.trail[stall.trail.length - 1].at : 0;
  const trail = stall.trail.map((entry) => `${Math.round(entry.at - end)} ${entry.code}`).join("\n");
  return {
    message: (stall.kind === "trace"
      ? `ink trace (field diagnosis); ${state}`
      : stall.kind === "busy"
      ? `ink input reached the page ${stall.silentMs}ms late (main thread busy); event ${stall.endedBy}; ${state}`
      : stall.kind === "deaf"
        ? `ink input deaf: ${stall.contacts ?? 0} pen contacts unheard while its hover kept arriving ` +
          `(last ${stall.silentMs}ms; browser withholding touches); ${state}`
        : `ink input silent ${stall.silentMs}ms with a pen down; timers late ${stall.lateMs}ms; ` +
          `ended by ${stall.endedBy}; ${state}`
    ).slice(0, 500),
    // The trail's end is the moment the report is about: a long trail keeps
    
    // trace came back cut 3 s before the moment that stalled).
    stack: (stall.diag ? `${stall.diag}\n` : "") + tailWithin(trail, 3800 - (stall.diag ? stall.diag.length + 1 : 0)),
  };
}

export function reportInkInputStall(stall: InkInputStall, state: string): void {
  if (typeof window === "undefined" || isDoNotTrackEnabled()) return;
  const { message, stack } = formatInkInputStall(stall, state);
  sendErrorBeacon({
    name: "InkInputStall",
    message,
    stack,
    path: window.location?.pathname || "/",
    source: "window",
    window_type: "note",
  });
}

/** The recorder of the window, while an ink surface holds the touch shield. */
let activeTrace: InkInputTrace | null = null;

export function setActiveInkInputTrace(trace: InkInputTrace | null): void {
  activeTrace = trace;
}

/**
 * The page's own work that ran `ms`, entered in the trail when it was long
 * enough to matter: a report of input reaching the page late then says what
 * held the main thread (a stroke landing, the float layout pass, the page
 * ledger). Costs nothing while no ink tool is armed.
 */
export function inkTraceWork(code: string, ms: number): void {
  if (activeTrace && ms >= MARK_MIN_MS) activeTrace.work(code, ms);
}

/** A surface's own decision on a pen contact, under the field diagnosis only. */
export function inkTraceNote(code: string): void {
  if (fieldSwitches.diag) activeTrace?.note(code);
}

/*
 * The black box: a page Safari kills while it is frozen (out of memory, or a
 * WebContent hang the watchdog ends) sends nothing, and comes back as a
 * reload; the stall reports above never leave it. While an ink surface is
 * armed, what the page looks like is written to storage after a lift (it
 * lives outside the killed process); a page that goes away properly (hidden,
 * closed, navigated) wipes it. One still there when the next page starts is
 * a page that died, and is reported then (InkPageLost).
 */
const BLACK_BOX_KEY = "mrnote.inkBlackBox";
const BLACK_BOX_EVERY_MS = 1000;
/** An older box is from another sitting, not the reload that followed a death. */
const BLACK_BOX_FRESH_MS = 15 * 60 * 1000;
let blackBoxWrittenAt = Number.NEGATIVE_INFINITY;
const pageStartedAt = typeof Date === "undefined" ? 0 : Date.now();

function storage(): Storage | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

/** What the page looks like now, for the box (counts only, no content). */
function pageShape(): string {
  if (typeof document === "undefined") return "";
  const memory = (performance as unknown as { memory?: { usedJSHeapSize?: number } }).memory;
  const heap = memory?.usedJSHeapSize ? ` heapMB=${Math.round(memory.usedJSHeapSize / 1048576)}` : "";
  return (
    `nodes=${document.getElementsByTagName("*").length} svgs=${document.getElementsByTagName("svg").length} ` +
    `ink=${document.querySelectorAll("[data-ink-tool]").length} canvases=${document.getElementsByTagName("canvas").length} ` +
    `windows=${document.querySelectorAll(".wm-window").length} upMin=${Math.round((Date.now() - pageStartedAt) / 60000)}${heap}`
  );
}

/** Write the box when the last write is a second old. */
export function writeInkBlackBox(trace: InkInputTrace | null, state: string): void {
  const now = performance.now();
  if (now - blackBoxWrittenAt < BLACK_BOX_EVERY_MS) return;
  blackBoxWrittenAt = now;
  try {
    storage()?.setItem(
      BLACK_BOX_KEY,
      JSON.stringify({
        at: Date.now(),
        path: window.location?.pathname || "/",
        shape: pageShape(),
        state,
        trail: trace?.recent(24) ?? "",
      }),
    );
  } catch {
    /* storage full or blocked: the box is best effort */
  }
}

/** The page went away properly: nothing to report next time. */
export function clearInkBlackBox(): void {
  blackBoxWrittenAt = Number.NEGATIVE_INFINITY;
  try {
    storage()?.removeItem(BLACK_BOX_KEY);
  } catch {
    /* ignore */
  }
}

let blackBoxRead = false;

/** Once a page: a box left by a page that died is reported, then wiped. */
export function reportLostInkPage(): void {
  if (blackBoxRead || typeof window === "undefined") return;
  blackBoxRead = true;
  let raw: string | null = null;
  try {
    raw = storage()?.getItem(BLACK_BOX_KEY) ?? null;
    storage()?.removeItem(BLACK_BOX_KEY);
  } catch {
    return;
  }
  if (!raw || isDoNotTrackEnabled()) return;
  try {
    const box = JSON.parse(raw) as { at?: number; path?: string; shape?: string; state?: string; trail?: string };
    const age = Date.now() - (box.at ?? 0);
    if (!(age >= 0 && age < BLACK_BOX_FRESH_MS)) return;
    const navigation = performance.getEntriesByType?.("navigation")[0] as PerformanceNavigationTiming | undefined;
    sendErrorBeacon({
      name: "InkPageLost",
      message: (
        `the page died while ink was armed, ${Math.round(age / 1000)}s before this ${navigation?.type ?? "load"}; ` +
        `${box.shape ?? ""}; ${box.state ?? ""}`
      ).slice(0, 500),
      stack: tailWithin(String(box.trail ?? ""), 3800),
      path: box.path || window.location?.pathname || "/",
      source: "window",
      window_type: "note",
    });
  } catch {
    /* a box from another build: drop it */
  }
}

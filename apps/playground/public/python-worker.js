/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/*
 * Pyodide host worker for in-browser Python execution.
 *
 * Runs on its own thread so a runaway snippet cannot freeze the UI — the main
 * thread stops it by terminating this worker. The interpreter is kept alive
 * across messages so cells in the same notebook share kernel state. Served with
 * its own permissive CSP (see proxy.ts: script-src allows cdn.jsdelivr.net +
 * 'wasm-unsafe-eval', connect-src cdn.jsdelivr.net) so the app origin's CSP does
 * not need those grants.
 *
 * Protocol:
 *   run:     postMessage({ id, code, stream? }) -> postMessage({ id, logs, logStreams?, result?, error?, images?, html? })
 *            (logStreams[i] is "stdout"|"stderr" for logs[i]; absent ⇒ all stdout)
 *            When stream:true, ALSO posts incremental { id, type:"stream", stream:"stdout"|"stderr", chunk }
 *            messages AS the kernel produces output (progressive rendering — Jupyter
 *            streaming parity), before the final aggregate message above. The final
 *            message still carries the complete logs and is the source of truth.
 *   inspect: postMessage({ id, op: "inspect" }) -> postMessage({ id, variables: [{name,type,preview}] })
 *   lint:    postMessage({ id, op: "lint", code }) -> postMessage({ id, lint: { ok, lineno?, offset?, endLineno?, endOffset?, msg? } })
 *            Parse-only `compile(code, "<lint>", "exec")` — never executes the snippet, so it
 *            cannot freeze the kernel; reports the first SyntaxError (1-based lineno/offset) or
 *            { ok: true }. Runs in a throwaway scope so it never touches kernel globals.
 *   hover:   postMessage({ id, op: "hover", code, line, column }) ->
 *            postMessage({ id, hover: { found, name?, kind?, signature?, doc? } })
 *            jedi help()/infer() at the position against the LIVE kernel namespace —
 *            type signature + docstring for the editor's hover tooltip.
 *   signature: postMessage({ id, op: "signature", code, line, column }) ->
 *            postMessage({ id, signature: { found, label?, params?, activeParameter? } })
 *            jedi get_signatures() at the position against the LIVE kernel namespace —
 *            the enclosing call's signature + active argument index for parameter hints.
 *   definition: postMessage({ id, op: "definition", code, line, column }) ->
 *            postMessage({ id, definition: { found, line?, column?, inBuffer? } })
 *            jedi goto() at the position against the LIVE kernel namespace — the
 *            symbol's definition location. `inBuffer` is true when it resolves to
 *            THIS source (navigable in a single-buffer editor); false ⇒ out of file.
 *   references: postMessage({ id, op: "references", code, line, column }) ->
 *            postMessage({ id, references: { found, refs: [{line,column,length}] } })
 *            jedi get_references() at the position against the LIVE kernel namespace,
 *            FILTERED to this buffer — every in-buffer usage of the symbol (for rename).
 *   format:    postMessage({ id, op: "format", code }) -> postMessage({ id, formatted: string, error?: string })
 *            Format Python source via autopep8 (lazily installed via micropip on first use).
 *            Exact `error: "mrnote.format.ok.v1"` is the bounded success receipt;
 *            failures carry the ORIGINAL code with no such receipt.
 *   debug-run: postMessage({ id, op: "debug-run", code, breakpoints, stdin?, stream? }) ->
 *            postMessage({ id, logs, logStreams, result?, error?, images?, html?,
 *                          debugEvents, debugTruncated, debugConditionNote? })
 *            Each breakpoint is { line, condition?, hit? }: `condition` is a
 *            Python expression eval'd in the frame (a raising condition is
 *            skipped — not a hit — and surfaced once via debugConditionNote);
 *            `hit` is a normalized hit-count predicate { op, value, rem? } that
 *            stops only on matching conditional-reach counts.
 *            RECORDED breakpoint debugging via sys.settrace: the code runs to
 *            completion (same stdout/stderr/stdin/auto-package behaviour as a
 *            normal run) while every breakpoint hit — plus a bounded window of
 *            line-steps after each hit — is recorded with locals previews and
 *            the user-frame call stack. The editor replays the recording with
 *            Continue/Step semantics. A worker cannot truly block mid-Python
 *            (no JSPI/SharedArrayBuffer guarantees), so recording is the
 *            correct-everywhere design: breakpoints work INSIDE functions
 *            (which JS instrumentation cannot do) and a paused kernel can
 *            never wedge. Only frames from the user's own source are traced.
 */

(function () {
  const WORKER_PROTOCOL_MARKER = "mrnote.execution-worker.v1";
  const PYTHON_FORMAT_SUCCESS_RECEIPT = "mrnote.format.ok.v1";
  const trustedPostMessage = self.postMessage.bind(self);
  // Authored Python can mutate Worker globals and JavaScript prototypes through
  // Pyodide's `js` FFI. Capture every intrinsic used by a resource/security
  // boundary before the interpreter exists; later checks must not dynamically
  // consult attacker-writable `Array`, `String`, `Number`, `Math`, `Object`, or
  // `JSON` properties.
  const trustedArrayIsArray = Array.isArray;
  const trustedArrayPush = Function.prototype.call.bind(Array.prototype.push);
  const trustedStringSlice = Function.prototype.call.bind(
    String.prototype.slice,
  );
  const trustedStringTrim = Function.prototype.call.bind(String.prototype.trim);
  const trustedStringSplit = Function.prototype.call.bind(String.prototype.split);
  const trustedArrayJoin = Function.prototype.call.bind(Array.prototype.join);
  const trustedRegExpTest = Function.prototype.call.bind(RegExp.prototype.test);
  const trustedNumberIsSafeInteger = Number.isSafeInteger;
  const trustedNumberIsFinite = Number.isFinite;
  const trustedNumber = Number;
  const trustedMathMin = Math.min;
  const trustedMathMax = Math.max;
  const trustedString = String;
  const trustedObjectFromEntries = Object.fromEntries;
  const trustedJsonStringify = JSON.stringify;
  const protocolWords = new Uint32Array(4);
  crypto.getRandomValues(protocolWords);
  const workerProtocolToken = Array.from(protocolWords, (word) =>
    word.toString(16).padStart(8, "0"),
  ).join("");

  // Python can reach WorkerGlobalScope through Pyodide's `js` FFI. Keep the
  // protocol capability in this closure, capture the native sender before any
  // authored code runs, and never expose either value on `self`.
  function postProtocolMessage(payload) {
    const isNonExecution =
      payload &&
      typeof payload === "object" &&
      ("variables" in payload ||
        "lint" in payload ||
        "formatted" in payload ||
        "completions" in payload ||
        "hover" in payload ||
        "signature" in payload ||
        "definition" in payload ||
        "references" in payload);
    const bounded =
      payload && payload.type === "protocol-ready"
        ? { type: "protocol-ready" }
        : isNonExecution
          ? boundNonExecutionResult(payload)
          : payload;
    if (!bounded) return;
    trustedPostMessage([WORKER_PROTOCOL_MARKER, workerProtocolToken, bounded]);
  }

  postProtocolMessage({ type: "protocol-ready" });

  const PYODIDE_VERSION = "0.26.2";
  // MRNote Classic serves this worker from its local Python sandbox host with
  // the Pyodide the app ships beside it, and binds `__mrnoteLocalPyodide`
  // ahead of this script (apps/web/desktop-sandbox.mjs): the runtime, its
  // packages and the lint/format helpers' wheels are then read from that
  // host, and nothing else is fetched. The site's worker has no binding and
  // loads the pinned CDN build. Read once, before any Python exists.
  const LOCAL_RUNTIME = (function () {
    const binding = self.__mrnoteLocalPyodide;
    try {
      delete self.__mrnoteLocalPyodide;
    } catch {
      /* non-configurable: still read once below */
    }
    if (!binding || typeof binding !== "object") return null;
    const sameOrigin = (path) => {
      if (typeof path !== "string" || path[0] !== "/") return null;
      const url = new URL(path, self.location.origin);
      return url.origin === self.location.origin ? url.href : null;
    };
    const indexURL = sameOrigin(binding.indexPath);
    if (!indexURL || indexURL[indexURL.length - 1] !== "/") return null;
    const wheels = Object.create(null);
    const listed =
      binding.wheels && typeof binding.wheels === "object" ? binding.wheels : {};
    for (const name of ["pyflakes", "autopep8"]) {
      const paths = trustedArrayIsArray(listed[name]) ? listed[name] : [];
      const urls = paths.map(sameOrigin);
      if (urls.length > 0 && urls.every(Boolean)) wheels[name] = urls;
    }
    return { indexURL, wheels };
  })();
  const PYODIDE_CDN = LOCAL_RUNTIME
    ? LOCAL_RUNTIME.indexURL
    : "https://cdn.jsdelivr.net/pyodide/v" + PYODIDE_VERSION + "/full/";
  /** The micropip call that installs a lint/format helper (pyflakes, autopep8). */
  function helperInstallSource(name) {
    const local = LOCAL_RUNTIME && LOCAL_RUNTIME.wheels[name];
    // The shipped wheels, dependencies listed alongside: no index is asked.
    if (local) {
      return (
        "import micropip; await micropip.install(" +
        trustedJsonStringify(local) +
        ", deps=False)"
      );
    }
    return "import micropip; await micropip.install('" + name + "')";
  }
  const MAX_OUTPUT_LINES = 2000;
  const MAX_OUTPUT_LINE_CHARS = 20000;
  const MAX_OUTPUT_LOG_CHARS = 300000;
  const MAX_OUTPUT_TEXT_CHARS = 50000;
  const MAX_OUTPUT_HTML_CHARS = 200000;
  const MAX_RICH_IMAGES = 20;
  const MAX_RICH_IMAGE_BASE64_CHARS = 1000000;
  const MAX_RICH_IMAGE_BYTES = Math.floor(
    (MAX_RICH_IMAGE_BASE64_CHARS * 3) / 4,
  );
  const MAX_RICH_FIGURE_EDGE = 4096;
  const MAX_RICH_FIGURE_PIXELS = 4000000;
  const MAX_PROTOCOL_ITEMS = 500;
  const MAX_PROTOCOL_COMPLETIONS = 80;
  const MAX_PROTOCOL_TOTAL_CHARS = 100000;
  const MAX_PROTOCOL_NAME_CHARS = 256;
  const MAX_PROTOCOL_TYPE_CHARS = 128;
  const MAX_PROTOCOL_PREVIEW_CHARS = 120;
  const MAX_PROTOCOL_FIELD_CHARS = 4096;
  const MAX_PROTOCOL_DOC_CHARS = 1500;
  const MAX_PROTOCOL_FORMATTED_CHARS = 2000000;
  const MAX_PROTOCOL_COORDINATE = 10000000;
  const OUTPUT_TRUNCATION_MARKER =
    "… [output truncated: execution budget reached]";

  function protocolString(value, maxChars, budget, allowEmpty) {
    if (typeof value !== "string" || budget.remaining <= 0) return null;
    const bounded = trustedStringSlice(
      value,
      0,
      trustedMathMin(maxChars, budget.remaining),
    );
    if (allowEmpty === false && bounded.length === 0) return null;
    budget.remaining -= bounded.length;
    return bounded;
  }

  function protocolError(data, target, budget) {
    if (typeof data.error !== "string") return;
    const error = protocolString(
      data.error,
      MAX_PROTOCOL_FIELD_CHARS,
      budget,
      true,
    );
    if (error !== null) target.error = error;
  }

  function protocolInteger(value, min, max) {
    return trustedNumberIsSafeInteger(value) && value >= min && value <= max
      ? value
      : null;
  }

  /**
   * Treat every Python-derived IDE reply as hostile even after it is wrapped in
   * the worker-owned capability. The persistent interpreter can be poisoned by
   * authored code; this fixed schema/count/string/aggregate budget prevents that
   * interpreter from using a trusted helper as a large-message confused deputy.
   */
  function boundNonExecutionResult(data) {
    if (!data || typeof data !== "object") return null;
    const id = protocolInteger(data.id, 0, Number.MAX_SAFE_INTEGER);
    if (id === null) return null;
    const budget = { remaining: MAX_PROTOCOL_TOTAL_CHARS };

    if (trustedArrayIsArray(data.variables)) {
      const variables = [];
      for (
        let index = 0;
        index < data.variables.length && index < MAX_PROTOCOL_ITEMS;
        index += 1
      ) {
        const candidate = data.variables[index];
        if (
          !candidate ||
          typeof candidate !== "object" ||
          trustedArrayIsArray(candidate)
        )
          continue;
        const name = protocolString(
          candidate.name,
          MAX_PROTOCOL_NAME_CHARS,
          budget,
          false,
        );
        if (name === null) continue;
        const type = protocolString(
          candidate.type,
          MAX_PROTOCOL_TYPE_CHARS,
          budget,
          true,
        );
        const preview = protocolString(
          candidate.preview,
          MAX_PROTOCOL_PREVIEW_CHARS,
          budget,
          true,
        );
        if (type === null || preview === null) break;
        trustedArrayPush(variables, { name, type, preview });
      }
      const result = { id, variables };
      protocolError(data, result, budget);
      return result;
    }

    if (
      data.lint &&
      typeof data.lint === "object" &&
      !trustedArrayIsArray(data.lint)
    ) {
      const rawLint = data.lint;
      const lint = rawLint.ok === false ? { ok: false } : { ok: true };
      if (lint.ok === false) {
        for (const key of ["lineno", "offset", "endLineno", "endOffset"]) {
          const value = protocolInteger(
            rawLint[key],
            0,
            MAX_PROTOCOL_COORDINATE,
          );
          if (value !== null) lint[key] = value;
        }
        const msg = protocolString(
          rawLint.msg,
          MAX_PROTOCOL_FIELD_CHARS,
          budget,
          true,
        );
        if (msg !== null) lint.msg = msg;
      }
      const warnings = [];
      const rawWarnings = trustedArrayIsArray(data.warnings)
        ? data.warnings
        : [];
      for (
        let index = 0;
        index < rawWarnings.length && index < MAX_PROTOCOL_ITEMS;
        index += 1
      ) {
        const candidate = rawWarnings[index];
        if (
          !candidate ||
          typeof candidate !== "object" ||
          trustedArrayIsArray(candidate)
        )
          continue;
        const row = protocolInteger(candidate.row, 1, MAX_PROTOCOL_COORDINATE);
        const col = protocolInteger(candidate.col, 1, MAX_PROTOCOL_COORDINATE);
        if (row === null || col === null) continue;
        const message = protocolString(
          candidate.message,
          MAX_PROTOCOL_FIELD_CHARS,
          budget,
          false,
        );
        const type = protocolString(
          candidate.type,
          MAX_PROTOCOL_TYPE_CHARS,
          budget,
          false,
        );
        if (message === null || type === null) break;
        trustedArrayPush(warnings, { row, col, message, type });
      }
      const result = { id, lint, warnings };
      protocolError(data, result, budget);
      return result;
    }

    if (typeof data.formatted === "string") {
      const result = {
        id,
        formatted:
          data.formatted.length <= MAX_PROTOCOL_FORMATTED_CHARS
            ? data.formatted
            : "",
      };
      protocolError(data, result, budget);
      return result;
    }

    if (trustedArrayIsArray(data.completions)) {
      const completions = [];
      for (
        let index = 0;
        index < data.completions.length && index < MAX_PROTOCOL_COMPLETIONS;
        index += 1
      ) {
        const candidate = data.completions[index];
        if (
          !candidate ||
          typeof candidate !== "object" ||
          trustedArrayIsArray(candidate)
        )
          continue;
        const name = protocolString(
          candidate.name,
          MAX_PROTOCOL_NAME_CHARS,
          budget,
          false,
        );
        if (name === null) continue;
        const kind =
          candidate.kind === null || candidate.kind === undefined
            ? null
            : protocolString(
                candidate.kind,
                MAX_PROTOCOL_TYPE_CHARS,
                budget,
                true,
              );
        if (candidate.kind != null && kind === null) break;
        trustedArrayPush(completions, { name, kind });
      }
      const result = { id, completions };
      protocolError(data, result, budget);
      return result;
    }

    if (
      data.hover &&
      typeof data.hover === "object" &&
      !trustedArrayIsArray(data.hover)
    ) {
      let hover = { found: false };
      if (data.hover.found === true) {
        const name = protocolString(
          data.hover.name,
          MAX_PROTOCOL_NAME_CHARS,
          budget,
          false,
        );
        if (name !== null) {
          const kind =
            data.hover.kind == null
              ? null
              : protocolString(
                  data.hover.kind,
                  MAX_PROTOCOL_TYPE_CHARS,
                  budget,
                  true,
                );
          const signature =
            data.hover.signature == null
              ? null
              : protocolString(
                  data.hover.signature,
                  MAX_PROTOCOL_FIELD_CHARS,
                  budget,
                  true,
                );
          const doc =
            data.hover.doc == null
              ? null
              : protocolString(
                  data.hover.doc,
                  MAX_PROTOCOL_DOC_CHARS,
                  budget,
                  true,
                );
          if (
            (data.hover.kind == null || kind !== null) &&
            (data.hover.signature == null || signature !== null) &&
            (data.hover.doc == null || doc !== null)
          ) {
            hover = { found: true, name, kind, signature, doc };
          }
        }
      }
      const result = { id, hover };
      protocolError(data, result, budget);
      return result;
    }

    if (
      data.signature &&
      typeof data.signature === "object" &&
      !trustedArrayIsArray(data.signature)
    ) {
      let signature = { found: false };
      if (
        data.signature.found === true &&
        trustedArrayIsArray(data.signature.params)
      ) {
        const label = protocolString(
          data.signature.label,
          MAX_PROTOCOL_FIELD_CHARS,
          budget,
          false,
        );
        if (label !== null) {
          const params = [];
          for (
            let index = 0;
            index < data.signature.params.length && index < MAX_PROTOCOL_ITEMS;
            index += 1
          ) {
            const candidate = data.signature.params[index];
            const param = protocolString(
              candidate,
              MAX_PROTOCOL_NAME_CHARS,
              budget,
              true,
            );
            if (param === null) break;
            trustedArrayPush(params, param);
          }
          const requested = protocolInteger(
            data.signature.activeParameter,
            0,
            MAX_PROTOCOL_ITEMS,
          );
          signature = {
            found: true,
            label,
            params,
            activeParameter: trustedMathMin(
              requested === null ? 0 : requested,
              trustedMathMax(params.length - 1, 0),
            ),
          };
        }
      }
      const result = { id, signature };
      protocolError(data, result, budget);
      return result;
    }

    if (
      data.definition &&
      typeof data.definition === "object" &&
      !trustedArrayIsArray(data.definition)
    ) {
      const line = protocolInteger(
        data.definition.line,
        1,
        MAX_PROTOCOL_COORDINATE,
      );
      const column = protocolInteger(
        data.definition.column,
        0,
        MAX_PROTOCOL_COORDINATE,
      );
      const definition =
        data.definition.found === true && line !== null && column !== null
          ? {
              found: true,
              line,
              column,
              inBuffer: data.definition.inBuffer === true,
            }
          : { found: false };
      const result = { id, definition };
      protocolError(data, result, budget);
      return result;
    }

    if (
      data.references &&
      typeof data.references === "object" &&
      !trustedArrayIsArray(data.references)
    ) {
      const refs = [];
      const rawRefs = trustedArrayIsArray(data.references.refs)
        ? data.references.refs
        : [];
      for (
        let index = 0;
        index < rawRefs.length && index < MAX_PROTOCOL_ITEMS;
        index += 1
      ) {
        const candidate = rawRefs[index];
        if (
          !candidate ||
          typeof candidate !== "object" ||
          trustedArrayIsArray(candidate)
        )
          continue;
        const line = protocolInteger(
          candidate.line,
          1,
          MAX_PROTOCOL_COORDINATE,
        );
        const column = protocolInteger(
          candidate.column,
          0,
          MAX_PROTOCOL_COORDINATE,
        );
        const length = protocolInteger(
          candidate.length,
          1,
          MAX_PROTOCOL_FIELD_CHARS,
        );
        if (line === null || column === null || length === null) continue;
        trustedArrayPush(refs, { line, column, length });
      }
      const result = {
        id,
        references: { found: refs.length > 0, refs },
      };
      protocolError(data, result, budget);
      return result;
    }

    return null;
  }

  /** Convert a PyProxy result to plain structured-clone data, then release it. */
  function pythonResultToJs(value) {
    if (!value || typeof value.toJs !== "function") return value;
    try {
      return value.toJs({
        dict_converter: trustedObjectFromEntries,
        create_pyproxies: false,
      });
    } finally {
      try {
        value.destroy();
      } catch {
        /* already released */
      }
    }
  }

  /**
   * A run's traceback opens with the frames of Pyodide's own runner
   * (`/lib/python3xx.zip/_pyodide/_base.py`, eval_code_async / run_async)
   * before the person's `<exec>` frames: noise that read as the first place
   * the error happened. Drop those frames (each with its source and caret
   * lines); every other frame, the standard library's included, stays.
   */
  const TRACEBACK_FRAME_START = /^  File "/;
  const RUNNER_FRAME_START = /^  File "\/lib\/python\d+\.zip\/_pyodide\//;
  const FRAME_DETAIL_LINE = /^    /;
  function withoutRunnerFrames(value) {
    const lines = trustedStringSplit(trustedString(value), "\n");
    const kept = [];
    let skipping = false;
    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index];
      if (trustedRegExpTest(TRACEBACK_FRAME_START, line)) {
        skipping = trustedRegExpTest(RUNNER_FRAME_START, line);
        if (skipping) continue;
      } else if (skipping && trustedRegExpTest(FRAME_DETAIL_LINE, line)) {
        continue;
      } else {
        skipping = false;
      }
      trustedArrayPush(kept, line);
    }
    return trustedArrayJoin(kept, "\n");
  }

  function truncateOutput(value, maxChars) {
    value = trustedString(value);
    if (value.length <= maxChars) return value;
    if (maxChars <= OUTPUT_TRUNCATION_MARKER.length) {
      return trustedStringSlice(OUTPUT_TRUNCATION_MARKER, 0, maxChars);
    }
    return (
      trustedStringSlice(value, 0, maxChars - OUTPUT_TRUNCATION_MARKER.length) +
      OUTPUT_TRUNCATION_MARKER
    );
  }

  /**
   * Bound stdout/stderr at production time, before it grows a worker array or is
   * cloned over postMessage. The returned aggregate and streamed chunks are kept
   * identical, including the one visible truncation marker.
   */
  function createOutputCapture(id, stream) {
    const logs = [];
    const logStreams = [];
    let chars = 0;
    let sealed = false;
    function marker() {
      if (sealed) return null;
      sealed = true;
      const remaining = MAX_OUTPUT_LOG_CHARS - chars;
      if (remaining <= 0 || logs.length >= MAX_OUTPUT_LINES) return null;
      const kept = trustedStringSlice(OUTPUT_TRUNCATION_MARKER, 0, remaining);
      trustedArrayPush(logs, kept);
      trustedArrayPush(logStreams, "stderr");
      chars += kept.length;
      if (stream) {
        postProtocolMessage({
          id,
          type: "stream",
          stream: "stderr",
          chunk: kept,
        });
      }
      return kept;
    }
    return {
      logs,
      logStreams,
      push(text, streamName) {
        if (sealed) return null;
        if (logs.length >= MAX_OUTPUT_LINES - 1) return marker();
        let kept;
        try {
          kept = truncateOutput(text, MAX_OUTPUT_LINE_CHARS);
        } catch {
          kept = "<unprintable>";
        }
        const remaining =
          MAX_OUTPUT_LOG_CHARS - chars - OUTPUT_TRUNCATION_MARKER.length;
        if (kept.length > remaining) {
          if (remaining <= 0) return marker();
          kept = truncateOutput(kept, remaining);
          sealed = true;
        }
        const tag = streamName === "stderr" ? "stderr" : "stdout";
        trustedArrayPush(logs, kept);
        trustedArrayPush(logStreams, tag);
        chars += kept.length;
        if (stream)
          postProtocolMessage({ id, type: "stream", stream: tag, chunk: kept });
        return kept;
      },
    };
  }

  function boundRichImages(value) {
    if (!trustedArrayIsArray(value)) return undefined;
    const images = [];
    let chars = 0;
    for (let index = 0; index < value.length; index += 1) {
      const image = value[index];
      if (images.length >= MAX_RICH_IMAGES) break;
      if (
        typeof image !== "string" ||
        image.length > MAX_RICH_IMAGE_BASE64_CHARS ||
        chars + image.length > MAX_RICH_IMAGE_BASE64_CHARS
      ) {
        continue;
      }
      trustedArrayPush(images, image);
      chars += image.length;
    }
    return images.length > 0 ? images : undefined;
  }

  // After running a cell, collect Jupyter-style rich outputs: any open matplotlib
  // figures rendered to PNG (headless Agg backend) and returned as base64. Runs in
  // the kernel namespace; failures are swallowed so they never mask the cell run.
  const COLLECT_RICH = [
    "import sys as _sys, io as _io, base64 as _b64, math as _math",
    '__rich = {"images": []}',
    `_MAX_IMAGES = ${MAX_RICH_IMAGES}`,
    `_MAX_IMAGE_BYTES = ${MAX_RICH_IMAGE_BYTES}`,
    `_MAX_FIGURE_EDGE = ${MAX_RICH_FIGURE_EDGE}`,
    `_MAX_FIGURE_PIXELS = ${MAX_RICH_FIGURE_PIXELS}`,
    "_image_bytes = 0",
    "if 'matplotlib' in _sys.modules:",
    "    try:",
    "        import matplotlib.pyplot as _plt",
    "        for _n in _plt.get_fignums()[:_MAX_IMAGES]:",
    "            _fig = _plt.figure(_n)",
    "            _size = _fig.get_size_inches()",
    "            _dpi = float(_fig.dpi)",
    "            _width = float(_size[0]) * _dpi",
    "            _height = float(_size[1]) * _dpi",
    "            if not (_math.isfinite(_width) and _math.isfinite(_height)):",
    "                continue",
    "            if _width <= 0 or _height <= 0:",
    "                continue",
    "            if _width > _MAX_FIGURE_EDGE or _height > _MAX_FIGURE_EDGE:",
    "                continue",
    "            if _width * _height > _MAX_FIGURE_PIXELS:",
    "                continue",
    "            _buf = _io.BytesIO()",
    "            _fig.savefig(_buf, format='png')",
    "            _payload = _buf.getvalue()",
    "            if len(_payload) > _MAX_IMAGE_BYTES:",
    "                continue",
    "            if _image_bytes + len(_payload) > _MAX_IMAGE_BYTES:",
    "                break",
    "            __rich['images'].append(_b64.b64encode(_payload).decode())",
    "            _image_bytes += len(_payload)",
    "        _plt.close('all')",
    "    except Exception:",
    "        pass",
    "__rich",
  ].join("\n");

  // Build the inspector once, before authored code runs, and keep the callable only
  // in this JavaScript closure. Its closure captures pristine builtins/type methods;
  // previews never invoke arbitrary __repr__/__str__, iterate user containers, or
  // import a user-mutable JSON serializer.
  const SAFE_INSPECT_FACTORY = [
    "def _mrai_make_safe_inspector():",
    "    import builtins as _b, types as _types",
    "    _safe_type = type",
    "    _safe_len = len",
    "    _safe_dict_items = dict.items",
    "    _safe_dict_get = dict.get",
    "    _safe_type_name = type.__getattribute__",
    "    _safe_str_repr = str.__repr__",
    "    _safe_int_repr = int.__repr__",
    "    _safe_float_repr = float.__repr__",
    "    _builtin_values = dict(vars(_b))",
    "    _skip_types = (_types.ModuleType, _types.FunctionType, _types.BuiltinFunctionType)",
    "    def _type_name(_value):",
    "        try:",
    "            _name = _safe_type_name(_safe_type(_value), '__name__')",
    "            return _name[:128] if _safe_type(_name) is str else '?'",
    "        except BaseException:",
    "            return '?'",
    "    def _preview(_value, _type_value):",
    "        try:",
    "            if _type_value is str:",
    "                _text = _safe_str_repr(_value)",
    "            elif _type_value is bool:",
    "                _text = 'True' if _value else 'False'",
    "            elif _type_value is int:",
    "                _text = _safe_int_repr(_value)",
    "            elif _type_value is float:",
    "                _text = _safe_float_repr(_value)",
    "            elif _value is None:",
    "                _text = 'None'",
    "            else:",
    "                return '<' + _type_name(_value) + '>'",
    "            return _text if _safe_len(_text) <= 120 else _text[:119] + '\\u2026'",
    "        except BaseException:",
    "            return '<' + _type_name(_value) + '>'",
    "    def _inspect(_user_globals):",
    "        _out = []",
    "        for _name, _value in _safe_dict_items(_user_globals):",
    "            if _safe_type(_name) is not str or not _name or _name[:1] == '_':",
    "                continue",
    "            _value_type = _safe_type(_value)",
    "            if _value_type in _skip_types:",
    "                continue",
    "            if _safe_dict_get(_builtin_values, _name, None) is _value:",
    "                continue",
    "            _out.append({'name': _name[:256], 'type': _type_name(_value), 'preview': _preview(_value, _value_type)})",
    "            if _safe_len(_out) >= 500:",
    "                break",
    "        return _out",
    "    return _inspect",
    "_mrai_make_safe_inspector()",
  ].join("\n");

  // Parse-only syntax check. `compile(..., "exec")` builds the AST/bytecode but
  // NEVER runs it, so timers / loops / network in the snippet never fire — this is
  // a pure syntax gate. Reports the first SyntaxError (1-based lineno/offset, plus
  // 3.10+ end positions) as a native record, or {"ok": true}. The user source is passed in a
  // dedicated builtin (`__mrnote_lint_src`) and the check runs in a throwaway dict
  // so it can never read or clobber kernel globals.
  const LINT_CHECK = [
    "def _mrnote_lint(_src):",
    "    try:",
    "        compile(_src, '<lint>', 'exec')",
    "        return {'ok': True}",
    "    except SyntaxError as _e:",
    "        return {",
    "            'ok': False,",
    "            'lineno': getattr(_e, 'lineno', None),",
    "            'offset': getattr(_e, 'offset', None),",
    "            'endLineno': getattr(_e, 'end_lineno', None),",
    "            'endOffset': getattr(_e, 'end_offset', None),",
    "            'msg': getattr(_e, 'msg', None) or str(_e),",
    "        }",
    "    except Exception:",
    "        return {'ok': True}",
    // The source is injected via pyodide's `globals` map (set from JS as data, not
    // code), so arbitrary text — including quotes/newlines/unicode — round-trips
    // safely without string-escaping pitfalls. Read it, then drop it.
    "_mrnote_lint_result = _mrnote_lint(globals().get('__mrnote_lint_src', ''))",
    "globals().pop('__mrnote_lint_src', None)",
    "_mrnote_lint_result",
  ].join("\n");

  // jedi-powered semantic completion at a (1-based line, 0-based column). The
  // source is injected as DATA via globals (same pattern as LINT_CHECK) so any
  // text round-trips; jedi sees the LIVE kernel namespace via Interpreter, so
  // `df.<cursor>` completes pandas members after the user ran `import pandas`.
  const COMPLETE_CHECK = [
    "def _mrnote_complete(_src, _line, _col):",
    "    try:",
    "        import jedi",
    "    except Exception:",
    "        return []",
    "    try:",
    "        _script = jedi.Interpreter(_src, [globals()])",
    "        _comps = _script.complete(int(_line), int(_col))",
    "        _out = []",
    "        for _c in _comps[:80]:",
    "            try:",
    "                _out.append({'name': _c.name, 'kind': _c.type or None})",
    "            except Exception:",
    "                continue",
    "        return _out",
    "    except Exception:",
    "        return []",
    "_mrnote_complete_result = _mrnote_complete(",
    "    globals().get('__mrnote_complete_src', ''),",
    "    globals().get('__mrnote_complete_line', 1),",
    "    globals().get('__mrnote_complete_col', 0),",
    ")",
    "globals().pop('__mrnote_complete_src', None)",
    "globals().pop('__mrnote_complete_line', None)",
    "globals().pop('__mrnote_complete_col', None)",
    "_mrnote_complete_result",
  ].join("\n");

  // jedi hover info at a (1-based line, 0-based column): help() resolves names,
  // keywords and call targets; infer() is the fallback for plain expressions.
  // Same data-injection pattern as COMPLETE_CHECK; returns native None when there
  // is nothing under the cursor so the editor shows no tooltip.
  const HOVER_CHECK = [
    "def _mrnote_hover(_src, _line, _col):",
    "    try:",
    "        import jedi",
    "    except Exception:",
    "        return None",
    "    try:",
    "        _script = jedi.Interpreter(_src, [globals()])",
    "        try:",
    "            _names = _script.help(int(_line), int(_col))",
    "        except Exception:",
    "            _names = []",
    "        if not _names:",
    "            try:",
    "                _names = _script.infer(int(_line), int(_col))",
    "            except Exception:",
    "                _names = []",
    "        for _n in _names[:1]:",
    "            _sig = None",
    "            try:",
    "                _sigs = _n.get_signatures()",
    "                if _sigs:",
    "                    _sig = _sigs[0].to_string()",
    "            except Exception:",
    "                pass",
    "            if not _sig:",
    "                try:",
    "                    _sig = _n.description or None",
    "                except Exception:",
    "                    pass",
    "            _doc = None",
    "            try:",
    "                _raw = _n.docstring(raw=True)",
    "                if _raw:",
    "                    _doc = _raw.strip()[:1500] or None",
    "            except Exception:",
    "                pass",
    "            try:",
    "                _kind = _n.type or None",
    "            except Exception:",
    "                _kind = None",
    "            try:",
    "                _name = _n.name or ''",
    "            except Exception:",
    "                _name = ''",
    "            return {'name': _name, 'kind': _kind, 'signature': _sig, 'doc': _doc}",
    "        return None",
    "    except Exception:",
    "        return None",
    "_mrnote_hover_result = _mrnote_hover(",
    "    globals().get('__mrnote_hover_src', ''),",
    "    globals().get('__mrnote_hover_line', 1),",
    "    globals().get('__mrnote_hover_col', 0),",
    ")",
    "globals().pop('__mrnote_hover_src', None)",
    "globals().pop('__mrnote_hover_line', None)",
    "globals().pop('__mrnote_hover_col', None)",
    "_mrnote_hover_result",
  ].join("\n");

  // jedi parameter hints at a (1-based line, 0-based column): get_signatures()
  // resolves the enclosing call's signature(s). The first signature's to_string()
  // is the label (e.g. `foo(a, b, *, c=1)`), its .params give the parameter names,
  // and .index is the active argument slot at that position (None outside any
  // param). Same data-injection pattern as HOVER_CHECK; returns {found:false} when
  // the cursor is not inside a call so the editor shows no hint.
  const SIGNATURE_CHECK = [
    "def _mrnote_signature(_src, _line, _col):",
    "    try:",
    "        import jedi",
    "    except Exception:",
    "        return {'found': False}",
    "    try:",
    "        _script = jedi.Interpreter(_src, [globals()])",
    "        try:",
    "            _sigs = _script.get_signatures(int(_line), int(_col))",
    "        except Exception:",
    "            _sigs = []",
    "        for _s in _sigs[:1]:",
    "            try:",
    "                _label = _s.to_string()",
    "            except Exception:",
    "                _label = None",
    "            if not _label:",
    "                return {'found': False}",
    "            _params = []",
    "            try:",
    "                for _p in _s.params:",
    "                    try:",
    "                        _params.append(_p.name)",
    "                    except Exception:",
    "                        _params.append('')",
    "            except Exception:",
    "                _params = []",
    "            _active = None",
    "            try:",
    "                _active = _s.index",
    "            except Exception:",
    "                _active = None",
    "            if not isinstance(_active, int):",
    "                _active = 0",
    "            return {'found': True, 'label': _label, 'params': _params, 'activeParameter': _active}",
    "        return {'found': False}",
    "    except Exception:",
    "        return {'found': False}",
    "_mrnote_signature_result = _mrnote_signature(",
    "    globals().get('__mrnote_signature_src', ''),",
    "    globals().get('__mrnote_signature_line', 1),",
    "    globals().get('__mrnote_signature_col', 0),",
    ")",
    "globals().pop('__mrnote_signature_src', None)",
    "globals().pop('__mrnote_signature_line', None)",
    "globals().pop('__mrnote_signature_col', None)",
    "_mrnote_signature_result",
  ].join("\n");

  // jedi go-to-definition at a (1-based line, 0-based column): goto() resolves the
  // name under the cursor to its definition. Each resolved Name has .line/.column
  // (the definition's position) and .module_path (the file it lives in). For a
  // single-buffer editor only WITHIN-buffer definitions are navigable, so the
  // result carries `inBuffer`: True when jedi resolved the def to THIS source
  // (module_path is None ⇒ the in-memory buffer, the common local-symbol case) so
  // the editor can jump there; False (stdlib / another module) ⇒ the editor shows
  // the same "not found / out of file" notice the TS path uses. Same data-
  // injection pattern as SIGNATURE_CHECK; always answers {found:false} on failure.
  const DEFINITION_CHECK = [
    "def _mrnote_definition(_src, _line, _col):",
    "    try:",
    "        import jedi",
    "    except Exception:",
    "        return {'found': False}",
    "    try:",
    "        _script = jedi.Interpreter(_src, [globals()])",
    "        try:",
    "            _names = _script.goto(int(_line), int(_col), follow_imports=True)",
    "        except Exception:",
    "            _names = []",
    "        for _n in _names[:1]:",
    "            try:",
    "                _dline = _n.line",
    "            except Exception:",
    "                _dline = None",
    "            try:",
    "                _dcol = _n.column",
    "            except Exception:",
    "                _dcol = None",
    "            if not isinstance(_dline, int) or not isinstance(_dcol, int):",
    "                return {'found': False}",
    // module_path is None for the in-memory buffer we passed to Interpreter; any
    // real path (stdlib / another module) means the definition is out of buffer
    // and not navigable in this single-buffer editor.
    "            try:",
    "                _path = _n.module_path",
    "            except Exception:",
    "                _path = None",
    "            _in_buffer = _path is None",
    "            return {'found': True, 'line': _dline, 'column': _dcol, 'inBuffer': _in_buffer}",
    "        return {'found': False}",
    "    except Exception:",
    "        return {'found': False}",
    "_mrnote_definition_result = _mrnote_definition(",
    "    globals().get('__mrnote_definition_src', ''),",
    "    globals().get('__mrnote_definition_line', 1),",
    "    globals().get('__mrnote_definition_col', 0),",
    ")",
    "globals().pop('__mrnote_definition_src', None)",
    "globals().pop('__mrnote_definition_line', None)",
    "globals().pop('__mrnote_definition_col', None)",
    "_mrnote_definition_result",
  ].join("\n");

  // jedi find-references at a (1-based line, 0-based column): get_references()
  // resolves every usage of the symbol under the cursor (the declaration plus all
  // reads/writes). Uses jedi.Script (NOT Interpreter): references is a buffer-local
  // syntactic query, and Interpreter.get_references() is unsupported (it raises on
  // the synthetic interpreter module). For a single-buffer rename only references
  // WITHIN this buffer are actionable, so we FILTER to refs whose .module_path is
  // None (the in-memory source) and emit each as {line, column, length} — length is
  // the identifier's text width (name end column minus start column) so the editor
  // can build the per-occurrence edit spans. Same data-injection pattern as
  // DEFINITION_CHECK; always answers {found:false, refs:[]} on failure.
  const REFERENCES_CHECK = [
    "def _mrnote_references(_src, _line, _col):",
    "    try:",
    "        import jedi",
    "    except Exception:",
    "        return {'found': False, 'refs': []}",
    "    try:",
    "        _script = jedi.Script(_src)",
    "        try:",
    "            _refs = _script.get_references(int(_line), int(_col), include_builtins=False)",
    "        except Exception:",
    "            _refs = []",
    "        _out = []",
    "        for _r in _refs:",
    "            try:",
    "                if _r.module_path is not None:",
    "                    continue",
    "            except Exception:",
    "                continue",
    "            try:",
    "                _rline = _r.line",
    "                _rcol = _r.column",
    "            except Exception:",
    "                continue",
    "            if not isinstance(_rline, int) or not isinstance(_rcol, int):",
    "                continue",
    // Prefer the precise name span (end - start); fall back to len(name).
    "            _len = None",
    "            try:",
    "                _end = _r.get_definition_end_position()",
    "            except Exception:",
    "                _end = None",
    "            try:",
    "                if _end is not None and _end[0] == _rline:",
    "                    _len = _end[1] - _rcol",
    "            except Exception:",
    "                _len = None",
    "            if not isinstance(_len, int) or _len <= 0:",
    "                try:",
    "                    _len = len(_r.name or '')",
    "                except Exception:",
    "                    _len = 0",
    "            if _len <= 0:",
    "                continue",
    "            _out.append({'line': _rline, 'column': _rcol, 'length': _len})",
    "        return {'found': len(_out) > 0, 'refs': _out}",
    "    except Exception:",
    "        return {'found': False, 'refs': []}",
    "_mrnote_references_result = _mrnote_references(",
    "    globals().get('__mrnote_references_src', ''),",
    "    globals().get('__mrnote_references_line', 1),",
    "    globals().get('__mrnote_references_col', 0),",
    ")",
    "globals().pop('__mrnote_references_src', None)",
    "globals().pop('__mrnote_references_line', None)",
    "globals().pop('__mrnote_references_col', None)",
    "_mrnote_references_result",
  ].join("\n");

  // Recorded breakpoint debugging. Runs the user source under sys.settrace in
  // the KERNEL globals (state parity with normal runs), recording each breakpoint
  // hit + a bounded step window after it: 1-based line, capped locals previews,
  // and the user-frame stack. Library/internal frames are never traced (the
  // callback returns None for any frame from another file — keeps pandas-sized
  // imports at full speed). Conditions are eval'd in the frame scope; a raising
  // condition records the hit with `conditionError` (VSCode semantics). The user
  // exception (if any) is re-formatted with driver frames stripped so the
  // renderer sees a normal-looking traceback. Source + breakpoints are injected
  // as DATA via globals (same pattern as the other ops).
  const DEBUG_RUN = [
    "import json as _json, sys as _sys, types as _types, traceback as _traceback, ast as _ast, inspect as _inspect",
    "async def _mrnote_debug_run(_src, _bp_json, _pause_uncaught):",
    "    _FILENAME = '<mrnote-debug>'",
    "    _MAX_EVENTS = 300",
    "    _STEP_WINDOW = 120",
    "    _safe_type = type",
    "    _safe_len = len",
    "    _safe_type_name = type.__getattribute__",
    "    _safe_str_repr = str.__repr__",
    "    _safe_int_repr = int.__repr__",
    "    _safe_float_repr = float.__repr__",
    "    _safe_skip_types = (_types.ModuleType, _types.FunctionType, _types.BuiltinFunctionType)",
    "    _bps = {}",
    "    try:",
    "        for _b in _json.loads(_bp_json):",
    "            try:",
    "                _line = int(_b.get('line'))",
    "            except Exception:",
    "                continue",
    "            _cond = _b.get('condition')",
    "            _cond = _cond if isinstance(_cond, str) and _cond.strip() else None",
    "            _hit = _b.get('hit')",
    "            _hit = _hit if isinstance(_hit, dict) else None",
    "            _bps[_line] = (_cond, _hit)",
    "    except Exception:",
    "        _bps = {}",
    "    _events = []",
    "    _counts = {}",
    "    _state = {'steps_left': 0, 'truncated': False, 'note': None}",
    "    def _hit_ok(_spec, _n):",
    "        if not _spec:",
    "            return True",
    "        _op = _spec.get('op')",
    "        try:",
    "            _val = int(_spec.get('value'))",
    "        except Exception:",
    "            return True",
    "        if _op == '>=':",
    "            return _n >= _val",
    "        if _op == '>':",
    "            return _n > _val",
    "        if _op == '<=':",
    "            return _n <= _val",
    "        if _op == '<':",
    "            return _n < _val",
    "        if _op == '==':",
    "            return _n == _val",
    "        if _op == '%':",
    "            if _val == 0:",
    "                return True",
    "            try:",
    "                _rem = int(_spec.get('rem') or 0)",
    "            except Exception:",
    "                _rem = 0",
    "            return (_n % _val) == _rem",
    "        return True",
    "    def _type_name(_v):",
    "        try:",
    "            _name = _safe_type_name(_safe_type(_v), '__name__')",
    "            return _name[:128] if _safe_type(_name) is str else '?'",
    "        except BaseException:",
    "            return '?'",
    "    def _preview(_v):",
    "        try:",
    "            _kind = _safe_type(_v)",
    "            if _kind is str:",
    "                _s = _safe_str_repr( _v )",
    "            elif _kind is bool:",
    "                _s = 'True' if _v else 'False'",
    "            elif _kind is int:",
    "                _s = _safe_int_repr( _v )",
    "            elif _kind is float:",
    "                _s = _safe_float_repr( _v )",
    "            elif _v is None:",
    "                _s = 'None'",
    "            else:",
    "                return '<' + _type_name(_v) + '>'",
    "            return _s if _safe_len(_s) <= 120 else _s[:119] + '\\u2026'",
    "        except BaseException:",
    "            return '<' + _type_name(_v) + '>'",
    "    def _exception_preview(_exc):",
    "        try:",
    "            _text = ''.join(_traceback.format_exception_only(_safe_type(_exc), _exc)).strip()",
    "            if _safe_type(_text) is not str or not _text:",
    "                return _type_name(_exc)",
    "            return _text if _safe_len(_text) <= 500 else _text[:499] + '\u2026'",
    "        except BaseException:",
    "            return _type_name(_exc)",
    "    def _locals_of(_frame):",
    "        _out = {}",
    "        for _name, _val in list(_frame.f_locals.items()):",
    "            if _name.startswith('_'):",
    "                continue",
    "            if _safe_type(_val) in _safe_skip_types:",
    "                continue",
    "            _out[_name] = _preview(_val)",
    "            if len(_out) >= 50:",
    "                break",
    "        return _out",
    "    def _stack_of(_frame):",
    "        _out = []",
    "        _f = _frame",
    "        while _f is not None and len(_out) < 20:",
    "            if _f.f_code.co_filename != _FILENAME:",
    "                break",
    "            _out.append({'name': _f.f_code.co_name, 'line': _f.f_lineno})",
    "            _f = _f.f_back",
    "        return _out",
    "    def _depth_of(_frame):",
    "        _n = 0",
    "        _f = _frame",
    "        while _f is not None:",
    "            if _f.f_code.co_filename != _FILENAME:",
    "                break",
    "            _n += 1",
    "            _f = _f.f_back",
    "        return _n",
    "    def _record(_frame, _kind, _exception=None):",
    "        if len(_events) >= _MAX_EVENTS:",
    "            _state['truncated'] = True",
    "            return False",
    "        _ev = {'kind': _kind, 'line': _frame.f_lineno,",
    "               'locals': _locals_of(_frame), 'stack': _stack_of(_frame),",
    "               'depth': _depth_of(_frame)}",
    "        if _exception is not None:",
    "            _ev['exception'] = _exception",
    "        _events.append(_ev)",
    "        return True",
    "    def _trace(_frame, _event, _arg):",
    "        if _frame.f_code.co_filename != _FILENAME:",
    "            return None",
    "        if _event != 'line':",
    "            return _trace",
    "        if _state['truncated']:",
    "            return _trace",
    "        _line = _frame.f_lineno",
    "        if _line in _bps:",
    "            _cond, _hitspec = _bps[_line]",
    "            if _cond is None:",
    "                _ok = True",
    "            else:",
    "                try:",
    "                    _ok = bool(eval(_cond, _frame.f_globals, _frame.f_locals))",
    "                except Exception as _e:",
    "                    _ok = False",
    "                    if _state['note'] is None:",
    "                        _state['note'] = _preview(_e)",
    "            if _ok:",
    "                _counts[_line] = _counts.get(_line, 0) + 1",
    "                if _hit_ok(_hitspec, _counts[_line]):",
    "                    if _record(_frame, 'breakpoint'):",
    "                        _state['steps_left'] = _STEP_WINDOW",
    "                return _trace",
    "        if _state['steps_left'] > 0:",
    "            if _record(_frame, 'step'):",
    "                _state['steps_left'] -= 1",
    "        return _trace",
    "    _g = globals()",
    "    _g.pop('__mrnote_debug_selected_test_awaitable__', None)",
    "    _result_repr = None",
    "    _html = None",
    "    _error = None",
    "    try:",
    "        _tree = _ast.parse(_src, _FILENAME, 'exec')",
    "        _last_code = None",
    "        if _tree.body and isinstance(_tree.body[-1], _ast.Expr):",
    "            _last_expr = _ast.Expression(_tree.body[-1].value)",
    "            _ast.copy_location(_last_expr.body, _tree.body[-1])",
    "            _tree.body = _tree.body[:-1]",
    "            _last_code = compile(_ast.fix_missing_locations(_last_expr), _FILENAME, 'eval')",
    "        _code = compile(_tree, _FILENAME, 'exec')",
    "        _sys.settrace(_trace)",
    "        try:",
    "            exec(_code, _g)",
    "            _v = eval(_last_code, _g) if _last_code is not None else None",
    "            _test_awaitable = _g.pop('__mrnote_debug_selected_test_awaitable__', None)",
    "            if _test_awaitable is not None:",
    "                if not _inspect.isawaitable(_test_awaitable):",
    "                    raise TypeError('Async test did not return an awaitable')",
    "                _v = await _test_awaitable",
    "        finally:",
    "            _sys.settrace(None)",
    "            _g.pop('__mrnote_debug_selected_test_awaitable__', None)",
    "        if _v is not None:",
    "            try:",
    "                _r = getattr(_v, '_repr_html_', None)",
    "                if callable(_r):",
    "                    _h = _r()",
    "                    if isinstance(_h, str) and _h:",
    "                        _html = _h[:524288]",
    "            except Exception:",
    "                pass",
    "            if _html is None:",
    "                try:",
    "                    _result_repr = str(_v)[:10000]",
    "                except Exception:",
    "                    _result_repr = '<unrepr>'",
    "    except BaseException as _exc:",
    "        _tb = _exc.__traceback__",
    "        _user_tb = None",
    "        _exception_frame = None",
    "        while _tb is not None:",
    "            if _tb.tb_frame.f_code.co_filename == _FILENAME:",
    "                if _user_tb is None:",
    "                    _user_tb = _tb",
    "                _exception_frame = _tb.tb_frame",
    "            _tb = _tb.tb_next",
    "        if _pause_uncaught and _exception_frame is not None:",
    "            if len(_events) >= _MAX_EVENTS:",
    "                _events.pop()",
    "                _state['truncated'] = True",
    "            _record(_exception_frame, 'exception', _exception_preview(_exc))",
    "        try:",
    "            _error = ''.join(_traceback.format_exception(type(_exc), _exc, _user_tb))",
    "        except Exception:",
    "            _error = repr(_exc)",
    "    return {'events': _events, 'truncated': _state['truncated'],",
    "            'conditionNote': _state['note'],",
    "            'result': _result_repr, 'html': _html, 'error': _error}",
    "_mrnote_debug_result = await _mrnote_debug_run(",
    "    globals().get('__mrnote_debug_src', ''),",
    "    globals().get('__mrnote_debug_bps', '[]'),",
    "    globals().get('__mrnote_debug_pause_uncaught', False) is True,",
    ")",
    "globals().pop('__mrnote_debug_src', None)",
    "globals().pop('__mrnote_debug_bps', None)",
    "globals().pop('__mrnote_debug_pause_uncaught', None)",
    "_mrnote_debug_result",
  ].join("\n");

  // pyflakes static-analysis snippet: runs over the source WITHOUT executing it
  // (AST-only pass) and reports undefined-name, unused-import, and unused/shadowed-
  // variable warnings. Each warning is { row, col, message, type } where row and col
  // are 1-based. pyflakes is a pure-Python stdlib-only package (no C extensions), so
  // it loads cleanly in Pyodide via micropip. The source is injected as DATA via
  // globals (same data-injection pattern as LINT_CHECK) so any text round-trips safely.
  const PYFLAKES_CHECK = [
    "def _mrnote_pyflakes(_src):",
    "    try:",
    "        import pyflakes.api as _pf_api",
    "        import pyflakes.reporter as _pf_rep",
    "        import io as _io",
    "    except Exception:",
    "        return []",
    "    try:",
    "        _warnings = []",
    "        class _Collector(_pf_rep.Reporter):",
    "            def unexpectedError(self, filename, msg):",
    "                pass",
    "            def syntaxError(self, filename, msg, lineno, offset, text):",
    "                pass",
    "            def flake(self, message):",
    "                _warnings.append({",
    "                    'row': message.lineno,",
    "                    'col': message.col + 1,",
    "                    'message': str(message),",
    "                    'type': type(message).__name__,",
    "                })",
    "        _pf_api.check(_src, '<pyflakes>', reporter=_Collector())",
    "        return _warnings",
    "    except Exception:",
    "        return []",
    "_mrnote_pyflakes_result = _mrnote_pyflakes(globals().get('__mrnote_pyflakes_src', ''))",
    "globals().pop('__mrnote_pyflakes_src', None)",
    "_mrnote_pyflakes_result",
  ].join("\n");

  // autopep8 formatter snippet: formats Python source with autopep8 (PEP 8 style).
  // autopep8 is a pure-Python package; loaded once via micropip. The source is
  // injected as DATA via globals so any text (including quotes/newlines) round-trips
  // safely. On success returns the formatted source; on any failure returns the
  // ORIGINAL text unchanged so the buffer is never corrupted.
  const AUTOPEP8_FORMAT = [
    "def _mrnote_autopep8(_src):",
    "    try:",
    "        import autopep8 as _ap8",
    "    except Exception:",
    "        return {'ok': False, 'formatted': _src}",
    "    try:",
    "        _out = _ap8.fix_code(_src, options={'max_line_length': 88})",
    "        return {'ok': True, 'formatted': _out}",
    "    except Exception:",
    "        return {'ok': False, 'formatted': _src}",
    "_mrnote_ap8_result = _mrnote_autopep8(globals().get('__mrnote_ap8_src', ''))",
    "globals().pop('__mrnote_ap8_src', None)",
    "_mrnote_ap8_result",
  ].join("\n");

  let pyodideReady = null;
  // Created before the first authored cell and retained only in this JS closure.
  let safeInspect = null;
  // jedi loads once per kernel, lazily on the first completion request; a load
  // failure (offline CDN) caches as null so we never re-attempt per keystroke.
  let jediReady = null;
  // pyflakes loads once per kernel, lazily on the first lint request that needs
  // semantic warnings; a load failure caches as false so we never re-attempt.
  let pyflakesReady = null;
  // autopep8 loads once per kernel, lazily on the first format request; a load
  // failure caches as false so we never re-attempt per keystroke.
  let autopep8Ready = null;

  async function getPyodide() {
    if (!pyodideReady) {
      if (typeof self.loadPyodide !== "function") {
        importScripts(PYODIDE_CDN + "pyodide.js");
      }
      pyodideReady = self
        .loadPyodide({ indexURL: PYODIDE_CDN })
        .then(async (py) => {
          // Force a headless matplotlib backend before any pyplot import so figures
          // can be saved to PNG without a DOM canvas (there is none in a worker).
          try {
            await py.runPythonAsync(
              "import os; os.environ.setdefault('MPLBACKEND', 'AGG')",
            );
          } catch {
            /* non-fatal */
          }
          safeInspect = await py.runPythonAsync(SAFE_INSPECT_FACTORY);
          return py;
        })
        .catch((error) => {
          // A CDN/runtime bootstrap failure is not permanent authority. Keep
          // this worker alive so the next explicit request can retry without a
          // page reload; successful kernels remain memoized as before.
          pyodideReady = null;
          throw error;
        });
    }
    return pyodideReady;
  }

  self.addEventListener("message", async (event) => {
    // Authored Python can synthesize MessageEvents via `js.self.dispatchEvent`.
    // Only browser-delivered requests from the worker owner are authoritative.
    if (event.isTrusted !== true) return;
    const data = event.data || {};
    const id = data.id;
    // Inspect op: snapshot user-defined globals without running a cell.
    if (data.op === "inspect") {
      try {
        const py = await getPyodide();
        let variables = [];
        try {
          const parsed = pythonResultToJs(safeInspect(py.globals));
          if (trustedArrayIsArray(parsed)) variables = parsed;
        } catch {
          /* non-fatal: report an empty snapshot */
        }
        postProtocolMessage({ id, variables });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({ id, variables: [], error: message });
      }
      return;
    }
    // Lint op: parse-check the source WITHOUT executing it (anti-freeze) and report
    // the first SyntaxError, plus pyflakes semantic warnings (undefined names,
    // unused imports, unused/shadowed variables). Both checks run in throwaway scopes
    // so they never touch or corrupt the live kernel namespace.
    if (data.op === "lint") {
      const src = typeof data.code === "string" ? data.code : "";
      try {
        const py = await getPyodide();
        // Step 1: syntax check (compile() only; never executes).
        // Inject the source as DATA via the globals map so any text round-trips
        // without escaping pitfalls; the snippet pops the key afterwards.
        py.globals.set("__mrnote_lint_src", src);
        let lint = { ok: true };
        try {
          const out = await py.runPythonAsync(LINT_CHECK);
          const parsed = pythonResultToJs(out);
          if (parsed && typeof parsed === "object") lint = parsed;
        } catch {
          /* non-fatal: a check failure reports "no syntax error" */
        } finally {
          try {
            if (py.globals.has("__mrnote_lint_src"))
              py.globals.delete("__mrnote_lint_src");
          } catch {
            /* ignore cleanup failure */
          }
        }
        // Step 2: semantic warnings via pyflakes (AST-only, never executes code).
        // Only run when the syntax check passed — pyflakes on broken syntax duplicates
        // the SyntaxError we already surfaced and adds noise. Load pyflakes lazily
        // once; a load failure caches as false so we never retry per keystroke.
        let warnings = [];
        if (lint.ok) {
          if (pyflakesReady === null) {
            pyflakesReady = py
              .loadPackage("micropip")
              .then(() =>
                py.runPythonAsync(helperInstallSource("pyflakes")),
              )
              .then(() => true)
              .catch(() => false);
          }
          const hasPyflakes = await pyflakesReady;
          if (hasPyflakes) {
            py.globals.set("__mrnote_pyflakes_src", src);
            try {
              const out = await py.runPythonAsync(PYFLAKES_CHECK);
              const parsed = pythonResultToJs(out);
              if (trustedArrayIsArray(parsed)) warnings = parsed;
            } catch {
              /* non-fatal: report no warnings */
            } finally {
              try {
                if (py.globals.has("__mrnote_pyflakes_src"))
                  py.globals.delete("__mrnote_pyflakes_src");
              } catch {
                /* ignore cleanup failure */
              }
            }
          }
        }
        postProtocolMessage({ id, lint, warnings });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({
          id,
          lint: { ok: true },
          warnings: [],
          error: message,
        });
      }
      return;
    }
    // Format op: format Python source via autopep8 (PEP 8 style).
    // autopep8 is a pure-Python package, loaded lazily via micropip on first use.
    // TOTAL: returns the ORIGINAL source unchanged on any failure so the buffer is
    // never corrupted (matches the never-throw contract of the TS formatCode helper).
    if (data.op === "format") {
      const src = typeof data.code === "string" ? data.code : "";
      try {
        const py = await getPyodide();
        // Load autopep8 lazily once; a load failure caches as false.
        if (autopep8Ready === null) {
          autopep8Ready = py
            .loadPackage("micropip")
            .then(() =>
              py.runPythonAsync(helperInstallSource("autopep8")),
            )
            .then(() => true)
            .catch(() => false);
        }
        const hasAutopep8 = await autopep8Ready;
        if (!hasAutopep8) {
          postProtocolMessage({ id, formatted: src });
          return;
        }
        py.globals.set("__mrnote_ap8_src", src);
        let formatted = src;
        let formatSucceeded = false;
        try {
          const out = await py.runPythonAsync(AUTOPEP8_FORMAT);
          const parsed = pythonResultToJs(out);
          if (
            parsed &&
            typeof parsed === "object" &&
            parsed.ok === true &&
            typeof parsed.formatted === "string" &&
            parsed.formatted.length <= MAX_PROTOCOL_FORMATTED_CHARS &&
            (trustedStringTrim(src).length === 0 ||
              trustedStringTrim(parsed.formatted).length > 0)
          ) {
            formatted = parsed.formatted;
            formatSucceeded = true;
          }
        } catch {
          /* non-fatal: return original */
        } finally {
          try {
            if (py.globals.has("__mrnote_ap8_src"))
              py.globals.delete("__mrnote_ap8_src");
          } catch {
            /* ignore cleanup failure */
          }
        }
        postProtocolMessage(
          formatSucceeded
            ? {
                id,
                formatted,
                error: PYTHON_FORMAT_SUCCESS_RECEIPT,
              }
            : { id, formatted: src },
        );
      } catch {
        // The `error` field carries the authenticated success receipt. Never
        // copy an ambient exception string into it: a forged matching message
        // must not turn a failed format into an idempotent success.
        postProtocolMessage({ id, formatted: src });
      }
      return;
    }
    // Complete op: jedi semantic completions against the live kernel namespace.
    // Always answers (possibly []) so the editor's completion source never hangs.
    if (data.op === "complete") {
      const src = typeof data.code === "string" ? data.code : "";
      const line = trustedNumberIsFinite(data.line)
        ? trustedNumber(data.line)
        : 1;
      const column = trustedNumberIsFinite(data.column)
        ? trustedNumber(data.column)
        : 0;
      try {
        const py = await getPyodide();
        if (jediReady === null) {
          jediReady = py
            .loadPackage("jedi")
            .then(() => true)
            .catch(() => false);
        }
        const hasJedi = await jediReady;
        if (!hasJedi) {
          postProtocolMessage({ id, completions: [] });
          return;
        }
        py.globals.set("__mrnote_complete_src", src);
        py.globals.set("__mrnote_complete_line", line);
        py.globals.set("__mrnote_complete_col", column);
        let completions = [];
        try {
          const out = await py.runPythonAsync(COMPLETE_CHECK);
          const parsed = pythonResultToJs(out);
          if (trustedArrayIsArray(parsed)) completions = parsed;
        } catch {
          /* non-fatal: report no completions */
        }
        postProtocolMessage({ id, completions });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({ id, completions: [], error: message });
      }
      return;
    }
    // Hover op: jedi type/doc info against the live kernel namespace. Always
    // answers ({found:false} on any failure) so the editor tooltip never hangs.
    if (data.op === "hover") {
      const src = typeof data.code === "string" ? data.code : "";
      const line = trustedNumberIsFinite(data.line)
        ? trustedNumber(data.line)
        : 1;
      const column = trustedNumberIsFinite(data.column)
        ? trustedNumber(data.column)
        : 0;
      try {
        const py = await getPyodide();
        if (jediReady === null) {
          jediReady = py
            .loadPackage("jedi")
            .then(() => true)
            .catch(() => false);
        }
        const hasJedi = await jediReady;
        if (!hasJedi) {
          postProtocolMessage({ id, hover: { found: false } });
          return;
        }
        py.globals.set("__mrnote_hover_src", src);
        py.globals.set("__mrnote_hover_line", line);
        py.globals.set("__mrnote_hover_col", column);
        let hover = { found: false };
        try {
          const out = await py.runPythonAsync(HOVER_CHECK);
          const parsed = pythonResultToJs(out);
          if (parsed && typeof parsed === "object")
            hover = { found: true, ...parsed };
        } catch {
          /* non-fatal: report nothing to show */
        }
        postProtocolMessage({ id, hover });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({ id, hover: { found: false }, error: message });
      }
      return;
    }
    // Signature op: jedi parameter hints (the enclosing call's signature + active
    // arg index) against the live kernel namespace. Always answers ({found:false}
    // on any failure) so the editor's parameter-hint tooltip never hangs.
    if (data.op === "signature") {
      const src = typeof data.code === "string" ? data.code : "";
      const line = trustedNumberIsFinite(data.line)
        ? trustedNumber(data.line)
        : 1;
      const column = trustedNumberIsFinite(data.column)
        ? trustedNumber(data.column)
        : 0;
      try {
        const py = await getPyodide();
        if (jediReady === null) {
          jediReady = py
            .loadPackage("jedi")
            .then(() => true)
            .catch(() => false);
        }
        const hasJedi = await jediReady;
        if (!hasJedi) {
          postProtocolMessage({ id, signature: { found: false } });
          return;
        }
        py.globals.set("__mrnote_signature_src", src);
        py.globals.set("__mrnote_signature_line", line);
        py.globals.set("__mrnote_signature_col", column);
        let signature = { found: false };
        try {
          const out = await py.runPythonAsync(SIGNATURE_CHECK);
          const parsed = pythonResultToJs(out);
          if (parsed && typeof parsed === "object") signature = parsed;
        } catch {
          /* non-fatal: report nothing to show */
        }
        postProtocolMessage({ id, signature });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({
          id,
          signature: { found: false },
          error: message,
        });
      }
      return;
    }
    // Definition op: jedi go-to-definition against the live kernel namespace. Always
    // answers ({found:false} on any failure) so the editor's jump never hangs. The
    // reply's `inBuffer` tells the editor whether the resolved definition lives in
    // THIS buffer (navigable) or out of file (show a notice).
    if (data.op === "definition") {
      const src = typeof data.code === "string" ? data.code : "";
      const line = trustedNumberIsFinite(data.line)
        ? trustedNumber(data.line)
        : 1;
      const column = trustedNumberIsFinite(data.column)
        ? trustedNumber(data.column)
        : 0;
      try {
        const py = await getPyodide();
        if (jediReady === null) {
          jediReady = py
            .loadPackage("jedi")
            .then(() => true)
            .catch(() => false);
        }
        const hasJedi = await jediReady;
        if (!hasJedi) {
          postProtocolMessage({ id, definition: { found: false } });
          return;
        }
        py.globals.set("__mrnote_definition_src", src);
        py.globals.set("__mrnote_definition_line", line);
        py.globals.set("__mrnote_definition_col", column);
        let definition = { found: false };
        try {
          const out = await py.runPythonAsync(DEFINITION_CHECK);
          const parsed = pythonResultToJs(out);
          if (parsed && typeof parsed === "object") definition = parsed;
        } catch {
          /* non-fatal: report nothing to jump to */
        }
        postProtocolMessage({ id, definition });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({
          id,
          definition: { found: false },
          error: message,
        });
      }
      return;
    }
    // References op: jedi find-references against the live kernel namespace, filtered
    // to THIS buffer (rename only edits the current source). Always answers
    // ({found:false, refs:[]} on any failure) so the editor's rename never hangs.
    if (data.op === "references") {
      const src = typeof data.code === "string" ? data.code : "";
      const line = trustedNumberIsFinite(data.line)
        ? trustedNumber(data.line)
        : 1;
      const column = trustedNumberIsFinite(data.column)
        ? trustedNumber(data.column)
        : 0;
      try {
        const py = await getPyodide();
        if (jediReady === null) {
          jediReady = py
            .loadPackage("jedi")
            .then(() => true)
            .catch(() => false);
        }
        const hasJedi = await jediReady;
        if (!hasJedi) {
          postProtocolMessage({ id, references: { found: false, refs: [] } });
          return;
        }
        py.globals.set("__mrnote_references_src", src);
        py.globals.set("__mrnote_references_line", line);
        py.globals.set("__mrnote_references_col", column);
        let references = { found: false, refs: [] };
        try {
          const out = await py.runPythonAsync(REFERENCES_CHECK);
          const parsed = pythonResultToJs(out);
          if (parsed && typeof parsed === "object") references = parsed;
        } catch {
          /* non-fatal: report no references */
        }
        postProtocolMessage({ id, references });
      } catch (err) {
        const message =
          err && err.message ? trustedString(err.message) : trustedString(err);
        postProtocolMessage({
          id,
          references: { found: false, refs: [] },
          error: message,
        });
      }
      return;
    }
    // Debug-run op: recorded breakpoint debugging (see protocol note above).
    // Mirrors the normal run path's I/O behaviour, then returns the recording.
    if (data.op === "debug-run") {
      const src = typeof data.code === "string" ? data.code : "";
      const breakpoints = trustedArrayIsArray(data.breakpoints)
        ? data.breakpoints
        : [];
      const stream = data.stream === true;
      const output = createOutputCapture(id, stream);
      const logs = output.logs;
      const logStreams = output.logStreams;
      try {
        const py = await getPyodide();
        py.setStdout({
          batched: (text) => {
            output.push(text, "stdout");
          },
        });
        py.setStderr({
          batched: (text) => {
            output.push(text, "stderr");
          },
        });
        if (typeof py.setStdin === "function") {
          const stdinText = typeof data.stdin === "string" ? data.stdin : "";
          if (stdinText !== "") {
            const lines = stdinText.split("\n");
            let lineIndex = 0;
            py.setStdin({
              stdin: () =>
                lineIndex < lines.length ? lines[lineIndex++] : null,
            });
          } else {
            py.setStdin({ stdin: () => null });
          }
        }
        if (typeof py.loadPackagesFromImports === "function") {
          try {
            // Scan the USER source (not the driver) for auto-installable imports.
            await py.loadPackagesFromImports(src);
          } catch {
            /* non-fatal: a truly missing import raises during execution */
          }
        }
        py.globals.set("__mrnote_debug_src", src);
        py.globals.set("__mrnote_debug_bps", trustedJsonStringify(breakpoints));
        py.globals.set(
          "__mrnote_debug_pause_uncaught",
          data.pauseOnUncaughtExceptions === true,
        );
        let debugEvents = [];
        let debugTruncated = false;
        let debugConditionNote;
        let result;
        let html;
        let error;
        try {
          const out = await py.runPythonAsync(DEBUG_RUN);
          const parsed = pythonResultToJs(out);
          if (parsed && typeof parsed === "object") {
            if (trustedArrayIsArray(parsed.events)) debugEvents = parsed.events;
            debugTruncated = parsed.truncated === true;
            if (
              typeof parsed.conditionNote === "string" &&
              parsed.conditionNote
            ) {
              debugConditionNote = parsed.conditionNote;
            }
            if (typeof parsed.result === "string") {
              result = truncateOutput(parsed.result, MAX_OUTPUT_TEXT_CHARS);
            }
            if (typeof parsed.html === "string" && parsed.html) {
              html = truncateOutput(parsed.html, MAX_OUTPUT_HTML_CHARS);
            }
            if (typeof parsed.error === "string" && parsed.error) {
              error = truncateOutput(parsed.error, MAX_OUTPUT_TEXT_CHARS);
            }
          }
        } catch (err) {
          // The driver itself failed (it never lets user errors escape) — report
          // like a normal run failure so the editor shows the message.
          error = truncateOutput(
            err && err.message
              ? trustedString(err.message)
              : trustedString(err),
            MAX_OUTPUT_TEXT_CHARS,
          );
        }
        let images;
        try {
          const rich = pythonResultToJs(await py.runPythonAsync(COLLECT_RICH));
          images = boundRichImages(rich.images);
        } catch {
          /* non-fatal */
        }
        postProtocolMessage({
          id,
          logs,
          logStreams,
          result: html ? undefined : result,
          html,
          error,
          images,
          debugEvents,
          debugTruncated,
          debugConditionNote,
        });
      } catch (err) {
        const message = truncateOutput(
          err && err.message ? trustedString(err.message) : trustedString(err),
          MAX_OUTPUT_TEXT_CHARS,
        );
        postProtocolMessage({
          id,
          logs,
          logStreams,
          error: message,
          debugEvents: [],
          debugTruncated: false,
        });
      }
      return;
    }
    const code = typeof data.code === "string" ? data.code : "";
    // When the caller opted into progressive output, mirror each captured chunk to
    // the main thread immediately so the cell renders live (Jupyter streaming).
    // The `batched` callback fires during runPythonAsync (which yields to the
    // event loop on awaits), so chunks pump out mid-run for typical output.
    const stream = data.stream === true;
    const output = createOutputCapture(id, stream);
    const logs = output.logs;
    // Parallel per-line stream tags ("stdout" | "stderr"), aligned to `logs` by
    // index, so the renderer can style stderr distinctly. `logs` stays string[].
    const logStreams = output.logStreams;
    try {
      const py = await getPyodide();
      py.setStdout({
        batched: (text) => {
          output.push(text, "stdout");
        },
      });
      py.setStderr({
        batched: (text) => {
          output.push(text, "stderr");
        },
      });
      // Pre-supplied stdin (the run panel's 输入 box): input() reads it line by
      // line; past the end returns EOF, which Python surfaces as EOFError with
      // its normal message. Reset per run so lines never leak across runs.
      if (typeof py.setStdin === "function") {
        const stdinText = typeof data.stdin === "string" ? data.stdin : "";
        if (stdinText !== "") {
          const lines = stdinText.split("\n");
          let lineIndex = 0;
          py.setStdin({
            stdin: () => (lineIndex < lines.length ? lines[lineIndex++] : null),
          });
        } else {
          py.setStdin({ stdin: () => null });
        }
      }
      if (typeof py.loadPackagesFromImports === "function") {
        try {
          // Auto-install imported stdlib-adjacent packages (numpy, pandas, …).
          await py.loadPackagesFromImports(code);
        } catch {
          /* non-fatal: a truly missing import raises during execution */
        }
      }
      const ret = await py.runPythonAsync(code);
      // Capture a rich HTML repr (e.g. a pandas DataFrame) of the last value.
      let html;
      try {
        if (ret && typeof ret._repr_html_ === "function") {
          const reprHtml = ret._repr_html_();
          // Cap the rich repr so a pathological object can't bloat the message/DOM.
          if (typeof reprHtml === "string" && reprHtml) {
            html = truncateOutput(reprHtml, MAX_OUTPUT_HTML_CHARS);
          }
        }
      } catch {
        /* non-fatal: fall back to the text repr */
      }
      const result =
        ret === undefined || ret === null
          ? undefined
          : truncateOutput(ret, MAX_OUTPUT_TEXT_CHARS);
      let images;
      try {
        const rich = pythonResultToJs(await py.runPythonAsync(COLLECT_RICH));
        images = boundRichImages(rich.images);
      } catch {
        /* non-fatal */
      }
      // When a rich HTML repr exists, suppress the redundant text repr.
      postProtocolMessage({
        id,
        logs,
        logStreams,
        result: html ? undefined : result,
        html,
        images,
      });
    } catch (err) {
      // Pyodide errors carry the full Python traceback in `.message` (it includes
      // the "Traceback (most recent call last):" frames + the final
      // "ErrorType: message" line), which the renderer parses into a styled block.
      const message = truncateOutput(
        withoutRunnerFrames(
          err && err.message ? trustedString(err.message) : trustedString(err),
        ),
        MAX_OUTPUT_TEXT_CHARS,
      );
      postProtocolMessage({ id, logs, logStreams, error: message });
    }
  });
})();

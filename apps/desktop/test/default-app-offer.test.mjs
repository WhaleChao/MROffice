/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// When MR Note Classic asks at launch to open documents by default (src/default-app-offer.ts).
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { test } from "node:test";

const require = createRequire(import.meta.url);
const { parseRegValue, readDefaultAppChoice, shouldOfferDefaultApp, writeDefaultAppChoice } = require("../dist/default-app-offer.js");

const all = ["docx", "xlsx", "pptx", "pdf", "md", "csv"];
const state = (defaultFor, extra = {}) => ({ supported: "direct", checked: true, defaultFor, extensions: all, ...extra });

test("asks when any type opens with another app, not when all open here", () => {
  assert.equal(shouldOfferDefaultApp(state(all.filter((ext) => ext !== "pdf")), { declined: false }), true);
  assert.equal(shouldOfferDefaultApp(state(all), { declined: false }), false);
  assert.equal(shouldOfferDefaultApp(state([], { supported: "settings" }), { declined: false }), true);
});

test("never after No thanks, nor without an answer from the system", () => {
  assert.equal(shouldOfferDefaultApp(state([]), { declined: true }), false);
  assert.equal(shouldOfferDefaultApp(state([], { checked: false }), { declined: false }), false);
  assert.equal(shouldOfferDefaultApp(state([], { supported: "none" }), { declined: false }), false);
});

test("No thanks is remembered across launches, and can be taken back", () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "mrai-default-app-"));
  const file = path.join(dir, "default-app-choice.json");
  assert.deepEqual(readDefaultAppChoice(file), { declined: false });
  writeDefaultAppChoice(file, { declined: true });
  assert.deepEqual(readDefaultAppChoice(file), { declined: true });
  writeDefaultAppChoice(file, { declined: false });
  assert.deepEqual(readDefaultAppChoice(file), { declined: false });
  fs.writeFileSync(file, "not json");
  assert.deepEqual(readDefaultAppChoice(file), { declined: false });
  fs.rmSync(dir, { recursive: true, force: true });
});

test("reads a ProgID out of reg query", () => {
  const out = "\r\nHKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.docx\\UserChoice\r\n    ProgId    REG_SZ    MROfficeEditors.Document\r\n\r\n";
  assert.equal(parseRegValue(out), "MROfficeEditors.Document");
  assert.equal(parseRegValue("\r\nHKEY_CLASSES_ROOT\\.pdf\r\n    (默认)    REG_SZ    AcroExch.Document.DC\r\n"), "AcroExch.Document.DC");
  assert.equal(parseRegValue("\r\nHKEY_CLASSES_ROOT\\.md\r\n"), null);
});

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where an off-screen text probe is laid out: a shadow root under one
 * persistent, hidden host on the body.
 *
 * A probe appended to the body changes the body's children, and WebKit then
 * restyles the whole document before the probe can be read: with four editor
 * windows open (about 10,000 elements) roughly 650 ms for every batch of
 * worksheet cells measured while a window opens. Inside the shadow root a
 * probe restyles only its own tree. The host inherits from the body exactly
 * as a probe on the body did, @font-face rules and the document's font set
 * reach the shadow tree, and the page's own rules never matched a probe's
 * text (every font and spacing property a probe depends on is set inline), so
 * a probe measures the same there.
 *
 * The worksheet's text probes and the Office font metrics' line-box probes
 * (components/office-ribbon/officeFontLineBox.ts) share this one host.
 */
const probeRoots = new WeakMap<Document, ShadowRoot>();

/** The parent to put a probe in: the shadow root, or the body where shadow
 * roots are unavailable. */
export function layoutProbeParent(doc: Document): ParentNode {
  const known = probeRoots.get(doc);
  if (known?.host.isConnected) return known;
  const host = doc.createElement("div");
  // No aria-hidden: visibility:hidden already keeps it out of the
  // accessibility tree, and the probes inside carry their own.
  host.dataset.layoutProbes = "true";
  host.style.cssText =
    "position:fixed;left:0;top:0;width:0;height:0;overflow:hidden;visibility:hidden;pointer-events:none;contain:strict;";
  let root: ShadowRoot;
  try {
    root = host.attachShadow({ mode: "open" });
  } catch {
    return doc.body;
  }
  doc.body.append(host);
  probeRoots.set(doc, root);
  return root;
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A small typed model of Office Math (OMML) shared by the UnicodeMath
 * serializer and parser. It covers every construct Word's linear format can
 * spell; anything else makes `astFromOmml` return null so callers fail closed.
 *
 * The model is deliberately close to OMML: one node per OMML object, argument
 * slots as node lists, and run text with its math run properties.
 */

import { defaultLimitLocation } from "./characters";

export type MathScript =
  | "roman"
  | "script"
  | "fraktur"
  | "double-struck"
  | "sans-serif"
  | "monospace";

export type MathSty = "p" | "b" | "i" | "bi";

export interface MathRunStyle {
  readonly sty?: MathSty;
  readonly scr?: MathScript;
  /** m:nor — ordinary (non-math) text. */
  readonly nor?: boolean;
  /** m:lit — a literal character that must not build up. */
  readonly lit?: boolean;
}

export type MathList = MathNode[];

export type MathNode =
  | { kind: "text"; text: string; style: MathRunStyle }
  | {
      kind: "f";
      type: "bar" | "skw" | "lin" | "noBar";
      num: MathList;
      den: MathList;
    }
  | { kind: "sSup"; e: MathList; sup: MathList }
  | { kind: "sSub"; e: MathList; sub: MathList }
  | { kind: "sSubSup"; e: MathList; sub: MathList; sup: MathList }
  | { kind: "sPre"; e: MathList; sub: MathList; sup: MathList }
  | { kind: "rad"; deg: MathList; e: MathList }
  | {
      kind: "nary";
      chr: string;
      limLoc: "undOvr" | "subSup" | null;
      sub: MathList | null;
      sup: MathList | null;
      e: MathList;
    }
  | { kind: "d"; beg: string; end: string; sep: string; args: MathList[] }
  | { kind: "m"; rows: MathList[][]; columnJc: readonly string[] }
  | { kind: "eqArr"; rows: MathList[] }
  | { kind: "func"; name: MathList; e: MathList }
  | { kind: "limLow"; e: MathList; lim: MathList }
  | { kind: "limUpp"; e: MathList; lim: MathList }
  | { kind: "acc"; chr: string; e: MathList }
  | { kind: "bar"; pos: "top" | "bot"; e: MathList }
  | {
      kind: "groupChr";
      chr: string;
      pos: "top" | "bot";
      vertJc: "top" | "bot";
      e: MathList;
    }
  | { kind: "borderBox"; flags: readonly string[]; e: MathList }
  | { kind: "box"; e: MathList }
  | {
      kind: "phant";
      show: boolean;
      zeroWid: boolean;
      zeroAsc: boolean;
      zeroDesc: boolean;
      transp: boolean;
      e: MathList;
    };

/* ------------------------------------------------------------------------ */
/* Minimal XML (OMML is produced by our own exporter or by this module).     */
/* ------------------------------------------------------------------------ */

export interface XmlNode {
  /** Qualified name as written, e.g. `m:r`. */
  readonly name: string;
  readonly attrs: Readonly<Record<string, string>>;
  readonly children: XmlNode[];
  /** Character data directly inside this element (m:t only). */
  text: string;
}

const XML_ENTITIES: Readonly<Record<string, string>> = {
  lt: "<",
  gt: ">",
  amp: "&",
  quot: '"',
  apos: "'",
};

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-z]+);/g, (full, body) => {
    if (body[0] === "#") {
      const code =
        body[1] === "x" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) && code >= 0 && code <= 0x10ffff
        ? String.fromCodePoint(code)
        : full;
    }
    return XML_ENTITIES[body] ?? full;
  });
}

/** Parse well-formed XML into a tree; returns null on any malformation. */
export function parseXml(source: string): XmlNode | null {
  const stack: XmlNode[] = [];
  let root: XmlNode | null = null;
  const pattern =
    /<\?[^>]*\?>|<!--[\s\S]*?-->|<\/([A-Za-z_][\w.:-]*)\s*>|<([A-Za-z_][\w.:-]*)((?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let consumed = 0;
  for (const match of source.matchAll(pattern)) {
    if ((match.index ?? 0) !== consumed) return null;
    consumed += match[0].length;
    if (match[0].startsWith("<?") || match[0].startsWith("<!--")) continue;
    if (match[1] !== undefined) {
      const open = stack.pop();
      if (!open || open.name !== match[1]) return null;
      continue;
    }
    if (match[2] !== undefined) {
      const attrs: Record<string, string> = {};
      for (const attr of match[3].matchAll(
        /([A-Za-z_][\w.:-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g,
      )) {
        attrs[attr[1]] = decodeXmlText(attr[2] ?? attr[3] ?? "");
      }
      const node: XmlNode = { name: match[2], attrs, children: [], text: "" };
      const parent = stack.at(-1);
      if (parent) parent.children.push(node);
      else if (root) return null;
      else root = node;
      if (match[4] !== "/") stack.push(node);
      continue;
    }
    const parent = stack.at(-1);
    if (parent) parent.text += decodeXmlText(match[5] ?? "");
    else if ((match[5] ?? "").trim()) return null;
  }
  if (consumed !== source.length || stack.length > 0) return null;
  return root;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const localName = (node: XmlNode) => node.name.slice(node.name.indexOf(":") + 1);
const isMath = (node: XmlNode) => node.name.startsWith("m:");
const mathChildren = (node: XmlNode) => node.children.filter(isMath);
const child = (node: XmlNode, name: string) =>
  node.children.find((item) => item.name === `m:${name}`);
const val = (node: XmlNode | undefined) =>
  node ? (node.attrs["m:val"] ?? node.attrs.val ?? null) : null;

/** OOXML on/off: an element present without a value is on. */
function flag(properties: XmlNode | undefined, name: string): boolean | null {
  const element = properties ? child(properties, name) : undefined;
  if (!element) return null;
  const value = val(element);
  return value === null || value === "1" || value === "on" || value === "true";
}

const ARGUMENT_SLOTS = new Set([
  "e",
  "num",
  "den",
  "sub",
  "sup",
  "deg",
  "lim",
  "fName",
]);

/* ------------------------------------------------------------------------ */
/* OMML → model.                                                              */
/* ------------------------------------------------------------------------ */

class Unsupported extends Error {}

function argument(node: XmlNode, slot: string): MathList {
  const element = child(node, slot);
  return element ? listFromOmml(element) : [];
}

function runStyle(run: XmlNode): MathRunStyle {
  const properties = child(run, "rPr");
  if (!properties) return {};
  const style: {
    sty?: MathSty;
    scr?: MathScript;
    nor?: boolean;
    lit?: boolean;
  } = {};
  const sty = val(child(properties, "sty"));
  if (sty === "p" || sty === "b" || sty === "i" || sty === "bi") style.sty = sty;
  const scr = val(child(properties, "scr"));
  if (
    scr === "roman" ||
    scr === "script" ||
    scr === "fraktur" ||
    scr === "double-struck" ||
    scr === "sans-serif" ||
    scr === "monospace"
  ) {
    style.scr = scr;
  }
  if (flag(properties, "nor")) style.nor = true;
  if (flag(properties, "lit")) style.lit = true;
  return style;
}

const IGNORED_ELEMENTS = new Set(["ctrlPr", "argPr"]);

function listFromOmml(container: XmlNode): MathList {
  const out: MathList = [];
  for (const element of mathChildren(container)) {
    const name = localName(element);
    if (IGNORED_ELEMENTS.has(name) || name.endsWith("Pr")) continue;
    out.push(nodeFromOmml(element));
  }
  return out;
}

function nodeFromOmml(element: XmlNode): MathNode {
  const name = localName(element);
  switch (name) {
    case "r": {
      const text = element.children
        .filter((item) => item.name === "m:t")
        .map((item) => item.text)
        .join("");
      return { kind: "text", text, style: runStyle(element) };
    }
    case "f": {
      const type = val(child(child(element, "fPr") ?? element, "type"));
      return {
        kind: "f",
        type:
          type === "skw" || type === "lin" || type === "noBar" ? type : "bar",
        num: argument(element, "num"),
        den: argument(element, "den"),
      };
    }
    case "sSup":
      return { kind: "sSup", e: argument(element, "e"), sup: argument(element, "sup") };
    case "sSub":
      return { kind: "sSub", e: argument(element, "e"), sub: argument(element, "sub") };
    case "sSubSup":
      return {
        kind: "sSubSup",
        e: argument(element, "e"),
        sub: argument(element, "sub"),
        sup: argument(element, "sup"),
      };
    case "sPre":
      return {
        kind: "sPre",
        e: argument(element, "e"),
        sub: argument(element, "sub"),
        sup: argument(element, "sup"),
      };
    case "rad": {
      const hidden = flag(child(element, "radPr"), "degHide");
      return {
        kind: "rad",
        deg: hidden ? [] : argument(element, "deg"),
        e: argument(element, "e"),
      };
    }
    case "nary": {
      const properties = child(element, "naryPr");
      const limLoc = val(properties ? child(properties, "limLoc") : undefined);
      return {
        kind: "nary",
        chr: val(properties ? child(properties, "chr") : undefined) || "∫",
        limLoc: limLoc === "undOvr" || limLoc === "subSup" ? limLoc : null,
        sub: flag(properties, "subHide") ? null : argument(element, "sub"),
        sup: flag(properties, "supHide") ? null : argument(element, "sup"),
        e: argument(element, "e"),
      };
    }
    case "d": {
      const properties = child(element, "dPr");
      const beg = properties ? child(properties, "begChr") : undefined;
      const end = properties ? child(properties, "endChr") : undefined;
      const sep = properties ? child(properties, "sepChr") : undefined;
      return {
        kind: "d",
        beg: beg ? (val(beg) ?? "") : "(",
        end: end ? (val(end) ?? "") : ")",
        sep: sep ? (val(sep) ?? "") : "|",
        args: element.children
          .filter((item) => item.name === "m:e")
          .map((item) => listFromOmml(item)),
      };
    }
    case "m": {
      const properties = child(element, "mPr");
      const columnJc: string[] = [];
      const columns = properties ? child(properties, "mcs") : undefined;
      for (const column of columns ? mathChildren(columns) : []) {
        const columnProperties = child(column, "mcPr");
        const count = Number(val(columnProperties ? child(columnProperties, "count") : undefined) ?? 1);
        const jc = val(columnProperties ? child(columnProperties, "mcJc") : undefined) ?? "center";
        for (let index = 0; index < Math.min(Math.max(count, 1), 256); index += 1) {
          columnJc.push(jc);
        }
      }
      return {
        kind: "m",
        rows: element.children
          .filter((item) => item.name === "m:mr")
          .map((row) =>
            row.children
              .filter((item) => item.name === "m:e")
              .map((cell) => listFromOmml(cell)),
          ),
        columnJc,
      };
    }
    case "eqArr":
      return {
        kind: "eqArr",
        rows: element.children
          .filter((item) => item.name === "m:e")
          .map((item) => listFromOmml(item)),
      };
    case "func":
      return {
        kind: "func",
        name: argument(element, "fName"),
        e: argument(element, "e"),
      };
    case "limLow":
      return { kind: "limLow", e: argument(element, "e"), lim: argument(element, "lim") };
    case "limUpp":
      return { kind: "limUpp", e: argument(element, "e"), lim: argument(element, "lim") };
    case "acc": {
      const properties = child(element, "accPr");
      return {
        kind: "acc",
        chr: val(properties ? child(properties, "chr") : undefined) || "̂",
        e: argument(element, "e"),
      };
    }
    case "bar": {
      const properties = child(element, "barPr");
      return {
        kind: "bar",
        pos: val(properties ? child(properties, "pos") : undefined) === "top" ? "top" : "bot",
        e: argument(element, "e"),
      };
    }
    case "groupChr": {
      const properties = child(element, "groupChrPr");
      return {
        kind: "groupChr",
        chr: val(properties ? child(properties, "chr") : undefined) || "⏟",
        pos: val(properties ? child(properties, "pos") : undefined) === "top" ? "top" : "bot",
        vertJc:
          val(properties ? child(properties, "vertJc") : undefined) === "bot" ? "bot" : "top",
        e: argument(element, "e"),
      };
    }
    case "borderBox": {
      const properties = child(element, "borderBoxPr");
      const flags = [
        "hideTop",
        "hideBot",
        "hideLeft",
        "hideRight",
        "strikeH",
        "strikeV",
        "strikeBLTR",
        "strikeTLBR",
      ].filter((name) => flag(properties, name));
      return { kind: "borderBox", flags, e: argument(element, "e") };
    }
    case "box":
      return { kind: "box", e: argument(element, "e") };
    case "phant": {
      const properties = child(element, "phantPr");
      return {
        kind: "phant",
        show: flag(properties, "show") ?? true,
        zeroWid: flag(properties, "zeroWid") ?? false,
        zeroAsc: flag(properties, "zeroAsc") ?? false,
        zeroDesc: flag(properties, "zeroDesc") ?? false,
        transp: flag(properties, "transp") ?? false,
        e: argument(element, "e"),
      };
    }
    default:
      if (ARGUMENT_SLOTS.has(name)) throw new Unsupported(name);
      throw new Unsupported(name);
  }
}

/** Read an `<m:oMath>` (or `<m:oMathPara>`) string; null when unsupported. */
export function astFromOmml(xml: string): MathList | null {
  const root = parseXml(xml);
  if (!root) return null;
  let zone = root;
  if (localName(zone) === "oMathPara") {
    const zones = zone.children.filter((item) => localName(item) === "oMath");
    if (zones.length !== 1) return null;
    zone = zones[0];
  }
  if (localName(zone) !== "oMath") return null;
  try {
    return listFromOmml(zone);
  } catch (error) {
    if (error instanceof Unsupported) return null;
    throw error;
  }
}

/* ------------------------------------------------------------------------ */
/* Model → OMML.                                                              */
/* ------------------------------------------------------------------------ */

const on = (name: string) => `<m:${name} m:val="1"/>`;
const valued = (name: string, value: string) =>
  `<m:${name} m:val="${escapeXml(value)}"/>`;

function slot(name: string, list: MathList): string {
  const body = listToOmml(list);
  return body ? `<m:${name}>${body}</m:${name}>` : `<m:${name}/>`;
}

function runToOmml(text: string, style: MathRunStyle): string {
  if (!text) return "";
  const properties =
    (style.lit ? "<m:lit/>" : "") +
    (style.nor ? "<m:nor/>" : "") +
    (style.scr ? valued("scr", style.scr) : "") +
    (style.sty ? valued("sty", style.sty) : "");
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
  return `<m:r>${properties ? `<m:rPr>${properties}</m:rPr>` : ""}<m:t${space}>${escapeXml(text)}</m:t></m:r>`;
}

function nodeToOmml(node: MathNode): string {
  switch (node.kind) {
    case "text":
      return runToOmml(node.text, node.style);
    case "f":
      return `<m:f>${node.type === "bar" ? "" : `<m:fPr>${valued("type", node.type)}</m:fPr>`}${slot("num", node.num)}${slot("den", node.den)}</m:f>`;
    case "sSup":
      return `<m:sSup>${slot("e", node.e)}${slot("sup", node.sup)}</m:sSup>`;
    case "sSub":
      return `<m:sSub>${slot("e", node.e)}${slot("sub", node.sub)}</m:sSub>`;
    case "sSubSup":
      return `<m:sSubSup>${slot("e", node.e)}${slot("sub", node.sub)}${slot("sup", node.sup)}</m:sSubSup>`;
    case "sPre":
      return `<m:sPre>${slot("sub", node.sub)}${slot("sup", node.sup)}${slot("e", node.e)}</m:sPre>`;
    case "rad":
      return `<m:rad>${node.deg.length === 0 ? `<m:radPr>${on("degHide")}</m:radPr>` : ""}${slot("deg", node.deg)}${slot("e", node.e)}</m:rad>`;
    case "nary": {
      const properties =
        valued("chr", node.chr) +
        valued("limLoc", node.limLoc ?? defaultLimitLocation(node.chr)) +
        (node.sub === null ? on("subHide") : "") +
        (node.sup === null ? on("supHide") : "");
      return `<m:nary><m:naryPr>${properties}</m:naryPr>${slot("sub", node.sub ?? [])}${slot("sup", node.sup ?? [])}${slot("e", node.e)}</m:nary>`;
    }
    case "d": {
      const properties =
        (node.beg !== "(" ? valued("begChr", node.beg) : "") +
        (node.sep !== "|" ? valued("sepChr", node.sep) : "") +
        (node.end !== ")" ? valued("endChr", node.end) : "");
      return `<m:d>${properties ? `<m:dPr>${properties}</m:dPr>` : ""}${node.args.map((arg) => slot("e", arg)).join("")}</m:d>`;
    }
    case "m": {
      const columns = Math.max(1, ...node.rows.map((row) => row.length));
      const jc = Array.from({ length: columns }, (_, index) => node.columnJc[index] ?? "center");
      // Adjacent columns of one justification share an m:mc, as Word and
      // the DOCX exporter write them.
      const groups: { jc: string; count: number }[] = [];
      for (const value of jc) {
        const last = groups.at(-1);
        if (last && last.jc === value) last.count += 1;
        else groups.push({ jc: value, count: 1 });
      }
      const mcs = groups
        .map((group) => `<m:mc><m:mcPr>${valued("count", String(group.count))}${valued("mcJc", group.jc)}</m:mcPr></m:mc>`)
        .join("");
      return `<m:m><m:mPr><m:mcs>${mcs}</m:mcs></m:mPr>${node.rows
        .map((row) => `<m:mr>${row.map((cell) => slot("e", cell)).join("")}</m:mr>`)
        .join("")}</m:m>`;
    }
    case "eqArr":
      return `<m:eqArr>${node.rows.map((row) => slot("e", row)).join("")}</m:eqArr>`;
    case "func":
      return `<m:func>${slot("fName", node.name)}${slot("e", node.e)}</m:func>`;
    case "limLow":
      return `<m:limLow>${slot("e", node.e)}${slot("lim", node.lim)}</m:limLow>`;
    case "limUpp":
      return `<m:limUpp>${slot("e", node.e)}${slot("lim", node.lim)}</m:limUpp>`;
    case "acc":
      return `<m:acc>${node.chr === "̂" ? "" : `<m:accPr>${valued("chr", node.chr)}</m:accPr>`}${slot("e", node.e)}</m:acc>`;
    case "bar":
      return `<m:bar>${node.pos === "top" ? `<m:barPr>${valued("pos", "top")}</m:barPr>` : ""}${slot("e", node.e)}</m:bar>`;
    case "groupChr":
      return `<m:groupChr><m:groupChrPr>${valued("chr", node.chr)}${valued("pos", node.pos)}${valued("vertJc", node.vertJc)}</m:groupChrPr>${slot("e", node.e)}</m:groupChr>`;
    case "borderBox":
      return `<m:borderBox>${node.flags.length ? `<m:borderBoxPr>${node.flags.map(on).join("")}</m:borderBoxPr>` : ""}${slot("e", node.e)}</m:borderBox>`;
    case "box":
      return `<m:box>${slot("e", node.e)}</m:box>`;
    case "phant": {
      const properties =
        (node.show ? "" : valued("show", "0")) +
        (node.zeroWid ? on("zeroWid") : "") +
        (node.zeroAsc ? on("zeroAsc") : "") +
        (node.zeroDesc ? on("zeroDesc") : "") +
        (node.transp ? on("transp") : "");
      return `<m:phant>${properties ? `<m:phantPr>${properties}</m:phantPr>` : ""}${slot("e", node.e)}</m:phant>`;
    }
  }
}

export function listToOmml(list: MathList): string {
  return list.map(nodeToOmml).join("");
}

export const OMML_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** A namespaced `<m:oMath>` element string for the model. */
export function astToOmml(list: MathList): string {
  return `<m:oMath xmlns:m="${OMML_NAMESPACE}">${listToOmml(list)}</m:oMath>`;
}

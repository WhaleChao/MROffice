/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A small XML reader for Office's SmartArt definitions (layout, colours, quick
 * style) and data models. They are well-formed, carry no DTD and no entities
 * beyond the five predefined ones, so a tokenizer is enough and the engine
 * runs the same in the browser, in Node hosts and in tests (no DOMParser).
 *
 * Names are kept as LOCAL names (`dgm:layoutNode` -> `layoutNode`): the
 * definitions never use the same local name for two meanings.
 */

export type XmlElement = {
  name: string;
  attrs: Record<string, string>;
  children: XmlElement[];
  text: string;
};

const ENTITIES: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

function decode(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-z]+);/g, (whole, body: string) => {
    if (body[0] === "#") {
      const code = body[1] === "x" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    return ENTITIES[body] ?? whole;
  });
}

function local(name: string): string {
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

const ATTRIBUTE = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

export function parseXml(source: string): XmlElement {
  const root: XmlElement = { name: "#document", attrs: {}, children: [], text: "" };
  const stack: XmlElement[] = [root];
  let index = 0;
  const length = source.length;
  while (index < length) {
    const open = source.indexOf("<", index);
    if (open < 0) break;
    if (open > index) stack[stack.length - 1].text += decode(source.slice(index, open));
    if (source.startsWith("<?", open)) {
      index = source.indexOf("?>", open) + 2;
      continue;
    }
    if (source.startsWith("<!--", open)) {
      index = source.indexOf("-->", open) + 3;
      continue;
    }
    if (source.startsWith("<![CDATA[", open)) {
      const end = source.indexOf("]]>", open);
      stack[stack.length - 1].text += source.slice(open + 9, end);
      index = end + 3;
      continue;
    }
    if (source.startsWith("<!", open)) {
      index = source.indexOf(">", open) + 1;
      continue;
    }
    const close = source.indexOf(">", open);
    if (close < 0) throw new Error("unterminated tag");
    const body = source.slice(open + 1, close);
    index = close + 1;
    if (body[0] === "/") {
      if (stack.length > 1) stack.pop();
      continue;
    }
    const selfClosing = body.endsWith("/");
    const content = selfClosing ? body.slice(0, -1) : body;
    const space = content.search(/\s/);
    const name = local(space < 0 ? content : content.slice(0, space));
    const attrs: Record<string, string> = {};
    if (space >= 0) {
      ATTRIBUTE.lastIndex = 0;
      const rest = content.slice(space);
      let match: RegExpExecArray | null;
      while ((match = ATTRIBUTE.exec(rest))) {
        const key = match[1];
        if (key === "xmlns" || key.startsWith("xmlns:")) continue;
        attrs[local(key)] = decode(match[3] ?? match[4] ?? "");
      }
    }
    const element: XmlElement = { name, attrs, children: [], text: "" };
    stack[stack.length - 1].children.push(element);
    if (!selfClosing) stack.push(element);
  }
  const top = root.children[0];
  if (!top) throw new Error("empty document");
  return top;
}

export function child(element: XmlElement | undefined, name: string): XmlElement | undefined {
  return element?.children.find((item) => item.name === name);
}

export function childrenNamed(element: XmlElement | undefined, name: string): XmlElement[] {
  return element ? element.children.filter((item) => item.name === name) : [];
}

export function descendants(element: XmlElement, name: string, out: XmlElement[] = []): XmlElement[] {
  for (const item of element.children) {
    if (item.name === name) out.push(item);
    descendants(item, name, out);
  }
  return out;
}

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office Math (OMML) to LaTeX, for the equation editor's Word-compatible
 * UnicodeMath (lib/unicode-math). Moved out of the deleted browser Word
 * importer (officeCompiler.ts) unchanged: it reads one <m:oMath> element tree,
 * never a package. The DOCX import itself is the server's
 * (app/services/file_adapters/docx/omml_to_latex.py).
 */
import { splitBehind } from "../regexp-behind";

type XmlElement = Element;

const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Descendant elements by local name, in document order. */
function descendantsByName(node: Document | Element, name: string): XmlElement[] {
  if (name === "*") return [];
  return Array.from(node.getElementsByTagNameNS("*", name));
}

function localName(node: Node): string {
  const maybeNamed = node as Node & { localName?: string };
  return (
    maybeNamed.localName || node.nodeName.split(":").pop() || node.nodeName
  );
}

function childrenByName(node: ParentNode, name: string): XmlElement[] {
  return Array.from(node.children).filter((child) => localName(child) === name);
}

/** Max characters of a converted equation's LaTeX (defends against a forged
 * deeply-nested OMML tree blowing up the formula string). */
const MAX_MATH_LATEX_CHARS = 4000;
/** Cap on OMML element recursion depth (a hostile tree cannot stack-overflow). */
const MAX_MATH_DEPTH = 64;

/**
 * Unicode → LaTeX for the symbols Word writes literally inside `<m:t>` (the
 * inverse of the exporter's LATEX_TO_UNICODE). Mapping back to a control word
 * keeps the formula KaTeX-renderable; an unmapped glyph is kept verbatim (KaTeX
 * renders most Unicode math glyphs directly), so this only needs the frequent
 * ones. Multi-character operators (≤, ≥, →, …) and Greek dominate real docs.
 */
const MATH_UNICODE_TO_LATEX: Record<string, string> = {
  α: "\\alpha",
  β: "\\beta",
  γ: "\\gamma",
  δ: "\\delta",
  ε: "\\varepsilon",
  ϵ: "\\epsilon",
  ζ: "\\zeta",
  η: "\\eta",
  θ: "\\theta",
  ϑ: "\\vartheta",
  ι: "\\iota",
  κ: "\\kappa",
  λ: "\\lambda",
  μ: "\\mu",
  ν: "\\nu",
  ξ: "\\xi",
  π: "\\pi",
  ϖ: "\\varpi",
  ρ: "\\rho",
  ϱ: "\\varrho",
  σ: "\\sigma",
  ς: "\\varsigma",
  τ: "\\tau",
  υ: "\\upsilon",
  φ: "\\varphi",
  ϕ: "\\phi",
  χ: "\\chi",
  ψ: "\\psi",
  ω: "\\omega",
  Γ: "\\Gamma",
  Δ: "\\Delta",
  Θ: "\\Theta",
  Λ: "\\Lambda",
  Ξ: "\\Xi",
  Π: "\\Pi",
  Σ: "\\Sigma",
  Υ: "\\Upsilon",
  Φ: "\\Phi",
  Ψ: "\\Psi",
  Ω: "\\Omega",
  "±": "\\pm",
  "∓": "\\mp",
  "×": "\\times",
  "÷": "\\div",
  "⋅": "\\cdot",
  "∗": "\\ast",
  "∘": "\\circ",
  "∙": "\\bullet",
  "⊕": "\\oplus",
  "⊖": "\\ominus",
  "⊗": "\\otimes",
  "⊙": "\\odot",
  "≤": "\\leq",
  "≥": "\\geq",
  "≠": "\\neq",
  "≈": "\\approx",
  "≡": "\\equiv",
  "≅": "\\cong",
  "≃": "\\simeq",
  "∼": "\\sim",
  "∝": "\\propto",
  "≪": "\\ll",
  "≫": "\\gg",
  "≺": "\\prec",
  "≻": "\\succ",
  "∈": "\\in",
  "∉": "\\notin",
  "∋": "\\ni",
  "⊂": "\\subset",
  "⊃": "\\supset",
  "⊆": "\\subseteq",
  "⊇": "\\supseteq",
  "∪": "\\cup",
  "∩": "\\cap",
  "∖": "\\setminus",
  "∅": "\\emptyset",
  "∧": "\\wedge",
  "∨": "\\vee",
  "¬": "\\neg",
  "∀": "\\forall",
  "∃": "\\exists",
  "∄": "\\nexists",
  "⊥": "\\perp",
  "∥": "\\parallel",
  "⊢": "\\vdash",
  "⊨": "\\models",
  "⊔": "\\sqcup",
  "⊓": "\\sqcap",
  "∑": "\\sum",
  "∏": "\\prod",
  "∐": "\\coprod",
  "∫": "\\int",
  "∬": "\\iint",
  "∭": "\\iiint",
  "∮": "\\oint",
  "⋀": "\\bigwedge",
  "⋁": "\\bigvee",
  "⋂": "\\bigcap",
  "⋃": "\\bigcup",
  "⨀": "\\bigodot",
  "⨁": "\\bigoplus",
  "⨂": "\\bigotimes",
  "⨄": "\\biguplus",
  "⨆": "\\bigsqcup",
  "∯": "\\oiint",
  "∰": "\\oiiint",
  "ℸ": "\\daleth",
  "ð": "\\eth",
  "≀": "\\wr",
  "∔": "\\dotplus",
  "⊞": "\\boxplus",
  "⊟": "\\boxminus",
  "⊠": "\\boxtimes",
  "⊡": "\\boxdot",
  "⋉": "\\ltimes",
  "⋊": "\\rtimes",
  "⋋": "\\leftthreetimes",
  "⋌": "\\rightthreetimes",
  "⋎": "\\curlyvee",
  "⋏": "\\curlywedge",
  "⊛": "\\circledast",
  "⊚": "\\circledcirc",
  "⊝": "\\circleddash",
  "⋖": "\\lessdot",
  "⋗": "\\gtrdot",
  "⋘": "\\lll",
  "⋙": "\\ggg",
  "⪅": "\\lessapprox",
  "⪆": "\\gtrapprox",
  "⋈": "\\bowtie",
  "⋔": "\\pitchfork",
  "↽": "\\leftharpoondown",
  "⇁": "\\rightharpoondown",
  "⊀": "\\nprec",
  "⊁": "\\nsucc",
  "⊄": "\\nsubset",
  "⊅": "\\nsupset",
  "⊬": "\\nvdash",
  "⊭": "\\nvDash",
  "∡": "\\measuredangle",
  "∢": "\\sphericalangle",
  "▲": "\\blacktriangle",
  "▽": "\\triangledown",
  "▼": "\\blacktriangledown",
  "→": "\\to",
  "←": "\\leftarrow",
  "↔": "\\leftrightarrow",
  "⇒": "\\Rightarrow",
  "⇐": "\\Leftarrow",
  "⇔": "\\Leftrightarrow",
  "↦": "\\mapsto",
  "↪": "\\hookrightarrow",
  "⇀": "\\rightharpoonup",
  "↑": "\\uparrow",
  "↓": "\\downarrow",
  "⟶": "\\longrightarrow",
  "⟹": "\\Longrightarrow",
  "↗": "\\nearrow",
  "↘": "\\searrow",
  "∂": "\\partial",
  "∇": "\\nabla",
  "∞": "\\infty",
  ℏ: "\\hbar",
  ℓ: "\\ell",
  "℘": "\\wp",
  ℜ: "\\Re",
  ℑ: "\\Im",
  ℵ: "\\aleph",
  "∠": "\\angle",
  "△": "\\triangle",
  "□": "\\square",
  "∎": "\\blacksquare",
  "…": "\\dots",
  "⋯": "\\cdots",
  "⋮": "\\vdots",
  "⋱": "\\ddots",
  "〈": "\\langle",
  "〉": "\\rangle",
  "‖": "\\Vert",
  "⌊": "\\lfloor",
  "⌋": "\\rfloor",
  "⌈": "\\lceil",
  "⌉": "\\rceil",
  "†": "\\dagger",
  "‡": "\\ddagger",
  "★": "\\star",
  "✓": "\\checkmark",
  "√": "\\sqrt",
  "∙ ": "\\cdot",
  // Unicode/typographic spaces have explicit width in Office Math.  Plain
  // ASCII token whitespace is handled separately below.
  " ": "\\;",
  " ": "\\;\\;",
  " ": "\\quad{}",
  " ": "\\;",
  " ": "\\;",
  " ": "\\,",
  " ": "\\phantom{0}",
  " ": "\\;",
  " ": "\\,",
  " ": "\\,",
  " ": "\\,",
  " ": "\\:",
  "　": "\\quad{}",
};

/** Whether an `<m:t>` explicitly asks XML readers to preserve edge spaces. */
function preservesXmlSpace(textNode: XmlElement): boolean {
  return (
    textNode.getAttribute("xml:space") === "preserve" ||
    textNode.getAttributeNS(XML_NS, "space") === "preserve"
  );
}

/** Translate one literal math run's text to LaTeX symbol-by-symbol.
 *
 * A single internal ASCII blank remains a non-rendering TeX token separator.
 * Repeated blanks, tabs, and every blank in an `xml:space="preserve"` run become visible
 * math spacing commands; otherwise a Word piecewise formula such as
 * `0    x∉E` collapses to `0x∉E` when KaTeX ignores the plain spaces.
 */
function mathTextToLatex(text: string, preserveAll = false): string {
  let out = "";
  for (let index = 0; index < text.length;) {
    const ch = text[index];
    if (ch === " " || ch === "\t") {
      const start = index;
      while (
        index < text.length &&
        (text[index] === " " || text[index] === "\t")
      ) {
        index += 1;
      }
      const raw = text.slice(start, index);
      const semantic = raw.length > 1 || raw.includes("\t") || preserveAll;
      if (semantic) {
        for (const blank of raw) out += blank === "\t" ? "\\quad{}" : "\\;";
      } else {
        out += " ";
      }
      continue;
    }
    index += 1;
    const mapped = MATH_UNICODE_TO_LATEX[ch];
    out += mapped ? `${mapped}${/^\\[A-Za-z]+$/.test(mapped) ? " " : ""}` : ch;
  }
  return out;
}

const MATH_SCRIPT_WRAPPERS: Record<string, string> = {
  script: "mathcal",
  "double-struck": "mathbb",
  fraktur: "mathfrak",
  "sans-serif": "mathsf",
  monospace: "mathtt",
};

/** Convert a complete `<m:r>`, including Office math run properties. */
function mathRunToLatex(run: XmlElement): string {
  const textNodes = childrenByName(run, "t");
  const rPr = childrenByName(run, "rPr")[0];
  if (rPr && childrenByName(rPr, "nor").length > 0) {
    const safe = textNodes
      .map((node) => node.textContent || "")
      .join("")
      .replaceAll("\\", "")
      .replaceAll("{", "(")
      .replaceAll("}", ")");
    return safe ? `\\text{${safe}}` : "";
  }

  const styleNode = rPr ? childrenByName(rPr, "sty")[0] : undefined;
  const scriptNode = rPr ? childrenByName(rPr, "scr")[0] : undefined;
  const style = styleNode ? mathAttr(styleNode, "val") : null;
  const script = scriptNode ? mathAttr(scriptNode, "val") : null;
  // Own keys only: a hostile m:scr value must not reach Object.prototype.
  const scriptWrapper =
    script && Object.prototype.hasOwnProperty.call(MATH_SCRIPT_WRAPPERS, script)
      ? MATH_SCRIPT_WRAPPERS[script]
      : undefined;
  const wrapper = script
    ? scriptWrapper
    : style === "b"
      ? "mathbf"
      : style === "bi"
        ? "boldsymbol"
        : style === "p"
          ? "mathrm"
          : null;
  const bold = style === "b" || style === "bi";

  return textNodes
    .map((node) => {
      const raw = node.textContent || "";
      const preserve = preservesXmlSpace(node);
      if (!wrapper) return mathTextToLatex(raw, preserve);
      let converted = "";
      for (const match of raw.matchAll(/[A-Za-z0-9]+|[^A-Za-z0-9]+/g)) {
        const token = match[0];
        if (/^[A-Za-z0-9]+$/.test(token)) {
          // A bold script run (𝓐) keeps both: \boldsymbol{\mathcal{A}}.
          const wrapped = `\\${wrapper}{${token}}`;
          converted += scriptWrapper && bold ? `\\boldsymbol{${wrapped}}` : wrapped;
          continue;
        }
        converted +=
          bold && !scriptWrapper
            ? boldGreekToLatex(token, preserve)
            : mathTextToLatex(token, preserve);
      }
      return converted;
    })
    .join("");
}

/** A bold run's Greek letters as `\boldsymbol{\theta}` (\mathbf takes no Greek). */
function boldGreekToLatex(text: string, preserve: boolean): string {
  let out = "";
  let pending = "";
  for (const character of text) {
    const mapped = MATH_UNICODE_TO_LATEX[character];
    if (/^[\u0370-\u03ff]$/u.test(character) && mapped?.startsWith("\\")) {
      if (pending) out += mathTextToLatex(pending, preserve);
      pending = "";
      out += `\\boldsymbol{${mapped}}`;
      continue;
    }
    pending += character;
  }
  return pending ? out + mathTextToLatex(pending, preserve) : out;
}

/** Read an OOXML-math attribute (e.g. `<m:chr m:val="∑"/>`). The value lives in
 * the MATH namespace (`m:val`), so the generic `attr()` (Word-ns aware only)
 * misses it; this checks the unqualified name AND the math-namespaced one. */
function mathAttr(node: Element, local: string): string | null {
  return node.getAttribute(local) || node.getAttributeNS(MATH_NS, local);
}

/** Concatenate the raw text of every `<m:t>` descendant, in document order —
 * the guaranteed-lossless fallback when structural conversion yields nothing. */
function ommlRawText(el: XmlElement): string {
  let buffer = "";
  for (const t of descendantsByName(el, "t")) buffer += t.textContent || "";
  return buffer;
}

/** True when a converted fragment is a single token (one char, or one `\cmd`)
 * that needs no `{}` wrapping when used as a sub/superscript or radical degree. */
function isMathAtom(fragment: string): boolean {
  const f = fragment.trim();
  if (f.length <= 1) return true;
  return /^\\[A-Za-z]+$/.test(f);
}

/** Wrap a converted fragment in braces unless it is already a single atom. */
function braceMath(fragment: string): string {
  const f = fragment.trim();
  if (f === "") return "{}";
  return isMathAtom(f) ? f : `{${f}}`;
}

const NATIVE_FUNCTION_APPLY_NAMES = new Set(["sin", "cos", "log"]);

function hasOnlyElementsAndWhitespace(node: XmlElement): boolean {
  return Array.from(node.childNodes).every(
    (child) =>
      child.nodeType === 1 ||
      (child.nodeType === 3 && !(child.textContent || "").trim()),
  );
}

function hasExactMathValue(el: XmlElement, value: string): boolean {
  const attributes = Array.from(el.attributes);
  return (
    attributes.length === 1 &&
    attributes[0]?.namespaceURI === MATH_NS &&
    attributes[0]?.localName === "val" &&
    attributes[0]?.value === value &&
    el.children.length === 0 &&
    Array.from(el.childNodes).every(
      (child) => child.nodeType === 3 && !(child.textContent || "").trim(),
    )
  );
}

function ommlSimpleNativeFunctionName(el: XmlElement): string | null {
  if (
    el.namespaceURI !== MATH_NS ||
    el.attributes.length !== 0 ||
    !hasOnlyElementsAndWhitespace(el)
  ) {
    return null;
  }
  const directChildren = Array.from(el.children);
  if (
    directChildren.length !== 1 ||
    directChildren[0]?.namespaceURI !== MATH_NS ||
    localName(directChildren[0]) !== "r"
  ) {
    return null;
  }
  const run = directChildren[0] as XmlElement;
  if (run.attributes.length !== 0 || !hasOnlyElementsAndWhitespace(run)) {
    return null;
  }
  const runChildren = Array.from(run.children);
  const textNodes = runChildren.filter(
    (child) => child.namespaceURI === MATH_NS && localName(child) === "t",
  ) as XmlElement[];
  const propertyNodes = runChildren.filter(
    (child) => child.namespaceURI === MATH_NS && localName(child) === "rPr",
  ) as XmlElement[];
  // Word writes a `w:rPr` font hint (Cambria Math) on every math run it saves,
  // including the function-name run. It carries no mathematical meaning, so a
  // single one between the math properties and the text keeps the run simple.
  const fontHintNodes = runChildren.filter(
    (child) => child.namespaceURI === WORD_NS && localName(child) === "rPr",
  ) as XmlElement[];
  if (
    textNodes.length !== 1 ||
    propertyNodes.length > 1 ||
    fontHintNodes.length > 1
  ) {
    return null;
  }
  if (
    runChildren.length !==
      textNodes.length + propertyNodes.length + fontHintNodes.length ||
    (propertyNodes.length === 1 && runChildren[0] !== propertyNodes[0]) ||
    runChildren.at(-1) !== textNodes[0]
  ) {
    return null;
  }
  const textNode = textNodes[0];
  if (
    textNode.attributes.length !== 0 ||
    textNode.children.length !== 0 ||
    Array.from(textNode.childNodes).some((child) => child.nodeType !== 3)
  ) {
    return null;
  }
  if (propertyNodes.length === 1) {
    const properties = propertyNodes[0];
    if (
      properties.attributes.length !== 0 ||
      !hasOnlyElementsAndWhitespace(properties)
    ) {
      return null;
    }
    const propertyChildren = Array.from(properties.children);
    const scriptNodes = propertyChildren.filter(
      (child) => child.namespaceURI === MATH_NS && localName(child) === "scr",
    ) as XmlElement[];
    const styleNodes = propertyChildren.filter(
      (child) => child.namespaceURI === MATH_NS && localName(child) === "sty",
    ) as XmlElement[];
    // The upright plain style is the whole identity of a function name. The
    // exporter spells it `<m:scr roman/><m:sty p/>`; Word itself omits `m:scr`
    // because `roman` is the ECMA-376 default (§22.1.2.98), so an absent script
    // is the same value. Anything else — bold, a second script, `m:nor` — is
    // deliberate styling that stays on the literal path.
    if (
      styleNodes.length !== 1 ||
      scriptNodes.length > 1 ||
      propertyChildren.length !== styleNodes.length + scriptNodes.length ||
      (scriptNodes.length === 1 &&
        !hasExactMathValue(scriptNodes[0], "roman")) ||
      !hasExactMathValue(styleNodes[0], "p")
    ) {
      return null;
    }
  }
  const rawName = textNode.textContent || "";
  return rawName === rawName.trim() && NATIVE_FUNCTION_APPLY_NAMES.has(rawName)
    ? rawName
    : null;
}

/** `m:count` is an ST_Integer255; larger grids stay generic. */
const NATIVE_MATRIX_MAX_COLUMNS = 255;
const NATIVE_MATRIX_MAX_CELLS = 1024;
/** Elements one matrix's cells may hold before it is read generically. */
const NATIVE_MATRIX_CELL_ELEMENT_BUDGET = 4096;
/** m:mcJc value -> LaTeX array column letter. */
const NATIVE_MATRIX_COLUMN_SPECS: ReadonlyMap<string, string> = new Map([
  ["left", "l"],
  ["center", "c"],
  ["right", "r"],
]);

/**
 * Accent characters the DOCX exporter writes, each with the macro the
 * exporter turns back into that same character (U+0305 serves both \bar and
 * \overline there).
 */
const OMML_ACCENT_MACROS: ReadonlyMap<string, string> = new Map([
  ["\u0302", "\\hat"],
  ["\u0303", "\\tilde"],
  ["\u0305", "\\overline"],
  ["\u20d7", "\\vec"],
  ["\u0307", "\\dot"],
  ["\u0308", "\\ddot"],
  ["\u20db", "\\dddot"],
  ["\u0301", "\\acute"],
  ["\u0300", "\\grave"],
  ["\u0306", "\\breve"],
  ["\u030c", "\\check"],
  ["\u030a", "\\mathring"],
]);

/** Accent characters other writers (Word itself) use for the same marks. */
const OMML_FOREIGN_ACCENT_MACROS: ReadonlyMap<string, string> = new Map([
  ["\u0304", "\\bar"],
  ["\u00af", "\\bar"],
  ["\u2192", "\\vec"],
  ["\u20d6", "\\overleftarrow"],
  ["\u20e1", "\\overleftrightarrow"],
  ["\u02c7", "\\check"],
  ["\u02d9", "\\dot"],
  ["^", "\\hat"],
  ["~", "\\tilde"],
]);

function hasOneMathValueAttribute(el: XmlElement): boolean {
  const attributes = Array.from(el.attributes);
  return (
    attributes.length === 1 &&
    attributes[0]?.namespaceURI === MATH_NS &&
    attributes[0]?.localName === "val"
  );
}

function exactMathValue(el: XmlElement): string | null {
  return hasOneMathValueAttribute(el)
    ? (el.attributes[0]?.value ?? null)
    : null;
}

function isEmptyMathElement(el: XmlElement, name: string): boolean {
  return (
    el.namespaceURI === MATH_NS &&
    localName(el) === name &&
    el.attributes.length === 0 &&
    el.children.length === 0 &&
    hasOnlyElementsAndWhitespace(el)
  );
}

function strictMathContainerChildren(
  el: XmlElement,
  name: string,
): XmlElement[] | null {
  if (
    el.namespaceURI !== MATH_NS ||
    localName(el) !== name ||
    el.attributes.length !== 0 ||
    !hasOnlyElementsAndWhitespace(el)
  ) {
    return null;
  }
  const children = Array.from(el.children) as XmlElement[];
  return children.every((child) => child.namespaceURI === MATH_NS)
    ? children
    : null;
}

function supportedMatrixValueProperty(el: XmlElement, name: string): boolean {
  if (el.namespaceURI !== MATH_NS || localName(el) !== name) return false;
  const value = exactMathValue(el);
  if (
    value === null ||
    el.children.length !== 0 ||
    !hasOnlyElementsAndWhitespace(el)
  ) {
    return false;
  }
  if (["chr", "begChr", "endChr", "sepChr"].includes(name)) {
    return value.length <= 2;
  }
  if (name === "sty") return ["p", "b", "i", "bi"].includes(value);
  if (name === "scr") {
    return [
      "roman",
      "script",
      "double-struck",
      "fraktur",
      "sans-serif",
      "monospace",
    ].includes(value);
  }
  if (name === "type") return value === "skw";
  if (name === "limLoc") return value === "subSup" || value === "undOvr";
  return value === "0" || value === "1";
}

function supportedMatrixPropertyContainer(
  el: XmlElement,
  name: string,
  orderedProperties: readonly string[],
): boolean {
  const children = strictMathContainerChildren(el, name);
  if (!children) return false;
  let lastIndex = -1;
  const seen = new Set<string>();
  for (const child of children) {
    const childName = localName(child);
    const propertyIndex = orderedProperties.indexOf(childName);
    if (propertyIndex < lastIndex || propertyIndex < 0 || seen.has(childName)) {
      return false;
    }
    seen.add(childName);
    lastIndex = propertyIndex;
    if (childName === "ctrlPr") {
      if (!isEmptyMathElement(child, "ctrlPr")) return false;
    } else if (!supportedMatrixValueProperty(child, childName)) {
      return false;
    }
  }
  return true;
}

type ExactMatrixPlainRun = {
  text: string;
};

function exactMatrixPlainRun(run: XmlElement): ExactMatrixPlainRun | null {
  const children = strictMathContainerChildren(run, "r");
  if (!children || children.length !== 1) return null;
  const text = children.at(-1);
  if (!text || localName(text) !== "t") return null;
  const textAttributes = Array.from(text.attributes);
  if (
    text.namespaceURI !== MATH_NS ||
    text.children.length !== 0 ||
    (text.textContent || "").length > MAX_MATH_LATEX_CHARS ||
    !Array.from(text.childNodes).every((child) => child.nodeType === 3) ||
    (textAttributes.length !== 0 &&
      !(
        textAttributes.length === 1 &&
        textAttributes[0]?.namespaceURI === XML_NS &&
        textAttributes[0]?.localName === "space" &&
        textAttributes[0]?.value === "preserve"
      ))
  ) {
    return null;
  }
  return {
    text: text.textContent || "",
  };
}

// A property-free native run is only an ordinary matrix cell when its source
// text cannot be reinterpreted by the browser LaTeX writer as structure. This
// check intentionally uses the original m:t spelling: checking the converted
// LaTeX would be too late because a literal `\frac` is indistinguishable from
// an imported native fraction at that point.
const AMBIGUOUS_MATRIX_LITERAL_TEX = /[\\{}_^&%#$~]/u;
const NON_ORDINARY_MATRIX_UNICODE_PROJECTIONS = new Set([
  "\\sqrt",
  "\\;",
  "\\;\\;",
  "\\quad{}",
  "\\,",
  "\\phantom{0}",
  "\\:",
]);

function matrixPlainRunIsUnambiguous(run: ExactMatrixPlainRun): boolean {
  if (
    AMBIGUOUS_MATRIX_LITERAL_TEX.test(run.text) ||
    /[\t\r\n]/u.test(run.text)
  ) {
    return false;
  }
  return !Array.from(run.text).some((character) => {
    // N-ary glyphs project to commands the writer promotes to m:nary. The
    // radical token and Office-width spaces likewise do not return to the same
    // property-free m:r/m:t. Greek, relation, arrow, and other mappings remain
    // admitted because the writer's inverse table emits the original glyph as
    // an ordinary run.
    const projected = MATH_UNICODE_TO_LATEX[character];
    return (
      NATIVE_MATRIX_NARY_LIMIT_LOCATIONS.has(character) ||
      (projected !== undefined &&
        NON_ORDINARY_MATRIX_UNICODE_PROJECTIONS.has(projected))
    );
  });
}

function supportedMatrixRun(run: XmlElement): boolean {
  const plainRun = exactMatrixPlainRun(run);
  if (plainRun) return matrixPlainRunIsUnambiguous(plainRun);
  return supportedMatrixStyledRun(run);
}

/** Text of an `m:t` that holds only text (xml:space="preserve" allowed). */
function exactMatrixRunText(text: XmlElement): string | null {
  const textAttributes = Array.from(text.attributes);
  if (
    text.namespaceURI !== MATH_NS ||
    localName(text) !== "t" ||
    text.children.length !== 0 ||
    (text.textContent || "").length > MAX_MATH_LATEX_CHARS ||
    !Array.from(text.childNodes).every((child) => child.nodeType === 3) ||
    (textAttributes.length !== 0 &&
      !(
        textAttributes.length === 1 &&
        textAttributes[0]?.namespaceURI === XML_NS &&
        textAttributes[0]?.localName === "space" &&
        textAttributes[0]?.value === "preserve"
      ))
  ) {
    return null;
  }
  return text.textContent || "";
}

const MATRIX_STYLED_RUN_SCRIPTS: ReadonlyMap<string, RegExp> = new Map([
  ["script", /^[A-Za-z]+$/u],
  ["fraktur", /^[A-Za-z]+$/u],
  ["double-struck", /^[A-Za-z0-9]+$/u],
  ["sans-serif", /^[A-Za-z0-9]+$/u],
  ["monospace", /^[A-Za-z0-9]+$/u],
]);

/**
 * A styled run exactly as the DOCX exporter writes `\text{…}` (m:nor),
 * `\mathrm` / `\mathbf` / `\boldsymbol` (m:sty) and the math alphabets
 * (m:scr + m:sty p). Each imports to LaTeX the exporter turns back into the
 * same single run, so a matrix holding one stays an exact fixed point.
 */
function supportedMatrixStyledRun(run: XmlElement): boolean {
  const children = strictMathContainerChildren(run, "r");
  if (
    !children ||
    children.length !== 2 ||
    localName(children[0]) !== "rPr" ||
    localName(children[1]) !== "t"
  ) {
    return false;
  }
  const text = exactMatrixRunText(children[1]);
  const properties = strictMathContainerChildren(children[0], "rPr");
  if (text === null || !text || !properties) return false;
  const names = properties.map((property) => localName(property)).join(" ");
  if (names === "nor") {
    return (
      isEmptyMathElement(properties[0], "nor") &&
      !AMBIGUOUS_MATRIX_LITERAL_TEX.test(text) &&
      !/[\t\r\n]/u.test(text)
    );
  }
  if (names === "sty") {
    const style = exactMathValue(properties[0]);
    return (
      (style === "p" || style === "b" || style === "bi") &&
      hasExactMathValue(properties[0], style) &&
      /^[A-Za-z0-9]+$/u.test(text)
    );
  }
  if (names === "scr sty") {
    const script = exactMathValue(properties[0]);
    const letters = script ? MATRIX_STYLED_RUN_SCRIPTS.get(script) : undefined;
    return Boolean(
      script &&
      letters &&
      hasExactMathValue(properties[0], script) &&
      hasExactMathValue(properties[1], "p") &&
      letters.test(text),
    );
  }
  return false;
}

const NATIVE_MATRIX_NARY_LIMIT_LOCATIONS = new Map<string, string>([
  ["∑", "undOvr"],
  ["∏", "undOvr"],
  ["∐", "subSup"],
  ["∫", "subSup"],
  ["∬", "subSup"],
  ["∭", "subSup"],
  ["∮", "subSup"],
  ["⋀", "subSup"],
  ["⋁", "subSup"],
  ["⋂", "subSup"],
  ["⋃", "subSup"],
  ["⨀", "subSup"],
  ["⨁", "subSup"],
  ["⨂", "subSup"],
]);

function matrixContainsAmbiguousLiteralRun(matrix: XmlElement): boolean {
  return descendantsByName(matrix, "r").some((run) => {
    const plainRun = exactMatrixPlainRun(run);
    return Boolean(plainRun && !matrixPlainRunIsUnambiguous(plainRun));
  });
}

function ambiguousMatrixLiteralRunToLatex(text: string): string {
  // Preserve the exact plain-run glyphs without handing TeX active characters
  // back to MathLive/KaTeX as syntax. Control symbols cover the ordinary TeX
  // metacharacters; the two active characters without control-symbol forms use
  // renderer-supported text-ascii controls that the DOCX writer maps back to a
  // property-free m:r/m:t. Structural Unicode glyphs are wrapped in \mathrm so
  // the writer emits the literal glyph rather than promoting it to m:nary/rad.
  // Safe ordinary characters stay direct instead of becoming fragile one-char
  // \text groups (notably, \text{_} and \text{%} are invalid/truncating input).
  const literalControls: Readonly<Record<string, string>> = {
    "\\": "\\backslash ",
    "{": "\\{",
    "}": "\\}",
    _: "\\_",
    "%": "\\%",
    $: "\\$",
    "&": "\\&",
    "#": "\\#",
    "^": "\\unicode{x005E}",
    "~": "\\unicode{x007E}",
  };
  return Array.from(text)
    .map((character) => {
      const literalControl = literalControls[character];
      if (literalControl) return literalControl;
      const projected = MATH_UNICODE_TO_LATEX[character];
      if (
        projected !== undefined &&
        NON_ORDINARY_MATRIX_UNICODE_PROJECTIONS.has(projected) &&
        /\s/u.test(character)
      ) {
        return `\\unicode{x${character
          .codePointAt(0)!
          .toString(16)
          .toUpperCase()}}`;
      }
      if (
        NATIVE_MATRIX_NARY_LIMIT_LOCATIONS.has(character) ||
        (projected !== undefined &&
          NON_ORDINARY_MATRIX_UNICODE_PROJECTIONS.has(projected))
      ) {
        return `\\mathrm{${character}}`;
      }
      if (
        character === " " ||
        character === "\t" ||
        character === "\r" ||
        character === "\n"
      ) {
        return "\\ ";
      }
      return character;
    })
    .join("");
}

function supportedCanonicalMatrixRadicalProperties(
  properties: XmlElement,
  degreePresent: boolean,
): boolean {
  const children = strictMathContainerChildren(properties, "radPr");
  return Boolean(
    children &&
    children.length === 2 &&
    localName(children[0]) === "degHide" &&
    hasExactMathValue(children[0], degreePresent ? "0" : "1") &&
    isEmptyMathElement(children[1], "ctrlPr"),
  );
}

function supportedCanonicalMatrixNaryProperties(
  properties: XmlElement,
): boolean {
  const children = strictMathContainerChildren(properties, "naryPr");
  if (
    !children ||
    children.length !== 3 ||
    localName(children[0]) !== "chr" ||
    localName(children[1]) !== "limLoc" ||
    !isEmptyMathElement(children[2], "ctrlPr")
  ) {
    return false;
  }
  const character = exactMathValue(children[0]);
  const limitLocation = exactMathValue(children[1]);
  return (
    character !== null &&
    limitLocation !== null &&
    NATIVE_MATRIX_NARY_LIMIT_LOCATIONS.get(character) === limitLocation
  );
}

function supportedMatrixSlot(
  slot: XmlElement,
  name: string,
  depth: number,
  budget: { remaining: number },
): boolean {
  const children = strictMathContainerChildren(slot, name);
  return (
    children !== null &&
    children.every((child) =>
      ommlMatrixCellContentSupported(child, depth + 1, budget),
    )
  );
}

function ommlMatrixCellContentSupported(
  el: XmlElement,
  depth: number,
  budget: { remaining: number },
): boolean {
  if (depth > 32 || budget.remaining-- <= 0 || el.namespaceURI !== MATH_NS) {
    return false;
  }
  const children = strictMathContainerChildren(el, localName(el));
  const slot = (child: XmlElement | undefined, name: string) =>
    Boolean(child && supportedMatrixSlot(child, name, depth, budget));
  const emptyProperties = (child: XmlElement | undefined, name: string) =>
    Boolean(child && supportedMatrixPropertyContainer(child, name, ["ctrlPr"]));

  switch (localName(el)) {
    case "r":
      return supportedMatrixRun(el);
    case "f": {
      if (!children) return false;
      const offset = children[0] && localName(children[0]) === "fPr" ? 1 : 0;
      if (children.length !== offset + 2) return false;
      if (
        offset === 1 &&
        !supportedMatrixPropertyContainer(children[0], "fPr", [
          "type",
          "ctrlPr",
        ])
      ) {
        return false;
      }
      return slot(children[offset], "num") && slot(children[offset + 1], "den");
    }
    case "sSup":
    case "sSub":
    case "sSubSup":
    case "sPre": {
      if (!children) return false;
      const name = localName(el);
      const propertyName = `${name}Pr`;
      const offset =
        children[0] && localName(children[0]) === propertyName ? 1 : 0;
      const slotNames =
        name === "sSup"
          ? ["e", "sup"]
          : name === "sSub"
            ? ["e", "sub"]
            : ["sub", "sup", "e"];
      if (name === "sSubSup") slotNames.splice(0, 3, "e", "sub", "sup");
      if (children.length !== offset + slotNames.length) return false;
      if (offset === 1 && !emptyProperties(children[0], propertyName))
        return false;
      return slotNames.every((slotName, index) =>
        slot(children[offset + index], slotName),
      );
    }
    case "rad": {
      if (
        !children ||
        children.length !== 3 ||
        localName(children[0]) !== "radPr"
      ) {
        return false;
      }
      const degree = strictMathContainerChildren(children[1], "deg");
      if (degree === null || !slot(children[1], "deg")) return false;
      const degreePresent =
        ommlChildrenToLatex(children[1], depth + 1).trim().length > 0;
      return (
        supportedCanonicalMatrixRadicalProperties(children[0], degreePresent) &&
        slot(children[2], "e")
      );
    }
    case "nary": {
      if (
        !children ||
        children.length !== 4 ||
        localName(children[0]) !== "naryPr" ||
        !supportedCanonicalMatrixNaryProperties(children[0])
      ) {
        return false;
      }
      const operand = strictMathContainerChildren(children[3], "e");
      if (operand === null || operand.length !== 0) return false;
      return (
        slot(children[1], "sub") &&
        slot(children[2], "sup") &&
        slot(children[3], "e")
      );
    }
    // The general delimiter importer intentionally flattens m:d to literal
    // delimiters. Until that path preserves scalable identity, admitting it
    // here would make an authenticated matrix lose m:d on the next export.
    case "d":
      return false;
    case "func": {
      if (!children) return false;
      const offset = children[0] && localName(children[0]) === "funcPr" ? 1 : 0;
      if (children.length !== offset + 2) return false;
      if (offset === 1 && !emptyProperties(children[0], "funcPr")) return false;
      return (
        ommlSimpleNativeFunctionName(children[offset]) !== null &&
        slot(children[offset + 1], "e")
      );
    }
    case "acc": {
      // An accent the exporter writes back with the same character.
      if (
        !children ||
        children.length !== 2 ||
        localName(children[0]) !== "accPr" ||
        !supportedMatrixPropertyContainer(children[0], "accPr", [
          "chr",
          "ctrlPr",
        ])
      ) {
        return false;
      }
      const character = childrenByName(children[0], "chr")[0];
      return (
        character !== undefined &&
        OMML_ACCENT_MACROS.has(exactMathValue(character) ?? "") &&
        slot(children[1], "e")
      );
    }
    // The current browser exporter serializes canonical lower/upper limits as
    // ordinary scripts, not m:limLow/m:limUpp. Keep the whole matrix generic.
    case "limLow":
    case "limUpp":
      return false;
    default:
      return false;
  }
}

/**
 * Each column's justification of an authenticated matrix: no m:mPr (all
 * centred), or an m:mcs whose m:mc groups are exactly what the DOCX exporter
 * writes — adjacent columns of one m:mcJc in one group, groups alternating,
 * counts summing to the column count. Null for anything else.
 */
function ommlMatrixColumnJustifications(
  matrixProperties: XmlElement | undefined,
  columnCount: number,
): string[] | null {
  if (!matrixProperties) return Array(columnCount).fill("center");
  const properties = strictMathContainerChildren(matrixProperties, "mPr");
  if (!properties) return null;
  if (properties.length === 0) return Array(columnCount).fill("center");
  if (properties.length !== 1) return null;
  const columns = strictMathContainerChildren(properties[0], "mcs");
  if (!columns || columns.length === 0) return null;
  const justifications: string[] = [];
  for (const column of columns) {
    const columnChildren = strictMathContainerChildren(column, "mc");
    const columnProperties =
      columnChildren?.length === 1
        ? strictMathContainerChildren(columnChildren[0], "mcPr")
        : null;
    if (
      !columnProperties ||
      columnProperties.length !== 2 ||
      localName(columnProperties[0]) !== "count" ||
      localName(columnProperties[1]) !== "mcJc" ||
      columnProperties.some(
        (property) =>
          property.children.length !== 0 ||
          !hasOnlyElementsAndWhitespace(property),
      )
    ) {
      return null;
    }
    const count = exactMathValue(columnProperties[0]);
    const justification = exactMathValue(columnProperties[1]);
    if (
      count === null ||
      !/^[1-9]\d{0,2}$/u.test(count) ||
      justification === null ||
      !NATIVE_MATRIX_COLUMN_SPECS.has(justification) ||
      justifications.at(-1) === justification ||
      justifications.length + Number(count) > columnCount
    ) {
      return null;
    }
    for (let index = 0; index < Number(count); index += 1) {
      justifications.push(justification);
    }
  }
  return justifications.length === columnCount ? justifications : null;
}

interface NativeOmmlMatrix {
  rows: string[][];
  justifications: string[];
}

/**
 * Read an m:m only when re-export writes the same matrix back: rectangular,
 * exporter-shaped column properties, and cells whose every element is one the
 * exporter reproduces exactly (see ommlMatrixCellContentSupported).
 */
function readNativeOmmlMatrix(
  matrix: XmlElement,
  depth: number,
): NativeOmmlMatrix | null {
  if (
    matrix.namespaceURI !== MATH_NS ||
    localName(matrix) !== "m" ||
    matrix.attributes.length !== 0 ||
    !hasOnlyElementsAndWhitespace(matrix)
  ) {
    return null;
  }
  const children = Array.from(matrix.children) as XmlElement[];
  const properties =
    children[0]?.namespaceURI === MATH_NS && localName(children[0]) === "mPr"
      ? children.shift()
      : undefined;
  const columnCount = children[0]?.children.length ?? 0;
  if (
    children.length === 0 ||
    columnCount === 0 ||
    columnCount > NATIVE_MATRIX_MAX_COLUMNS ||
    children.length * columnCount > NATIVE_MATRIX_MAX_CELLS
  ) {
    return null;
  }
  const justifications = ommlMatrixColumnJustifications(
    properties,
    columnCount,
  );
  if (!justifications) return null;

  const rows: string[][] = [];
  const budget = { remaining: NATIVE_MATRIX_CELL_ELEMENT_BUDGET };
  for (const row of children) {
    if (
      row.namespaceURI !== MATH_NS ||
      localName(row) !== "mr" ||
      row.attributes.length !== 0 ||
      !hasOnlyElementsAndWhitespace(row)
    ) {
      return null;
    }
    const cells = Array.from(row.children) as XmlElement[];
    if (cells.length !== columnCount) return null;
    const converted: string[] = [];
    for (const cell of cells) {
      if (
        cell.namespaceURI !== MATH_NS ||
        localName(cell) !== "e" ||
        cell.attributes.length !== 0 ||
        !hasOnlyElementsAndWhitespace(cell) ||
        !Array.from(cell.children).every((child) =>
          ommlMatrixCellContentSupported(
            child as XmlElement,
            depth + 1,
            budget,
          ),
        )
      ) {
        return null;
      }
      const latex = ommlChildrenToLatex(cell, depth + 1).trim();
      if (
        latex.length > MAX_MATH_LATEX_CHARS ||
        latex.includes("&") ||
        latex.includes("\\\\")
      ) {
        return null;
      }
      converted.push(latex);
    }
    rows.push(converted);
  }
  return { rows, justifications };
}

/** `\begin{matrix}` (all centred), `\begin{array}{lcr}`, or cases. */
function nativeOmmlMatrixLatex(
  matrix: NativeOmmlMatrix,
  environment?: "cases" | "rcases",
): string {
  const body = matrix.rows.map((row) => row.join("&")).join("\\\\");
  if (environment) return `\\begin{${environment}}${body}\\end{${environment}}`;
  if (matrix.justifications.every((justification) => justification === "center")) {
    return `\\begin{matrix}${body}\\end{matrix}`;
  }
  const spec = matrix.justifications
    .map((justification) => NATIVE_MATRIX_COLUMN_SPECS.get(justification))
    .join("");
  return `\\begin{array}{${spec}}${body}\\end{array}`;
}

function ommlNativeMatrixToLatex(
  matrix: XmlElement,
  depth: number,
): string | null {
  const native = readNativeOmmlMatrix(matrix, depth);
  return native ? nativeOmmlMatrixLatex(native) : null;
}

/** Fence characters the DOCX exporter maps its `\left…\right` tokens back to. */
const NATIVE_MATRIX_FENCE_CHARACTERS = new Set([
  "",
  "(",
  ")",
  "[",
  "]",
  "{",
  "}",
  "|",
  "‖",
  "〈",
  "〉",
  "⌊",
  "⌋",
  "⌈",
  "⌉",
]);

function ommlNativeMatrixDelimiterToLatex(
  delimiter: XmlElement,
  depth: number,
): string | null {
  if (
    delimiter.namespaceURI !== MATH_NS ||
    delimiter.attributes.length !== 0 ||
    !hasOnlyElementsAndWhitespace(delimiter)
  ) {
    return null;
  }
  const children = Array.from(delimiter.children) as XmlElement[];
  if (
    children.length !== 2 ||
    localName(children[0]) !== "dPr" ||
    localName(children[1]) !== "e" ||
    children.some((child) => child.namespaceURI !== MATH_NS) ||
    children[1].attributes.length !== 0 ||
    !hasOnlyElementsAndWhitespace(children[1])
  ) {
    return null;
  }
  const matrices = Array.from(children[1].children) as XmlElement[];
  if (matrices.length !== 1 || localName(matrices[0]) !== "m") return null;
  const matrix = readNativeOmmlMatrix(matrices[0], depth + 1);
  if (!matrix) return null;

  const delimiterProperties = children[0];
  if (
    delimiterProperties.attributes.length !== 0 ||
    !hasOnlyElementsAndWhitespace(delimiterProperties)
  ) {
    return null;
  }
  const propertyChildren = Array.from(
    delimiterProperties.children,
  ) as XmlElement[];
  if (
    (propertyChildren.length !== 2 && propertyChildren.length !== 3) ||
    propertyChildren.some((child) => child.namespaceURI !== MATH_NS) ||
    localName(propertyChildren[0]) !== "begChr" ||
    localName(propertyChildren[1]) !== "endChr" ||
    (propertyChildren.length === 3 &&
      !isEmptyMathElement(propertyChildren[2], "ctrlPr"))
  ) {
    return null;
  }
  const begin = exactMathValue(propertyChildren[0]);
  const end = exactMathValue(propertyChildren[1]);
  if (
    begin === null ||
    end === null ||
    !NATIVE_MATRIX_FENCE_CHARACTERS.has(begin) ||
    !NATIVE_MATRIX_FENCE_CHARACTERS.has(end)
  ) {
    return null;
  }
  // A brace on one side of left-justified value/condition columns is how
  // the exporter (and Word) write cases.
  const leftJustified = matrix.justifications.every(
    (justification) => justification === "left",
  );
  if (leftJustified && matrix.justifications.length <= 2) {
    if (begin === "{" && end === "") return nativeOmmlMatrixLatex(matrix, "cases");
    if (begin === "" && end === "}") return nativeOmmlMatrixLatex(matrix, "rcases");
  }
  return (
    `\\left${mathDelimiterToLatexToken(begin)}` +
    `${nativeOmmlMatrixLatex(matrix)}` +
    `\\right${mathDelimiterToLatexToken(end)}`
  );
}

function mathDelimiterToLatexToken(value: string): string {
  if (!value) return ".";
  if (value === "{") return "\\{";
  if (value === "}") return "\\}";
  return mathTextToLatex(value).trim() || ".";
}

/** Characters TeX can stretch with `\middle` between `\left…\right`. */
const MATH_MIDDLE_DELIMITERS = new Set([
  "|",
  "‖",
  "/",
  "(",
  ")",
  "[",
  "]",
  "{",
  "}",
  "〈",
  "〉",
  "⟨",
  "⟩",
  "⌊",
  "⌋",
  "⌈",
  "⌉",
]);

/** A fence glyph as literal LaTeX: braces escaped so the group stays balanced. */
function mathDelimiterGlyphToLatex(value: string): string {
  if (value === "{") return "\\{";
  if (value === "}") return "\\}";
  return mathTextToLatex(value);
}

/** An argument's math content, without its argPr/ctrlPr properties. */
function mathContentChildren(argument: XmlElement | undefined): XmlElement[] {
  if (!argument) return [];
  return (Array.from(argument.children) as XmlElement[]).filter(
    (child) => localName(child) !== "argPr" && localName(child) !== "ctrlPr",
  );
}

/** An ST_OnOff math property: present and not switched off. */
function ommlMathFlagOn(
  properties: XmlElement | undefined,
  name: string,
): boolean {
  const flag = properties ? childrenByName(properties, name)[0] : undefined;
  if (!flag) return false;
  const value = mathAttr(flag, "val");
  return !value || value === "1" || value === "true" || value === "on";
}

const MATH_RELATION_START =
  /^(?:[=<>]|\\(?:le|ge|leq|geq|ne|neq|approx|equiv|sim|simeq|cong|propto|to|in|ni|subset|subseteq|supset|supseteq|Rightarrow|Leftarrow|Leftrightarrow|coloneqq|ll|gg)(?![A-Za-z]))/u;

/**
 * A brace on one side of an equation array whose rows hold a value and a
 * condition is Word's own cases layout. Rows aligned on a relation are a
 * system of equations instead, and stay an aligned array.
 */
function ommlEquationArrayCasesToLatex(
  equationArray: XmlElement,
  depth: number,
  environment: "cases" | "rcases",
  escapeAmbiguousMatrixRuns: boolean,
): string | null {
  // Cells split at "&" not after a backslash (tested in code: Safari before
  // 16.4 cannot compile a lookbehind).
  const rows = childrenByName(equationArray, "e").map((row) =>
    splitBehind(ommlChildrenToLatex(row, depth + 1, escapeAmbiguousMatrixRuns), /&/u, { char: /\\/, not: true })
      .map((cell) => cell.trim()),
  );
  if (
    rows.length === 0 ||
    rows.some((cells) => cells.length > 2) ||
    rows.some((cells) => cells.length === 2 && MATH_RELATION_START.test(cells[1]))
  ) {
    return null;
  }
  const body = rows.map((cells) => cells.join("&")).join("\\\\");
  return `\\begin{${environment}}${body}\\end{${environment}}`;
}

function ommlDelimiterToLatex(
  el: XmlElement,
  depth: number,
  preserveScalableIdentity = false,
  escapeAmbiguousMatrixRuns = false,
): string {
  const dPr = childrenByName(el, "dPr")[0];
  const begEl = dPr ? childrenByName(dPr, "begChr")[0] : undefined;
  const endEl = dPr ? childrenByName(dPr, "endChr")[0] : undefined;
  const sepEl = dPr ? childrenByName(dPr, "sepChr")[0] : undefined;
  const beg = begEl ? (mathAttr(begEl, "val") ?? "(") : "(";
  const end = endEl ? (mathAttr(endEl, "val") ?? ")") : ")";
  const sep = sepEl ? (mathAttr(sepEl, "val") ?? "|") : "|";
  const entries = childrenByName(el, "e");
  const content = entries.length === 1 ? mathContentChildren(entries[0]) : [];
  const tallContent =
    content.length === 1 &&
    (localName(content[0]) === "m" || localName(content[0]) === "eqArr");
  if (tallContent && localName(content[0]) === "eqArr") {
    const environment =
      beg === "{" && end === "" ? "cases" : beg === "" && end === "}" ? "rcases" : null;
    const cases = environment
      ? ommlEquationArrayCasesToLatex(
          content[0],
          depth + 1,
          environment,
          escapeAmbiguousMatrixRuns,
        )
      : null;
    if (cases) return cases;
  }
  const pieces = entries.map((entry) =>
    ommlChildrenToLatex(entry, depth + 1, escapeAmbiguousMatrixRuns),
  );
  const separated = pieces.length > 1 && MATH_MIDDLE_DELIMITERS.has(sep);
  // A one-sided or empty delimiter has no plain-glyph spelling: "(" alone
  // reads back as text, not as Word's fence with an invisible side.
  const oneSided = beg === "" || end === "";
  const empty = pieces.every((piece) => piece.trim() === "");
  // A matrix or equation array keeps Word's stretching fence, and so does a
  // separated delimiter (its bar is a \middle between \left and \right).
  if (preserveScalableIdentity || tallContent || separated || oneSided || empty) {
    const inner = separated
      ? pieces.join(`\\middle${mathDelimiterToLatexToken(sep)} `)
      : pieces.join(mathTextToLatex(sep));
    return `\\left${mathDelimiterToLatexToken(beg)}${inner}\\right${mathDelimiterToLatexToken(end)}`;
  }
  const inner = pieces.join(mathTextToLatex(sep));
  return `${mathDelimiterGlyphToLatex(beg)}${inner}${mathDelimiterGlyphToLatex(end)}`;
}

/** Stretchy arrows Word draws on the baseline with a label (m:groupChr). */
const MATH_X_ARROW_MACROS: ReadonlyMap<string, string> = new Map([
  ["→", "\\xrightarrow"],
  ["←", "\\xleftarrow"],
  ["↔", "\\xleftrightarrow"],
  ["⇒", "\\xRightarrow"],
  ["⇐", "\\xLeftarrow"],
  ["⇔", "\\xLeftrightarrow"],
  ["↦", "\\xmapsto"],
]);

/** Stretchy arrows over or under text on the baseline: [over, under]. */
const MATH_OVER_UNDER_ARROW_MACROS: ReadonlyMap<
  string,
  readonly [string, string]
> = new Map([
  ["→", ["\\overrightarrow", "\\underrightarrow"]],
  ["←", ["\\overleftarrow", "\\underleftarrow"]],
  ["↔", ["\\overleftrightarrow", "\\underleftrightarrow"]],
]);

/** Operator names whose m:limLow limit sits under them. */
const MATH_LIMIT_BASE_MACROS: ReadonlyMap<string, string> = new Map([
  ["lim", "\\lim"],
  ["limsup", "\\limsup"],
  ["liminf", "\\liminf"],
  ["max", "\\max"],
  ["min", "\\min"],
  ["sup", "\\sup"],
  ["inf", "\\inf"],
  ["argmax", "\\operatorname*{argmax}"],
  ["argmin", "\\operatorname*{argmin}"],
]);

/** (character, its position, the part on the baseline) of an m:groupChr. */
function ommlGroupCharacterShape(groupCharacter: XmlElement): {
  character: string;
  position: string;
  baseline: string;
} {
  const properties = childrenByName(groupCharacter, "groupChrPr")[0];
  const value = (name: string, fallback: string): string => {
    const property = properties ? childrenByName(properties, name)[0] : undefined;
    return (property && mathAttr(property, "val")) || fallback;
  };
  // ECMA-376 §22.1.2.41: an underbrace below, on the baseline, by default.
  return {
    character: value("chr", "⏟"),
    position: value("pos", "bot"),
    baseline: value("vertJc", "bot"),
  };
}

function ommlGroupCharacterToLatex(
  el: XmlElement,
  depth: number,
  escapeAmbiguousMatrixRuns: boolean,
): string {
  const { character, position, baseline } = ommlGroupCharacterShape(el);
  const base = ommlChildContainer(el, "e", depth, escapeAmbiguousMatrixRuns);
  if (character === "⏟" || character === "︸") return `\\underbrace{${base}}`;
  if (character === "⏞" || character === "︷") return `\\overbrace{${base}}`;
  const xArrow = MATH_X_ARROW_MACROS.get(character);
  if (xArrow && position === baseline) {
    // The arrow on the baseline with its label above (or only below).
    return position === "bot" ? `${xArrow}{${base}}` : `${xArrow}[${base}]{}`;
  }
  const overUnder = MATH_OVER_UNDER_ARROW_MACROS.get(character);
  if (overUnder) {
    return `${position === "top" ? overUnder[0] : overUnder[1]}{${base}}`;
  }
  const decoration = mathTextToLatex(character).trim() || character;
  return `${position === "top" ? "\\overset" : "\\underset"}{${decoration}}{${base}}`;
}

function ommlLowerLimitToLatex(
  el: XmlElement,
  depth: number,
  escapeAmbiguousMatrixRuns: boolean,
): string {
  const content = mathContentChildren(childrenByName(el, "e")[0]);
  const lim = ommlChildContainer(el, "lim", depth, escapeAmbiguousMatrixRuns);
  if (content.length === 1 && localName(content[0]) === "groupChr") {
    // A brace or a stretchy arrow carrying a label below it.
    const { character, position, baseline } = ommlGroupCharacterShape(
      content[0],
    );
    const body = ommlChildContainer(
      content[0],
      "e",
      depth + 1,
      escapeAmbiguousMatrixRuns,
    );
    if (character === "⏟" || character === "︸") {
      return `\\underbrace{${body}}_{${lim}}`;
    }
    const xArrow = MATH_X_ARROW_MACROS.get(character);
    if (xArrow && position === "bot" && baseline === "bot") {
      return `${xArrow}[${lim}]{${body}}`;
    }
  }
  const limitBase =
    content.length === 1 && localName(content[0]) === "r"
      ? MATH_LIMIT_BASE_MACROS.get(ommlRawText(content[0]).trim())
      : undefined;
  if (limitBase) return `${limitBase}_{${lim}}`;
  const base = ommlChildContainer(el, "e", depth, escapeAmbiguousMatrixRuns);
  // Word's empty lower-limit template keeps both slots.
  if (!lim && !base.trim()) return "\\underset{}{}";
  return lim ? `\\underset{${lim}}{${base}}` : base;
}

function ommlUpperLimitToLatex(
  el: XmlElement,
  depth: number,
  escapeAmbiguousMatrixRuns: boolean,
): string {
  const content = mathContentChildren(childrenByName(el, "e")[0]);
  const lim = ommlChildContainer(el, "lim", depth, escapeAmbiguousMatrixRuns);
  if (content.length === 1 && localName(content[0]) === "groupChr") {
    const { character } = ommlGroupCharacterShape(content[0]);
    if (character === "⏞" || character === "︷") {
      const body = ommlChildContainer(
        content[0],
        "e",
        depth + 1,
        escapeAmbiguousMatrixRuns,
      );
      return `\\overbrace{${body}}^{${lim}}`;
    }
  }
  const base = ommlChildContainer(el, "e", depth, escapeAmbiguousMatrixRuns);
  // Word's empty upper-limit template keeps both slots.
  if (!lim && !base.trim()) return "\\overset{}{}";
  return lim ? `\\overset{${lim}}{${base}}` : base;
}

/**
 * An equation array: one row per m:e. Word reads an "&" in a row's text as
 * an alignment mark, LaTeX's column separator, so rows that carry one are an
 * aligned environment and the others a gathered one.
 */
function ommlEquationArrayToLatex(
  el: XmlElement,
  depth: number,
  escapeAmbiguousMatrixRuns: boolean,
): string {
  const rows = childrenByName(el, "e");
  const aligned = rows.some((row) =>
    childrenByName(row, "r").some((run) => ommlRawText(run).includes("&")),
  );
  const environment = aligned ? "aligned" : "gathered";
  const body = rows
    .map((row) =>
      ommlChildrenToLatex(row, depth + 1, escapeAmbiguousMatrixRuns).trim(),
    )
    .join("\\\\");
  return `\\begin{${environment}}${body}\\end{${environment}}`;
}

/**
 * Convert one OMML element subtree to LaTeX. Dispatches on the OOXML math
 * element's local name; unknown elements fall through to converting their
 * children (so their run text is never lost). `depth` bounds recursion.
 */
function ommlElementToLatex(
  el: XmlElement,
  depth: number,
  escapeAmbiguousMatrixRuns = false,
): string {
  if (depth > MAX_MATH_DEPTH) return ommlRawText(el);
  const name = localName(el);
  switch (name) {
    case "t":
      return mathTextToLatex(el.textContent || "", preservesXmlSpace(el));
    case "r": {
      const plainRun = exactMatrixPlainRun(el);
      if (
        escapeAmbiguousMatrixRuns &&
        plainRun &&
        !matrixPlainRunIsUnambiguous(plainRun)
      ) {
        return ambiguousMatrixLiteralRunToLatex(plainRun.text);
      }
      return mathRunToLatex(el);
    }
    case "f": {
      // Fraction: <m:f><m:num/><m:den/></m:f>.
      const num = ommlChildContainer(
        el,
        "num",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const den = ommlChildContainer(
        el,
        "den",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const properties = childrenByName(el, "fPr")[0];
      const type = properties
        ? childrenByName(properties, "type")[0]
        : undefined;
      // Word stores a diagonal/skewed fraction as the same m:f container with
      // an explicit m:fPr/m:type="skw" discriminator. Preserve that native
      // identity with the exact canonical spelling the existing DOCX exporter
      // recognises; absent/ordinary/unknown types retain the long-standing
      // stacked-fraction import path below.
      if (type && mathAttr(type, "val") === "skw") {
        return `{${num}}\\!/\\!{${den}}`;
      }
      // A bar-less fraction is Word's binomial-coefficient stack (its own
      // "Binomial Theorem" gallery entry is built from one). Importing it as an
      // ordinary `\frac` drew a rule Word never draws, so keep the native
      // identity with the spelling both the DOCX exporter and the server-side
      // OMML reader already recognise.
      if (type && mathAttr(type, "val") === "noBar") {
        return `\\genfrac{}{}{0pt}{}{${num}}{${den}}`;
      }
      return `\\frac{${num}}{${den}}`;
    }
    case "sSup": {
      // Superscript: base <m:e>, exponent <m:sup>.
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sup = ommlChildContainer(
        el,
        "sup",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      return `${braceMath(base)}^${braceMath(sup)}`;
    }
    case "sSub": {
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sub = ommlChildContainer(
        el,
        "sub",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      return `${braceMath(base)}_${braceMath(sub)}`;
    }
    case "sSubSup": {
      // Combined sub+superscript (e.g. summation-style x_i^2).
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sub = ommlChildContainer(
        el,
        "sub",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sup = ommlChildContainer(
        el,
        "sup",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      return `${braceMath(base)}_${braceMath(sub)}^${braceMath(sup)}`;
    }
    case "sPre": {
      // Pre-sub/superscript: keep every slot explicitly grouped, including
      // single atoms. The DOCX exporter uses this exact canonical shape to
      // distinguish native m:sPre from an ordinary empty-base sSubSup.
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sub = ommlChildContainer(
        el,
        "sub",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sup = ommlChildContainer(
        el,
        "sup",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      return `{}_{${sub}}^{${sup}}{${base}}`;
    }
    case "rad": {
      // Radical: <m:deg/> (degree, optional) + <m:e/> (radicand). An empty
      // degree → \sqrt{e}; a present degree → \sqrt[deg]{e}.
      const deg = ommlChildContainer(
        el,
        "deg",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const radicand = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      return deg ? `\\sqrt[${deg}]{${radicand}}` : `\\sqrt{${radicand}}`;
    }
    case "nary": {
      // N-ary operator (∑ ∏ ∫ …) with optional lower (m:sub) / upper (m:sup)
      // limits and an operand (m:e). The operator char itself is in
      // <m:naryPr><m:chr m:val="…"/>; default is ∫ when absent (per OOXML).
      const naryPr = childrenByName(el, "naryPr")[0];
      const chrEl = naryPr ? childrenByName(naryPr, "chr")[0] : undefined;
      const chr = (chrEl && mathAttr(chrEl, "val")) || "∫";
      const op = mathTextToLatex(chr).trim() || chr;
      const sub = ommlChildContainer(
        el,
        "sub",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const sup = ommlChildContainer(
        el,
        "sup",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const operand = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      let out = op;
      // An explicit limit placement other than the exporter's default for
      // this operator (limits under/over ∑ ∏, beside the others) is written
      // as \limits / \nolimits so it survives the next export.
      const limLocEl = naryPr ? childrenByName(naryPr, "limLoc")[0] : undefined;
      const limLoc = limLocEl ? mathAttr(limLocEl, "val") : null;
      const defaultLimLoc = chr === "∑" || chr === "∏" ? "undOvr" : "subSup";
      if (
        (limLoc === "undOvr" || limLoc === "subSup") &&
        limLoc !== defaultLimLoc &&
        /^\\[A-Za-z]+$/u.test(op)
      ) {
        out += limLoc === "undOvr" ? "\\limits" : "\\nolimits";
      }
      if (sub) out += `_${braceMath(sub)}`;
      if (sup) out += `^${braceMath(sup)}`;
      // A space separates the operator (with its limits) from the operand.
      return operand ? `${out} ${operand}` : out;
    }
    case "d": {
      const matrix = ommlNativeMatrixDelimiterToLatex(el, depth);
      if (matrix) return matrix;
      return ommlDelimiterToLatex(el, depth, false, escapeAmbiguousMatrixRuns);
    }
    case "m": {
      const matrix = ommlNativeMatrixToLatex(el, depth);
      if (matrix) return matrix;
      if (matrixContainsAmbiguousLiteralRun(el)) {
        return ommlChildrenToLatex(el, depth, true);
      }
      return ommlChildrenToLatex(el, depth, escapeAmbiguousMatrixRuns);
    }
    case "func": {
      // Named function apply: <m:fName/> (the name, e.g. sin) + <m:e/> (arg).
      const fnameElement = childrenByName(el, "fName")[0];
      const argumentElement = childrenByName(el, "e")[0];
      const fname = fnameElement
        ? ommlChildrenToLatex(
            fnameElement,
            depth + 1,
            escapeAmbiguousMatrixRuns,
          )
        : "";
      let arg = argumentElement
        ? ommlChildrenToLatex(
            argumentElement,
            depth + 1,
            escapeAmbiguousMatrixRuns,
          )
        : "";
      const rawName = fnameElement
        ? ommlSimpleNativeFunctionName(fnameElement)
        : null;
      if (argumentElement && rawName) {
        const argumentChildren = Array.from(argumentElement.children);
        if (
          argumentChildren.length === 1 &&
          localName(argumentChildren[0]) === "d"
        ) {
          arg = ommlDelimiterToLatex(
            argumentChildren[0] as XmlElement,
            depth + 1,
            true,
            escapeAmbiguousMatrixRuns,
          );
        }
        return `\\${rawName}{${arg}}`;
      }
      return arg ? `${fname} ${arg}` : fname;
    }
    case "limLow":
      return ommlLowerLimitToLatex(el, depth, escapeAmbiguousMatrixRuns);
    case "limUpp":
      return ommlUpperLimitToLatex(el, depth, escapeAmbiguousMatrixRuns);
    case "eqArr":
      return ommlEquationArrayToLatex(el, depth, escapeAmbiguousMatrixRuns);
    case "groupChr":
      return ommlGroupCharacterToLatex(el, depth, escapeAmbiguousMatrixRuns);
    case "acc": {
      const properties = childrenByName(el, "accPr")[0];
      const characterElement = properties
        ? childrenByName(properties, "chr")[0]
        : undefined;
      // ECMA-376 §22.1.2.20: an absent m:chr is the combining circumflex.
      const character =
        (characterElement && mathAttr(characterElement, "val")) || "̂";
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const macro =
        OMML_ACCENT_MACROS.get(character) ??
        OMML_FOREIGN_ACCENT_MACROS.get(character);
      if (macro) return `${macro}{${base}}`;
      const decoration = mathTextToLatex(character).trim() || character;
      return `\\overset{${decoration}}{${base}}`;
    }
    case "bar": {
      const properties = childrenByName(el, "barPr")[0];
      const positionElement = properties
        ? childrenByName(properties, "pos")[0]
        : undefined;
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      // ECMA-376 §22.1.2.7: the bar sits below unless m:pos says top.
      return positionElement && mathAttr(positionElement, "val") === "top"
        ? `\\overline{${base}}`
        : `\\underline{${base}}`;
    }
    case "borderBox": {
      const properties = childrenByName(el, "borderBoxPr")[0];
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const rising = ommlMathFlagOn(properties, "strikeBLTR");
      const falling = ommlMathFlagOn(properties, "strikeTLBR");
      // A borderless box with a diagonal strike is how \cancel is written.
      if (
        (rising || falling) &&
        ["hideTop", "hideBot", "hideLeft", "hideRight"].every((side) =>
          ommlMathFlagOn(properties, side),
        ) &&
        !ommlMathFlagOn(properties, "strikeH") &&
        !ommlMathFlagOn(properties, "strikeV")
      ) {
        const macro =
          rising && falling ? "\\xcancel" : rising ? "\\cancel" : "\\bcancel";
        return `${macro}{${base}}`;
      }
      return `\\boxed{${base}}`;
    }
    case "phant": {
      const properties = childrenByName(el, "phantPr")[0];
      const base = ommlChildContainer(
        el,
        "e",
        depth,
        escapeAmbiguousMatrixRuns,
      );
      const hidden =
        properties !== undefined &&
        childrenByName(properties, "show").length > 0 &&
        !ommlMathFlagOn(properties, "show");
      // A shown phantom is just its contents (with adjusted spacing).
      if (!hidden) return base;
      const noWidth = ommlMathFlagOn(properties, "zeroWid");
      const noHeight =
        ommlMathFlagOn(properties, "zeroAsc") &&
        ommlMathFlagOn(properties, "zeroDesc");
      if (noWidth && !noHeight) return `\\vphantom{${base}}`;
      if (noHeight && !noWidth) return `\\hphantom{${base}}`;
      return `\\phantom{${base}}`;
    }
    case "rPr":
    case "ctrlPr":
    case "fPr":
    case "naryPr":
    case "radPr":
    case "dPr":
    case "sSupPr":
    case "sSubPr":
    case "sSubSupPr":
    case "funcPr":
    case "sPrePr":
    case "accPr":
    case "barPr":
    case "limLowPr":
    case "limUppPr":
    case "groupChrPr":
    case "borderBoxPr":
    case "phantPr":
    case "eqArrPr":
    case "mPr":
      // Property elements carry no visible content.
      return "";
    default:
      // Unknown / unsupported construct (matrix, accent, box, equation array,
      // grouped char, …): transparently convert its children so nested run text
      // and any recognized sub-constructs still surface. Never drops text.
      return ommlChildrenToLatex(el, depth, escapeAmbiguousMatrixRuns);
  }
}

/** Convert every element child of an OMML node in document order, joined. */
function ommlChildrenToLatex(
  el: XmlElement,
  depth: number,
  escapeAmbiguousMatrixRuns = false,
): string {
  if (depth > MAX_MATH_DEPTH) return ommlRawText(el);
  let out = "";
  for (const child of Array.from(el.children)) {
    out += ommlElementToLatex(child, depth + 1, escapeAmbiguousMatrixRuns);
  }
  return out;
}

/** Convert the named single argument container (m:num/m:den/m:e/m:sup/…). The
 * first matching child is used; multiple (forged) duplicates are ignored. */
function ommlChildContainer(
  el: XmlElement,
  childName: string,
  depth: number,
  escapeAmbiguousMatrixRuns = false,
): string {
  const container = childrenByName(el, childName)[0];
  return container
    ? ommlChildrenToLatex(container, depth + 1, escapeAmbiguousMatrixRuns)
    : "";
}

/**
 * Convert an `<m:oMath>` zone to the LaTeX an editor math node stores. Returns
 * "" only when the zone is genuinely empty. A zone whose structural conversion
 * collapses to nothing (all-unsupported constructs with no text) falls back to
 * its raw concatenated text so the equation is never lost; the result is
 * whitespace-normalized and clamped.
 */
export function ommlToLatex(omath: XmlElement): string {
  let latex = ommlChildrenToLatex(omath, 0).replace(/\s+/g, " ").trim();
  if (latex === "") {
    // Nothing structural converted — degrade to the literal symbols.
    latex = ommlRawText(omath).replace(/\s+/g, " ").trim();
  }
  return latex.slice(0, MAX_MATH_LATEX_CHARS);
}

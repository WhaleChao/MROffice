/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { transform } from "sucrase";
import type * as TS from "typescript";
import {
  normalizeCodeWorkspacePath,
  codeWorkspacePathIdentity,
} from "./code-workspace";
import { AWAIT_COMPLETION_SOURCE_PREFIX } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import type { CodeTestHarness, CodeTestSource } from "./code-test-explorer";

const MODULE_LIMIT = 16;
const SOURCE_LIMIT = 1_000_000;

export interface CodeTestModulePlan {
  readonly sources: readonly CodeTestSource[];
  bundle(harness: CodeTestHarness): CodeTestHarness;
}

function relativePath(from: string, specifier: string): string {
  if (!specifier.startsWith("./") && !specifier.startsWith("../")) {
    throw new Error(
      `Only local relative module imports are available: ${specifier}`,
    );
  }
  const parts = from.split("/").slice(0, -1);
  for (const part of specifier.split("/")) {
    if (part === ".") continue;
    if (part === "..") {
      if (!parts.length)
        throw new Error("Module import leaves the notebook workspace");
      parts.pop();
    } else parts.push(part);
  }
  const result = normalizeCodeWorkspacePath(parts.join("/"));
  if (!result) throw new Error(`Invalid local module path: ${specifier}`);
  return result;
}

/** Parse actual import syntax, not occurrences in comments or strings. The
 * compiler is already the Code editor's lazy JS/TS intelligence dependency. */
function moduleSpecifiers(
  ts: typeof TS,
  source: CodeTestSource,
): { specs: string[]; module: boolean } {
  const file = ts.createSourceFile(
    source.path,
    source.source,
    ts.ScriptTarget.Latest,
    true,
    source.language === "typescript" ? ts.ScriptKind.TS : ts.ScriptKind.JS,
  );
  const specs = new Set<string>();
  let usesModules = false;
  let localRequire = false;
  const findBindings = (node: TS.Node) => {
    if (
      (ts.isVariableDeclaration(node) ||
        ts.isParameter(node) ||
        ts.isBindingElement(node) ||
        ts.isFunctionDeclaration(node) ||
        ts.isFunctionExpression(node)) &&
      node.name &&
      ts.isIdentifier(node.name) &&
      node.name.text === "require"
    )
      localRequire = true;
    ts.forEachChild(node, findBindings);
  };
  findBindings(file);
  const add = (expression: TS.Expression | undefined) => {
    if (!expression || !ts.isStringLiteralLike(expression))
      throw new Error("Dynamic module imports are unavailable");
    specs.add(expression.text);
    usesModules = true;
  };
  const visit = (node: TS.Node, functionDepth: number) => {
    if (ts.isImportDeclaration(node)) {
      const clause = node.importClause;
      const names = clause?.namedBindings;
      const typeOnly =
        clause?.isTypeOnly ||
        (names &&
          ts.isNamedImports(names) &&
          !clause?.name &&
          names.elements.length > 0 &&
          names.elements.every((name) => name.isTypeOnly));
      if (!typeOnly) add(node.moduleSpecifier);
    } else if (ts.isExportDeclaration(node)) {
      usesModules = true;
      const typeOnly =
        node.isTypeOnly ||
        (node.exportClause &&
          ts.isNamedExports(node.exportClause) &&
          node.exportClause.elements.length > 0 &&
          node.exportClause.elements.every((name) => name.isTypeOnly));
      if (!typeOnly && node.moduleSpecifier) add(node.moduleSpecifier);
    } else if (ts.isImportEqualsDeclaration(node)) {
      if (
        !node.isTypeOnly &&
        ts.isExternalModuleReference(node.moduleReference)
      )
        add(node.moduleReference.expression);
    } else if (ts.isCallExpression(node)) {
      if (node.expression.kind === ts.SyntaxKind.ImportKeyword)
        throw new Error("Dynamic module imports are unavailable");
      if (
        !localRequire &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === "require"
      ) {
        if (node.arguments.length !== 1)
          throw new Error("A local require needs one literal module path");
        add(node.arguments[0]);
      }
    }
    if (
      (ts.isAwaitExpression(node) ||
        (ts.isForOfStatement(node) && node.awaitModifier)) &&
      functionDepth === 0
    ) {
      throw new Error(
        "Top-level await is unavailable in local test modules; use an async test function",
      );
    }
    if (
      ts.canHaveModifiers(node) &&
      ts
        .getModifiers(node)
        ?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword)
    )
      usesModules = true;
    const depth = functionDepth + (ts.isFunctionLike(node) ? 1 : 0);
    ts.forEachChild(node, (child) => visit(child, depth));
  };
  visit(file, 0);
  return { specs: [...specs], module: usesModules };
}

/** Snapshot-only packaging for the EXISTING credential-free execution provider.
 * There is no filesystem, URL loader, npm lookup or additional execution engine. */
export async function prepareCodeTestModules(
  entry: CodeTestSource,
  available: readonly CodeTestSource[],
): Promise<CodeTestModulePlan> {
  const passthrough = {
    sources: [entry],
    bundle: (harness: CodeTestHarness) => harness,
  };
  if (
    !["javascript", "typescript"].includes(entry.language) ||
    !/\b(?:import|export|require)\b/.test(entry.source)
  )
    return passthrough;
  const ts = await import("typescript");
  const byPath = new Map<string, CodeTestSource[]>();
  for (const source of available) {
    const path = normalizeCodeWorkspacePath(source.path);
    if (!path) continue;
    const key = codeWorkspacePathIdentity(path);
    byPath.set(key, [...(byPath.get(key) ?? []), source]);
  }
  const sources: CodeTestSource[] = [];
  const edges: Record<string, string>[] = [];
  const ids = new Map<string, number>();
  let size = 0;
  const resolve = (source: CodeTestSource, specifier: string) => {
    const relative = relativePath(source.path, specifier);
    const exact = byPath.get(codeWorkspacePathIdentity(relative));
    const names = /\.[^/]+$/.test(relative)
      ? relative.endsWith(".js")
        ? [relative.slice(0, -3) + ".ts"]
        : []
      : [
          relative + ".js",
          relative + ".ts",
          relative + "/index.js",
          relative + "/index.ts",
        ];
    const candidates =
      exact ??
      names.flatMap(
        (name) => byPath.get(codeWorkspacePathIdentity(name)) ?? [],
      );
    if (candidates.length !== 1)
      throw new Error(`Missing or ambiguous local module: ${specifier}`);
    return candidates[0];
  };
  const add = (source: CodeTestSource): number => {
    const existing = ids.get(source.pageId);
    if (existing !== undefined) {
      const captured = sources[existing];
      if (
        captured.path !== source.path ||
        captured.language !== source.language ||
        captured.contentVersion !== source.contentVersion ||
        captured.source !== source.source
      ) {
        throw new Error("Ambiguous local module identity");
      }
      return existing;
    }
    if (
      !["javascript", "typescript"].includes(source.language) ||
      !normalizeCodeWorkspacePath(source.path) ||
      sources.length >= MODULE_LIMIT ||
      (size += source.source.length) > SOURCE_LIMIT
    )
      throw new Error("Local test module scope is unsupported or too large");
    const id = sources.length;
    ids.set(source.pageId, id);
    sources.push(Object.freeze({ ...source }));
    edges.push(Object.create(null));
    const parsed = moduleSpecifiers(ts, source);
    for (const spec of parsed.specs)
      edges[id][spec] = String(add(resolve(source, spec)));
    return id;
  };
  const parsedEntry = moduleSpecifiers(ts, entry);
  if (!parsedEntry.module) return passthrough;
  add(entry);
  const compile = (source: string, language: string) =>
    transform(source, {
      transforms:
        language === "typescript" ? ["typescript", "imports"] : ["imports"],
      disableESTransforms: true,
    }).code;
  const dependencies = sources
    .slice(1)
    .map((source) => compile(source.source, source.language));
  return {
    sources: Object.freeze(sources),
    bundle(harness) {
      const compiled = [
        compile(harness.source, entry.language),
        ...dependencies,
      ];
      // The entry factory is first and starts on line1, preserving the existing
      // harness's entry-file line mapping. Dependency frames are beyond it.
      const factories = compiled
        .map(
          (code, index) =>
            `${JSON.stringify(String(index))}:function(module,exports,require){${code}\n}`,
        )
        .join(",");
      const loader = `const __mrai_test_edges=${JSON.stringify(edges)},__mrai_test_cache=Object.create(null);function __mrai_test_load(id){if(Object.hasOwn(__mrai_test_cache,id))return __mrai_test_cache[id];if(!Object.hasOwn(__mrai_test_factories,id))throw new Error("Unknown local module");const module={exports:Object.create(null),completion:undefined};__mrai_test_cache[id]=module;module.completion=__mrai_test_factories[id](module,module.exports,function(spec){const map=__mrai_test_edges[Number(id)];if(typeof spec!=="string"||!Object.hasOwn(map,spec))throw new Error("Module was not captured for this test: "+String(spec));return __mrai_test_load(map[spec]).exports;});return module;}`;
      const source = `${harness.execution === "async" ? AWAIT_COMPLETION_SOURCE_PREFIX : ""}const __mrai_test_factories={${factories}};${loader}return __mrai_test_load("0").completion;`;
      if (new TextEncoder().encode(source).length > 2_000_000)
        throw new Error("Compiled test module scope is too large");
      return { ...harness, source };
    },
  };
}

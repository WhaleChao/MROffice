/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Solver Parameters model (XLS-A073) — what the dialog holds, its Options,
// and its Load / Save Model form.
//
// Microsoft Support, "Define and solve a problem by using Solver": Load/Save
// asks for "a cell range for the model area"; Save writes the model down a
// vertical range from its first cell, Load reads the whole range back.
//
// The layout is Excel's own (read from a model the local Excel saved with
// SolverSave, and from Excel's Solver add-in, which reads it back), so a
// model saved here loads in Excel and one Excel saved loads here:
//
//   =MAX($B$5) | =MIN($B$5) | =$B$5=100      objective and goal
//   =COUNT($A$1:$A$3,$C$1)                   variable cells
//   =$C$1<=10  =$C$2>=$D$1  =$C$3=4          relational constraints
//   =$A$1=INT($A$1)                          int
//   =($C$1=0)+($C$1=1)=1                     bin
//   ={MaxTime,Iterations,Precision,IntegerOptimality/100,AssumeLinear,
//     ShowIterations,Scaling,1,Derivatives,1,Convergence,NonNegative}
//   =$A$1:$A$3=COUNT($A$1:$A$3)              dif (Excel keeps these after
//                                             the first options row)
//   ={MaxTime,Iterations,Engine,PopulationSize,RandomSeed,Multistart,
//     RequireBounds,MutationRate,MaxSubproblems,MaxFeasibleSolutions,
//     IgnoreIntegers,MaxTimeWithoutImprovement}
//
// A blank Max Time / Iterations is 32767 in the first row and 0 in the
// second; Engine is 0 GRG Nonlinear, 1 Simplex LP, 2 Evolutionary; Derivatives
// 1 Forward, 2 Central. A localized Excel writes int / bin as words
// (`=$A$1=整数`), which Load accepts too, as it does the layout MRAI wrote
// before (`=OR(A=0,A=1)`, `SUMPRODUCT(1/COUNTIF…)`, `={"grg",…}`).
// ---------------------------------------------------------------------------

export type SolverDialogOp = "<=" | ">=" | "==" | "int" | "bin" | "dif";
export type SolverDialogMethod = "grg" | "simplex" | "evolutionary";

export interface SolverDialogConstraint {
  cellRef: string;
  op: SolverDialogOp;
  /** The right-hand side (a number or a cell); empty for int / bin / dif. */
  bound: string;
}

/** Solver Options beyond precision / iterations / Ignore Integer
 *  Constraints (Microsoft Support, "Solver options"); text as typed. */
export interface SolverDialogExtraOptions {
  /** All Methods: Max Time (seconds); "" = no limit. */
  maxTime: string;
  scaling: boolean;
  showIterations: boolean;
  /** Integer Optimality (%). */
  integerOptimality: string;
  /** Evolutionary and Integer Constraints; "" = no limit. */
  maxSubproblems: string;
  maxFeasibleSolutions: string;
  /** GRG Nonlinear / Evolutionary. */
  convergence: string;
  derivatives: "forward" | "central";
  multistart: boolean;
  populationSize: string;
  randomSeed: string;
  requireBounds: boolean;
  mutationRate: string;
  maxTimeNoImprovement: string;
}

export interface SolverDialogModel {
  objective: string;
  goal: "max" | "min" | "value";
  target: string;
  changing: string;
  constraints: SolverDialogConstraint[];
  nonNegative: boolean;
  method: SolverDialogMethod;
  precision: string;
  iterations: string;
  ignoreIntegers: boolean;
  extra: SolverDialogExtraOptions;
}

/** Excel's defaults (Solver Options as a reset Solver shows them). */
export const DEFAULT_SOLVER_EXTRA_OPTIONS: SolverDialogExtraOptions = {
  maxTime: "",
  scaling: false,
  showIterations: false,
  integerOptimality: "1",
  maxSubproblems: "",
  maxFeasibleSolutions: "",
  convergence: "0.0001",
  derivatives: "forward",
  multistart: false,
  populationSize: "100",
  randomSeed: "0",
  requireBounds: true,
  mutationRate: "0.075",
  maxTimeNoImprovement: "30",
};

/** Excel's defaults: Max, non-negative on, GRG Nonlinear, precision 0.000001, 100 iterations. */
export const DEFAULT_SOLVER_OPTIONS = {
  nonNegative: true,
  method: "grg" as SolverDialogMethod,
  precision: "0.000001",
  iterations: "100",
  ignoreIntegers: false,
  extra: DEFAULT_SOLVER_EXTRA_OPTIONS,
};

const REF = "\\$?[A-Za-z]{1,3}\\$?\\d+(?::\\$?[A-Za-z]{1,3}\\$?\\d+)?";
const ENGINES: readonly SolverDialogMethod[] = ["grg", "simplex", "evolutionary"];
/** The words a localized Excel saves int / bin / dif constraints with. */
const INTEGER_WORDS = new Set(["int", "integer", "整数", "整數", "ganzzahl", "entier", "entero", "intero", "inteiro", "整数値"]);
const BINARY_WORDS = new Set(["bin", "binary", "二进制", "二進位", "binär", "binaire", "binario", "binário", "バイナリ"]);
const DIFFERENT_WORDS = new Set(["dif", "alldifferent", "alldiff", "所有不同", "全部不同", "全不同"]);

const bool = (value: boolean) => (value ? "TRUE" : "FALSE");
const or = (text: string, fallback: string) => (text.trim() === "" ? fallback : text.trim());

/** The formulas Save Model writes, top to bottom (Excel's layout). */
export function solverModelFormulas(model: SolverDialogModel): string[] {
  const extra = model.extra ?? DEFAULT_SOLVER_EXTRA_OPTIONS;
  const lines: string[] = [];
  const objective = model.objective.trim();
  lines.push(
    model.goal === "max" ? `=MAX(${objective})`
      : model.goal === "min" ? `=MIN(${objective})`
        : `=${objective}=${model.target.trim() || "0"}`,
  );
  lines.push(`=COUNT(${model.changing.trim()})`);
  const different: string[] = [];
  for (const constraint of model.constraints) {
    const ref = constraint.cellRef.trim();
    if (constraint.op === "int") lines.push(`=${ref}=INT(${ref})`);
    else if (constraint.op === "bin") lines.push(`=(${ref}=0)+(${ref}=1)=1`);
    else if (constraint.op === "dif") different.push(`=${ref}=COUNT(${ref})`);
    else lines.push(`=${ref}${constraint.op === "==" ? "=" : constraint.op}${constraint.bound.trim()}`);
  }
  const integerOptimality = Number(or(extra.integerOptimality, "1")) / 100;
  lines.push(
    `={${or(extra.maxTime, "32767")},${or(model.iterations, "32767")},${or(model.precision, DEFAULT_SOLVER_OPTIONS.precision)},` +
      `${Number.isFinite(integerOptimality) ? integerOptimality : 0.01},${bool(model.method === "simplex")},` +
      `${bool(extra.showIterations)},${bool(extra.scaling)},1,${extra.derivatives === "central" ? 2 : 1},1,` +
      `${or(extra.convergence, "0.0001")},${bool(model.nonNegative)}}`,
  );
  lines.push(...different);
  lines.push(
    `={${or(extra.maxTime, "0")},${or(model.iterations, "0")},${Math.max(0, ENGINES.indexOf(model.method))},` +
      `${or(extra.populationSize, "100")},${or(extra.randomSeed, "0")},${bool(extra.multistart)},` +
      `${bool(extra.requireBounds)},${or(extra.mutationRate, "0.075")},${or(extra.maxSubproblems, "0")},` +
      `${or(extra.maxFeasibleSolutions, "0")},${bool(model.ignoreIntegers)},${or(extra.maxTimeNoImprovement, "30")}}`,
  );
  return lines;
}

const blankAt = (value: string, blank: string) => (value === blank ? "" : value);

/** Read a model area back (Load Model), or null when it is not one. */
export function parseSolverModel(formulas: readonly string[]): SolverDialogModel | null {
  const lines = formulas.map((line) => line.trim()).filter((line) => line !== "");
  if (lines.length < 2) return null;
  const model: SolverDialogModel = {
    objective: "",
    goal: "max",
    target: "",
    changing: "",
    constraints: [],
    ...DEFAULT_SOLVER_OPTIONS,
    extra: { ...DEFAULT_SOLVER_EXTRA_OPTIONS },
  };
  const head = lines[0];
  let match = new RegExp(`^=(MAX|MIN)\\((${REF})\\)$`, "i").exec(head);
  if (match) {
    model.goal = match[1].toUpperCase() === "MAX" ? "max" : "min";
    model.objective = match[2];
  } else {
    match = new RegExp(`^=(${REF})=(.+)$`, "i").exec(head);
    if (!match) return null;
    model.goal = "value";
    model.objective = match[1];
    model.target = match[2];
  }
  match = /^=COUNT\((.+)\)$/i.exec(lines[1]);
  if (!match) return null;
  model.changing = match[1];
  const same = (a: string, b: string) => a.toUpperCase() === b.toUpperCase();
  let optionRows = 0;
  for (const line of lines.slice(2)) {
    // Excel's two options rows: an array constant, no "=" after the first.
    const array = /^=\{(.*)\}$/.exec(line);
    if (array && !array[1].includes("=")) {
      const values = array[1].split(",").map((value) => value.trim());
      const legacy = /^"(grg|simplex|evolutionary)"$/i.exec(values[0] ?? "");
      if (legacy && values.length === 5) {
        // The layout MRAI saved before: {"grg",nonNeg,precision,iterations,ignore}.
        model.method = legacy[1].toLowerCase() as SolverDialogMethod;
        model.nonNegative = values[1].toUpperCase() === "TRUE";
        model.precision = values[2];
        model.iterations = values[3];
        model.ignoreIntegers = values[4].toUpperCase() === "TRUE";
        continue;
      }
      if (values.length < 11) return null;
      if (optionRows === 0) {
        model.extra.maxTime = blankAt(values[0], "32767");
        model.iterations = blankAt(values[1], "32767") || model.iterations;
        model.precision = values[2];
        const gap = Number(values[3]);
        if (Number.isFinite(gap)) model.extra.integerOptimality = String(Math.round(gap * 100 * 1e6) / 1e6);
        if (values[4].toUpperCase() === "TRUE") model.method = "simplex";
        model.extra.showIterations = values[5].toUpperCase() === "TRUE";
        model.extra.scaling = values[6].toUpperCase() === "TRUE";
        model.extra.derivatives = values[8] === "2" ? "central" : "forward";
        model.extra.convergence = values[10];
        if (values[11] !== undefined) model.nonNegative = values[11].toUpperCase() === "TRUE";
      } else {
        const engine = Number(values[2]);
        if (Number.isInteger(engine) && ENGINES[engine]) model.method = ENGINES[engine];
        model.extra.populationSize = values[3];
        model.extra.randomSeed = values[4];
        model.extra.multistart = values[5].toUpperCase() === "TRUE";
        model.extra.requireBounds = values[6].toUpperCase() === "TRUE";
        model.extra.mutationRate = values[7];
        model.extra.maxSubproblems = blankAt(values[8], "0");
        model.extra.maxFeasibleSolutions = blankAt(values[9], "0");
        model.ignoreIntegers = values[10].toUpperCase() === "TRUE";
        if (values[11] !== undefined) model.extra.maxTimeNoImprovement = values[11];
      }
      optionRows += 1;
      continue;
    }
    const integer = new RegExp(`^=(${REF})=INT\\((${REF})\\)$`, "i").exec(line);
    if (integer && same(integer[1], integer[2])) {
      model.constraints.push({ cellRef: integer[1], op: "int", bound: "" });
      continue;
    }
    const binary = new RegExp(`^=\\((${REF})=0\\)\\+\\((${REF})=1\\)=1$`, "i").exec(line) ??
      new RegExp(`^=OR\\((${REF})=0,(${REF})=1\\)$`, "i").exec(line);
    if (binary && same(binary[1], binary[2])) {
      model.constraints.push({ cellRef: binary[1], op: "bin", bound: "" });
      continue;
    }
    const different = new RegExp(`^=(${REF})=COUNT\\((${REF})\\)$`, "i").exec(line);
    if (different && same(different[1], different[2])) {
      model.constraints.push({ cellRef: different[1], op: "dif", bound: "" });
      continue;
    }
    const legacyDifferent = new RegExp(`^=SUMPRODUCT\\(1/COUNTIF\\((${REF}),(${REF})\\)\\)=COUNT\\((${REF})\\)$`, "i").exec(line);
    if (legacyDifferent) {
      model.constraints.push({ cellRef: legacyDifferent[1], op: "dif", bound: "" });
      continue;
    }
    const worded = new RegExp(`^=(${REF})=([^=<>]+)$`, "i").exec(line);
    const word = worded?.[2].trim().toLowerCase() ?? "";
    if (worded && (INTEGER_WORDS.has(word) || BINARY_WORDS.has(word) || DIFFERENT_WORDS.has(word))) {
      model.constraints.push({
        cellRef: worded[1],
        op: INTEGER_WORDS.has(word) ? "int" : BINARY_WORDS.has(word) ? "bin" : "dif",
        bound: "",
      });
      continue;
    }
    const relation = new RegExp(`^=(${REF})(<=|>=|=)(.+)$`, "i").exec(line);
    if (!relation) return null;
    model.constraints.push({
      cellRef: relation[1],
      op: relation[2] === "=" ? "==" : (relation[2] as "<=" | ">="),
      bound: relation[3],
    });
  }
  return model.objective && model.changing ? model : null;
}

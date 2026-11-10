/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { createSelectionAggregateAccumulator, type SelectionAggregates } from "./grid-ops";
import type { SelectionObservationRange } from "./selection-observations";

export interface SelectionStatisticsSource {
  range: SelectionObservationRange;
  cells: Readonly<Record<string, unknown>>;
  getComputed(row:number,col:number): unknown;
  spilled: ReadonlyMap<string, unknown>;
  pivot?: ReadonlyMap<string, {value:unknown}> | null;
  /** A row a filter hides. Excel's status bar sums what the filter left
   * visible (a filtered table's total row and the status bar agree), so a
   * hidden row's cells are left out; rows hidden by hand are still counted. */
  isRowHidden?: (row:number)=>boolean;
}

/** Whether a stored cell holds a formula: its "" result is text Excel's
 * Count tallies, where an empty cell is not counted. */
function holdsFormula(cell:unknown):boolean{
  if(!cell||typeof cell!=="object")return false;
  const {formula,value}=cell as {formula?:unknown;value?:unknown};
  return (typeof formula==="string"&&formula!=="")||(typeof value==="string"&&value.startsWith("="));
}

/** The existing dense/sparse summary algorithm, with bounded preparation and
 * ordered accumulation. A cell's value is what the grid shows there: a pivot
 * output over a spilled element over the cell's own computed value, in both
 * the dense scan and the sparse one; merged cells follow getComputed. */
export function* selectionStatisticsSteps(source: SelectionStatisticsSource): Generator<void, SelectionAggregates> {
  const {range,cells,getComputed,spilled,pivot,isRowHidden}=source;
  const accumulator=createSelectionAggregateAccumulator();let work=0;
  const selected=(range.r2-range.r1+1)*(range.c2-range.c1+1);
  if(selected<=20_000){
    for(let row=range.r1;row<=range.r2;row++){
      if(isRowHidden?.(row))continue;
      for(let col=range.c1;col<=range.c2;col++){
        const key=`${row}:${col}`;
        const output=pivot?.get(key);
        accumulator.add(output?output.value:spilled.has(key)?spilled.get(key):getComputed(row,col),!output&&(spilled.has(key)||holdsFormula(cells[key])));
        if(++work%128===0)yield;
      }
    }
    return accumulator.result();
  }
  const coordinate=(key:string)=>{
    const separator=key.indexOf(":");if(separator<=0||separator===key.length-1)return null;
    const row=Number(key.slice(0,separator)),col=Number(key.slice(separator+1));
    return Number.isSafeInteger(row)&&Number.isSafeInteger(col)&&row>=range.r1&&row<=range.r2&&col>=range.c1&&col<=range.c2&&key===`${row}:${col}`&&!isRowHidden?.(row)?{row,col}:null;
  };
  // Map overwrite order is part of the existing contract: raw cells retain
  // their original insertion position when a spill or pivot replaces a value.
  const values=new Map<string,{value:unknown;formulaResult:boolean}>();
  for(const key in cells){
    if(!Object.prototype.hasOwnProperty.call(cells,key))continue;
    const at=coordinate(key);if(at)values.set(key,{value:getComputed(at.row,at.col),formulaResult:holdsFormula(cells[key])});
    if(++work%128===0)yield;
  }
  for(const [key,value] of spilled){
    if(coordinate(key))values.set(key,{value,formulaResult:true});
    if(++work%128===0)yield;
  }
  for(const [key,cell] of pivot??[]){
    if(coordinate(key))values.set(key,{value:cell.value,formulaResult:false});
    if(++work%128===0)yield;
  }
  for(const {value,formulaResult} of values.values()){
    accumulator.add(value,formulaResult);
    if(++work%128===0)yield;
  }
  return accumulator.result();
}

export function selectionStatisticsSync(source: SelectionStatisticsSource): SelectionAggregates {
  const steps=selectionStatisticsSteps(source);
  for(;;){const next=steps.next();if(next.done)return next.value;}
}

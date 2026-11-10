/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface SelectionObservationRange { r1:number; r2:number; c1:number; c2:number }
export interface SelectionObservations {
  readonly valid: boolean;
  /** Raw style identities, including an implicit default when any selected
   * coordinate is unmaterialized. Each family retains its own validation. */
  readonly styles: readonly unknown[];
  readonly locked: boolean | "mixed" | null;
}
const INVALID: SelectionObservations = Object.freeze({valid:false,styles:Object.freeze([]),locked:null});

/** Every stored cell of one map observed once, as counted identities so an
 * accepted edit can add or remove one cell's contribution exactly. `bounds`
 * covers the keys the large-selection walk would parse and may only grow after
 * removals, so a selection containing it observes exactly `count` cells. */
export interface StoredSelectionSummary {
  readonly styles: readonly (readonly [unknown, number])[];
  readonly locks: readonly (readonly [boolean, number])[];
  readonly malformedLock: number;
  readonly count: number;
  readonly bounds: SelectionObservationRange | null;
}

/** How one raw stored cell contributes: the same optional style read as the
 * original primary loops, and a lock class. A malformed cell does not
 * automatically revoke unrelated style families' authority. */
function classify(raw:unknown):{style:unknown;lock:boolean|"malformed"}{
  const style=(raw as {style?:unknown}|null|undefined)?.style;
  if(raw===undefined)return {style,lock:true};
  if(raw===null||typeof raw!=="object"||Array.isArray(raw)||![Object.prototype,null].includes(Object.getPrototypeOf(raw)))return {style,lock:"malformed"};
  const locked=(raw as {locked?:unknown}).locked;
  if(locked!==undefined&&typeof locked!=="boolean")return {style,lock:"malformed"};
  return {style,lock:locked!==false};
}
/** True when two raw cells contribute identically to a summary. */
export function sameObservation(a:unknown,b:unknown):boolean{
  const left=classify(a),right=classify(b);
  return left.style===right.style&&left.lock===right.lock;
}

/** The observation one stored cell contributes; shared by the range walk
 * below and by the per-map summary a cell index builds in its own pass. */
export function createSelectionObserver(seed?: StoredSelectionSummary) {
  const styles=new Map<unknown,number>(seed?.styles ?? []);
  const locks=new Map<boolean,number>(seed?.locks ?? []);
  let malformedLock=seed?.malformedLock ?? 0;
  const bump=<K,>(map:Map<K,number>,key:K,by:number)=>{
    const next=(map.get(key)??0)+by;
    if(next<=0)map.delete(key);else map.set(key,next);
  };
  const record=(raw:unknown,by:number)=>{
    const {style,lock}=classify(raw);
    bump(styles,style,by);
    if(lock==="malformed")malformedLock+=by;else bump(locks,lock,by);
  };
  return {
    observe(raw:unknown){record(raw,1);},
    unobserve(raw:unknown){record(raw,-1);},
    summary(count:number,bounds:SelectionObservationRange|null):StoredSelectionSummary{
      return Object.freeze({styles:Object.freeze([...styles.entries()]),locks:Object.freeze([...locks.entries()]),malformedLock,count,bounds});
    },
    result():SelectionObservations{
      return Object.freeze({valid:true,styles:Object.freeze([...styles.keys()]),locked:malformedLock>0?null:locks.size>1?"mixed":locks.keys().next().value??true});
    },
  };
}

/** One coordinate pass shared by the six primary formatting families. This
 * deliberately observes stored cells, including merged/hidden coordinates:
 * formatting and lock authority do not use the visible-row projection.
 * A large selection that contains every stored key reads the map's summary
 * instead of walking the map again. */
export function collectSelectionObservations(
  cells: Readonly<Record<string, unknown>>, range: SelectionObservationRange,
  rowCount:number, columnCount:number, summary?: StoredSelectionSummary,
): SelectionObservations {
  const {r1,r2,c1,c2}=range;
  if (![r1,r2,c1,c2].every(Number.isSafeInteger) || r1<0 || c1<0 || r2<r1 || c2<c1 || r2>=rowCount || c2>=columnCount) return INVALID;
  const selected=(r2-r1+1)*(c2-c1+1);
  const bounds=summary?.bounds;
  if(selected>20_000 && summary && (bounds===null || (bounds!==undefined && bounds.r1>=r1 && bounds.r2<=r2 && bounds.c1>=c1 && bounds.c2<=c2))){
    const seeded=createSelectionObserver(summary);
    if(summary.count<selected)seeded.observe(undefined);
    return seeded.result();
  }
  const observer=createSelectionObserver();
  const observe=(raw:unknown)=>observer.observe(raw);
  if(selected<=20_000){
    for(let row=r1;row<=r2;row++)for(let col=c1;col<=c2;col++)observe(cells[`${row}:${col}`]);
  } else {
    let materialized=0;
    for(const key in cells){
      if(!Object.prototype.hasOwnProperty.call(cells,key))continue;
      const separator=key.indexOf(":");if(separator<=0||separator===key.length-1)continue;
      const row=Number(key.slice(0,separator)),col=Number(key.slice(separator+1));
      // Preserve the existing primary parser; do not impose the stricter
      // canonical-key policy used separately by numeric selection statistics.
      if(!Number.isSafeInteger(row)||!Number.isSafeInteger(col)||row<r1||row>r2||col<c1||col>c2)continue;
      materialized++;observe(cells[key]);
    }
    if(materialized<selected)observe(undefined);
  }
  return observer.result();
}

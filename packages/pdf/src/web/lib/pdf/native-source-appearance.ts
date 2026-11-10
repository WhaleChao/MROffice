/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionProof, PdfNativeSourceTextAppearanceRun } from "./engine/native/list-text-fidelity";
import type { PdfPageSourceAppearance } from "./source-page-appearance";
import { matchPdfTextSourceAppearance } from "./source-text-appearance";
import { pdfNativeReaderSpelling } from "./native-text-reader-spelling";

/** Display color comes from the mature renderer. It never replaces raw PDF
 * color operands or upgrades the run's property mutation authority. */
export async function inheritPdfSourceAppearance(proof:PdfNativeListTextSelectionProof,
  readPage:(pageNumber:number)=>Promise<PdfPageSourceAppearance>):Promise<PdfNativeListTextSelectionProof>{
 const appearance=proof.sourceAppearance,selection=proof.evidence.selection;
 if(!selection)return proof;
 const scene=await readPage(selection.pageNumber);
 if(scene.sourceSha256!==selection.sourceSha256)throw new Error("PDF source appearance belongs to another source.");
 const [a,b,c,d,e,f]=scene.pageTransform,det=a*d-b*c;
 if(!Number.isFinite(det)||Math.abs(det)<1e-12)throw new Error("PDF source page transform is singular.");
 const toPdf=(point:readonly number[])=>[(d*(point[0]-e)-c*(point[1]-f))/det,(-b*(point[0]-e)+a*(point[1]-f))/det];
 // The box is edited in the reader's spelling when the source spells no
 // space where the reader reads one in a kerning gap: its caret geometry is
 // that spelling's, a gap standing for each such space.
 const reader=pdfNativeReaderSpelling(proof.evidence.originalDisplayText,selection.item.str)?.reader;
 const sourceGeometry=(reader?matchPdfTextSourceAppearance(scene,{...selection,item:{...selection.item,str:reader}},toPdf):null)??
  matchPdfTextSourceAppearance(scene,{...selection,
  item:{...selection.item,str:proof.evidence.originalDisplayText}},toPdf)??undefined;
 if(!sourceGeometry&&process.env.NODE_ENV!=="production"){
  // Why the line's glyphs were not read (dev builds): each reader item
  // against the page characters near its origin.
  const items=selection.items??[selection.item];
  console.debug("[pdf-text-prepare] geometry miss "+JSON.stringify({box:selection.item.str.slice(0,40),reader:reader?.slice(0,40)??null,
   items:items.map(item=>({str:item.str.slice(0,24),x:Math.round(item.transform[4]*100)/100,y:Math.round(item.transform[5]*100)/100,
    near:scene.characters.filter(character=>{const o=toPdf(character.origin);return Math.abs(o[0]-item.transform[4])<3&&Math.abs(o[1]-item.transform[5])<30;})
     .slice(0,4).map(character=>{const o=toPdf(character.origin);return [character.text,Math.round(o[0]*100)/100,Math.round(o[1]*100)/100,Math.round(character.fontSize*10)/10];})}))}));
 }
 const lineEvidence=proof.evidence.positionedLines?.lines.map(line=>line.evidence)??[proof.evidence];
 const leaves=lineEvidence.flatMap(line=>line.positionedRunGroup?.runs.map(run=>run.evidence.selection!)??[line.selection!]);
 if(!appearance?.length||leaves.length!==appearance.length)return {...proof,...(sourceGeometry?{sourceGeometry}:{})};
 const runs:PdfNativeSourceTextAppearanceRun[]=appearance.map((run,index)=>{
  if(!run.text.trim())return run;
  const matched=matchPdfTextSourceAppearance(scene,leaves[index],toPdf);
  const colors=matched?.characters.filter(character=>character.text.trim()).map(character=>character.color);
  if(!colors?.length||colors.some(color=>color.length!==3||color.some(value=>!Number.isFinite(value)||value<0||value>1))||
    colors.some(color=>color.some((value,i)=>Math.abs(value-colors[0][i])>1/255)))return run;
  return Object.freeze({...run,displayColor:Object.freeze({kind:"rgb" as const,components:Object.freeze([...colors[0]]) as readonly[number,number,number]})});
 });
 return Object.freeze({...proof,sourceAppearance:Object.freeze(runs),...(sourceGeometry?{sourceGeometry}:{})});
}

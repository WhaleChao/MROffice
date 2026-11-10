/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionEvidence, PdfNativeListTextFidelityEvidence } from "./engine/native/list-text-fidelity";
import type { PdfPageSourceAppearance } from "./source-page-appearance";

type SourceCharacter=PdfPageSourceAppearance["characters"][number];
/** v5 may redistribute replacement text across original font slots. Those
 * slots no longer have the old glyph origins. Match each proved visual row
 * from its original row origin, then read every candidate glyph's true quad. */
export function pdfNativeCandidateAppearanceSelection(evidence:PdfNativeListTextFidelityEvidence,
  outputSha256:string):PdfNativeListTextSelectionEvidence|null {
  const source=evidence.paragraphReflow?.replacementSelection??evidence.selection;
  if(!source)return null;
  const linePlan=evidence.linePlan;
  if(source.items&&linePlan?.lines.length&&linePlan.lines.every(line=>line.lineIndex!==undefined)) {
    const rows=new Map<number,{item:PdfNativeListTextSelectionEvidence["item"];text:string}>();
    for(const line of linePlan.lines){
      const index=line.lineIndex!;
      let row=rows.get(index);
      if(!row){
        const item=source.items.find(item=>item.itemIndex===line.itemIndex);
        if(!item)return null;
        row={item,text:""};rows.set(index,row);
      }
      row.text+=line.replacementText;
    }
    const items=[...rows.values()].map(row=>({...row.item,str:row.text.trimEnd()}));
    if(items.some(item=>!item.str))return null;
    return {...source,sourceSha256:outputSha256,item:{...source.item,str:evidence.replacementDisplayText},items};
  }
  const lines=evidence.positionedLines?.lines.map(line=>line.evidence.replacementDisplayText.trimEnd())??linePlan?.lines.map(line=>line.replacementText);
  const items=source.items&&lines?.length===source.items.length?source.items.map((item,index)=>({...item,str:lines[index]})):source.items;
  return {...source,sourceSha256:outputSha256,item:{...source.item,str:evidence.replacementDisplayText},...(items?{items}:{})};
}
export interface PdfTextSourceAppearance {
  readonly sourceSha256:string;
  readonly pageNumber:number;
  readonly text:string;
  /** Native PDF coordinates to the coordinates used by these quads. */
  readonly pageTransform:readonly number[];
  readonly characters:readonly SourceCharacter[];
  readonly bounds:Readonly<{left:number;top:number;right:number;bottom:number}>;
}
const sameField=(left:unknown,right:unknown):boolean=>left===right||(Array.isArray(left)&&Array.isArray(right)&&
 left.length===right.length&&left.every((value,index)=>value===right[index]));
/** The same glyph boxes, value for value: a geometry measured again from an
 * unchanged draft is this one, and keeping the one already held spares the
 * editor a render (several per key, each before the key's glyph painted). */
export function samePdfTextSourceAppearance(left:PdfTextSourceAppearance|null,right:PdfTextSourceAppearance|null):boolean{
 if(left===right)return true;
 if(!left||!right||left.sourceSha256!==right.sourceSha256||left.pageNumber!==right.pageNumber||left.text!==right.text||
  !sameField(left.pageTransform,right.pageTransform)||left.characters.length!==right.characters.length)return false;
 const a=left.bounds,b=right.bounds;
 if(a.left!==b.left||a.top!==b.top||a.right!==b.right||a.bottom!==b.bottom)return false;
 return left.characters.every((character,index)=>{
  const other=right.characters[index] as unknown as Record<string,unknown>;
  const fields=character as unknown as Record<string,unknown>;
  const keys=Object.keys(fields);
  return keys.length===Object.keys(other).length&&keys.every(key=>sameField(fields[key],other[key]));
 });
}
const sameBoxes=(left:readonly SourceCharacter[],right:readonly SourceCharacter[])=>left.length===right.length&&
 left.every((character,index)=>character.text===right[index].text&&character.quad.every((value,at)=>Math.abs(value-right[index].quad[at])<=.05+.02*Math.hypot(character.quad[0]-character.quad[4],character.quad[1]-character.quad[5])));
const blank=(value:string)=>/^\s+$/u.test(value);
/** Read one reader item out of the page's characters, from the character the
 * item starts on. The reader's text and the page's glyphs differ only in
 * white space and compatibility forms: a gap the reader reads as a space has
 * no glyph (it gets the gap between its neighbours as its box), a space glyph
 * the reader left out is passed over, any white-space glyph reads as the
 * reader's space, and a ligature glyph reads as the letters the reader
 * normalised it to. Anything else is not this item. */
function readItemCharacters(scene:readonly SourceCharacter[],start:number,text:string,offset:number,
 before?:SourceCharacter):{end:number;characters:SourceCharacter[]}|null{
 const characters:SourceCharacter[]=[];
 let at=start,position=0,last=start-1;
 while(position<text.length){
  const character=scene[at];
  if(character){
   const value=character.text;
   if(blank(value)){
    if(blank(text[position])){
     characters.push(Object.freeze({...character,text:text[position],start:offset+position,end:offset+position+1}));
     position+=1;last=at;at+=1;continue;
    }
    if(characters.length){at+=1;continue;}
    return null;
   }
   const form=text.startsWith(value,position)?value:value.normalize("NFKC");
   if(form&&text.startsWith(form,position)){
    characters.push(Object.freeze({...character,start:offset+position,end:offset+position+form.length}));
    position+=form.length;last=at;at+=1;continue;
   }
  }
  const previous=characters[characters.length-1]??before;
  if(!previous||!blank(text[position]))return null;
  // A space the reader read from a gap: its box is the gap, when the next
  // glyph follows on the same line; otherwise it has no width.
  const from=previous.quad,to=character?.quad;
  const height=Math.hypot(from[0]-from[4],from[1]-from[5]);
  const near=to!==undefined&&Math.hypot(to[0]-from[2],to[1]-from[3])<=Math.max(height*4,1);
  const quad=near?[from[2],from[3],to[0],to[1],from[6],from[7],to[4],to[5]]:[from[2],from[3],from[2],from[3],from[6],from[7],from[6],from[7]];
  characters.push(Object.freeze({...previous,text:text[position],start:offset+position,end:offset+position+1,
   origin:Object.freeze([previous.origin[0]+from[2]-from[0],previous.origin[1]+from[3]-from[1]]),quad:Object.freeze(quad)}));
  position+=1;
 }
 return characters.length?{end:last,characters}:null;
}
/** Map a renderer selection to actual source characters. No width scaling,
 * font substitution or writer permission is inferred by this read adapter. */
export function matchPdfTextSourceAppearance(scene:PdfPageSourceAppearance,selection:PdfNativeListTextSelectionEvidence,
  pagePointToPdf?:(point:readonly number[])=>readonly number[]):PdfTextSourceAppearance|null{
 if(scene.sourceSha256!==selection.sourceSha256||scene.pageNumber!==selection.pageNumber)return null;
 const [a,b,c,d,e,f]=scene.pageTransform,det=a*d-b*c;
 if(!Number.isFinite(det)||Math.abs(det)<1e-12)return null;
 const toPdf=pagePointToPdf??((point:readonly number[])=>[(d*(point[0]-e)-c*(point[1]-f))/det,(-b*(point[0]-e)+a*(point[1]-f))/det]);
 const items=selection.items??[selection.item],matched:SourceCharacter[]=[];
 let canonicalOffset=0,previousSceneEnd=-1;
 for(const item of items){
  const offset=selection.item.str.indexOf(item.str,canonicalOffset);if(offset<0)return null;
  const candidates:Array<{end:number;characters:SourceCharacter[]}>=[];
  // An item the reader starts with a space it took from the run before (a
  // raised run after a space: the item's origin carries the raised run's
  // rise, the space glyph sits on the baseline): the space may start it from
  // the same pen position on another baseline, within the item's height.
  const leadingBlank=blank(item.str[0]??"");
  const rowReach=Math.max(Number(item.height)||0,1);
  for(let start=previousSceneEnd+1;start<scene.characters.length;start++){
   const first=scene.characters[start],origin=toPdf(first.origin);
   if(origin.length!==2||Math.abs(origin[0]-item.transform[4])>.05)continue;
   if(Math.abs(origin[1]-item.transform[5])>.05&&!(leadingBlank&&blank(first.text)&&Math.abs(origin[1]-item.transform[5])<=rowReach))continue;
   const read=readItemCharacters(scene.characters,start,item.str,offset);
   if(read)candidates.push(read);
  }
  // An item the reader starts with a space read from a gap (the reader puts
  // its origin at the gap's start, where no glyph is): it starts on the next
  // glyph, on its baseline within its width, the space boxed in the gap from
  
  const before=matched[matched.length-1];
  if(!candidates.length&&leadingBlank&&before){
   let start=previousSceneEnd+1;
   while(start<scene.characters.length&&blank(scene.characters[start].text))start++;
   const first=scene.characters[start],origin=first?toPdf(first.origin):[];
   const along=(origin[0]??NaN)-item.transform[4];
   if(origin.length===2&&Math.abs(origin[1]-item.transform[5])<=.05&&along>=-.05&&along<=(Number(item.width)||0)+.05){
    const read=readItemCharacters(scene.characters,start,item.str,offset,before);
    if(read)candidates.push(read);
   }
  }
  // Text shown twice in one place (a drawn shadow or a faked bold) is one
  // reading: every candidate then has the same boxes, and the first stands.
  if(!candidates.length||candidates.slice(1).some(candidate=>!sameBoxes(candidate.characters,candidates[0].characters)))return null;
  matched.push(...candidates[0].characters);previousSceneEnd=candidates[candidates.length-1].end;canonicalOffset=offset+item.str.length;
 }
 if(!matched.length||canonicalOffset!==selection.item.str.length)return null;
 const xs=matched.flatMap(character=>character.quad.filter((_,i)=>i%2===0)),ys=matched.flatMap(character=>character.quad.filter((_,i)=>i%2===1));
 return Object.freeze({sourceSha256:scene.sourceSha256,pageNumber:scene.pageNumber,text:selection.item.str,pageTransform:scene.pageTransform,
  characters:Object.freeze(matched),bounds:Object.freeze({left:Math.min(...xs),top:Math.min(...ys),right:Math.max(...xs),bottom:Math.max(...ys)})});
}
/** Internal boundaries use the next source glyph's own edge, retaining Tc/TJ
 * placement instead of accumulating nominal browser character widths. */
export function pdfSourceCaretEdge(appearance:PdfTextSourceAppearance,offset:number):readonly [number,number,number,number]|null{
 const characters=appearance.characters;
 let low=0,high=characters.length;
 while(low<high){const middle=(low+high)>>>1;if(characters[middle].start<offset)low=middle+1;else high=middle;}
 const next=characters[low];
 low=0;high=characters.length;
 while(low<high){const middle=(low+high)>>>1;if(characters[middle].end<=offset)low=middle+1;else high=middle;}
 const previous=characters[low-1];
 // A virtual separator between source lines has no glyph. Its left boundary
 // remains at the preceding line's end rather than jumping to the next line.
 const atStart=Boolean(next && (next.start===offset || !previous));
 const character=atStart?next:previous??next;if(!character)return null;
 const rtl=(character.bidi&1)===1;
 const edge=(atStart!==rtl)?[0,1,4,5]:[2,3,6,7];
 return Object.freeze(edge.map(index=>character.quad[index])) as readonly [number,number,number,number];
}

const caretBoundaries=new WeakMap<PdfTextSourceAppearance,readonly {offset:number;edge:readonly[number,number,number,number]}[]>();

/** Reproject the engine's real quads through the reader's current viewport.
 * Crop origin, page rotation and outer display scale must never be replaced
 * with browser text widths or a whole-line scaleX fit. */
export function projectPdfTextSourceAppearance(appearance:PdfTextSourceAppearance,
  pdfToViewport:readonly number[]):PdfTextSourceAppearance|null {
 const source=appearance.pageTransform;
 if(source.length!==6||pdfToViewport.length!==6||[...source,...pdfToViewport].some(value=>!Number.isFinite(value)))return null;
 const [a,b,c,d,e,f]=source,det=a*d-b*c;
 if(Math.abs(det)<1e-12)return null;
 const point=(x:number,y:number)=>{
  const px=(d*(x-e)-c*(y-f))/det,py=(-b*(x-e)+a*(y-f))/det;
  return [pdfToViewport[0]*px+pdfToViewport[2]*py+pdfToViewport[4],pdfToViewport[1]*px+pdfToViewport[3]*py+pdfToViewport[5]];
 };
 const characters=appearance.characters.map(character=>Object.freeze({...character,
  origin:Object.freeze(point(character.origin[0],character.origin[1])),
  quad:Object.freeze(character.quad.flatMap((value,index)=>index%2===0?point(value,character.quad[index+1]):[])),
 }));
 const xs=characters.flatMap(character=>character.quad.filter((_,index)=>index%2===0));
 const ys=characters.flatMap(character=>character.quad.filter((_,index)=>index%2===1));
 if(!characters.length)return null;
 return Object.freeze({...appearance,pageTransform:Object.freeze([...pdfToViewport]),characters:Object.freeze(characters),
  bounds:Object.freeze({left:Math.min(...xs),top:Math.min(...ys),right:Math.max(...xs),bottom:Math.max(...ys)})});
}

const canvasGeometry=new WeakMap<HTMLCanvasElement,{key:string;width:number;height:number;geometry:PdfTextSourceAppearance}>();
/** The last geometry each box's previews were drawn with, kept for what
 * follows the text once it is committed (its links, pdfLinksFollowingText):
 * the canvas's own entry is the editor's to take. */
const latestGeometry=new Map<string,PdfTextSourceAppearance>();
export function rememberPdfNativeTextSourceGeometry(canvas:HTMLCanvasElement,key:string,geometry:PdfTextSourceAppearance):void{
 canvasGeometry.set(canvas,{key,width:canvas.width,height:canvas.height,geometry});
 latestGeometry.delete(key);latestGeometry.set(key,geometry);
 while(latestGeometry.size>16)latestGeometry.delete(latestGeometry.keys().next().value!);
}
/** The geometry a preview of `text` in the box `key` was last drawn with. */
export function readLatestPdfNativeTextSourceGeometry(key:string,text:string):PdfTextSourceAppearance|null{
 const geometry=latestGeometry.get(key);
 return geometry&&geometry.text===text?geometry:null;
}
export function readPdfNativeTextSourceGeometry(canvas:HTMLCanvasElement,key:string,text:string):PdfTextSourceAppearance|null{
 const entry=canvasGeometry.get(canvas);canvasGeometry.delete(canvas);
 return entry&&entry.key===key&&entry.geometry.text===text&&entry.width===canvas.width&&entry.height===canvas.height&&canvas.width>0&&canvas.height>0?entry.geometry:null;
}

function sourceCaretBoundaries(appearance:PdfTextSourceAppearance):readonly {offset:number;edge:readonly[number,number,number,number]}[]{
 let boundaries=caretBoundaries.get(appearance);
 if(!boundaries){
  const offsets=[...new Set([0,appearance.text.length,...appearance.characters.flatMap(character=>[character.start,character.end])])];
  boundaries=offsets.flatMap(offset=>{const edge=pdfSourceCaretEdge(appearance,offset);return edge?[{offset,edge}]:[];});
  caretBoundaries.set(appearance,boundaries);
 }
 return boundaries;
}

export function pdfSourceCaretOffset(appearance:PdfTextSourceAppearance,x:number,y:number):number{
 const boundaries=sourceCaretBoundaries(appearance);
 let best=0,distance=Infinity;
 for(const {offset,edge} of boundaries){
  const dx=edge[2]-edge[0],dy=edge[3]-edge[1],length=dx*dx+dy*dy;
  const t=length?Math.max(0,Math.min(1,((x-edge[0])*dx+(y-edge[1])*dy)/length)):0;
  const next=(x-edge[0]-t*dx)**2+(y-edge[1]-t*dy)**2;
  if(next<distance){distance=next;best=offset;}
 }
 return best;
}










export function pdfSourceVerticalCaretOffset(appearance:PdfTextSourceAppearance,offset:number,direction:-1|1,
  goal?:number):{offset:number;goal:number}|null{
 const current=pdfSourceCaretEdge(appearance,offset);if(!current)return null;
 const vx=current[2]-current[0],vy=current[3]-current[1],height=Math.hypot(vx,vy);
 if(!(height>0))return null;
 const ux=vx/height,uy=vy/height;
 const across=(edge:readonly number[])=>((edge[0]+edge[2])/2)*ux+((edge[1]+edge[3])/2)*uy;
 const along=(edge:readonly number[])=>-((edge[0]+edge[2])/2)*uy+((edge[1]+edge[3])/2)*ux;
 const boundaries=sourceCaretBoundaries(appearance);
 const first=boundaries.find(boundary=>boundary.offset===0),last=boundaries.find(boundary=>boundary.offset===appearance.text.length);
 if(!first||!last)return null;
 // Later rows of the text are "down", whichever way the page is turned.
 const sign=Math.sign(across(last.edge)-across(first.edge));
 if(Math.abs(across(last.edge)-across(first.edge))<height/2||sign===0)return null;
 const down=(edge:readonly number[])=>sign*across(edge);
 const rows:number[]=[];
 for(const value of boundaries.map(boundary=>down(boundary.edge)).sort((left,right)=>left-right)){
  if(!rows.length||value-rows[rows.length-1]>height/2)rows.push(value);
 }
 const here=down(current);
 let row=0;
 for(let index=1;index<rows.length;index++)if(Math.abs(rows[index]-here)<Math.abs(rows[row]-here))row=index;
 const target=row+direction;
 const goalAlong=goal??along(current);
 if(target<0)return {offset:0,goal:goalAlong};
 if(target>=rows.length)return {offset:appearance.text.length,goal:goalAlong};
 let best:{offset:number;distance:number}|null=null;
 for(const boundary of boundaries){
  if(Math.abs(down(boundary.edge)-rows[target])>height/2)continue;
  const distance=Math.abs(along(boundary.edge)-goalAlong);
  if(!best||distance<best.distance-1e-6||(Math.abs(distance-best.distance)<=1e-6&&boundary.offset<best.offset))best={offset:boundary.offset,distance};
 }
 return best?{offset:best.offset,goal:goalAlong}:null;
}

const BREAKS_AROUND=/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}　-〿＀-￯]/u;
/** The widest run of a paragraph's characters a rewrap cannot break inside,
 * in the appearance's horizontal units: a word (runs between spaces), a CJK
 * character on its own, never across a row the source already breaks. A frame
 * narrower than this sets the paragraph exactly as it was. */
export function pdfTextWidestUnbreakableExtent(appearance:PdfTextSourceAppearance):number{
 let widest=0,left=Infinity,right=-Infinity,previous:SourceCharacter|null=null;
 const flush=()=>{if(right>left)widest=Math.max(widest,right-left);left=Infinity;right=-Infinity;};
 for(const character of appearance.characters){
  const xs=[character.quad[0],character.quad[2],character.quad[4],character.quad[6]];
  const size=Math.abs(character.fontSize)||Math.abs(character.quad[1]-character.quad[5])||1;
  if(previous&&(character.start!==previous.end||Math.abs(character.origin[1]-previous.origin[1])>size*.5))flush();
  previous=character;
  if(/^\s+$/u.test(character.text)){flush();continue;}
  if(BREAKS_AROUND.test(character.text)){flush();left=Math.min(...xs);right=Math.max(...xs);flush();continue;}
  left=Math.min(left,...xs);right=Math.max(right,...xs);
 }
 flush();
 return widest;
}

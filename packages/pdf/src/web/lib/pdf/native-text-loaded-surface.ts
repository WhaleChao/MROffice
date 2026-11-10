/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionEvidence, PdfNativeListTextSelectionProof } from "./engine/native/list-text-fidelity";
import { multiplyMatrix } from "./text-boxes";
import type { PDFPageProxy } from "pdfjs-dist";
import { validPdfTextSurfaceSize } from "./native-text-ink-bounds";
import {
  pdfPageSourceAppearanceFromOperators,
  type PdfPageSourceAppearance,
  type PdfSourceFont,
} from "./source-page-appearance";
import { matchPdfTextSourceAppearance, type PdfTextSourceAppearance } from "./source-text-appearance";
import { createPdfRenderOperationsFilter } from "./pdfjs-render-operator-filter";

type Operators = Readonly<{fnArray: readonly number[]; argsArray: readonly unknown[]}>;
type TextRun = Pick<NonNullable<PdfNativeListTextSelectionProof["evidence"]["linePlan"]>["lines"][number],
  "itemIndex"|"originalText"|"textX"|"baseline"|"originalAdvance"|"layoutState"|"verticalCapacity">;
type State = {ctm:number[];tm:number[];x:number;y:number;lineX:number;lineY:number;known:boolean;
  font:string;size:number;spacing:number;wordSpacing:number;hscale:number;rise:number;leading:number;mode:number;annotation:number};
const identity=()=>[1,0,0,1,0,0];
const initial=():State=>({ctm:identity(),tm:identity(),x:0,y:0,lineX:0,lineY:0,known:false,
  font:"",size:0,spacing:0,wordSpacing:0,hscale:100,rise:0,leading:0,mode:0,annotation:0});
const number=(v:unknown):v is number=>typeof v==="number"&&Number.isFinite(v);
function vector(value:unknown):number[]|null {
  const a=Array.isArray(value)||ArrayBuffer.isView(value)?Array.from(value as ArrayLike<unknown>):null;
  return a?.length===6&&a.every(number)?a as number[]:null;
}
function argsAt(list:Operators,index:number):unknown[] {
  return Array.isArray(list.argsArray[index])?list.argsArray[index] as unknown[]:[];
}
function glyphText(args:unknown[]):string|null {
  if(!Array.isArray(args[0])||args[0].length>65_536)return null;
  let text="";
  for(const glyph of args[0]) {
    if(!glyph||typeof glyph!=="object"||typeof glyph.unicode!=="string")return null;
    text+=glyph.unicode;
  }
  return text;
}
const close=(a:number,b:number)=>Math.abs(a-b)<=0.0001;

/** Interpret only text-position state, using the same matrix primitive as the
 * reader. Unknown text/group state declines this optimization. Drawing and
 * writer authority remain with PDF.js and the native source proof. */
function step(state:State,stack:State[],op:number,args:unknown[],ops:Readonly<Record<string,number>>):boolean {
  const save=()=>{if(stack.length>=64)return false;stack.push({...state});return true;};
  const restore=()=>{const previous=stack.pop();if(!previous)return false;Object.assign(state,previous);return true;};
  const move=(x:number,y:number)=>{state.x=state.lineX+=x;state.y=state.lineY+=y;state.known=true;};
  if(op===ops.save)return save();
  if(op===ops.restore)return restore();
  if(op===ops.beginAnnotation){if(!save())return false;state.annotation++;state.known=false;return true;}
  if(op===ops.endAnnotation)return restore();
  if(op===ops.paintFormXObjectBegin){
    if(!save())return false;
    if(args[0]){const matrix=vector(args[0]);if(!matrix)return false;state.ctm=multiplyMatrix(state.ctm,matrix);}
    return true;
  }
  if(op===ops.paintFormXObjectEnd)return restore();
  if(op===ops.beginGroup||op===ops.endGroup||op===ops.showSpacedText||
    op===ops.nextLineShowText||op===ops.nextLineSetSpacingShowText)return false;
  if(op===ops.transform){const m=vector(args);if(!m)return false;state.ctm=multiplyMatrix(state.ctm,m);}
  else if(op===ops.beginText){state.tm=identity();state.x=state.y=state.lineX=state.lineY=0;state.known=true;}
  else if(op===ops.endText)state.known=false;
  else if(op===ops.setTextMatrix){const m=vector(args[0])??vector(args);if(!m)return false;
    state.tm=m;state.x=state.y=state.lineX=state.lineY=0;state.known=true;}
  else if(op===ops.moveText||op===ops.setLeadingMoveText){
    if(!number(args[0])||!number(args[1]))return false;
    if(op===ops.setLeadingMoveText)state.leading=-args[1];move(args[0],args[1]);
  }else if(op===ops.nextLine)move(0,-state.leading);
  else if(op===ops.setFont){if(typeof args[0]!=="string"||!number(args[1]))return false;state.font=args[0];state.size=args[1];}
  else if(op===ops.setLeading||op===ops.setCharSpacing||op===ops.setWordSpacing||op===ops.setHScale||
    op===ops.setTextRise||op===ops.setTextRenderingMode){
    if(!number(args[0]))return false;
    if(op===ops.setLeading)state.leading=args[0];
    else if(op===ops.setCharSpacing)state.spacing=args[0];
    else if(op===ops.setWordSpacing)state.wordSpacing=args[0];
    else if(op===ops.setHScale)state.hscale=args[0];
    else if(op===ops.setTextRise)state.rise=args[0];else state.mode=args[0];
  }else if(op===ops.setGState){
    // A GState can change the font and text metrics; do not infer its contents.
    if(!Array.isArray(args[0])||args[0].some((entry:unknown)=>!Array.isArray(entry)||entry[0]==="Font"))return false;
  }
  return true;
}

/** Removing showText also removes its cursor advance. Prove that every later
 * unselected show reads an independent text position, including saved states
 * restored after an intervening reset. Never shift adjacent text to hide it. */
function independentFollowingText(list:Operators,selected:ReadonlySet<number>,ops:Readonly<Record<string,number>>):boolean {
  let changed=false;const stack:boolean[]=[];
  for(let i=0;i<list.fnArray.length;i++){
    const op=list.fnArray[i];
    if(op===ops.save||op===ops.paintFormXObjectBegin||op===ops.beginAnnotation)stack.push(changed);
    else if(op===ops.restore||op===ops.paintFormXObjectEnd||op===ops.endAnnotation){if(!stack.length)return false;changed=stack.pop()!;}
    else if(op===ops.beginText||op===ops.setTextMatrix||op===ops.moveText||op===ops.setLeadingMoveText||op===ops.nextLine)changed=false;
    else if(op===ops.showText){if(selected.has(i))changed=true;else if(changed&&glyphText(argsAt(list,i))!=="")return false;}
  }
  return stack.length===0;
}

/** UI-only erasure plan for a source-proved text selection. It consumes the
 * loaded page's operator list, never edits it, and grants no writer authority.
 * Null preserves the existing native blanked-document rendering path. */
/** Why the last erasure plan declined (dev diagnostics; the viewer logs it
 * when the line is drawn over its own ink instead of on a clean surface). */
let lastErasureDecline:string|null=null;
export function readPdfLoadedTextErasureDecline():string|null{return lastErasureDecline;}
function decline(reason:string):null{lastErasureDecline=reason;return null;}

export function planPdfLoadedTextErasure(list:Operators,proof:PdfNativeListTextSelectionProof,
  ops:Readonly<Record<string,number>>):ReadonlySet<number>|null {
  const selection=proof.evidence.selection,plan=proof.evidence.linePlan;
  if(!selection||list.fnArray.length>200_000||list.fnArray.length!==list.argsArray.length||
    selection.startOffset!==0||selection.endOffset!==selection.item.str.length||
    proof.evidence.positionedRunGroup||proof.evidence.positionedLines||proof.evidence.paragraphReflow||
    proof.evidence.fontEncoding?.kind.startsWith("identity-v"))return decline("selection-or-list");
  const original=proof.properties?.original;
  const lines:readonly TextRun[]|null=plan?.lines??(original?[{
    itemIndex:selection.item.itemIndex,originalText:proof.evidence.originalText,
    textX:proof.evidence.textX,baseline:proof.evidence.baseline,originalAdvance:proof.evidence.originalAdvance,
    layoutState:original,
  }]:null);
  if(!lines?.length||lines.length>256)return decline("rows");
  const items=selection.items??[selection.item];
  const firstByItem=new Map<number,TextRun>();
  const rows=lines.map(line=>{
    const item=items.find(item=>item.itemIndex===line.itemIndex),layout=line.layoutState;
    if(!item||!layout||!line.originalText||line.verticalCapacity!==undefined||
      !number(line.originalAdvance)||line.originalAdvance<0||!number(layout.fontSize)||layout.fontSize<=0||
      item.transform.length!==6||!item.transform.every(number))return decline("rows-incomplete");
    if(!firstByItem.has(line.itemIndex))firstByItem.set(line.itemIndex,line);
    return {line,item,layout,first:firstByItem.get(line.itemIndex)!};
  });
  if(rows.some(row=>!row))return decline("too-many-attempts");
  const matches=(state:State,text:string|null,row:NonNullable<typeof rows[number]>)=>{
    if(!state.known||state.annotation||state.mode!==0||state.font!==row.item.fontName||
      text!==row.line.originalText||!close(state.size,row.layout.fontSize)||
      !close(state.spacing,row.layout.characterSpacing)||!close(state.wordSpacing,row.layout.wordSpacing)||
      !close(state.hscale,row.layout.horizontalScale)||!close(state.rise,0))return false;
    const m=multiplyMatrix(multiplyMatrix(state.ctm,state.tm),[state.size*state.hscale/100,0,0,state.size,state.x,state.y]);
    const dx=row.line.textX-row.first.textX,dy=row.line.baseline-row.first.baseline;
    const expected=[...row.item.transform];expected[4]+=state.ctm[0]*dx+state.ctm[2]*dy;expected[5]+=state.ctm[1]*dx+state.ctm[3]*dy;
    return m.every((value,index)=>close(value,expected[index]));
  };
  let found:Set<number>|null=null,attempts=0;
  const state=initial(),stack:State[]=[];
  for(let i=0;i<list.fnArray.length;i++){
    const op=list.fnArray[i],args=argsAt(list,i);
    if(op!==ops.showText){if(!step(state,stack,op,args,ops))return decline("second-match");continue;}
    const text=glyphText(args);
    if(matches(state,text,rows[0]!)){
      if(++attempts>64)return decline("null-5");
      const candidate=new Set<number>(),local={...state},saved=stack.map(s=>({...s}));let row=0;
      for(let j=i;j<list.fnArray.length&&row<rows.length;j++){
        const next=list.fnArray[j],values=argsAt(list,j);
        if(next===ops.showText){
          const value=glyphText(values);if(value==="")continue;
          if(!matches(local,value,rows[row]!))break;
          candidate.add(j);local.x+=rows[row]!.line.originalAdvance;row++;
        }else if(!step(local,saved,next,values,ops))break;
      }
      if(row===rows.length){
        if(found)return decline("null-6");found=candidate;
      }
    }
    if(text!=="")state.known=false;
  }
  if(stack.length!==0)return decline("unbalanced-stack");if(!found)return decline("no-match");return independentFollowingText(list,found,ops)?found:decline("following-text-depends");
}

const compactText=(text:string)=>text.replace(/\s+/gu,"");

/** The same UI-only erasure before the source proof arrives, from the reader's
 * own selection witness (pdf.js's item: font, start matrix, text), so the line
 * can be taken over the moment it is clicked, as WPS does. The first show of
 * an item must read with the item's font and size at the item's start matrix
 * and its glyph text must open the item's text (pdf.js's gap spaces removed);
 * the shows that follow it in the same text object complete the text. Spacing
 * and scale are not witnessed, so they are not checked. An item a text object
 * does not finish, a second reading, or text that would shift declines: null
 * keeps the source ink until the proof, as before. */
export function planPdfLoadedTextErasureForSelection(list:Operators,selection:PdfNativeListTextSelectionEvidence,
  ops:Readonly<Record<string,number>>):ReadonlySet<number>|null {
  if(list.fnArray.length>200_000||list.fnArray.length!==list.argsArray.length||
    selection.startOffset!==0||selection.endOffset!==selection.item.str.length)return decline("selection-partial");
  const items=selection.items??[selection.item];
  if(!items.length||items.length>256||items.some(item=>item.transform.length!==6||!item.transform.every(number)||
    !compactText(item.str)))return decline("items");
  const opens=(state:State,text:string|null,item:typeof items[number])=>{
    if(!state.known||state.annotation||state.mode!==0||state.font!==item.fontName||text===null||
      !(state.size>0)||!close(state.rise,0))return false;
    const want=compactText(item.str),got=compactText(text);
    if(!got||!want.startsWith(got))return false;
    const m=multiplyMatrix(multiplyMatrix(state.ctm,state.tm),[state.size*state.hscale/100,0,0,state.size,state.x,state.y]);
    return m.every((value,index)=>close(value,item.transform[index]));
  };
  const found=new Set<number>();
  let from=0,attempts=0;
  for(const item of items){
    const want=compactText(item.str);
    let reading:{indexes:number[];end:number}|null=null;
    const state=initial(),stack:State[]=[];
    for(let i=0;i<list.fnArray.length;i++){
      const op=list.fnArray[i],args=argsAt(list,i);
      if(op!==ops.showText){if(!step(state,stack,op,args,ops))return decline("state-op");continue;}
      const text=glyphText(args);
      if(i>=from&&opens(state,text,item)){
        if(++attempts>64)return decline("attempts");
        const indexes=[i];let got=compactText(text!);
        const local={...state},saved=stack.map(s=>({...s}));
        for(let j=i+1;j<list.fnArray.length&&got.length<want.length;j++){
          const next=list.fnArray[j],values=argsAt(list,j);
          if(next===ops.showText){
            const value=glyphText(values);if(value===null)break;
            const piece=compactText(value);if(!piece)continue;
            if(!want.startsWith(got+piece))break;
            got+=piece;indexes.push(j);
          }else if(next===ops.endText||next===ops.beginText||!step(local,saved,next,values,ops))break;
        }
        if(got===want){
          if(reading)return decline("ambiguous");
          reading={indexes,end:indexes[indexes.length-1]};
        }
      }
      if(text!=="")state.known=false;
    }
    if(!reading)return decline("not-found");
    for(const index of reading.indexes)found.add(index);
    from=reading.end+1;
  }
  return found.size&&independentFollowingText(list,found,ops)?found:decline("dependent-following");
}

/** Replay a current loaded page with only the proved glyph-show operations
 * omitted. All graphics/font state still executes on an independent canvas.
 * Before the proof, the reader's selection witness names the shows instead. */
const PATH_PAINTS=["stroke","closeStroke","fill","eoFill","fillStroke","eoFillStroke","closeFillStroke","closeEOFillStroke"];
const OTHER_PAINTS=["showText","showSpacedText","nextLineShowText","nextLineSetSpacingShowText","paintImageXObject",
  "paintImageXObjectRepeat","paintImageMaskXObject","paintImageMaskXObjectGroup","paintImageMaskXObjectRepeat",
  "paintInlineImageXObject","paintInlineImageXObjectGroup","paintJpegXObject","paintSolidColorImageMask","shadingFill"];

/** The operations that draw the block alone: every state operation of the
 * page content (so its text is drawn in the state the page draws it in) and,
 * of its paints, only the block's own shows. Null when leaving a paint out
 * would change how the block is drawn: a path painted and made the clip at
 * once, or text that clips (render modes 4 to 7), before the block's last
 * show. */
export function planPdfLoadedTextBlockLayer(list:Operators,block:ReadonlySet<number>,
  ops:Readonly<Record<string,number|undefined>>):((index:number)=>boolean)|null{
  const constructPath=ops.constructPath;
  if(typeof constructPath!=="number"||!block.size)return null;
  const pathPaints=new Set(PATH_PAINTS.map(name=>ops[name]).filter((id):id is number=>typeof id==="number"));
  const otherPaints=new Set(OTHER_PAINTS.map(name=>ops[name]).filter((id):id is number=>typeof id==="number"));
  let contentEnd=list.fnArray.length;
  if(typeof ops.beginAnnotation==="number"){const first=list.fnArray.indexOf(ops.beginAnnotation);if(first>=0)contentEnd=first;}
  const last=Math.max(...block);
  if(last>=contentEnd)return null;
  const paint=new Uint8Array(contentEnd);
  for(let index=0;index<contentEnd;index++){
    const fn=list.fnArray[index];
    if(fn===constructPath){
      const args=argsAt(list,index);
      if(pathPaints.has(Number(args[0]))){
        const before=index>0?list.fnArray[index-1]:-1;
        if(index<last&&(before===ops.clip||before===ops.eoClip))return null;
        paint[index]=1;
      }
    }else if(otherPaints.has(fn)){
      paint[index]=1;
    }else if(fn===ops.setTextRenderingMode&&index<last){
      const mode=Number(argsAt(list,index)[0]);
      if(mode>=4&&mode<=7)return null;
    }
  }
  return index=>index<contentEnd&&(block.has(index)||paint[index]===0);
}

const blockLayers=new WeakMap<HTMLCanvasElement,Promise<HTMLCanvasElement|null>>();
const drawnFrom=new WeakMap<HTMLCanvasElement,()=>boolean>();
/** Whether the page drawing `surface` was rendered against is still the
 * page's drawing (see renderPdfLoadedTextEditingSurface); false for a surface
 * drawn elsewhere. A surface drawn after the page was painted again (a line
 * opened while its revision was going up) is as good as one drawn before it,
 * and is not drawn a second time from the bytes. */
export function pdfLoadedTextSurfaceStillCurrent(surface:HTMLCanvasElement|null|undefined):boolean{
  if(!surface)return false;
  try{return drawnFrom.get(surface)?.()??false;}catch{return false;}
}
/** The selected block drawn alone, on a transparent canvas the size of the
 * editing surface `surface` was rendered with (see renderPdfLoadedTextEditingSurface);
 * null when none is drawn for it. A moved block is carried as this layer:
 * the whole of it, as the page draws it, and nothing of its neighbours. */
export function pdfLoadedTextBlockLayer(surface:HTMLCanvasElement|null|undefined):Promise<HTMLCanvasElement|null>|null{
  return surface?blockLayers.get(surface)??null:null;
}

export async function renderPdfLoadedTextEditingSurface(input:{
  page:PDFPageProxy;sourceSha256:string;proof?:PdfNativeListTextSelectionProof;
  selection?:PdfNativeListTextSelectionEvidence;
  ops:Readonly<Record<string,number>>;sourceCanvas:HTMLCanvasElement;rotation:number;
  signal:AbortSignal;isCurrent:()=>boolean;
  optionalContentConfigPromise?:Parameters<PDFPageProxy["render"]>[0]["optionalContentConfigPromise"];
  /** Only this strip of the page (source canvas pixels): the canvas is the
   * strip's size, drawn at the strip's place, and no block layer is drawn. */
  region?:Readonly<{x:number;y:number;width:number;height:number}>;
}):Promise<HTMLCanvasElement|null> {
  const {page,sourceCanvas,signal,proof}=input;
  const strip=input.region;
  const region=strip?(()=>{
    const x=Math.max(0,Math.floor(strip.x)),y=Math.max(0,Math.floor(strip.y));
    const width=Math.min(sourceCanvas.width,Math.ceil(strip.x+strip.width))-x;
    const height=Math.min(sourceCanvas.height,Math.ceil(strip.y+strip.height))-y;
    return width>0&&height>0?{x,y,width,height}:null;
  })():null;
  if(strip&&!region)return null;
  const selection=proof?proof.evidence.selection:input.selection;
  const current=()=>!signal.aborted&&input.isCurrent();
  if(!current()||selection?.sourceSha256!==input.sourceSha256||
    selection.pageNumber!==page.pageNumber||![0,90,180,270].includes(input.rotation)||
    !validPdfTextSurfaceSize(sourceCanvas.width,sourceCanvas.height))return null;
  // Widgets paint from the viewer's annotation storage here too, so the editing
  // surface shows the same filled controls as the page under it; hidden
  // comments are noView entries in that storage, set by the viewer's render.
  const annotationMode=3;
  const deadline=performance.now()+3_000;
  // A cancelled optimization must not cancel the shared document or leave
  // activation waiting for a stalled operator-list request.
  const operators=await new Promise<Awaited<ReturnType<PDFPageProxy["getOperatorList"]>>|null>(resolve=>{
    let settled=false;
    const finish=(value:Awaited<ReturnType<PDFPageProxy["getOperatorList"]>>|null)=>{
      if(settled)return;settled=true;clearTimeout(timer);signal.removeEventListener("abort",abort);resolve(value);
    };
    const abort=()=>finish(null);
    const timer=setTimeout(abort,Math.max(0,deadline-performance.now()));
    signal.addEventListener("abort",abort,{once:true});
    if(!current()){abort();return;}
    void Promise.resolve().then(()=>current()?page.getOperatorList({annotationMode}):null).then(finish,abort);
  });
  if(!operators||!current())return null;
  const erase=proof?planPdfLoadedTextErasure(operators,proof,input.ops):
    input.selection?planPdfLoadedTextErasureForSelection(operators,input.selection,input.ops):null;
  if(!erase&&process.env.NODE_ENV!=="production"){
    // The line is then drawn over its own ink until the candidate lands
    
    try{console.debug("[pdf-erase] declined",{proof:Boolean(proof),reason:lastErasureDecline,item:(proof?.evidence.selection??(input as {selection?:{item?:{itemIndex?:number}}}).selection)?.item?.itemIndex??null});}catch{/* diagnostics never throw */}
  }
  if(!erase)return null;
  if(process.env.NODE_ENV!=="production"){try{console.debug("[pdf-erase] planned",{count:erase.size,proof:Boolean(proof),region:Boolean(region)});}catch{/* diagnostics never throw */}}
  const base=page.getViewport({scale:1,rotation:input.rotation});
  const viewport=page.getViewport({scale:sourceCanvas.width/base.width,rotation:input.rotation});
  if(!Number.isFinite(viewport.height)||Math.abs(viewport.height-sourceCanvas.height)>2)return null;
  const canvas=sourceCanvas.ownerDocument.createElement("canvas");
  canvas.dataset.pdfSurfaceOrigin=proof?"erased-proof":"erased-selection";
  canvas.width=region?.width??sourceCanvas.width;canvas.height=region?.height??sourceCanvas.height;
  for(const property of ["filter","opacity","transform","transformOrigin"] as const)canvas.style[property]=sourceCanvas.style[property];
  // The block alone, drawn beside the surface from the same operations.
  const blockFilter=region?null:planPdfLoadedTextBlockLayer(operators,erase,input.ops);
  const blockLayer=blockFilter?(async():Promise<HTMLCanvasElement|null>=>{
    const layer=sourceCanvas.ownerDocument.createElement("canvas");
    layer.width=sourceCanvas.width;layer.height=sourceCanvas.height;
    const layerContext=layer.getContext("2d");
    if(!layerContext||!current())return null;
    // (Indices of the list read above; render runs pdf.js's folded list.)
    const layerFilter=createPdfRenderOperationsFilter({list:operators,ops:input.ops,keep:blockFilter});
    const layerTask=page.render({canvas:layer,canvasContext:layerContext,viewport,background:"rgba(0,0,0,0)",
      optionalContentConfigPromise:input.optionalContentConfigPromise,annotationMode,operationsFilter:layerFilter.filter});
    layerFilter.bind(layerTask);
    const stop=()=>layerTask.cancel();
    signal.addEventListener("abort",stop,{once:true});
    try{await layerTask.promise;return current()&&layerFilter.finish()?layer:null;}
    catch{return null;}
    finally{signal.removeEventListener("abort",stop);}
  })():null;
  let task:ReturnType<PDFPageProxy["render"]>|undefined,accepted=false,timedOut=false;
  const cancel=()=>task?.cancel();
  const timer=setTimeout(()=>{timedOut=true;cancel();},Math.max(0,deadline-performance.now()));
  signal.addEventListener("abort",cancel,{once:true});
  try{
    const context=canvas.getContext("2d");if(!context||!current())return null;
    // The erased shows are indices of the list read above (getOperatorList,
    // unfolded); render runs pdf.js's folded list, whose indices differ on
    // any page the optimizer folded paths or images on before the line: the
    // line was then never erased and the new face drawn over its old ink
    
    // align for is dropped (the editor keeps the page's ink, as on decline).
    const eraseFilter=createPdfRenderOperationsFilter({list:operators,ops:input.ops,keep:index=>!erase.has(index)});
    task=page.render({canvas,canvasContext:context,viewport,optionalContentConfigPromise:input.optionalContentConfigPromise,
      annotationMode,
      ...(region?{transform:[1,0,0,1,-region.x,-region.y]}:{}),
      operationsFilter:eraseFilter.filter});
    eraseFilter.bind(task);
    await task.promise;
    if(!current()||timedOut)return null;
    if(!eraseFilter.finish()){
      if(process.env.NODE_ENV!=="production"){try{console.debug("[pdf-erase] declined",{proof:Boolean(proof),reason:"render-list-unaligned"});}catch{/* diagnostics never throw */}}
      return null;
    }
    accepted=true;
    if(blockLayer)blockLayers.set(canvas,blockLayer);
    drawnFrom.set(canvas,input.isCurrent);
    return canvas;
  }catch{return null;}
  finally{clearTimeout(timer);signal.removeEventListener("abort",cancel);if(!accepted){canvas.width=0;canvas.height=0;}}
}

/** The line erased from the page as the reader sees it while the revision
 * an edit just staged is read in (`shown`: that page as edited, drawn over
 * the loaded one). The line's own ink -- the pixels where the loaded page
 * and the loaded page without the line (`erased`) differ, within `rect`
 * (canvas pixels) -- is taken from `erased`; every other pixel, the edit
 * just made included, stays as shown. Null when the three are not one
 * page's drawings at one size. */
export function erasePdfLoadedLineOnShownPage(input:{
  shown:HTMLCanvasElement;loaded:HTMLCanvasElement;erased:HTMLCanvasElement;
  rect:Readonly<{x:number;y:number;width:number;height:number}>;
  /** Where `erased` lies on the page when it is a strip of it (canvas
   * pixels; renderPdfLoadedTextEditingSurface `region`). */
  erasedAt?:Readonly<{x:number;y:number}>;
}):HTMLCanvasElement|null{
  const {shown,loaded,erased}=input;
  const width=shown.width,height=shown.height;
  const ox=input.erasedAt?.x??0,oy=input.erasedAt?.y??0;
  if(!validPdfTextSurfaceSize(width,height)||loaded.width!==width||loaded.height!==height||
    (!input.erasedAt&&(erased.width!==width||erased.height!==height)))return null;
  const x=Math.max(0,Math.floor(input.rect.x),ox),y=Math.max(0,Math.floor(input.rect.y),oy);
  const w=Math.min(width,Math.ceil(input.rect.x+input.rect.width),ox+erased.width)-x;
  const h=Math.min(height,Math.ceil(input.rect.y+input.rect.height),oy+erased.height)-y;
  if(!(w>0&&h>0))return null;
  const out=shown.ownerDocument.createElement("canvas");
  out.width=width;out.height=height;
  for(const property of ["filter","opacity","transform","transformOrigin"] as const)out.style[property]=erased.style[property];
  const context=out.getContext("2d");
  const loadedContext=loaded.getContext("2d");
  const erasedContext=erased.getContext("2d");
  if(!context||!loadedContext||!erasedContext){out.width=0;out.height=0;return null;}
  try{
    context.drawImage(shown,0,0);
    const page=context.getImageData(x,y,w,h);
    const before=loadedContext.getImageData(x,y,w,h).data;
    const after=erasedContext.getImageData(x-ox,y-oy,w,h).data;
    const pixels=page.data;
    for(let index=0;index<pixels.length;index+=4){
      if(Math.abs(before[index]-after[index])+Math.abs(before[index+1]-after[index+1])+
        Math.abs(before[index+2]-after[index+2])+Math.abs(before[index+3]-after[index+3])<=6)continue;
      pixels[index]=after[index];pixels[index+1]=after[index+1];
      pixels[index+2]=after[index+2];pixels[index+3]=after[index+3];
    }
    context.putImageData(page,x,y);
    return out;
  }catch{out.width=0;out.height=0;return null;}
}

const loadedPageScenes=new WeakMap<PDFPageProxy,Map<string,Promise<PdfPageSourceAppearance|null>>>();








/** The fonts an operator list names, as the reader resolves them: it answers
 * the list first and each font once its face is bound, so a page read at once
 * had none of its shows some of the time (the candidate then had no glyphs for
 * the caret and the selection's band). */
async function pageFontsResolved(page:PDFPageProxy,operators:Readonly<{fnArray:ArrayLike<number>;argsArray:ArrayLike<unknown>}>,ops:Readonly<Record<string,number>>):Promise<void>{
  const names=new Set<string>();
  for(let index=0;index<operators.fnArray.length;index+=1){
    if(operators.fnArray[index]!==ops.setFont)continue;
    const name=(operators.argsArray[index] as readonly unknown[]|null)?.[0];
    if(typeof name==="string")names.add(name);
  }
  await Promise.all([...names].map(name=>new Promise<void>(resolve=>{
    if(page.commonObjs.has(name)){resolve();return;}
    const timer=setTimeout(resolve,3_000);
    try{page.commonObjs.get(name,()=>{clearTimeout(timer);resolve();});}catch{clearTimeout(timer);resolve();}
  })));
}

export async function readPdfLoadedTextSourceGeometry(input:{
  page:PDFPageProxy;sourceSha256:string;selection:PdfNativeListTextSelectionEvidence;text:string;
  ops:Readonly<Record<string,number>>;signal?:AbortSignal;
}):Promise<PdfTextSourceAppearance|null>{
  const {page,selection}=input;
  if(selection.sourceSha256!==input.sourceSha256||selection.pageNumber!==page.pageNumber||input.signal?.aborted)return null;
  let scenes=loadedPageScenes.get(page);
  if(!scenes){scenes=new Map();loadedPageScenes.set(page,scenes);}
  let scene=scenes.get(input.sourceSha256);
  if(!scene){
    scene=page.getOperatorList({annotationMode:0}).then(async operators=>{await pageFontsResolved(page,operators,input.ops);return operators;}).then(operators=>pdfPageSourceAppearanceFromOperators({
      sourceSha256:input.sourceSha256,pageNumber:page.pageNumber,operators,ops:input.ops,
      font:name=>{try{return page.commonObjs.has(name)?page.commonObjs.get(name) as PdfSourceFont:null;}catch{return null;}},
    }),()=>null);
    scenes.set(input.sourceSha256,scene);
  }
  let read:PdfPageSourceAppearance|null;
  try{read=await scene;}catch{read=null;}
  if(!read||input.signal?.aborted)return null;
  try{return matchPdfTextSourceAppearance(read,{...selection,item:{...selection.item,str:input.text}});}catch{return null;}
}

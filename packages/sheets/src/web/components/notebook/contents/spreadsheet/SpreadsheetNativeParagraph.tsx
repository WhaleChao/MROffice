"use client";
/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore, type CSSProperties } from "react";
import { flushSync } from "react-dom";
import { layoutNativeText, layoutNativeTextBatch, nativeTextLayoutIsKnown, nativeTextTouchedBy, resolveRelativePartFontSizes, type NativeTextLayoutRequest, nativeTextFontLead, nativeParagraphLineStyle, nativeParagraphLinePitch, measureNativeLineAdvance, nativeTextAlignmentOffset, nativeStrutBaselinePx, type NativeTextAlignmentBox, type NativeTextPart, type NativeTextLine } from "../../../../lib/spreadsheet/native-text-layout";
import { nativeTextLayoutQuantum, nativeTypographyGeneration, resolveSpreadsheetTextStyle, subscribeNativeTypography } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { fontLoadReach } from "@mrnote/core/web/lib/spreadsheet/font-load-reach.ts";
import { distributedLineUnits } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { SpreadsheetNativeTextParts, spreadsheetNativeTextPaint } from "./SpreadsheetNativeText";

/** A paragraph whose lines the commit's batch has yet to lay out. */
const PENDING_LAYOUT={lines:[] as readonly NativeTextLine[],offsets:[] as readonly (number|undefined)[]};

// The paragraphs that mounted in one commit and are not laid out yet: their
// probes go into the document together, once the commit's layout effects are
// done, and they repaint together before the browser paints.
const batchRequests:NativeTextLayoutRequest[]=[];
const batchListeners=new Set<()=>void>();
let batchScheduled=false;
function requestNativeTextLayout(request:NativeTextLayoutRequest,done:()=>void):()=>void{
  batchRequests.push(request);batchListeners.add(done);
  if(!batchScheduled){
    batchScheduled=true;
    queueMicrotask(()=>{
      batchScheduled=false;
      const requests=batchRequests.splice(0),listeners=[...batchListeners];
      batchListeners.clear();
      layoutNativeTextBatch(requests);
      const notify=()=>{for(const listener of listeners)listener()};
      try{flushSync(notify)}catch{notify()}
    });
  }
  return()=>{batchListeners.delete(done)};
}

const noTypographySubscription=()=>()=>{};
const noTypographyGeneration=()=>0;
export function SpreadsheetNativeParagraph({parts,style,dpi,scale=1,wrap=true,layoutWidth,alignmentBox,boxStyle,drawSingleRun=false,spread,distributed=false,officeBaseline}: {
  parts:readonly NativeTextPart[];style:CSSProperties;dpi:number;scale?:number;wrap?:boolean;layoutWidth?:number;alignmentBox?:NativeTextAlignmentBox;
  /** Vertical Justify / Distributed: the lines spread over the paragraph's
   * box, which fills the cell (the page's text box; lookwin probe p22). Lines
   * taller than the box keep to its top, as Excel prints them. */
  spread?:"space-between"|"space-evenly";
  /** The paragraph is its own box in the cell (a box around it set these). */
  boxStyle?:CSSProperties;
  /** A single line of one plain run in the paragraph's own font draws its
   * glyph runs in the line itself rather than in a text span inside it. */
  drawSingleRun?:boolean;
  /** Horizontal Distributed (Indent): each wrapped line of one plain run
   * spreads across the layout width (the cell less its indent at both
   * sides) by its words, or by character with CJK, and a line of one unit
   * is centred (distributedLineUnits; the API's print draws the same). */
  distributed?:boolean;
  /** Where Windows Excel stands the line's baseline, px below the
   * paragraph's top edge (a Windows grid's one printed line, set by its
   * vertical alignment and its faces' automatic line, as the API's writer
   * stands it): the paragraph moves so its baseline is there, the browser's
   * own (nativeStrutBaselinePx) read for its font and line height. */
  officeBaseline?:number;
}) {
  const ref=useRef<HTMLSpanElement>(null);
  // A paragraph standing at Excel's baseline reads the browser's again when
  // a font arrives (only those subscribe).
  const typographyGeneration=useSyncExternalStore(officeBaseline!==undefined?subscribeNativeTypography:noTypographySubscription,
    officeBaseline!==undefined?nativeTypographyGeneration:noTypographyGeneration,noTypographyGeneration);
  // A part sized relative to the paragraph is laid out at its pixels.
  const signature=JSON.stringify([resolveRelativePartFontSizes(parts,style),style,dpi,scale,wrap,layoutWidth,alignmentBox]);
  const stable=useMemo(()=>JSON.parse(signature) as [NativeTextPart[],CSSProperties,number,number,boolean,number|null,NativeTextAlignmentBox|null],[signature]);
  const [layout,setLayout]=useState<{signature:string;lines:readonly NativeTextLine[];offsets:readonly (number|undefined)[]}|null>(null);
  // A paragraph whose wrap width is given (a native grid's cell) needs nothing
  // from its own box: its lines are laid out while it renders, from the
  // layout cache when the text was seen before. Laying them out in an effect
  // cost every such cell a second render, a forced layout of the grid to read
  // a width it already had, and a ResizeObserver; a sheet of wrapped cells
  // took seconds to show. Fonts still loading keep the effect's path.
  const [fontEpoch,setFontEpoch]=useState(0);
  const renderLayout=useMemo(()=>{
    void fontEpoch;
    if(typeof document==="undefined")return undefined;
    const align=stable[1].textAlign;
    const nativeAlignment=stable[2]===72&&stable[6]&&(align==="center"||align==="right"||align==="end");
    if(stable[4]?stable[5]==null:!nativeAlignment)return undefined;
    // While a font is loading nothing measured would last: the paragraph
    // waits for the font, then for the batch.
    if(document.fonts?.status==="loading")return PENDING_LAYOUT;
    // A text not laid out before waits for the commit's batch (below): one
    // layout of the page serves every paragraph that mounted with it.
    if(stable[4]&&!nativeTextLayoutIsKnown(stable[0],stable[1],stable[5]!,stable[2]))return PENDING_LAYOUT;
    const lines:readonly NativeTextLine[]|undefined=stable[4]?layoutNativeText(stable[0],stable[1],stable[5]!,stable[2]):[{parts:stable[0],hardBreak:""}];
    if(!lines)return null;
    const offsets=lines.map(line=>nativeAlignment?nativeTextAlignmentOffset(measureNativeLineAdvance(line.parts,stable[1],stable[2]),stable[1].textAlign,stable[6]??undefined,stable[2],stable[3]):undefined);
    return {lines,offsets};
  },[stable,fontEpoch]);
  const laidOutInRender=renderLayout!==undefined;
  const awaitingBatch=renderLayout===PENDING_LAYOUT;
  useLayoutEffect(()=>{
    if(!awaitingBatch||document.fonts?.status==="loading")return;
    // Unwrapped alignment needs no line layout: the font's arrival alone
    // (the listener below) repaints it.
    if(!stable[4])return;
    return requestNativeTextLayout({parts:stable[0],style:stable[1],width:stable[5]!,dpi:stable[2]},()=>setFontEpoch(epoch=>epoch+1));
  },[awaitingBatch,stable,fontEpoch]);
  useLayoutEffect(()=>{
    if(!laidOutInRender)return;
    // Laid out again only for a load that reaches this paragraph's text
    
    const refresh=(event:Event)=>{
      const reach=fontLoadReach(event);
      if(reach&&!nativeTextTouchedBy(reach,stable[0],stable[1],stable[2]))return;
      setFontEpoch(epoch=>epoch+1);
    };
    document.fonts?.addEventListener("loadingdone",refresh);
    return()=>document.fonts?.removeEventListener("loadingdone",refresh);
  },[laidOutInRender,stable]);
  useLayoutEffect(()=>{
    if(laidOutInRender)return;
    const node=ref.current;
    const align=stable[1].textAlign;
    const nativeAlignment=stable[2]===72&&stable[6]&&(align==="center"||align==="right"||align==="end");
    if(!node||(!stable[4]&&!nativeAlignment))return;
    let cancelled=false;
    const unwrappedLines:readonly NativeTextLine[]=[{parts:stable[0],hardBreak:""}];
    const update=async()=>{
      // Reading fonts.ready lays the page out; it is awaited only while a
      // font is actually loading.
      if(document.fonts?.status==="loading")await document.fonts.ready;if(cancelled)return;
      const rawWidth=parseFloat(getComputedStyle(node).width);
      const quantum=nativeTextLayoutQuantum();
      const width=quantum?Math.round(rawWidth/quantum)*quantum:rawWidth;
      const lines=stable[4]?layoutNativeText(stable[0],stable[1],stable[5]??width/stable[3],stable[2]):unwrappedLines;
      const offsets=lines?.map(line=>nativeAlignment?nativeTextAlignmentOffset(measureNativeLineAdvance(line.parts,stable[1],stable[2]),stable[1].textAlign,stable[6]??undefined,stable[2],stable[3]):undefined)??[];
      if(!cancelled)setLayout(previous=>lines?previous?.signature===signature&&previous.lines===lines&&JSON.stringify(previous.offsets)===JSON.stringify(offsets)?previous:{signature,lines,offsets}:null);
    };
    void update();
    const observer=new ResizeObserver(()=>{void update()});observer.observe(node);
    const fontsLoaded=(event:Event)=>{
      const reach=fontLoadReach(event);
      if(reach&&!nativeTextTouchedBy(reach,stable[0],stable[1],stable[2]))return;
      void update();
    };
    document.fonts?.addEventListener("loadingdone",fontsLoaded);
    return()=>{cancelled=true;observer.disconnect();document.fonts?.removeEventListener("loadingdone",fontsLoaded)};
  },[signature,stable,laidOutInRender]);
  const lineHeight=nativeParagraphLinePitch(stable[0],style,dpi);
  // Only a fixed line height needs the wrapper (its zero line height keeps
  // the runs from stretching the line box); otherwise the runs sit in the
  // line directly, one element less for every cell on screen.
  const paintParts=(lineParts:readonly NativeTextPart[])=>lineHeight===undefined
    ?<SpreadsheetNativeTextParts parts={lineParts} style={style} dpi={dpi} scale={scale}/>
    :<span style={{lineHeight:0}}><SpreadsheetNativeTextParts parts={lineParts} style={style} dpi={dpi} scale={scale}/></span>;
  const shown=laidOutInRender?(awaitingBatch?null:renderLayout):layout?.signature===signature?layout:null;
  const lines=shown?.lines;
  const base=resolveSpreadsheetTextStyle(style,stable[0][0]?nativeTextFontLead(stable[0][0]):"",dpi);
  const spreading=Boolean(spread)&&lines!==undefined&&lines.length>1;
  const fixedLines=lineHeight!==undefined&&lines!==undefined&&!spreading;
  // One line stood at Excel's baseline: the browser's own, read for this
  // font and line height, and the difference taken up above the paragraph.
  // Where nothing lays out a probe, a line box's usual split stands in.
  const strutSize=typeof base.fontSize==="number"?base.fontSize*scale:undefined;
  const strut=officeBaseline!==undefined&&lineHeight!==undefined&&(lines===undefined||lines.length===1)&&!spreading
    ?nativeStrutBaselinePx({fontFamily:base.fontFamily,fontSize:strutSize??base.fontSize,fontWeight:base.fontWeight,fontStyle:base.fontStyle},lineHeight*scale,typographyGeneration)
      ??(strutSize!==undefined?(lineHeight*scale-1.12*strutSize)/2+0.9*strutSize:undefined):undefined;
  const paragraphStyle:CSSProperties={...boxStyle,display:"block",width:"100%",minWidth:0,...(fixedLines?{position:"relative",height:lineHeight*scale*lines.length}:{}),
    ...(strut!==undefined&&officeBaseline!==undefined?{marginTop:officeBaseline-strut,flex:"none"}:{}),
    ...(spreading?{display:"flex",flexDirection:"column",justifyContent:spread,height:"100%"}:{}),
    lineHeight:lineHeight===undefined?"normal":`${lineHeight*scale}px`,fontFamily:base.fontFamily,fontSize:typeof base.fontSize==="number"?base.fontSize*scale:base.fontSize,fontWeight:base.fontWeight,fontStyle:base.fontStyle,textAlign:base.textAlign};
  // One line that sits where the paragraph does (no fixed line box, no
  // alignment offset, no line break to carry): the paragraph and its line are
  // two blocks of the same box, so one element draws both. Most cells on a
  // sheet are this, and the grid mounts one node less for each.
  const only=lines?.length===1&&!fixedLines&&shown?.offsets[0]===undefined&&!lines[0].hardBreak&&!distributed?lines[0]:undefined;
  if(only){
    const lineStyle={...paragraphStyle,...nativeParagraphLineStyle(only,style,scale,dpi,lineHeight)};
    const run=drawSingleRun&&lineHeight===undefined&&only.parts.length===1&&!Object.keys(only.parts[0].style??{}).length?resolveRelativePartFontSizes(only.parts,style)[0]:undefined;
    if(run){
      // Exactly what the text span would draw (SpreadsheetNativeTextParts'
      // arguments): in the same font as the line, the line's strut and the
      // span's box are one, so the line can carry the runs itself.
      const lead=nativeTextFontLead(run);
      const paint=spreadsheetNativeTextPaint({text:run.text,style:resolveSpreadsheetTextStyle({...style,...run.style},lead,dpi),dpi,scale,advanceOffset:0,fontLeadingText:lead,partStyle:run.style});
      if(paint.style.fontFamily===lineStyle.fontFamily&&paint.style.fontSize===lineStyle.fontSize&&paint.style.fontStyle===lineStyle.fontStyle&&paint.style.fontWeight===undefined)
        return <span ref={ref} data-spreadsheet-native-paragraph="" data-native-text-line="" data-spreadsheet-native-text="" style={{...lineStyle,...paint.style}}>{paint.children}</span>;
    }
    return <span ref={ref} data-spreadsheet-native-paragraph="" data-native-text-line="" style={lineStyle}>{paintParts(only.parts)}</span>;
  }
  // Distributed: each line's units at their places, the gaps between them
  // equal (or the line centred, or at the start where it overfills).
  const distributedRoom=distributed&&wrap&&stable[5]!=null&&parts.length===1&&!Object.keys(parts[0]?.style??{}).length?stable[5]:undefined;
  const distributedLine=(line:NativeTextLine)=>{
    if(distributedRoom===undefined)return null;
    const text=line.parts.map(part=>part.text).join("");
    if(!text.trim()||line.parts.some(part=>Object.keys(part.style??{}).length))return null;
    const {mode,units}=distributedLineUnits(text);
    const widths=units.map(unit=>measureNativeLineAdvance([{text:unit}],style,dpi));
    if(widths.some(width=>width===undefined))return null;
    const total=widths.reduce<number>((sum,width)=>sum+(width??0),0);
    if(mode==="center"||total>=distributedRoom){
      const trimmed=text.trim();
      const advance=measureNativeLineAdvance([{text:trimmed}],style,dpi);
      const start=mode==="center"&&advance!==undefined?(distributedRoom-advance)/2:0;
      return <span style={{marginLeft:start*scale}}>{paintParts([{text:trimmed}])}</span>;
    }
    const gap=(distributedRoom-total)/(units.length-1);
    return units.map((unit,index)=><span key={index} style={index?{marginLeft:gap*scale}:undefined}>{paintParts([{text:unit}])}</span>);
  };
  return <span ref={ref} data-spreadsheet-native-paragraph="" style={paragraphStyle}>
    {lines?lines.map((line,index)=>{
      const terminal=resolveSpreadsheetTextStyle({...style,...line.breakStyle},"",dpi);
      const offset=shown?.offsets[index];
      return <span key={index} data-native-text-line="" style={{...nativeParagraphLineStyle(line,style,scale,dpi,lineHeight),...(fixedLines?{position:"absolute",top:index*lineHeight*scale,left:offset??0,width:"100%",...(offset!==undefined?{textAlign:"left"}:{})}:offset!==undefined?{textAlign:"left",position:"relative",left:offset}:{}),
        ...(spreading?{flex:"none",...(lineHeight!==undefined?{height:lineHeight*scale}:{})}:{}),
        ...(distributedRoom!==undefined?{textAlign:"left"}:{})}}>
        {distributedLine(line)??paintParts(line.parts)}
        {line.hardBreak?<><span style={{fontFamily:terminal.fontFamily,fontWeight:terminal.fontWeight,fontStyle:terminal.fontStyle,fontSize:typeof terminal.fontSize==="number"?terminal.fontSize*scale:terminal.fontSize,lineHeight:lineHeight===undefined?undefined:0}}/><span style={{display:"none"}}>{line.hardBreak}</span></>:null}
      </span>;
    }):paintParts(parts)}
  </span>;
}

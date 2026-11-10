/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ElementPicture } from "@mrnote/core/web/lib/export/element-picture.ts";

export interface HtmlShowSlide {
  readonly picture: ElementPicture;
  readonly notes: string;
  readonly title: string;
}

export interface HtmlShowStrings {
  readonly previous: string;
  readonly next: string;
  readonly notes: string;
  readonly noNotes: string;
  readonly fullScreen: string;
  /** "Slide {current} of {total}", the two places left as written. */
  readonly counter: string;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function counterText(template: string, current: number, total: number): string {
  return template.replace("{current}", String(current)).replace("{total}", String(total));
}

/** CSS that cannot close its `<style>` element. */
function styleText(css: string): string {
  return css.replace(/</g, "\\3c ");
}

/** CSS rules of the pictures, each kept once (faces repeat across slides). */
function pictureRules(slides: readonly HtmlShowSlide[]): string {
  const rules = new Set<string>();
  for (const slide of slides) {
    for (const rule of slide.picture.css.split("\n")) {
      if (rule.trim()) rules.add(rule);
    }
  }
  return [...rules].join("\n");
}

const SHOW_STYLE = `
*{box-sizing:border-box}
html,body{margin:0;height:100%;background:#000;overflow:hidden}
body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,"PingFang SC","Microsoft YaHei",sans-serif}
.show{position:fixed;inset:0;display:flex;align-items:center;justify-content:center}
.show.with-notes{right:min(34vw,420px)}
.slide{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;opacity:0;transition:opacity .25s ease;pointer-events:none}
.slide.current{opacity:1;pointer-events:auto}
.slide__frame{flex:none;transform-origin:center center;box-shadow:0 0 0 1px rgba(255,255,255,.06)}
.notes{position:fixed;top:0;right:0;bottom:0;width:min(34vw,420px);background:#16181c;color:#e8eaed;padding:20px 20px 72px;overflow:auto;border-left:1px solid #2a2d31;font-size:15px;line-height:1.55}
[data-notes-for]{white-space:pre-wrap}
.notes[hidden]{display:none}
.notes__empty{color:#8a8f96;font-style:italic}
.bar{position:fixed;left:50%;bottom:14px;transform:translateX(-50%);display:flex;align-items:center;gap:6px;padding:6px 8px;border-radius:999px;background:rgba(22,24,28,.86);color:#e8eaed;opacity:0;transition:opacity .2s ease;z-index:2}
body:hover .bar,.bar:focus-within{opacity:1}
.bar button{appearance:none;border:0;background:transparent;color:inherit;font:inherit;font-size:14px;padding:6px 10px;border-radius:999px;cursor:pointer}
.bar button:hover{background:rgba(255,255,255,.12)}
.bar button:focus-visible{outline:2px solid #5b9dff;outline-offset:1px}
.bar__count{font-variant-numeric:tabular-nums;font-size:13px;padding:0 6px;color:#c3c7cc}
@media print{html,body{overflow:visible;background:#fff}.bar,.notes{display:none!important}.show{position:static;display:block}.slide{position:relative;opacity:1;break-after:page;height:auto}.slide__frame{transform:none!important}}
`.trim();

/** The player: the only script the file runs (its hash is in the policy). */
const PLAYER = `(function(){
var slides=[].slice.call(document.querySelectorAll(".slide"));
var notes=[].slice.call(document.querySelectorAll("[data-notes-for]"));
var panel=document.querySelector(".notes");
var show=document.querySelector(".show");
var count=document.querySelector(".bar__count");
var total=slides.length;var current=0;
function fit(){var box=show.getBoundingClientRect();slides.forEach(function(slide){var frame=slide.firstElementChild;var w=+frame.getAttribute("data-width"),h=+frame.getAttribute("data-height");var s=Math.min(box.width/w,box.height/h);frame.style.transform="scale("+s+")";});}
var viewed=0;function go(index){if(!total)return;var next=Math.max(0,Math.min(total-1,index));if(next!==current)viewed=current;current=next;slides.forEach(function(slide,i){var on=i===current;slide.classList.toggle("current",on);slide.setAttribute("aria-hidden",on?"false":"true");});notes.forEach(function(note){note.hidden=+note.getAttribute("data-notes-for")!==current;});count.textContent=count.getAttribute("data-template").replace("{current}",current+1).replace("{total}",total);if(location.hash!=="#"+(current+1))history.replaceState(null,"","#"+(current+1));}
function toggleNotes(){panel.hidden=!panel.hidden;show.classList.toggle("with-notes",!panel.hidden);fit();}
function fullScreen(){var root=document.documentElement;if(document.fullscreenElement){document.exitFullscreen();}else if(root.requestFullscreen){root.requestFullscreen();}}
document.addEventListener("keydown",function(event){if(event.metaKey||event.ctrlKey||event.altKey)return;var key=event.key;var target=event.target;if((key==="Enter"||key===" ")&&target&&target.closest&&target.closest("button,a[href],input,select,textarea,[role=button]"))return;if(key==="ArrowRight"||key==="ArrowDown"||key==="PageDown"||key===" "||key==="Enter"){go(current+1);}else if(key==="ArrowLeft"||key==="ArrowUp"||key==="PageUp"||key==="Backspace"){go(current-1);}else if(key==="Home"){go(0);}else if(key==="End"){go(total-1);}else if(key==="n"||key==="N"){toggleNotes();}else if(key==="f"||key==="F"){fullScreen();}else{return;}event.preventDefault();});
show.addEventListener("click",function(event){var target=event.target;if(target.closest&&target.closest("a"))return;var action=target.closest&&target.closest("[data-action-target]");var jump=action?action.getAttribute("data-action-target"):null;if(jump==="first"){go(0);}else if(jump==="last"){go(total-1);}else if(jump==="previous"){go(current-1);}else if(jump==="last-viewed"){go(viewed);}else if(jump==="end-show"){return;}else{go(current+1);}});
var startX=null;show.addEventListener("touchstart",function(event){startX=event.touches[0].clientX;},{passive:true});
show.addEventListener("touchend",function(event){if(startX===null)return;var dx=event.changedTouches[0].clientX-startX;startX=null;if(Math.abs(dx)>40)go(current+(dx<0?1:-1));});
document.querySelector("[data-action=prev]").addEventListener("click",function(){go(current-1);});
document.querySelector("[data-action=next]").addEventListener("click",function(){go(current+1);});
document.querySelector("[data-action=notes]").addEventListener("click",toggleNotes);
document.querySelector("[data-action=full]").addEventListener("click",fullScreen);
window.addEventListener("resize",fit);
var start=parseInt(location.hash.slice(1),10);fit();go(isFinite(start)?start-1:0);
})();`;

async function sha256Base64(text: string): Promise<string | null> {
  const subtle = typeof crypto !== "undefined" ? crypto.subtle : undefined;
  if (!subtle || typeof TextEncoder === "undefined") return null;
  try {
    const digest = await subtle.digest("SHA-256", new TextEncoder().encode(text));
    let binary = "";
    for (const byte of new Uint8Array(digest)) binary += String.fromCharCode(byte);
    return btoa(binary);
  } catch {
    return null;
  }
}

/** The standalone HTML show for `slides`, in order. */
export async function buildPresentationHtmlShow(input: {
  title: string;
  lang: string;
  slides: readonly HtmlShowSlide[];
  strings: HtmlShowStrings;
}): Promise<string> {
  const { title, lang, slides, strings } = input;
  const total = slides.length;
  const hash = await sha256Base64(PLAYER);
  const policy = [
    "default-src 'none'",
    "img-src data:",
    "font-src data:",
    "style-src 'unsafe-inline'",
    hash ? `script-src 'sha256-${hash}'` : "script-src 'unsafe-inline'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join("; ");
  const sections = slides
    .map((slide, index) => {
      const { width, height, markup } = slide.picture;
      return (
        `<section class="slide" role="group" aria-roledescription="slide" aria-label="${escapeHtml(
          counterText(strings.counter, index + 1, total),
        )}${slide.title ? ` · ${escapeHtml(slide.title)}` : ""}" aria-hidden="true">` +
        `<div class="slide__frame" data-width="${width}" data-height="${height}" style="width:${width}px;height:${height}px;position:relative;overflow:hidden">` +
        markup +
        `</div></section>`
      );
    })
    .join("\n");
  const notes = slides
    .map((slide, index) => {
      const text = slide.notes.trim();
      return text
        ? `<div data-notes-for="${index}" hidden>${escapeHtml(text)}</div>`
        : `<div data-notes-for="${index}" class="notes__empty" hidden>${escapeHtml(strings.noNotes)}</div>`;
    })
    .join("\n");
  return `<!doctype html>
<html lang="${escapeHtml(lang)}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="${escapeHtml(policy)}">
<meta name="generator" content="MRNote">
<title>${escapeHtml(title)}</title>
<style>${SHOW_STYLE}</style>
<style>${styleText(pictureRules(slides))}</style>
</head>
<body>
<main class="show" aria-live="polite">
${sections}
</main>
<aside class="notes" aria-label="${escapeHtml(strings.notes)}" hidden>
${notes}
</aside>
<nav class="bar">
<button type="button" data-action="prev" aria-label="${escapeHtml(strings.previous)}" title="${escapeHtml(strings.previous)}">&#8249;</button>
<span class="bar__count" data-template="${escapeHtml(strings.counter)}"></span>
<button type="button" data-action="next" aria-label="${escapeHtml(strings.next)}" title="${escapeHtml(strings.next)}">&#8250;</button>
<button type="button" data-action="notes" title="${escapeHtml(strings.notes)} (N)">${escapeHtml(strings.notes)}</button>
<button type="button" data-action="full" title="${escapeHtml(strings.fullScreen)} (F)">${escapeHtml(strings.fullScreen)}</button>
</nav>
<script>${PLAYER}</script>
</body>
</html>
`;
}

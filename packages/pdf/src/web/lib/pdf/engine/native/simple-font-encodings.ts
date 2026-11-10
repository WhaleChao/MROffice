/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Simple-font base encodings and the Latin glyph names they use, frozen from
 * pdfminer.six (latin_enc + Adobe Glyph List) so the server replay and this
 * client resolve exactly the same character for every byte. WinAnsi carries
 * the cp1252 values the Standard 14 contract already uses (0xA0 U+00A0,
 * 0xAD U+00AD). The API contract test pins these tables against pdfminer.
 */

export type PdfSimpleFontBaseEncodingName =
  | "StandardEncoding"
  | "MacRomanEncoding"
  | "WinAnsiEncoding";

export const PDF_SIMPLE_FONT_BASE_ENCODINGS: Readonly<
  Record<PdfSimpleFontBaseEncodingName, Readonly<Record<number, string>>>
> = Object.freeze({
  StandardEncoding: Object.freeze({32:" ",33:"!",34:"\"",35:"#",36:"$",37:"%",38:"&",39:"’",40:"(",41:")",42:"*",43:"+",44:",",45:"-",46:".",47:"/",48:"0",49:"1",50:"2",51:"3",52:"4",53:"5",54:"6",55:"7",56:"8",57:"9",58:":",59:";",60:"<",61:"=",62:">",63:"?",64:"@",65:"A",66:"B",67:"C",68:"D",69:"E",70:"F",71:"G",72:"H",73:"I",74:"J",75:"K",76:"L",77:"M",78:"N",79:"O",80:"P",81:"Q",82:"R",83:"S",84:"T",85:"U",86:"V",87:"W",88:"X",89:"Y",90:"Z",91:"[",92:"\\",93:"]",94:"^",95:"_",96:"‘",97:"a",98:"b",99:"c",100:"d",101:"e",102:"f",103:"g",104:"h",105:"i",106:"j",107:"k",108:"l",109:"m",110:"n",111:"o",112:"p",113:"q",114:"r",115:"s",116:"t",117:"u",118:"v",119:"w",120:"x",121:"y",122:"z",123:"{",124:"|",125:"}",126:"~",161:"¡",162:"¢",163:"£",164:"⁄",165:"¥",166:"ƒ",167:"§",168:"¤",169:"'",170:"“",171:"«",172:"‹",173:"›",174:"ﬁ",175:"ﬂ",177:"–",178:"†",179:"‡",180:"·",182:"¶",183:"•",184:"‚",185:"„",186:"”",187:"»",188:"…",189:"‰",191:"¿",193:"`",194:"´",195:"ˆ",196:"˜",197:"¯",198:"˘",199:"˙",200:"¨",202:"˚",203:"¸",205:"˝",206:"˛",207:"ˇ",208:"—",225:"Æ",227:"ª",232:"Ł",233:"Ø",234:"Œ",235:"º",241:"æ",245:"ı",248:"ł",249:"ø",250:"œ",251:"ß"}),
  MacRomanEncoding: Object.freeze({32:" ",33:"!",34:"\"",35:"#",36:"$",37:"%",38:"&",39:"'",40:"(",41:")",42:"*",43:"+",44:",",45:"-",46:".",47:"/",48:"0",49:"1",50:"2",51:"3",52:"4",53:"5",54:"6",55:"7",56:"8",57:"9",58:":",59:";",60:"<",61:"=",62:">",63:"?",64:"@",65:"A",66:"B",67:"C",68:"D",69:"E",70:"F",71:"G",72:"H",73:"I",74:"J",75:"K",76:"L",77:"M",78:"N",79:"O",80:"P",81:"Q",82:"R",83:"S",84:"T",85:"U",86:"V",87:"W",88:"X",89:"Y",90:"Z",91:"[",92:"\\",93:"]",94:"^",95:"_",96:"`",97:"a",98:"b",99:"c",100:"d",101:"e",102:"f",103:"g",104:"h",105:"i",106:"j",107:"k",108:"l",109:"m",110:"n",111:"o",112:"p",113:"q",114:"r",115:"s",116:"t",117:"u",118:"v",119:"w",120:"x",121:"y",122:"z",123:"{",124:"|",125:"}",126:"~",128:"Ä",129:"Å",130:"Ç",131:"É",132:"Ñ",133:"Ö",134:"Ü",135:"á",136:"à",137:"â",138:"ä",139:"ã",140:"å",141:"ç",142:"é",143:"è",144:"ê",145:"ë",146:"í",147:"ì",148:"î",149:"ï",150:"ñ",151:"ó",152:"ò",153:"ô",154:"ö",155:"õ",156:"ú",157:"ù",158:"û",159:"ü",160:"†",161:"°",162:"¢",163:"£",164:"§",165:"•",166:"¶",167:"ß",168:"®",169:"©",170:"™",171:"´",172:"¨",174:"Æ",175:"Ø",177:"±",180:"¥",181:"µ",187:"ª",188:"º",190:"æ",191:"ø",192:"¿",193:"¡",194:"¬",196:"ƒ",199:"«",200:"»",201:"…",202:" ",203:"À",204:"Ã",205:"Õ",206:"Œ",207:"œ",208:"–",209:"—",210:"“",211:"”",212:"‘",213:"’",214:"÷",216:"ÿ",217:"Ÿ",218:"⁄",219:"¤",220:"‹",221:"›",222:"ﬁ",223:"ﬂ",224:"‡",225:"·",226:"‚",227:"„",228:"‰",229:"Â",230:"Ê",231:"Á",232:"Ë",233:"È",234:"Í",235:"Î",236:"Ï",237:"Ì",238:"Ó",239:"Ô",241:"Ò",242:"Ú",243:"Û",244:"Ù",245:"ı",246:"ˆ",247:"˜",248:"¯",249:"˘",250:"˙",251:"˚",252:"¸",253:"˝",254:"˛",255:"ˇ"}),
  WinAnsiEncoding: Object.freeze({32:" ",33:"!",34:"\"",35:"#",36:"$",37:"%",38:"&",39:"'",40:"(",41:")",42:"*",43:"+",44:",",45:"-",46:".",47:"/",48:"0",49:"1",50:"2",51:"3",52:"4",53:"5",54:"6",55:"7",56:"8",57:"9",58:":",59:";",60:"<",61:"=",62:">",63:"?",64:"@",65:"A",66:"B",67:"C",68:"D",69:"E",70:"F",71:"G",72:"H",73:"I",74:"J",75:"K",76:"L",77:"M",78:"N",79:"O",80:"P",81:"Q",82:"R",83:"S",84:"T",85:"U",86:"V",87:"W",88:"X",89:"Y",90:"Z",91:"[",92:"\\",93:"]",94:"^",95:"_",96:"`",97:"a",98:"b",99:"c",100:"d",101:"e",102:"f",103:"g",104:"h",105:"i",106:"j",107:"k",108:"l",109:"m",110:"n",111:"o",112:"p",113:"q",114:"r",115:"s",116:"t",117:"u",118:"v",119:"w",120:"x",121:"y",122:"z",123:"{",124:"|",125:"}",126:"~",128:"€",130:"‚",131:"ƒ",132:"„",133:"…",134:"†",135:"‡",136:"ˆ",137:"‰",138:"Š",139:"‹",140:"Œ",142:"Ž",145:"‘",146:"’",147:"“",148:"”",149:"•",150:"–",151:"—",152:"˜",153:"™",154:"š",155:"›",156:"œ",158:"ž",159:"Ÿ",160:" ",161:"¡",162:"¢",163:"£",164:"¤",165:"¥",166:"¦",167:"§",168:"¨",169:"©",170:"ª",171:"«",172:"¬",173:"­",174:"®",175:"¯",176:"°",177:"±",178:"²",179:"³",180:"´",181:"µ",182:"¶",183:"·",184:"¸",185:"¹",186:"º",187:"»",188:"¼",189:"½",190:"¾",191:"¿",192:"À",193:"Á",194:"Â",195:"Ã",196:"Ä",197:"Å",198:"Æ",199:"Ç",200:"È",201:"É",202:"Ê",203:"Ë",204:"Ì",205:"Í",206:"Î",207:"Ï",208:"Ð",209:"Ñ",210:"Ò",211:"Ó",212:"Ô",213:"Õ",214:"Ö",215:"×",216:"Ø",217:"Ù",218:"Ú",219:"Û",220:"Ü",221:"Ý",222:"Þ",223:"ß",224:"à",225:"á",226:"â",227:"ã",228:"ä",229:"å",230:"æ",231:"ç",232:"è",233:"é",234:"ê",235:"ë",236:"ì",237:"í",238:"î",239:"ï",240:"ð",241:"ñ",242:"ò",243:"ó",244:"ô",245:"õ",246:"ö",247:"÷",248:"ø",249:"ù",250:"ú",251:"û",252:"ü",253:"ý",254:"þ",255:"ÿ"}),
});

/** Latin glyph names (Adobe Glyph List subset used by the base encodings). */
export const PDF_SIMPLE_FONT_GLYPH_NAMES: Readonly<Record<string, string>> =
  Object.freeze({"A":"A","AE":"Æ","Aacute":"Á","Acircumflex":"Â","Adieresis":"Ä","Agrave":"À","Aring":"Å","Atilde":"Ã","B":"B","C":"C","Ccedilla":"Ç","D":"D","E":"E","Eacute":"É","Ecircumflex":"Ê","Edieresis":"Ë","Egrave":"È","Eth":"Ð","Euro":"€","F":"F","G":"G","H":"H","I":"I","Iacute":"Í","Icircumflex":"Î","Idieresis":"Ï","Igrave":"Ì","J":"J","K":"K","L":"L","Lslash":"Ł","M":"M","N":"N","Ntilde":"Ñ","O":"O","OE":"Œ","Oacute":"Ó","Ocircumflex":"Ô","Odieresis":"Ö","Ograve":"Ò","Oslash":"Ø","Otilde":"Õ","P":"P","Q":"Q","R":"R","S":"S","Scaron":"Š","T":"T","Thorn":"Þ","U":"U","Uacute":"Ú","Ucircumflex":"Û","Udieresis":"Ü","Ugrave":"Ù","V":"V","W":"W","X":"X","Y":"Y","Yacute":"Ý","Ydieresis":"Ÿ","Z":"Z","Zcaron":"Ž","a":"a","aacute":"á","acircumflex":"â","acute":"´","adieresis":"ä","ae":"æ","agrave":"à","ampersand":"&","aring":"å","asciicircum":"^","asciitilde":"~","asterisk":"*","at":"@","atilde":"ã","b":"b","backslash":"\\","bar":"|","braceleft":"{","braceright":"}","bracketleft":"[","bracketright":"]","breve":"˘","brokenbar":"¦","bullet":"•","c":"c","caron":"ˇ","ccedilla":"ç","cedilla":"¸","cent":"¢","circumflex":"ˆ","colon":":","comma":",","copyright":"©","currency":"¤","d":"d","dagger":"†","daggerdbl":"‡","degree":"°","dieresis":"¨","divide":"÷","dollar":"$","dotaccent":"˙","dotlessi":"ı","e":"e","eacute":"é","ecircumflex":"ê","edieresis":"ë","egrave":"è","eight":"8","ellipsis":"…","emdash":"—","endash":"–","equal":"=","eth":"ð","exclam":"!","exclamdown":"¡","f":"f","fi":"ﬁ","five":"5","fl":"ﬂ","florin":"ƒ","four":"4","fraction":"⁄","g":"g","germandbls":"ß","grave":"`","greater":">","guillemotleft":"«","guillemotright":"»","guilsinglleft":"‹","guilsinglright":"›","h":"h","hungarumlaut":"˝","hyphen":"-","i":"i","iacute":"í","icircumflex":"î","idieresis":"ï","igrave":"ì","j":"j","k":"k","l":"l","less":"<","logicalnot":"¬","lslash":"ł","m":"m","macron":"¯","minus":"−","mu":"µ","multiply":"×","n":"n","nbspace":" ","nine":"9","ntilde":"ñ","numbersign":"#","o":"o","oacute":"ó","ocircumflex":"ô","odieresis":"ö","oe":"œ","ogonek":"˛","ograve":"ò","one":"1","onehalf":"½","onequarter":"¼","onesuperior":"¹","ordfeminine":"ª","ordmasculine":"º","oslash":"ø","otilde":"õ","p":"p","paragraph":"¶","parenleft":"(","parenright":")","percent":"%","period":".","periodcentered":"·","perthousand":"‰","plus":"+","plusminus":"±","q":"q","question":"?","questiondown":"¿","quotedbl":"\"","quotedblbase":"„","quotedblleft":"“","quotedblright":"”","quoteleft":"‘","quoteright":"’","quotesinglbase":"‚","quotesingle":"'","r":"r","registered":"®","ring":"˚","s":"s","scaron":"š","section":"§","semicolon":";","seven":"7","six":"6","slash":"/","space":" ","sterling":"£","t":"t","thorn":"þ","three":"3","threequarters":"¾","threesuperior":"³","tilde":"˜","trademark":"™","two":"2","twosuperior":"²","u":"u","uacute":"ú","ucircumflex":"û","udieresis":"ü","ugrave":"ù","underscore":"_","v":"v","w":"w","x":"x","y":"y","yacute":"ý","ydieresis":"ÿ","yen":"¥","z":"z","zcaron":"ž","zero":"0"});

export function isPdfSimpleFontBaseEncodingName(
  value: string,
): value is PdfSimpleFontBaseEncodingName {
  return (
    value === "StandardEncoding" ||
    value === "MacRomanEncoding" ||
    value === "WinAnsiEncoding"
  );
}

/**
 * Glyph name behind every StandardEncoding code, frozen from the Adobe
 * Standard 14 AFM encoding vector (as ReportLab ships it) so a Standard 14
 * font drawn without /Encoding measures the same glyph here and on the
 * server. Codes match `PDF_SIMPLE_FONT_BASE_ENCODINGS.StandardEncoding`.
 */
export const PDF_STANDARD_ENCODING_GLYPH_NAMES: Readonly<Record<number, string>> =
  Object.freeze({32:"space",33:"exclam",34:"quotedbl",35:"numbersign",36:"dollar",37:"percent",38:"ampersand",39:"quoteright",40:"parenleft",41:"parenright",42:"asterisk",43:"plus",44:"comma",45:"hyphen",46:"period",47:"slash",48:"zero",49:"one",50:"two",51:"three",52:"four",53:"five",54:"six",55:"seven",56:"eight",57:"nine",58:"colon",59:"semicolon",60:"less",61:"equal",62:"greater",63:"question",64:"at",65:"A",66:"B",67:"C",68:"D",69:"E",70:"F",71:"G",72:"H",73:"I",74:"J",75:"K",76:"L",77:"M",78:"N",79:"O",80:"P",81:"Q",82:"R",83:"S",84:"T",85:"U",86:"V",87:"W",88:"X",89:"Y",90:"Z",91:"bracketleft",92:"backslash",93:"bracketright",94:"asciicircum",95:"underscore",96:"quoteleft",97:"a",98:"b",99:"c",100:"d",101:"e",102:"f",103:"g",104:"h",105:"i",106:"j",107:"k",108:"l",109:"m",110:"n",111:"o",112:"p",113:"q",114:"r",115:"s",116:"t",117:"u",118:"v",119:"w",120:"x",121:"y",122:"z",123:"braceleft",124:"bar",125:"braceright",126:"asciitilde",161:"exclamdown",162:"cent",163:"sterling",164:"fraction",165:"yen",166:"florin",167:"section",168:"currency",169:"quotesingle",170:"quotedblleft",171:"guillemotleft",172:"guilsinglleft",173:"guilsinglright",174:"fi",175:"fl",177:"endash",178:"dagger",179:"daggerdbl",180:"periodcentered",182:"paragraph",183:"bullet",184:"quotesinglbase",185:"quotedblbase",186:"quotedblright",187:"guillemotright",188:"ellipsis",189:"perthousand",191:"questiondown",193:"grave",194:"acute",195:"circumflex",196:"tilde",197:"macron",198:"breve",199:"dotaccent",200:"dieresis",202:"ring",203:"cedilla",205:"hungarumlaut",206:"ogonek",207:"caron",208:"emdash",225:"AE",227:"ordfeminine",232:"Lslash",233:"Oslash",234:"OE",235:"ordmasculine",241:"ae",245:"dotlessi",248:"lslash",249:"oslash",250:"oe",251:"germandbls"});

/**
 * Glyph name behind every WinAnsiEncoding code, frozen from
 * @pdf-lib/standard-fonts 1.0.0 (byte 0x9F is the uppercase Ydieresis the
 * Standard 14 metrics name). Codes match
 * `PDF_SIMPLE_FONT_BASE_ENCODINGS.WinAnsiEncoding`, whose characters keep the
 * cp1252 values 0xA0 U+00A0 and 0xAD U+00AD that this vector draws as space
 * and hyphen.
 */
export const PDF_WINANSI_ENCODING_GLYPH_NAMES: Readonly<Record<number, string>> =
  Object.freeze({32:"space",33:"exclam",34:"quotedbl",35:"numbersign",36:"dollar",37:"percent",38:"ampersand",39:"quotesingle",40:"parenleft",41:"parenright",42:"asterisk",43:"plus",44:"comma",45:"hyphen",46:"period",47:"slash",48:"zero",49:"one",50:"two",51:"three",52:"four",53:"five",54:"six",55:"seven",56:"eight",57:"nine",58:"colon",59:"semicolon",60:"less",61:"equal",62:"greater",63:"question",64:"at",65:"A",66:"B",67:"C",68:"D",69:"E",70:"F",71:"G",72:"H",73:"I",74:"J",75:"K",76:"L",77:"M",78:"N",79:"O",80:"P",81:"Q",82:"R",83:"S",84:"T",85:"U",86:"V",87:"W",88:"X",89:"Y",90:"Z",91:"bracketleft",92:"backslash",93:"bracketright",94:"asciicircum",95:"underscore",96:"grave",97:"a",98:"b",99:"c",100:"d",101:"e",102:"f",103:"g",104:"h",105:"i",106:"j",107:"k",108:"l",109:"m",110:"n",111:"o",112:"p",113:"q",114:"r",115:"s",116:"t",117:"u",118:"v",119:"w",120:"x",121:"y",122:"z",123:"braceleft",124:"bar",125:"braceright",126:"asciitilde",128:"Euro",130:"quotesinglbase",131:"florin",132:"quotedblbase",133:"ellipsis",134:"dagger",135:"daggerdbl",136:"circumflex",137:"perthousand",138:"Scaron",139:"guilsinglleft",140:"OE",142:"Zcaron",145:"quoteleft",146:"quoteright",147:"quotedblleft",148:"quotedblright",149:"bullet",150:"endash",151:"emdash",152:"tilde",153:"trademark",154:"scaron",155:"guilsinglright",156:"oe",158:"zcaron",159:"Ydieresis",160:"space",161:"exclamdown",162:"cent",163:"sterling",164:"currency",165:"yen",166:"brokenbar",167:"section",168:"dieresis",169:"copyright",170:"ordfeminine",171:"guillemotleft",172:"logicalnot",173:"hyphen",174:"registered",175:"macron",176:"degree",177:"plusminus",178:"twosuperior",179:"threesuperior",180:"acute",181:"mu",182:"paragraph",183:"periodcentered",184:"cedilla",185:"onesuperior",186:"ordmasculine",187:"guillemotright",188:"onequarter",189:"onehalf",190:"threequarters",191:"questiondown",192:"Agrave",193:"Aacute",194:"Acircumflex",195:"Atilde",196:"Adieresis",197:"Aring",198:"AE",199:"Ccedilla",200:"Egrave",201:"Eacute",202:"Ecircumflex",203:"Edieresis",204:"Igrave",205:"Iacute",206:"Icircumflex",207:"Idieresis",208:"Eth",209:"Ntilde",210:"Ograve",211:"Oacute",212:"Ocircumflex",213:"Otilde",214:"Odieresis",215:"multiply",216:"Oslash",217:"Ugrave",218:"Uacute",219:"Ucircumflex",220:"Udieresis",221:"Yacute",222:"Thorn",223:"germandbls",224:"agrave",225:"aacute",226:"acircumflex",227:"atilde",228:"adieresis",229:"aring",230:"ae",231:"ccedilla",232:"egrave",233:"eacute",234:"ecircumflex",235:"edieresis",236:"igrave",237:"iacute",238:"icircumflex",239:"idieresis",240:"eth",241:"ntilde",242:"ograve",243:"oacute",244:"ocircumflex",245:"otilde",246:"odieresis",247:"divide",248:"oslash",249:"ugrave",250:"uacute",251:"ucircumflex",252:"udieresis",253:"yacute",254:"thorn",255:"ydieresis"});

/**
 * Glyph name behind every MacRomanEncoding code, frozen from the Adobe
 * encoding vector (as ReportLab ships it). Codes match
 * `PDF_SIMPLE_FONT_BASE_ENCODINGS.MacRomanEncoding`.
 */
export const PDF_MACROMAN_ENCODING_GLYPH_NAMES: Readonly<Record<number, string>> =
  Object.freeze({32:"space",33:"exclam",34:"quotedbl",35:"numbersign",36:"dollar",37:"percent",38:"ampersand",39:"quotesingle",40:"parenleft",41:"parenright",42:"asterisk",43:"plus",44:"comma",45:"hyphen",46:"period",47:"slash",48:"zero",49:"one",50:"two",51:"three",52:"four",53:"five",54:"six",55:"seven",56:"eight",57:"nine",58:"colon",59:"semicolon",60:"less",61:"equal",62:"greater",63:"question",64:"at",65:"A",66:"B",67:"C",68:"D",69:"E",70:"F",71:"G",72:"H",73:"I",74:"J",75:"K",76:"L",77:"M",78:"N",79:"O",80:"P",81:"Q",82:"R",83:"S",84:"T",85:"U",86:"V",87:"W",88:"X",89:"Y",90:"Z",91:"bracketleft",92:"backslash",93:"bracketright",94:"asciicircum",95:"underscore",96:"grave",97:"a",98:"b",99:"c",100:"d",101:"e",102:"f",103:"g",104:"h",105:"i",106:"j",107:"k",108:"l",109:"m",110:"n",111:"o",112:"p",113:"q",114:"r",115:"s",116:"t",117:"u",118:"v",119:"w",120:"x",121:"y",122:"z",123:"braceleft",124:"bar",125:"braceright",126:"asciitilde",128:"Adieresis",129:"Aring",130:"Ccedilla",131:"Eacute",132:"Ntilde",133:"Odieresis",134:"Udieresis",135:"aacute",136:"agrave",137:"acircumflex",138:"adieresis",139:"atilde",140:"aring",141:"ccedilla",142:"eacute",143:"egrave",144:"ecircumflex",145:"edieresis",146:"iacute",147:"igrave",148:"icircumflex",149:"idieresis",150:"ntilde",151:"oacute",152:"ograve",153:"ocircumflex",154:"odieresis",155:"otilde",156:"uacute",157:"ugrave",158:"ucircumflex",159:"udieresis",160:"dagger",161:"degree",162:"cent",163:"sterling",164:"section",165:"bullet",166:"paragraph",167:"germandbls",168:"registered",169:"copyright",170:"trademark",171:"acute",172:"dieresis",174:"AE",175:"Oslash",177:"plusminus",180:"yen",181:"mu",187:"ordfeminine",188:"ordmasculine",190:"ae",191:"oslash",192:"questiondown",193:"exclamdown",194:"logicalnot",196:"florin",199:"guillemotleft",200:"guillemotright",201:"ellipsis",202:"space",203:"Agrave",204:"Atilde",205:"Otilde",206:"OE",207:"oe",208:"endash",209:"emdash",210:"quotedblleft",211:"quotedblright",212:"quoteleft",213:"quoteright",214:"divide",216:"ydieresis",217:"Ydieresis",218:"fraction",219:"currency",220:"guilsinglleft",221:"guilsinglright",222:"fi",223:"fl",224:"daggerdbl",225:"periodcentered",226:"quotesinglbase",227:"quotedblbase",228:"perthousand",229:"Acircumflex",230:"Ecircumflex",231:"Aacute",232:"Edieresis",233:"Egrave",234:"Iacute",235:"Icircumflex",236:"Idieresis",237:"Igrave",238:"Oacute",239:"Ocircumflex",241:"Ograve",242:"Uacute",243:"Ucircumflex",244:"Ugrave",245:"dotlessi",246:"circumflex",247:"tilde",248:"macron",249:"breve",250:"dotaccent",251:"ring",252:"cedilla",253:"hungarumlaut",254:"ogonek",255:"caron"});

/** The glyph each base encoding names for a code. A Standard 14 font is
 * measured by these names: the AFM keeps its widths by glyph, not by code. */
export const PDF_SIMPLE_FONT_BASE_ENCODING_GLYPH_NAMES: Readonly<
  Record<PdfSimpleFontBaseEncodingName, Readonly<Record<number, string>>>
> = Object.freeze({
  StandardEncoding: PDF_STANDARD_ENCODING_GLYPH_NAMES,
  MacRomanEncoding: PDF_MACROMAN_ENCODING_GLYPH_NAMES,
  WinAnsiEncoding: PDF_WINANSI_ENCODING_GLYPH_NAMES,
});

const BASE_ENCODING_GLYPH_NAMES: ReadonlySet<string> = new Set(
  Object.values(PDF_SIMPLE_FONT_BASE_ENCODING_GLYPH_NAMES).flatMap((table) =>
    Object.values(table),
  ),
);

/**
 * Whether the Standard 14 AFM width of one glyph name is authority on both
 * sides of the engine. The browser reads the AFMs through pdf-lib and the
 * server through ReportLab; the names the three base encodings use are the
 * glyphs both stores carry with the same widths, so a /Differences name
 * outside them costs its own codes rather than the whole font.
 */
export function isPdfStandard14MeasurableGlyphName(name: string): boolean {
  return BASE_ENCODING_GLYPH_NAMES.has(name);
}

/** Unicode character for one /Differences glyph name, or null when unknown. */
/** Resolve one glyph name, optionally against a wider list than the base
 * encodings use.
 *
 * A variant suffix after `.` and a `_` joining several glyphs are part of the
 * Adobe Glyph List algorithm rather than of any table, so they are handled
 * here once for every caller: the /Differences of an ordinary font and the
 * built-in encoding of a symbolic one resolve a name the same way. */
export function pdfSimpleFontGlyphNameToCharacter(
  name: string,
  wider?: Readonly<Record<string, string>>,
): string | null {
  const base = name.split(".")[0];
  if (!base) return null;
  let text = "";
  for (const part of base.split("_")) {
    const resolved = glyphPartToCharacter(part, wider);
    if (resolved === null) return null;
    text += resolved;
  }
  return text || null;
}

function glyphPartToCharacter(
  part: string,
  wider?: Readonly<Record<string, string>>,
): string | null {
  const known = PDF_SIMPLE_FONT_GLYPH_NAMES[part];
  if (known !== undefined) return known;
  const uni = /^uni((?:[0-9A-F]{4})+)$/.exec(part);
  if (uni) {
    let text = "";
    for (const hex of uni[1].match(/.{4}/gu) ?? []) {
      const codePoint = Number.parseInt(hex, 16);
      if (codePoint >= 0xd800 && codePoint <= 0xdfff) return null;
      text += String.fromCodePoint(codePoint);
    }
    return text;
  }
  const u = /^u([0-9A-F]{4,6})$/.exec(part);
  if (u) {
    const codePoint = Number.parseInt(u[1], 16);
    return codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff)
      ? null
      : String.fromCodePoint(codePoint);
  }
  if (wider && Object.prototype.hasOwnProperty.call(wider, part)) {
    const text = wider[part];
    return typeof text === "string" && text ? text : null;
  }
  return null;
}

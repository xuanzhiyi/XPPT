// pptx_viewer.js — all logic extracted from pptx_viewer.html
// Required because Chrome MV3 extensions block inline scripts.

let slides = [];
let currentSlide = 0;
let slideW = 960, slideH = 540;
let themeColors = {};

// ── colour helpers ────────────────────────────────────────────────────────────

function hexToRgb(hex) {
  hex = hex.replace('#','');
  if (hex.length === 3) hex = hex.split('').map(c=>c+c).join('');
  const n = parseInt(hex, 16);
  return [n>>16&255, n>>8&255, n&255];
}

function rgbToHex(r,g,b) {
  return '#' + [r,g,b].map(v=>Math.max(0,Math.min(255,Math.round(v))).toString(16).padStart(2,'0')).join('');
}

function applyColorMods(hex, mods) {
  let [r,g,b] = hexToRgb(hex);
  const rn=r/255, gn=g/255, bn=b/255;
  const max=Math.max(rn,gn,bn), min=Math.min(rn,gn,bn);
  let h=0, s=0, l=(max+min)/2;
  if (max!==min) {
    const d=max-min;
    s=l>0.5?d/(2-max-min):d/(max+min);
    switch(max){case rn:h=(gn-bn)/d+(gn<bn?6:0);break;case gn:h=(bn-rn)/d+2;break;case bn:h=(rn-gn)/d+4;break;}
    h/=6;
  }
  if (mods.lumMod !== undefined) l = l * (mods.lumMod/100000);
  if (mods.lumOff !== undefined) l = l + (mods.lumOff/100000);
  if (mods.shade  !== undefined) l = l * (mods.shade/100000);
  if (mods.tint   !== undefined) l = l + (1-l)*(mods.tint/100000);
  l = Math.max(0, Math.min(1, l));
  let nr,ng,nb;
  if (s===0) { nr=ng=nb=l; }
  else {
    const hue2rgb=(p,q,t)=>{if(t<0)t+=1;if(t>1)t-=1;if(t<1/6)return p+(q-p)*6*t;if(t<1/2)return q;if(t<2/3)return p+(q-p)*(2/3-t)*6;return p;};
    const q=l<0.5?l*(1+s):l+s-l*s, p=2*l-q;
    nr=hue2rgb(p,q,h+1/3); ng=hue2rgb(p,q,h); nb=hue2rgb(p,q,h-1/3);
  }
  return rgbToHex(nr*255, ng*255, nb*255);
}

function resolveColorEl(el, theme) {
  if (!el) return null;
  const srgb = el.querySelector('srgbClr');
  if (srgb) {
    const val = srgb.getAttribute('val');
    const mods = extractColorMods(srgb);
    return Object.keys(mods).length ? applyColorMods('#'+val, mods) : '#'+val;
  }
  const sc = el.querySelector('schemeClr');
  if (sc) {
    const raw = sc.getAttribute('val');
    const aliasMap = {bg1:'lt1', bg2:'lt2', tx1:'dk1', tx2:'dk2'};
    const name = aliasMap[raw] || raw;
    const base = theme[name] || defaultScheme[name] || '#000000';
    const mods = extractColorMods(sc);
    return Object.keys(mods).length ? applyColorMods(base, mods) : base;
  }
  const pc = el.querySelector('prstClr');
  if (pc) return cssNamedColor(pc.getAttribute('val'));
  const sysc = el.querySelector('sysClr');
  if (sysc) return '#' + (sysc.getAttribute('lastClr') || '000000');
  return null;
}

function extractColorMods(el) {
  const mods = {};
  for (const m of ['lumMod','lumOff','shade','tint','alpha']) {
    const ch = el.querySelector(m);
    if (ch) mods[m] = parseInt(ch.getAttribute('val')||'0');
  }
  return mods;
}

const defaultScheme = {
  dk1:'#000000', lt1:'#FFFFFF', dk2:'#1F3864', lt2:'#E7E6E6',
  accent1:'#4472C4', accent2:'#ED7D31', accent3:'#A9D18E',
  accent4:'#FFC000', accent5:'#5B9BD5', accent6:'#70AD47',
  hlink:'#0563C1', folHlink:'#954F72'
};

// Map Symbol / Wingdings Private-Use-Area bullet glyphs → real Unicode characters.
// PowerPoint stores bullets as PUA codepoints (U+F0xx) from those legacy fonts.
const PUA_BULLET_MAP = {
  0xF0B7: '•', // • bullet
  0xF0FC: '✓', // ✓ check mark
  0xF0A7: '▪', // ▪ small square
  0xF076: '❖', // ❖ diamond
  0xF0D8: '➤', // ➤ right arrow
  0xF0E8: '✔', // ✔ heavy check
  0xF06C: '●', // ● filled circle
  0xF0D0: '◆', // ◆ diamond
  0xF0A8: '▫', // ▫ small empty square
  0xF0AE: '►', // ► right-pointing triangle
  0xF0B2: '◄', // ◄ left-pointing triangle
  0xF0BC: '★', // ★ star
};

function normalizeBulletChar(char, fontFace) {
  if (!char) return null;
  const code = char.codePointAt(0);
  // PUA range used by Symbol / Wingdings
  if (code >= 0xE000 && code <= 0xF8FF) {
    return PUA_BULLET_MAP[code] || '•'; // fall back to standard bullet
  }
  return char;
}

function cssNamedColor(name) {
  const m = {black:'#000000',white:'#FFFFFF',red:'#FF0000',green:'#00FF00',blue:'#0000FF',
    yellow:'#FFFF00',cyan:'#00FFFF',magenta:'#FF00FF',gray:'#808080',grey:'#808080',
    orange:'#FFA500',purple:'#800080',brown:'#A52A2A',pink:'#FFC0CB',lime:'#00FF00',
    navy:'#000080',teal:'#008080',maroon:'#800000',silver:'#C0C0C0',gold:'#FFD700'};
  return m[name?.toLowerCase()] || '#000000';
}

// ── path helper ───────────────────────────────────────────────────────────────

function resolveTarget(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

// ── theme parsing ─────────────────────────────────────────────────────────────

async function parseTheme(zip) {
  const colors = {};
  try {
    const presRels = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
    if (!presRels) return colors;
    const rDoc = new DOMParser().parseFromString(presRels,'application/xml');
    for (const rel of rDoc.querySelectorAll('Relationship')) {
      if (rel.getAttribute('Type')?.includes('/theme')) {
        const tPath = resolveTarget('ppt', rel.getAttribute('Target'));
        const tXml = await zip.file(tPath)?.async('string');
        if (!tXml) continue;
        const tDoc = new DOMParser().parseFromString(tXml,'application/xml');
        const slots = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
        const clrScheme = tDoc.querySelector('clrScheme');
        if (!clrScheme) continue;
        for (const slot of slots) {
          const el = clrScheme.querySelector(slot);
          if (!el) continue;
          const srgb = el.querySelector('srgbClr');
          if (srgb) { colors[slot] = '#' + srgb.getAttribute('val'); continue; }
          const sys = el.querySelector('sysClr');
          if (sys) { colors[slot] = '#' + (sys.getAttribute('lastClr') || '000000'); continue; }
        }
        break;
      }
    }
  } catch(e) {}
  return colors;
}

// ── slide background ──────────────────────────────────────────────────────────

async function parseSlideBg(zip, slideXmlPath, theme) {
  const result = {color:'#FFFFFF', gradStops:null, imgData:null};
  const chain = [slideXmlPath];
  let cur = slideXmlPath;
  for (let depth = 0; depth < 1; depth++) {
    const dir = cur.split('/').slice(0,-1).join('/');
    const fname = cur.split('/').pop();
    const relsPath = `${dir}/_rels/${fname}.rels`;
    const rFile = zip.file(relsPath);
    if (!rFile) break;
    const rx = await rFile.async('string');
    const rd = new DOMParser().parseFromString(rx,'application/xml');
    let next = null;
    for (const rel of rd.querySelectorAll('Relationship')) {
      if ((rel.getAttribute('Type')||'').includes('slideLayout')) {
        next = resolveTarget(dir, rel.getAttribute('Target'));
        break;
      }
    }
    if (!next) break;
    chain.push(next);
    cur = next;
  }
  for (const path of chain) {
    const f = zip.file(path);
    if (!f) continue;
    const xml = await f.async('string');
    const doc = new DOMParser().parseFromString(xml,'application/xml');
    const bg = doc.querySelector('bg');
    if (!bg) continue;
    const bgPr = bg.querySelector('bgPr');
    if (bgPr) {
      const solid = bgPr.querySelector('solidFill');
      if (solid) { result.color = resolveColorEl(solid, theme) || '#FFFFFF'; result.gradStops = null; return result; }
      const grad = bgPr.querySelector('gradFill');
      if (grad) { result.gradStops = parseGradStops(grad, theme); result.color = null; return result; }
      const blip = bgPr.querySelector('blipFill blip');
      if (blip) {
        const rId = blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','embed');
        const relsPath2 = path.replace(/(slides[^/]*)\/(slide[^/]+)\.xml/, '$1/_rels/$2.xml.rels');
        const imgData = await resolveRIdToDataUrl(zip, relsPath2, rId);
        if (imgData) { result.imgData = imgData; result.color = null; return result; }
      }
    }
    const bgRef = bg.querySelector('bgRef');
    if (bgRef) {
      const c = resolveColorEl(bgRef, theme);
      if (c && c !== '#FFFFFF') { result.color = c; return result; }
    }
  }
  return result;
}

async function resolveRIdToDataUrl(zip, relsPath, rId) {
  if (!rId) return null;
  const rf = zip.file(relsPath);
  if (!rf) return null;
  const rx = await rf.async('string');
  const rd = new DOMParser().parseFromString(rx,'application/xml');
  for (const rel of rd.querySelectorAll('Relationship')) {
    if (rel.getAttribute('Id') !== rId) continue;
    const tgt = rel.getAttribute('Target');
    const relsDir = relsPath.split('/').slice(0,-1).join('/').replace('/_rels','');
    const imgPath = resolveTarget(relsDir, tgt);
    const imgFile = zip.file(imgPath);
    if (!imgFile) return null;
    const b64 = await imgFile.async('base64');
    const ext2 = tgt.split('.').pop().toLowerCase();
    const mimeMap = {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',svg:'image/svg+xml',emf:'image/x-emf',wmf:'image/x-wmf'};
    return `data:${mimeMap[ext2]||'image/png'};base64,${b64}`;
  }
  return null;
}

// ── gradient helpers ──────────────────────────────────────────────────────────

function parseGradStops(gradFill, theme) {
  const stops = [];
  for (const gs of gradFill.querySelectorAll('gs')) {
    const pos = parseInt(gs.getAttribute('pos')||'0') / 100000;
    const color = resolveColorEl(gs, theme) || '#000000';
    stops.push({pos, color});
  }
  stops.sort((a,b)=>a.pos-b.pos);
  const lin = gradFill.querySelector('lin');
  const angle = lin ? parseInt(lin.getAttribute('ang')||'0') / 60000 : 0;
  return {stops, angle};
}

function gradStopsToCss(gradStops) {
  if (!gradStops || !gradStops.stops.length) return null;
  const {stops, angle} = gradStops;
  const cssAngle = (angle + 90) % 360;
  const parts = stops.map(s => `${s.color} ${(s.pos*100).toFixed(1)}%`);
  return `linear-gradient(${cssAngle}deg, ${parts.join(', ')})`;
}

// ── fill resolver ─────────────────────────────────────────────────────────────

function resolveFill(el, theme) {
  for (const child of el.children) {
    const n = child.localName;
    if (n === 'noFill') return {none: true};
    if (n === 'solidFill') return {color: resolveColorEl(child, theme)};
    if (n === 'gradFill')  return {grad: parseGradStops(child, theme)};
    if (n === 'grpFill')   return {grpFill: true};
    if (n === 'pattFill')  {
      const fg = child.querySelector('fgClr');
      return {color: resolveColorEl(fg, theme) || '#000000'};
    }
  }
  return {};
}

// ── EMU conversion ────────────────────────────────────────────────────────────

function emu(v) { return parseInt(v||0) / 914400 * 96; }

// ── shape text parser ─────────────────────────────────────────────────────────

function parseTextBody(txBody, theme) {
  if (!txBody) return null;
  const bodyPr = txBody.querySelector('bodyPr');
  const vAlign = bodyPr?.getAttribute('anchor') || 't';
  const insetL = emu(bodyPr?.getAttribute('lIns') ?? 91440);
  const insetR = emu(bodyPr?.getAttribute('rIns') ?? 91440);
  const insetT = emu(bodyPr?.getAttribute('tIns') ?? 45720);
  const insetB = emu(bodyPr?.getAttribute('bIns') ?? 45720);
  const wordWrap = bodyPr?.getAttribute('wrap') !== 'none';
  const vert = bodyPr?.getAttribute('vert');
  const paras = [];
  for (const p of txBody.querySelectorAll(':scope > p')) {
    const pPr = p.querySelector('pPr');
    let pAlign = pPr?.getAttribute('algn') || 'l';
    pAlign = {l:'left',ctr:'center',r:'right',just:'justify'}[pAlign] || 'left';
    const lvl = parseInt(pPr?.getAttribute('lvl')||'0');
    const buChar = normalizeBulletChar(
      pPr?.querySelector('buChar')?.getAttribute('char') || null,
      pPr?.querySelector('buFont')?.getAttribute('typeface') || null
    );
    const buFont = buChar ? (pPr?.querySelector('buFont')?.getAttribute('typeface') || null) : null;
    const buNone = !!pPr?.querySelector('buNone');
    const buAutoNum = pPr?.querySelector('buAutoNum') || null;
    const buAutoNumType = buAutoNum?.getAttribute('type') || 'arabicPeriod';
    const buClrEl = pPr?.querySelector('buClr');
    const buColor = buClrEl ? resolveColorEl(buClrEl, theme) : null;
    const buSzPct = pPr?.querySelector('buSzPct') ? parseInt(pPr.querySelector('buSzPct').getAttribute('val')||'100000')/1000 : 100;
    const marL  = emu(pPr?.getAttribute('marL') || '0');
    const lnSpc = p.querySelector('pPr lnSpc spcPct');
    const lineSpacing = lnSpc ? parseInt(lnSpc.getAttribute('val')||'120000')/100000 : 1.2;
    const spcBef = p.querySelector('pPr spcBef spcPts');
    const spaceBeforePt = spcBef ? parseInt(spcBef.getAttribute('val')||'0')/100 : 0;
    const spcAft = p.querySelector('pPr spcAft spcPts');
    const spaceAfterPt  = spcAft  ? parseInt(spcAft.getAttribute('val')||'0')/100  : 0;
    const defRPr = pPr?.querySelector('defRPr');
    const runs = [];
    for (const child of p.childNodes) {
      const ln = child.localName;
      if (ln === 'r') {
        const rPr = child.querySelector('rPr');
        const run = parseRunProps(rPr || defRPr, theme);
        run.t = child.querySelector('t')?.textContent || '';
        runs.push(run);
      } else if (ln === 'br') {
        runs.push({t:'\n', isBr:true});
      } else if (ln === 'fld') {
        const rPr = child.querySelector('rPr');
        const run = parseRunProps(rPr, theme);
        run.t = child.querySelector('t')?.textContent || '';
        runs.push(run);
      }
    }
    paras.push({runs, align:pAlign, lvl, buChar, buFont, buNone, buAutoNum: !!buAutoNum, buAutoNumType, buColor, buSzPct, marL, lineSpacing, spaceBeforePt, spaceAfterPt});
  }
  return {paras, vAlign, insetL, insetR, insetT, insetB, wordWrap, vert};
}

function parseRunProps(rPr, theme) {
  if (!rPr) return {fc:null, bold:false, italic:false, underline:false, strike:false, fontSize:null, fontFace:null};
  const solidFill = rPr.querySelector('solidFill');
  const fc = resolveColorEl(solidFill, theme) || resolveColorEl(rPr.querySelector('solidFill'), theme);
  const sz = rPr.getAttribute('sz');
  const latin = rPr.querySelector('latin');
  return {
    fc,
    bold:      rPr.getAttribute('b') === '1',
    italic:    rPr.getAttribute('i') === '1',
    underline: rPr.getAttribute('u') && rPr.getAttribute('u') !== 'none',
    strike:    rPr.getAttribute('strike') && rPr.getAttribute('strike') !== 'noStrike',
    fontSize:  sz ? parseInt(sz)/100 : null,
    fontFace:  latin?.getAttribute('typeface') || null,
    caps:      rPr.getAttribute('cap') === 'all' || rPr.getAttribute('cap') === 'small',
    smallCaps: rPr.getAttribute('cap') === 'small',
    baseline:  parseInt(rPr.getAttribute('baseline')||'0'),
  };
}

// ── connector / line ──────────────────────────────────────────────────────────

function parseCxnSp(cxnSp, theme) {
  const xfrm = cxnSp.querySelector('xfrm');
  if (!xfrm) return null;
  const off = xfrm.querySelector('off'), ext = xfrm.querySelector('ext');
  if (!off||!ext) return null;
  const x = emu(off.getAttribute('x')), y = emu(off.getAttribute('y'));
  const w = emu(ext.getAttribute('cx')), h = emu(ext.getAttribute('cy'));
  const rot = parseInt(xfrm.getAttribute('rot')||'0')/60000;
  const flipH = xfrm.getAttribute('flipH')==='1', flipV = xfrm.getAttribute('flipV')==='1';
  const spPr = cxnSp.querySelector('spPr');
  const ln = spPr?.querySelector('ln');
  const strokeColor = resolveColorEl(ln?.querySelector('solidFill'), theme) || '#000000';
  const strokeW = Math.max(0.5, parseInt(ln?.getAttribute('w')||'12700')/12700);
  const geom = cxnSp.querySelector('prstGeom')?.getAttribute('prst') || 'line';
  return {type:'cxn', x,y,w,h,rot,flipH,flipV,strokeColor,strokeW,geom};
}

// ── table ─────────────────────────────────────────────────────────────────────

function parseTbl(graphicFrame, theme) {
  const xfrm = graphicFrame.querySelector('xfrm');
  if (!xfrm) return null;
  const off = xfrm.querySelector('off'), ext = xfrm.querySelector('ext');
  if (!off||!ext) return null;
  const x = emu(off.getAttribute('x')), y = emu(off.getAttribute('y'));
  const w = emu(ext.getAttribute('cx')), h = emu(ext.getAttribute('cy'));
  const tbl = graphicFrame.querySelector('tbl');
  if (!tbl) return null;
  const rows = [];
  for (const tr of tbl.querySelectorAll('tr')) {
    const rowH = emu(tr.getAttribute('h'));
    const cells = [];
    for (const tc of tr.querySelectorAll('tc')) {
      const tcPr = tc.querySelector('tcPr');
      const fillColor = resolveColorEl(tcPr?.querySelector('solidFill'), theme);
      const lnEl = tcPr?.querySelector('lnL, lnR, lnT, lnB');
      const borderColor = resolveColorEl(lnEl?.querySelector('solidFill'), theme) || '#cccccc';
      cells.push({
        fillColor, borderColor,
        gridSpan: parseInt(tc.getAttribute('gridSpan')||'1'),
        rowSpan:  parseInt(tc.getAttribute('rowSpan')||'1'),
        hMerge: tc.getAttribute('hMerge')==='1',
        vMerge: tc.getAttribute('vMerge')==='1',
        text: parseTextBody(tc.querySelector('txBody'), theme),
      });
    }
    rows.push({rowH, cells});
  }
  const colWidths = [...tbl.querySelectorAll('gridCol')].map(c=>emu(c.getAttribute('w')));
  return {type:'tbl', x,y,w,h,rows,colWidths};
}

// ── ancestor shapes (layout + master) ────────────────────────────────────────

const _ancestorCache = new Map();

async function parseAncestorShapes(zip, slideXmlPath, theme) {
  const result = [];
  let cur = slideXmlPath;
  for (let depth = 0; depth < 2; depth++) {
    const dir  = cur.split('/').slice(0, -1).join('/');
    const fname = cur.split('/').pop();
    const relsPath = `${dir}/_rels/${fname}.rels`;
    const rFile = zip.file(relsPath);
    if (!rFile) break;
    const rx = await rFile.async('string');
    const rd = new DOMParser().parseFromString(rx, 'application/xml');
    let next = null;
    for (const rel of rd.querySelectorAll('Relationship')) {
      const t = rel.getAttribute('Type') || '';
      if (t.includes('slideLayout') || t.includes('slideMaster')) {
        next = resolveTarget(dir, rel.getAttribute('Target'));
        break;
      }
    }
    if (!next) break;
    if (_ancestorCache.has(next)) {
      result.unshift(..._ancestorCache.get(next));
    } else {
      const f = zip.file(next);
      if (!f) { cur = next; continue; }
      const xml = await f.async('string');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const aDir = next.split('/').slice(0, -1).join('/');
      const aFname = next.split('/').pop();
      const imgMap = {};
      const arf = zip.file(`${aDir}/_rels/${aFname}.rels`);
      if (arf) {
        const arx = await arf.async('string');
        const ard = new DOMParser().parseFromString(arx, 'application/xml');
        for (const rel of ard.querySelectorAll('Relationship')) {
          const rid = rel.getAttribute('Id');
          const target = rel.getAttribute('Target');
          if (/\.(png|jpg|jpeg|gif|svg|emf|wmf)$/i.test(target)) {
            const imgFile = zip.file(resolveTarget(aDir, target));
            if (imgFile) {
              const b64 = await imgFile.async('base64');
              const ext2 = target.split('.').pop().toLowerCase();
              const mm = {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',svg:'image/svg+xml',emf:'image/x-emf',wmf:'image/x-wmf'};
              imgMap[rid] = `data:${mm[ext2]||'image/png'};base64,${b64}`;
            }
          }
        }
      }
      const spTree = doc.querySelector('spTree');
      const aShapes = [];
      if (spTree) {
        for (const child of spTree.children) {
          const tag = child.tagName?.split(':').pop();
          if (tag === 'sp') {
            if (child.querySelector('ph')) continue;
            const sh = parseSp(child, theme);
            if (sh) aShapes.push(sh);
          } else if (tag === 'pic') {
            const sh = parsePic(child, imgMap);
            if (sh) aShapes.push(sh);
          } else if (tag === 'grpSp') {
            parseSpTree(child, aShapes, theme, imgMap, 0, 0, 1);
          }
        }
      }
      _ancestorCache.set(next, aShapes);
      result.unshift(...aShapes);
    }
    cur = next;
  }
  return result;
}

// ── shape parsers ─────────────────────────────────────────────────────────────

function parseSpTree(spTree, shapes, theme, imgMap, offX, offY, scaleGroup, groupFill) {
  for (const child of spTree.children) {
    const tag = child.tagName?.split(':').pop();
    if (tag === 'sp') {
      const sh = parseSp(child, theme, groupFill);
      if (sh) { sh.x += offX; sh.y += offY; shapes.push(sh); }
    } else if (tag === 'pic') {
      const sh = parsePic(child, imgMap);
      if (sh) { sh.x += offX; sh.y += offY; shapes.push(sh); }
    } else if (tag === 'cxnSp') {
      const sh = parseCxnSp(child, theme);
      if (sh) { sh.x += offX; sh.y += offY; shapes.push(sh); }
    } else if (tag === 'graphicFrame') {
      const sh = parseTbl(child, theme);
      if (sh) { sh.x += offX; sh.y += offY; shapes.push(sh); }
    } else if (tag === 'grpSp') {
      const grpSpPr = [...child.children].find(c => c.localName === 'grpSpPr');
      const grpFillInfo = grpSpPr ? resolveFill(grpSpPr, theme) : {};
      const inheritedFill = (grpFillInfo.color || grpFillInfo.grad) ? grpFillInfo : groupFill;
      const grpXfrm = grpSpPr ? [...grpSpPr.children].find(c => c.localName === 'xfrm') : null;
      let gOffX = offX, gOffY = offY;
      if (grpXfrm) {
        const gOff   = [...grpXfrm.children].find(c => c.localName === 'off');
        const gChOff = [...grpXfrm.children].find(c => c.localName === 'chOff');
        if (gOff && gChOff) {
          gOffX += emu(gOff.getAttribute('x')) - emu(gChOff.getAttribute('x'));
          gOffY += emu(gOff.getAttribute('y')) - emu(gChOff.getAttribute('y'));
        }
      }
      parseSpTree(child, shapes, theme, imgMap, gOffX, gOffY, scaleGroup, inheritedFill);
    }
  }
}

function parseCustGeomPath(custGeom, shW, shH) {
  let d = '';
  for (const pathEl of custGeom.querySelectorAll('pathLst path')) {
    const pw = parseFloat(pathEl.getAttribute('w') || '1') || 1;
    const ph = parseFloat(pathEl.getAttribute('h') || '1') || 1;
    const sx = shW / pw, sy = shH / ph;
    for (const cmd of pathEl.children) {
      const ln = cmd.localName;
      const pts = [...cmd.querySelectorAll('pt')].map(p => [
        (parseFloat(p.getAttribute('x')) * sx).toFixed(2),
        (parseFloat(p.getAttribute('y')) * sy).toFixed(2)
      ]);
      if      (ln === 'moveTo'     && pts[0])           d += `M${pts[0][0]},${pts[0][1]} `;
      else if (ln === 'lnTo'       && pts[0])           d += `L${pts[0][0]},${pts[0][1]} `;
      else if (ln === 'cubicBezTo' && pts.length === 3) d += `C${pts[0][0]},${pts[0][1]} ${pts[1][0]},${pts[1][1]} ${pts[2][0]},${pts[2][1]} `;
      else if (ln === 'quadBezTo'  && pts.length === 2) d += `Q${pts[0][0]},${pts[0][1]} ${pts[1][0]},${pts[1][1]} `;
      else if (ln === 'close')                          d += 'Z ';
    }
  }
  return d.trim() || null;
}

function parseSp(sp, theme, groupFill) {
  const xfrm = sp.querySelector('xfrm');
  if (!xfrm) return null;
  const off = xfrm.querySelector('off'), ext = xfrm.querySelector('ext');
  if (!off||!ext) return null;
  const x = emu(off.getAttribute('x')), y = emu(off.getAttribute('y'));
  const w = emu(ext.getAttribute('cx')), h = emu(ext.getAttribute('cy'));
  const rot = parseInt(xfrm.getAttribute('rot')||'0')/60000;
  const flipH = xfrm.getAttribute('flipH')==='1', flipV = xfrm.getAttribute('flipV')==='1';
  const spPr = sp.querySelector('spPr');
  let fillInfo = spPr ? resolveFill(spPr, theme) : {};
  if (fillInfo.grpFill) fillInfo = groupFill || {};
  if (!fillInfo.color && !fillInfo.grad && !fillInfo.none) {
    const fillRef = sp.querySelector('style fillRef');
    if (fillRef) {
      const idx = parseInt(fillRef.getAttribute('idx') || '0');
      if (idx > 0) { const c = resolveColorEl(fillRef, theme); if (c) fillInfo = {color: c}; }
    }
  }
  const ln = spPr?.querySelector('ln');
  let strokeColor = null, strokeW = 0;
  if (ln) {
    if (!ln.querySelector('noFill')) {
      strokeColor = resolveColorEl(ln.querySelector('solidFill'), theme) || '#000000';
      strokeW = Math.max(0.5, parseInt(ln.getAttribute('w')||'12700')/12700);
    }
  }
  const geom = spPr?.querySelector('prstGeom')?.getAttribute('prst') || 'rect';
  const custGeomEl = spPr?.querySelector('custGeom');
  const svgPath = custGeomEl ? parseCustGeomPath(custGeomEl, w, h) : null;
  const text = parseTextBody(sp.querySelector('txBody'), theme);
  const shadow = parseShadow(spPr, theme);
  return {type:'sp', x,y,w,h,rot,flipH,flipV, fill:fillInfo, strokeColor, strokeW, geom, svgPath, text, shadow};
}

function parseShadow(spPr, theme) {
  if (!spPr) return null;
  const ef = spPr.querySelector('effectLst outerShdw, effectLst innerShdw');
  if (!ef) return null;
  const blurRad = emu(ef.getAttribute('blurRad')||'0');
  const dist    = emu(ef.getAttribute('dist')||'0');
  const dir     = parseInt(ef.getAttribute('dir')||'0')/60000;
  const color   = resolveColorEl(ef, theme) || '#000000';
  const dx = dist * Math.cos(dir * Math.PI / 180);
  const dy = dist * Math.sin(dir * Math.PI / 180);
  const alpha = parseInt(ef.querySelector('srgbClr alpha, schemeClr alpha')?.getAttribute('val')||'0');
  const [r,g,b] = hexToRgb(color.startsWith('#') ? color : '#000000');
  const a = alpha ? alpha/100000 : 0.35;
  return `${dx.toFixed(1)}px ${dy.toFixed(1)}px ${Math.round(blurRad)}px rgba(${r},${g},${b},${a})`;
}

function parsePic(pic, imgMap) {
  const blip = pic.querySelector('blipFill blip');
  const rId = blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','embed');
  const xfrm = pic.querySelector('xfrm');
  if (!xfrm||!rId) return null;
  const off = xfrm.querySelector('off'), ext = xfrm.querySelector('ext');
  if (!off||!ext) return null;
  const x = emu(off.getAttribute('x')), y = emu(off.getAttribute('y'));
  const w = emu(ext.getAttribute('cx')), h = emu(ext.getAttribute('cy'));
  const rot = parseInt(xfrm.getAttribute('rot')||'0')/60000;
  const flipH = xfrm.getAttribute('flipH')==='1', flipV = xfrm.getAttribute('flipV')==='1';
  return {type:'pic', x,y,w,h,rot,flipH,flipV, rId, imgData: imgMap[rId]||null, stretch: !!pic.querySelector('blipFill stretch')};
}

// ── main parse ────────────────────────────────────────────────────────────────

async function loadFileBuffer(file) {
  // Prefer the modern Promise-based API; fall back to FileReader for older browsers.
  if (typeof file.arrayBuffer === 'function') {
    return file.arrayBuffer();
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => resolve(ev.target.result);
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

function validateBuffer(buffer, filename) {
  if (!buffer || buffer.byteLength === 0) {
    throw new Error(`File is empty (0 bytes). The file "${filename}" could not be read.`);
  }
  const bytes = new Uint8Array(buffer, 0, Math.min(8, buffer.byteLength));
  const hex = Array.from(bytes).map(b => b.toString(16).padStart(2,'0')).join(' ');

  // ZIP / PPTX magic: PK (50 4B)
  if (bytes[0] === 0x50 && bytes[1] === 0x4B) return; // ✓ valid ZIP

  // CFB / Compound File Binary (D0 CF 11 E0) — encrypted or legacy .ppt
  if (bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) {
    throw new Error(
      `"${filename}" is password-protected or encrypted (Microsoft CFB format).\n` +
      `Remove the password in PowerPoint → File → Info → Protect Presentation, then try again.`
    );
  }

  // PDF
  if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) {
    throw new Error(`"${filename}" is a PDF, not a PPTX file.`);
  }

  throw new Error(
    `"${filename}" does not appear to be a valid PPTX file (not a ZIP archive).\n` +
    `First bytes: ${hex} (${buffer.byteLength} bytes total).\n` +
    `Make sure you selected a .pptx file saved in the modern Office Open XML format.`
  );
}

async function parsePptx(buffer, filename) {
  filename = filename || 'presentation.pptx';
  _ancestorCache.clear();
  validateBuffer(buffer, filename);
  const zip = await JSZip.loadAsync(buffer);
  themeColors = await parseTheme(zip);
  const presXml = await zip.file('ppt/presentation.xml').async('string');
  const presDoc = new DOMParser().parseFromString(presXml, 'application/xml');
  const sldSz = presDoc.querySelector('sldSz');
  if (sldSz) {
    slideW = parseInt(sldSz.getAttribute('cx') || 9144000) / 9525;
    slideH = parseInt(sldSz.getAttribute('cy') || 5143500) / 9525;
  } else { slideW = 960; slideH = 540; }

  const slideFiles = [];
  const presRels = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
  if (presRels) {
    const rd = new DOMParser().parseFromString(presRels,'application/xml');
    for (const rel of rd.querySelectorAll('Relationship')) {
      const type = rel.getAttribute('Type') || '';
      if (type.includes('/slide') && !type.includes('Layout') && !type.includes('Master')) {
        const tgt = rel.getAttribute('Target');
        slideFiles.push(tgt.startsWith('/') ? tgt.slice(1) : tgt.startsWith('../') ? tgt.slice(3) : 'ppt/' + tgt);
      }
    }
  }
  if (!slideFiles.length) {
    for (const key of Object.keys(zip.files)) {
      if (/^ppt\/slides\/slide\d+\.xml$/.test(key)) slideFiles.push(key);
    }
    slideFiles.sort((a,b)=>parseInt(a.match(/\d+/)[0])-parseInt(b.match(/\d+/)[0]));
  }

  slides = [];
  for (const sf of slideFiles) {
    const xml = await zip.file(sf)?.async('string');
    if (!xml) continue;
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const bg = await parseSlideBg(zip, sf, themeColors);
    const shapes = [];
    const spTree = doc.querySelector('spTree');
    if (!spTree) { slides.push({bg, shapes}); continue; }
    const slideDir = sf.split('/').slice(0,-1).join('/');
    const relsPath = sf.replace(/slides\/(slide[^/]+\.xml)/, 'slides/_rels/$1.rels');
    const imgMap = {};
    const rFile = zip.file(relsPath);
    if (rFile) {
      const rx = await rFile.async('string');
      const rd = new DOMParser().parseFromString(rx,'application/xml');
      for (const rel of rd.querySelectorAll('Relationship')) {
        const rid = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        if (/\.(png|jpg|jpeg|gif|svg|emf|wmf)$/i.test(target)) {
          const imgFile = zip.file(resolveTarget(slideDir, target));
          if (imgFile) {
            const b64 = await imgFile.async('base64');
            const ext2 = target.split('.').pop().toLowerCase();
            const mimeMap = {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',svg:'image/svg+xml',emf:'image/x-emf',wmf:'image/x-wmf'};
            imgMap[rid] = `data:${mimeMap[ext2]||'image/png'};base64,${b64}`;
          }
        }
      }
    }
    parseSpTree(spTree, shapes, themeColors, imgMap, 0, 0, 1);
    const ancestorShapes = await parseAncestorShapes(zip, sf, themeColors);
    ancestorShapes.forEach(s => s.layer = 0);
    shapes.forEach(s => s.layer = 1);
    slides.push({bg, shapes: [...ancestorShapes, ...shapes]});
  }

  currentSlide = 0;
  renderThumbnails();
  renderSlide(currentSlide);
  document.getElementById('drop-zone').style.display = 'none';
  document.getElementById('slide-canvas-wrap').style.display = 'block';
  document.getElementById('nav-controls').style.display = 'flex';
  document.getElementById('thumbnails-panel').style.display = 'flex';
}

// ── rendering ─────────────────────────────────────────────────────────────────

function renderSlide(idx) {
  if (idx < 0 || idx >= slides.length) return;
  currentSlide = idx;
  const sl = slides[idx];
  const canvas = document.getElementById('slide-canvas');
  const areaEl = document.getElementById('slide-area');
  const scale = Math.min((areaEl.clientWidth - 48) / slideW, (areaEl.clientHeight - 48) / slideH, 1.5);
  canvas.style.width  = Math.round(slideW * scale) + 'px';
  canvas.style.height = Math.round(slideH * scale) + 'px';
  canvas.innerHTML = '';
  const bgDiv = document.createElement('div');
  bgDiv.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;';
  if (sl.bg.imgData) {
    bgDiv.style.backgroundImage = `url(${sl.bg.imgData})`;
    bgDiv.style.backgroundSize = 'cover';
    bgDiv.style.backgroundPosition = 'center';
  } else if (sl.bg.gradStops) {
    const css = gradStopsToCss(sl.bg.gradStops);
    if (css) bgDiv.style.background = css;
  } else {
    bgDiv.style.background = sl.bg.color || '#FFFFFF';
  }
  canvas.appendChild(bgDiv);
  const inner = document.createElement('div');
  inner.style.cssText = `position:absolute;top:0;left:0;width:${slideW}px;height:${slideH}px;transform:scale(${scale});transform-origin:top left;`;
  for (const sh of sl.shapes) renderShape(sh, inner);
  canvas.appendChild(inner);
  document.getElementById('slide-counter').textContent = `${idx+1} / ${slides.length}`;
  document.getElementById('btn-prev').disabled = idx === 0;
  document.getElementById('btn-next').disabled = idx === slides.length - 1;
  updateThumbnailHighlight();
}

function renderShape(sh, parent) {
  if (sh.type === 'pic') { renderPic(sh, parent); return; }
  if (sh.type === 'cxn') { renderCxn(sh, parent); return; }
  if (sh.type === 'tbl') { renderTable(sh, parent); return; }
  if (sh.type === 'sp')  { renderSp(sh, parent); return; }
}

function applyTransform(el, sh) {
  const t = [];
  if (sh.rot)   t.push(`rotate(${sh.rot}deg)`);
  if (sh.flipH) t.push('scaleX(-1)');
  if (sh.flipV) t.push('scaleY(-1)');
  if (t.length) el.style.transform = t.join(' ');
}

function renderPic(sh, parent) {
  if (!sh.imgData) return;
  const el = document.createElement('div');
  el.style.cssText = `position:absolute;left:${sh.x}px;top:${sh.y}px;width:${sh.w}px;height:${sh.h}px;overflow:hidden;z-index:${sh.layer||0};`;
  applyTransform(el, sh);
  const img = document.createElement('img');
  img.src = sh.imgData;
  img.style.cssText = `width:100%;height:100%;object-fit:${sh.stretch?'fill':'cover'};display:block;`;
  el.appendChild(img);
  parent.appendChild(el);
}

function renderCxn(sh, parent) {
  const svg = document.createElementNS('http://www.w3.org/2000/svg','svg');
  svg.style.cssText = `position:absolute;left:${sh.x}px;top:${sh.y}px;width:${sh.w||2}px;height:${sh.h||2}px;overflow:visible;`;
  applyTransform(svg, sh);
  const line = document.createElementNS('http://www.w3.org/2000/svg','line');
  line.setAttribute('x1','0'); line.setAttribute('y1','0');
  line.setAttribute('x2',sh.w||0); line.setAttribute('y2',sh.h||0);
  line.setAttribute('stroke', sh.strokeColor||'#000');
  line.setAttribute('stroke-width', sh.strokeW||1);
  svg.appendChild(line);
  parent.appendChild(svg);
}

function renderSp(sh, parent) {
  const el = document.createElement('div');
  el.style.cssText = `position:absolute;left:${sh.x}px;top:${sh.y}px;width:${sh.w}px;height:${sh.h}px;box-sizing:border-box;overflow:hidden;z-index:${sh.layer||0};`;
  applyTransform(el, sh);
  if (!sh.fill?.none) {
    if (sh.fill?.color) el.style.backgroundColor = sh.fill.color;
    else if (sh.fill?.grad) { const css = gradStopsToCss(sh.fill.grad); if (css) el.style.background = css; }
  }
  if (sh.strokeColor && sh.strokeW > 0) el.style.border = `${sh.strokeW}px solid ${sh.strokeColor}`;
  // preset geometry
  if      (sh.geom === 'ellipse' || sh.geom === 'oval') el.style.borderRadius = '50%';
  else if (sh.geom === 'roundRect')  el.style.borderRadius = '8px';
  else if (['triangle','rtTriangle','diamond','parallelogram'].includes(sh.geom)) {
    el.style.background = 'none'; el.style.border = 'none'; el.style.overflow = 'visible';
    const svg = document.createElementNS('http://www.w3.org/2000/svg','svg');
    svg.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;';
    const poly = document.createElementNS('http://www.w3.org/2000/svg','polygon');
    const W = sh.w, H = sh.h;
    const pts = sh.geom === 'triangle'      ? `${W/2},0 ${W},${H} 0,${H}`
              : sh.geom === 'rtTriangle'    ? `0,0 ${W},${H} 0,${H}`
              : sh.geom === 'diamond'       ? `${W/2},0 ${W},${H/2} ${W/2},${H} 0,${H/2}`
              : /* parallelogram */          `${W*0.2},0 ${W},0 ${W*0.8},${H} 0,${H}`;
    poly.setAttribute('points', pts);
    poly.setAttribute('fill', sh.fill?.color||'none');
    if (sh.strokeColor) { poly.setAttribute('stroke',sh.strokeColor); poly.setAttribute('stroke-width',sh.strokeW||1); }
    svg.appendChild(poly); el.appendChild(svg);
  }
  // custom geometry
  if (sh.svgPath) {
    el.style.background = 'none'; el.style.border = 'none'; el.style.overflow = 'visible';
    const svg = document.createElementNS('http://www.w3.org/2000/svg','svg');
    svg.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;overflow:visible;';
    svg.setAttribute('viewBox', `0 0 ${sh.w} ${sh.h}`);
    svg.setAttribute('preserveAspectRatio','none');
    const pathEl = document.createElementNS('http://www.w3.org/2000/svg','path');
    pathEl.setAttribute('d', sh.svgPath);
    pathEl.setAttribute('fill', sh.fill?.color||'none');
    if (sh.strokeColor && sh.strokeW > 0) { pathEl.setAttribute('stroke',sh.strokeColor); pathEl.setAttribute('stroke-width',sh.strokeW); }
    else pathEl.setAttribute('stroke','none');
    svg.appendChild(pathEl); el.appendChild(svg);
  }
  if (sh.shadow && !sh.svgPath) el.style.boxShadow = sh.shadow;
  if (sh.text) renderTextBody(sh.text, el, sh.w, sh.h);
  parent.appendChild(el);
}

function renderTextBody(text, container, shW, shH) {
  if (!text?.paras) return;
  const {paras, vAlign, insetL, insetR, insetT, insetB, wordWrap, vert} = text;
  if (!paras.some(p => p.runs.some(r => r.t?.trim()))) return;
  const txDiv = document.createElement('div');
  txDiv.style.cssText = `position:absolute;left:${insetL}px;top:${insetT}px;right:${insetR}px;bottom:${insetB}px;overflow:hidden;`;
  if (vert === 'vert' || vert === 'eaVert') { txDiv.style.writingMode='vertical-rl'; txDiv.style.textOrientation='mixed'; }
  else if (vert === 'vert270') { txDiv.style.writingMode='vertical-lr'; txDiv.style.transform='rotate(180deg)'; }
  if (vAlign === 'ctr') { txDiv.style.display='flex'; txDiv.style.flexDirection='column'; txDiv.style.justifyContent='center'; }
  else if (vAlign === 'b') { txDiv.style.display='flex'; txDiv.style.flexDirection='column'; txDiv.style.justifyContent='flex-end'; }
  for (const para of paras) {
    const pEl = document.createElement('p');
    pEl.style.cssText = `margin:0;padding:0;text-align:${para.align};word-wrap:${wordWrap?'break-word':'nowrap'};white-space:${wordWrap?'normal':'nowrap'};`;
    if (para.spaceBeforePt > 0) pEl.style.marginTop = para.spaceBeforePt + 'pt';
    if (para.spaceAfterPt  > 0) pEl.style.marginBottom = para.spaceAfterPt + 'pt';
    const hasBullet = (para.buChar || para.buAutoNum) && !para.buNone;
    const hangPx = hasBullet ? Math.max(para.marL || 0, 14) : 0;
    if (hasBullet) {
      pEl.style.position    = 'relative';
      pEl.style.paddingLeft = hangPx + 'px';
    } else if (para.marL > 0) {
      pEl.style.paddingLeft = para.marL + 'px';
    } else if (para.lvl > 0) {
      pEl.style.paddingLeft = (para.lvl * 18) + 'px';
    }
    if (!para.runs.some(r=>r.t?.trim())) { pEl.innerHTML = '&nbsp;'; txDiv.appendChild(pEl); continue; }
    if (hasBullet) {
      const b = document.createElement('span');
      // Absolutely position bullet at left:0, text indented by hangPx
      b.style.cssText = 'position:absolute;left:0;top:0;width:' + hangPx + 'px;' +
        'text-align:right;padding-right:3px;box-sizing:border-box;white-space:nowrap;';

      let bulletText;
      if (para.buAutoNum) {
        const paraIdx = text.paras.indexOf(para);
        const n = paraIdx + 1;
        const sm = {arabicPeriod: n + '.', arabicParenR: n + ')',
          romanLcPeriod: toRoman(n).toLowerCase() + '.', romanUcPeriod: toRoman(n) + '.',
          alphaLcParenR: String.fromCharCode(96 + n) + ')', alphaUcParenR: String.fromCharCode(64 + n) + ')' };
        bulletText = sm[para.buAutoNumType] || (n + '.');
      } else {
        bulletText = para.buChar;
      }
      b.textContent = bulletText;

      // Colour: explicit buClr, then first run with a colour, then inherit
      const bColor = para.buColor || (para.runs.find(function(r){ return r.fc; }) || {}).fc || null;
      if (bColor) b.style.color = bColor;

      // Size: use explicit size from first run, or default to 24pt
      const baseFontPt = (para.runs.find(function(r){ return r.fontSize; }) || {}).fontSize || 24;
      const bSize = baseFontPt * (para.buSzPct / 100);
      b.style.fontSize = bSize + 'pt';

      // Font (skip Symbol/Wingdings, char already normalised)
      const safeFont = para.buFont && !/(symbol|wingdings)/i.test(para.buFont) ? para.buFont : null;
      b.style.fontFamily = safeFont ? ('"' + safeFont + '", Arial, sans-serif') : 'Arial, "Segoe UI", sans-serif';

      pEl.appendChild(b);
    }
    for (const run of para.runs) {
      if (run.isBr) { pEl.appendChild(document.createElement('br')); continue; }
      if (!run.t) continue;
      const span = document.createElement('span');
      // Default to 24pt if no size specified; PowerPoint default is typically 18pt but rendered larger
      const sz = run.fontSize || 24;
      span.style.fontSize = sz + 'pt';
      // Use run colour if set, otherwise default to black for better readability
      span.style.color = run.fc || '#000000';
      if (run.bold)      span.style.fontWeight = 'bold';
      if (run.italic)    span.style.fontStyle = 'italic';
      if (run.underline) span.style.textDecoration = 'underline';
      if (run.strike)    span.style.textDecoration = 'line-through';
      if (run.caps)      span.style.textTransform = run.smallCaps ? 'none' : 'uppercase';
      if (run.smallCaps) span.style.fontVariant = 'small-caps';
      if (run.baseline > 0) { span.style.verticalAlign='super'; span.style.fontSize=(sz*0.6)+'pt'; }
      else if (run.baseline < 0) { span.style.verticalAlign='sub'; span.style.fontSize=(sz*0.6)+'pt'; }
      const ff = run.fontFace;
      span.style.fontFamily = (ff && ff !== '+mj-lt' && ff !== '+mn-lt') ? `"${ff}", Calibri, sans-serif` : 'Calibri, "Segoe UI", sans-serif';
      span.style.lineHeight = para.lineSpacing;
      span.textContent = run.t;
      pEl.appendChild(span);
    }
    txDiv.appendChild(pEl);
  }
  container.appendChild(txDiv);
}

function renderTable(sh, parent) {
  const el = document.createElement('div');
  el.style.cssText = `position:absolute;left:${sh.x}px;top:${sh.y}px;width:${sh.w}px;height:${sh.h}px;overflow:hidden;box-sizing:border-box;`;
  const tbl = document.createElement('table');
  tbl.style.cssText = 'border-collapse:collapse;width:100%;height:100%;table-layout:fixed;';
  for (const row of sh.rows) {
    const tr = document.createElement('tr');
    tr.style.height = row.rowH ? row.rowH+'px' : 'auto';
    let colI = 0;
    for (const cell of row.cells) {
      if (cell.hMerge || cell.vMerge) { colI++; continue; }
      const td = document.createElement('td');
      td.style.cssText = `border:1px solid ${cell.borderColor||'#ccc'};box-sizing:border-box;overflow:hidden;vertical-align:middle;padding:4px;`;
      if (cell.fillColor) td.style.background = cell.fillColor;
      if (cell.gridSpan > 1) td.setAttribute('colspan', cell.gridSpan);
      if (cell.rowSpan  > 1) td.setAttribute('rowspan', cell.rowSpan);
      if (sh.colWidths[colI]) td.style.width = sh.colWidths[colI]+'px';
      if (cell.text) {
        const inner = document.createElement('div');
        inner.style.position = 'relative';
        renderTextBody(cell.text, inner, sh.colWidths[colI]||100, row.rowH||40);
        const textDiv = inner.querySelector('div');
        if (textDiv) { textDiv.style.position='static'; textDiv.style.inset='unset'; td.appendChild(textDiv); }
      }
      tr.appendChild(td);
      colI += cell.gridSpan || 1;
    }
    tbl.appendChild(tr);
  }
  el.appendChild(tbl);
  parent.appendChild(el);
}

// ── thumbnails ────────────────────────────────────────────────────────────────

function renderThumbnails() {
  const panel = document.getElementById('thumbnails-panel');
  panel.innerHTML = '';

  // Panel header: total slide count
  const header = document.createElement('div');
  header.style.cssText =
    'font-size:11px;font-weight:600;color:#aaa;text-align:center;' +
    'padding:6px 0 8px;letter-spacing:0.05em;flex-shrink:0;' +
    'border-bottom:1px solid #444;margin-bottom:6px;';
  header.textContent = slides.length + ' slide' + (slides.length !== 1 ? 's' : '');
  panel.appendChild(header);

  const TW = 136, TH = 77;

  slides.forEach(function(sl, i) {
    const thumb = document.createElement('div');
    thumb.id = 'thumb-' + i;
    thumb.style.cssText =
      'cursor:pointer;flex-shrink:0;border-radius:5px;padding:2px;' +
      'border:2px solid transparent;transition:border-color 0.1s;';
    thumb.title = 'Slide ' + (i + 1);
    thumb.addEventListener('click', function() { renderSlide(i); });

    // Mini slide canvas
    const miniSlide = document.createElement('div');
    miniSlide.style.cssText =
      'width:' + TW + 'px;height:' + TH + 'px;position:relative;overflow:hidden;' +
      'border-radius:2px;box-shadow:0 1px 4px rgba(0,0,0,0.45);';

    // Background
    if (sl.bg.imgData) {
      miniSlide.style.backgroundImage    = 'url(' + sl.bg.imgData + ')';
      miniSlide.style.backgroundSize     = 'cover';
      miniSlide.style.backgroundPosition = 'center';
    } else if (sl.bg.gradStops) {
      var css = gradStopsToCss(sl.bg.gradStops);
      miniSlide.style.background = css || '#ffffff';
    } else {
      miniSlide.style.background = sl.bg.color || '#ffffff';
    }

    // Scaled shape layer (fills + images, no text)
    var scale = TW / slideW;
    const tInner = document.createElement('div');
    tInner.style.cssText =
      'position:absolute;top:0;left:0;width:' + slideW + 'px;height:' + slideH + 'px;' +
      'transform:scale(' + scale + ');transform-origin:top left;pointer-events:none;';

    for (var si = 0; si < sl.shapes.length; si++) {
      var sh = sl.shapes[si];
      if (sh.type === 'pic' && sh.imgData) {
        var d = document.createElement('div');
        d.style.cssText =
          'position:absolute;left:' + sh.x + 'px;top:' + sh.y + 'px;' +
          'width:' + sh.w + 'px;height:' + sh.h + 'px;overflow:hidden;';
        var img = document.createElement('img');
        img.src = sh.imgData;
        img.style.cssText = 'width:100%;height:100%;object-fit:cover;display:block;';
        d.appendChild(img);
        tInner.appendChild(d);
        continue;
      }
      if (sh.type === 'sp') {
        var hasFill   = sh.fill && !sh.fill.none && (sh.fill.color || sh.fill.grad);
        var hasStroke = sh.strokeColor && sh.strokeW > 0;
        if (!hasFill && !hasStroke && !sh.svgPath) continue;
        var d2 = document.createElement('div');
        d2.style.cssText =
          'position:absolute;left:' + sh.x + 'px;top:' + sh.y + 'px;' +
          'width:' + sh.w + 'px;height:' + sh.h + 'px;box-sizing:border-box;overflow:hidden;';
        if (hasFill) {
          if (sh.fill.color) d2.style.background = sh.fill.color;
          else if (sh.fill.grad) {
            var gc = gradStopsToCss(sh.fill.grad);
            if (gc) d2.style.background = gc;
          }
        }
        if (hasStroke) {
          d2.style.border = Math.max(1, Math.round(sh.strokeW * scale)) + 'px solid ' + sh.strokeColor;
        }
        if (sh.geom === 'ellipse' || sh.geom === 'oval') d2.style.borderRadius = '50%';
        else if (sh.geom === 'roundRect') d2.style.borderRadius = '6px';
        tInner.appendChild(d2);
      }
    }

    // Text indicators: thin lines showing where text exists
    for (var ti = 0; ti < sl.shapes.length; ti++) {
      var sh = sl.shapes[ti];
      if (!sh.text || !sh.text.paras) continue;
      var hasText = sh.text.paras.some(function(p) {
        return p.runs && p.runs.some(function(r) { return r.t && r.t.trim(); });
      });
      if (!hasText) continue;
      // Draw thin text indicator lines (positioned in slide coordinate space, tInner scales them)
      var lineH = 3;  // line height in slide coordinates
      var lineW = sh.w * 0.65;  // 65% of shape width
      var lineX = sh.x + (sh.w - lineW) / 2;  // center horizontally in shape
      var lineStartY = sh.y + sh.h * 0.3;  // start at 30% from top of shape
      var lineGap = sh.h * 0.15;  // gap between lines
      for (var ln = 0; ln < 3; ln++) {
        var line = document.createElement('div');
        line.style.cssText =
          'position:absolute;left:' + lineX + 'px;top:' + (lineStartY + ln * lineGap) + 'px;' +
          'width:' + lineW + 'px;height:' + lineH + 'px;' +
          'background:#666;opacity:0.7;border-radius:1px;pointer-events:none;';
        tInner.appendChild(line);
      }
    }
    miniSlide.appendChild(tInner);

    // Slide number badge (bottom-right)
    const badge = document.createElement('div');
    badge.textContent = i + 1;
    badge.style.cssText =
      'position:absolute;bottom:3px;right:4px;pointer-events:none;' +
      'font-size:9px;font-weight:700;color:#fff;line-height:14px;' +
      'background:rgba(0,0,0,0.55);border-radius:3px;padding:0 4px;';
    miniSlide.appendChild(badge);

    thumb.appendChild(miniSlide);
    panel.appendChild(thumb);
  });
  updateThumbnailHighlight();
}

function updateThumbnailHighlight() {
  slides.forEach((_,i) => {
    const t = document.getElementById('thumb-'+i);
    if (t) t.style.border = i===currentSlide ? '2px solid #2B7CD3' : '2px solid transparent';
  });
  document.getElementById('thumb-'+currentSlide)?.scrollIntoView({block:'nearest', behavior:'smooth'});
}

function toRoman(n) {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let s = '';
  for (let i = 0; i < vals.length; i++) { while (n >= vals[i]) { s += syms[i]; n -= vals[i]; } }
  return s || 'I';
}

function navigate(dir) { renderSlide(currentSlide + dir); }

function showError(msg) {
  const dz = document.getElementById('drop-zone');
  dz.style.display = 'flex';
  document.getElementById('slide-canvas-wrap').style.display = 'none';
  document.getElementById('nav-controls').style.display = 'none';
  document.getElementById('thumbnails-panel').style.display = 'none';
  dz.innerHTML = `
    <div style="color:#f87171;font-size:14px;font-weight:600;margin-bottom:8px;">⚠ Could not open file</div>
    <div style="color:#ccc;font-size:12px;white-space:pre-wrap;max-width:440px;line-height:1.6;">${msg.replace(/</g,'&lt;')}</div>
    <div style="margin-top:16px;font-size:11px;color:#777;">Click to try another file</div>`;
}

function toggleFullscreen() {
  const el = document.getElementById('slide-canvas');
  if (!document.fullscreenElement) el.requestFullscreen?.();
  else document.exitFullscreen?.();
}

// ── event wiring (replaces all inline onclick/onchange) ───────────────────────

document.getElementById('btn-prev').addEventListener('click', () => navigate(-1));
document.getElementById('btn-next').addEventListener('click', () => navigate(1));
document.getElementById('btn-fullscreen').addEventListener('click', toggleFullscreen);
document.getElementById('drop-zone').addEventListener('click', () => document.getElementById('file-input').click());

document.getElementById('file-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('filename').textContent = file.name;
  try { await parsePptx(await loadFileBuffer(file), file.name); }
  catch(err) { showError(err.message); }
});

document.getElementById('slide-area').addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect='copy'; });
document.getElementById('slide-area').addEventListener('drop', async (e) => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (!file || !file.name.endsWith('.pptx')) return;
  document.getElementById('filename').textContent = file.name;
  try { await parsePptx(await loadFileBuffer(file), file.name); }
  catch(err) { showError(err.message); }
});

document.addEventListener('keydown', e => {
  if (e.key==='ArrowRight'||e.key==='ArrowDown') navigate(1);
  if (e.key==='ArrowLeft' ||e.key==='ArrowUp')   navigate(-1);
});

window.addEventListener('resize', () => { if (slides.length) renderSlide(currentSlide); });

// ── auto-load from ?url= (opened by XPPT-Open.ps1 or background.js) ──────────

(async () => {
  const params = new URLSearchParams(window.location.search);
  const pptxUrl = params.get('url');
  if (!pptxUrl) return;
  const rawName = params.get('name') || decodeURIComponent(pptxUrl).split('/').pop().split('?')[0] || 'presentation.pptx';
  const name = decodeURIComponent(rawName);  // normalise any %20 etc. from URL-encoded path segments
  document.getElementById('filename').textContent = name;
  document.getElementById('drop-zone').innerHTML =
    `<div style="font-size:13px;color:#ccc;">Loading <strong>${name}</strong>…</div>
     <div style="margin-top:8px;width:200px;height:4px;background:#555;border-radius:2px;overflow:hidden;">
       <div id="load-bar" style="width:0%;height:100%;background:#2B7CD3;transition:width 0.3s;"></div>
     </div>`;
  const bar = document.getElementById('load-bar');
  let p = 0;
  const tick = setInterval(() => { p = Math.min(p+3, 85); if (bar) bar.style.width = p+'%'; }, 120);
  try {
    let buffer;
    if (pptxUrl.startsWith('file://')) {
      buffer = await new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', pptxUrl, true);
        xhr.responseType = 'arraybuffer';
        xhr.onload = () => (xhr.status === 0 || xhr.status === 200) ? resolve(xhr.response) : reject(new Error('XHR status ' + xhr.status));
        xhr.onerror = () => reject(new Error('Cannot read local file — enable "Allow access to file URLs" for XPPT in chrome://extensions'));
        xhr.send();
      });
    } else {
      const resp = await fetch(pptxUrl);
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      buffer = await resp.arrayBuffer();
    }
    clearInterval(tick);
    if (bar) bar.style.width = '100%';
    await parsePptx(buffer, name);
  } catch(err) {
    clearInterval(tick);
    showError(err.message);
  }
})();

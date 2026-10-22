
// ==UserScript==
// @name         AZGrab
// @namespace    https://amazon.sharepoint.com/
// @version      4.1.1
// @updateURL    https://raw.githubusercontent.com/morakbri/WorkloadTool//main/azgrab.user.js
// @downloadURL  https://raw.githubusercontent.com/morakbri/WorkloadTool//main/azgrab.user.js
// @description  Auto-extracts rack positions from SharePoint Excel files — supports 1/62AZ, 6AZ, 7/55AZ, 12AZ with AZ selector. Groups by site with 2-row gaps. Midday rows (above thick border separator) included in optics only, excluded from cabling. Per-row midday override.
// @author       @morakbri @tngujona
// @match        https://*.sharepoint.com/*
// @match        https://amazon.sharepoint.com/*
// @match        https://*.officeapps.live.com/*
// @match        https://view.officeapps.live.com/*
// @grant        GM_setClipboard
// @grant        GM_xmlhttpRequest
// @connect      *.sharepoint.com
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';
    if (window.top !== window.self) return;
    if (document.getElementById('azg-fab')) return;

    const AZ_MAP = {
        '1/62AZ': { label: '1/62AZ',  match: ['1AZ', '62AZ', 'IAD1', 'IAD62'] },
        '6AZ':    { label: '6AZ',     match: ['6AZ', 'IAD6'] },
        '7/55AZ': { label: '7/55AZ',  match: ['7AZ', '55AZ', 'IAD7', 'IAD55'] },
        '12AZ':   { label: '12AZ',    match: ['12AZ', 'IAD12'] }
    };
    let selectedAZ = '6AZ';
    const VALID_RACK_PREFIXES = ['EC2', 'MV.BONSAI', 'BONSAI', 'EBS', 'S3'];
    const ALERT_LOST_RES = ['lost reservation', 'lost res'];
    const ALERT_REJECTED = ['rejected', 'reject'];
    const SITE_GAP_ROWS = 2;
    const DEST = {
        brickHandoffDate:'', dateWorkloadReceived:'',
        copper:'Not Installed', copperSig:'Tech Alias',
        fiber:'Not Installed', fiberSig:'Tech Alias',
        mnSig:'Tech Alias', optics:'Check for Optics',
        brickPatching:'Tech Alias', sitePOC:'', posStatus:''
    };
    const THICK_BORDERS = ['thick', 'medium', 'double'];

    // === MIDDAY DETECTION HELPERS ===
    function isThickBorder(edge) {
        return edge ? THICK_BORDERS.includes((edge.style||'').toLowerCase()) : false;
    }
    function isDarkFill(cell) {
        if (!cell || !cell.s) return false;
        const src = [];
        if (cell.s.fill) { if (cell.s.fill.fgColor) src.push(cell.s.fill.fgColor); if (cell.s.fill.bgColor) src.push(cell.s.fill.bgColor); }
        if (cell.s.fgColor) src.push(cell.s.fgColor);
        if (cell.s.bgColor) src.push(cell.s.bgColor);
        if (cell.s.patternFill && cell.s.patternFill.fgColor) src.push(cell.s.patternFill.fgColor);
        for (const c of src) {
            if (c.rgb) { let h=c.rgb.replace(/^#/,''); if(h.length===8)h=h.substring(2); if(h.length>=6){ const r=parseInt(h.substring(0,2),16),g=parseInt(h.substring(2,4),16),b=parseInt(h.substring(4,6),16); if(r<40&&g<40&&b<40)return true; } }
            if (c.theme!==undefined && Number(c.theme)===1) return true;
            if (c.indexed!==undefined && (Number(c.indexed)===0||Number(c.indexed)===8)) return true;
        }
        return false;
    }
    function rowHasData(sheet, r, maxC) {
        for (let c=0;c<=maxC;c++) { const cell=sheet[XLSX.utils.encode_cell({r,c})]; if(cell&&cell.v!==undefined&&cell.v!==null&&String(cell.v).trim()!=='') return true; }
        return false;
    }
    function hasDataAboveAndBelow(sheet, sep, range, maxC) {
        let above=false, below=false;
        for(let r=range.s.r+1;r<sep;r++){if(rowHasData(sheet,r,maxC)){above=true;break;}}
        for(let r=sep+1;r<=range.e.r;r++){if(rowHasData(sheet,r,maxC)){below=true;break;}}
        return above&&below;
    }
    function findMiddaySeparatorRow(sheet) {
        const ref=sheet['!ref']; if(!ref) return -1;
        const range=XLSX.utils.decode_range(ref), maxC=Math.min(range.e.c,14), start=range.s.r+1;
        for(let r=start;r<=Math.min(start+15,range.e.r);r++){const c=sheet[XLSX.utils.encode_cell({r,c:0})];if(c&&c.s)console.log('[azg] Row',r,'style:',JSON.stringify(c.s).substring(0,300));}
        const cands=[];
        for(let r=start;r<=range.e.r;r++){for(let c=range.s.c;c<=maxC;c++){const cell=sheet[XLSX.utils.encode_cell({r,c})];if(cell&&cell.v!==undefined){const v=String(cell.v).trim().toUpperCase();if(v==='MD'||v==='MIDDAY'||v==='MID-DAY'||v==='MID DAY'){console.log('[azg] MD marker row',r,'col',c);cands.push({row:r-1,type:'MD-text',score:100});}}}}
        for(let r=start;r<=range.e.r;r++){let bb=0,tb=0,df=0,empty=true;for(let c=range.s.c;c<=maxC;c++){const addr=XLSX.utils.encode_cell({r,c}),cell=sheet[addr];if(cell){if(cell.v!==undefined&&cell.v!==null&&String(cell.v).trim()!=='')empty=false;if(cell.s&&cell.s.border&&isThickBorder(cell.s.border.bottom))bb++;if(isDarkFill(cell))df++;}if(r+1<=range.e.r){const nx=sheet[XLSX.utils.encode_cell({r:r+1,c})];if(nx&&nx.s&&nx.s.border&&isThickBorder(nx.s.border.top))tb++;}}
            if(bb>=3)cands.push({row:r,type:'bottom-border',score:bb+10});if(tb>=3)cands.push({row:r,type:'top-border-next',score:tb+10});if(df>=3)cands.push({row:r-1,type:'dark-fill',score:df+5});if(empty&&r>start&&r<range.e.r)cands.push({row:r-1,type:'empty-row',score:1});}
        console.log('[azg] Candidates:',cands.length,cands.map(c=>'row '+c.row+' ('+c.type+')').join(', '));
        const valid=cands.filter(c=>hasDataAboveAndBelow(sheet,c.row+1,range,maxC));
        if(!valid.length){console.log('[azg] No midday separator.');return -1;}
        valid.sort((a,b)=>b.score-a.score);
        console.log('[azg] Best separator:',valid[0].type,'row',valid[0].row);
        return valid[0].row;
    }

    // === STYLES ===
    const sty=document.createElement('style');
    sty.textContent=`
#azg-fab{position:fixed;bottom:30px;right:30px;width:56px;height:56px;border-radius:50%;background:#2563eb;color:#fff;font-size:14px;font-weight:700;border:none;cursor:pointer;z-index:2147483647;box-shadow:0 4px 14px rgba(37,99,235,.4);display:flex;align-items:center;justify-content:center;transition:transform .2s,box-shadow .2s;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
#azg-fab:hover{transform:scale(1.1);box-shadow:0 6px 20px rgba(37,99,235,.6)}
#azg-fab.azg-loading{animation:azg-pulse 1.2s infinite;background:#1d4ed8}
@keyframes azg-pulse{0%,100%{box-shadow:0 4px 14px rgba(37,99,235,.4)}50%{box-shadow:0 4px 24px rgba(37,99,235,.8)}}
#azg-popup-overlay{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,.6);z-index:2147483646;display:none;align-items:center;justify-content:center}
#azg-popup-overlay.azg-visible{display:flex}
#azg-popup{background:#111827;border:1px solid #374151;border-radius:12px;width:960px;max-width:94vw;max-height:88vh;overflow-y:auto;padding:24px;color:#e5e7eb;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;box-shadow:0 25px 60px rgba(0,0,0,.5)}
#azg-popup h2{margin:0 0 4px;font-size:20px;color:#fff}
#azg-popup .azg-sub{color:#9ca3af;font-size:12px;margin-bottom:16px}
#azg-file-info{background:#1f2937;border:1px solid #374151;border-radius:6px;padding:10px 14px;margin-bottom:16px;display:flex;align-items:center;gap:10px}
#azg-file-info .azg-fi{font-size:20px}
#azg-file-info .azg-fd{flex:1}
#azg-file-info .azg-fn{color:#34d399;font-weight:600;font-size:13px}
#azg-file-info .azg-fs{color:#6b7280;font-size:11px;margin-top:2px}
.azg-ctrl{display:flex;gap:12px;align-items:center;margin-bottom:16px;flex-wrap:wrap}
.azg-ctrl select{background:#1f2937;color:#e5e7eb;border:1px solid #374151;border-radius:6px;padding:8px 12px;font-size:13px;min-width:120px;cursor:pointer}
#azg-az-sel{min-width:90px;font-weight:700;border-color:#2563eb;color:#60a5fa}
.azg-ctrl select:focus{outline:none;border-color:#2563eb}
.azg-btn{padding:8px 16px;border-radius:6px;border:none;font-size:13px;font-weight:600;cursor:pointer;transition:background .2s}
.azg-btn-p{background:#2563eb;color:#fff}.azg-btn-p:hover{background:#1d4ed8}.azg-btn-p:disabled{background:#374151;color:#6b7280;cursor:not-allowed}
.azg-btn-s{background:#059669;color:#fff}.azg-btn-s:hover{background:#047857}
.azg-btn-c{background:#374151;color:#e5e7eb}.azg-btn-c:hover{background:#4b5563}
.azg-btn-r{background:#374151;color:#e5e7eb;font-size:12px;padding:6px 12px}.azg-btn-r:hover{background:#4b5563}
.azg-btn-o{background:#0891b2;color:#fff}.azg-btn-o:hover{background:#0e7490}
#azg-st{font-size:13px;margin-bottom:12px;min-height:20px;padding:6px 0}
#azg-st.azg-ok{color:#34d399}#azg-st.azg-err{color:#f87171}#azg-st.azg-inf{color:#60a5fa}
.azg-ab{border-radius:8px;padding:12px 16px;margin-bottom:12px;font-size:13px;display:none}
.azg-ab-l{background:#78350f;border:1px solid #d97706;color:#fde68a}
.azg-ab-r{background:#7f1d1d;border:1px solid #dc2626;color:#fca5a5}
.azg-ab-m{background:#064e3b;border:1px solid #059669;color:#6ee7b7}
.azg-ab strong{display:block;margin-bottom:4px;font-size:14px}
.azg-ab ul{margin:4px 0 0 16px;padding:0}.azg-ab li{margin:2px 0}
#azg-tbl{width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px}
#azg-tbl th{background:#1f2937;color:#9ca3af;font-weight:600;text-align:left;padding:8px 10px;border-bottom:1px solid #374151;position:sticky;top:0}
#azg-tbl td{padding:6px 10px;border-bottom:1px solid #1f2937;color:#d1d5db}
#azg-tbl tr:hover td{background:#1f2937}
#azg-tbl tr.azg-rl td{background:#422006;color:#fde68a}
#azg-tbl tr.azg-rr td{background:#450a0a;color:#fca5a5;text-decoration:line-through}
#azg-tbl tr.azg-rm td{background:#1e1b4b;color:#c4b5fd}
#azg-tbl tr.azg-rm:hover td{background:#312e81}
#azg-tbl tr.azg-sd td{background:#0f172a;padding:2px;border-bottom:2px solid #374151}
.azg-tag{display:inline-block;font-size:10px;font-weight:700;padding:1px 5px;border-radius:3px;margin-left:4px;vertical-align:middle;color:#fff}
.azg-tag-m{background:#7c3aed}.azg-tag-l{background:#d97706}.azg-tag-j{background:#dc2626}
.azg-mt{background:none;border:1px solid #4b5563;border-radius:4px;color:#9ca3af;font-size:11px;padding:2px 6px;cursor:pointer;transition:all .15s;vertical-align:middle}
.azg-mt:hover{border-color:#7c3aed;color:#c4b5fd;background:#1e1b4b}
.azg-mt.azg-im{border-color:#7c3aed;color:#c4b5fd;background:#312e81}
.azg-mt.azg-im:hover{border-color:#ef4444;color:#fca5a5;background:#450a0a}
#azg-tc{max-height:300px;overflow-y:auto;border:1px solid #374151;border-radius:6px;margin-bottom:16px;display:none}
#azg-ob{background:#064e3b;border:1px solid #065f46;border-radius:6px;padding:12px;font-family:'Consolas','Monaco',monospace;font-size:12px;color:#6ee7b7;white-space:pre;overflow-x:auto;max-height:180px;overflow-y:auto;margin-bottom:12px;display:none}
#azg-cr{display:none;gap:12px;align-items:center;margin-bottom:8px;flex-wrap:wrap}
#azg-stats{font-size:12px;color:#9ca3af;margin-bottom:12px;display:none}
.azg-hr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px}
.azg-ch{font-size:11px;color:#6b7280;width:100%;margin-top:4px}
#azg-ocr{display:none;gap:12px;align-items:center;margin-bottom:8px;flex-wrap:wrap;padding:10px 12px;background:#0c4a6e;border:1px solid #0369a1;border-radius:6px}
#azg-ocr .azg-oh{font-size:11px;color:#7dd3fc;width:100%;margin-top:4px}
.azg-oc{font-size:11px;color:#db2777;margin-left:8px}`;
    document.head.appendChild(sty);

    // === UI ===
    const fab=document.createElement('button');fab.id='azg-fab';fab.textContent='AZ';fab.title='AZGrab — Click to grab positions';document.body.appendChild(fab);
    const overlay=document.createElement('div');overlay.id='azg-popup-overlay';
    overlay.innerHTML=`<div id="azg-popup"><div class="azg-hr"><div><h2>AZGrab</h2><div class="azg-sub">Auto-extract rack positions — grouped by site, 2-row gaps — v4.1.1 (multi-AZ, midday detection + per-row override)</div></div><button class="azg-btn azg-btn-c" id="azg-close">Close</button></div><div id="azg-file-info" style="display:none"><div class="azg-fi">📊</div><div class="azg-fd"><div class="azg-fn" id="azg-fname">—</div><div class="azg-fs" id="azg-fsrc">—</div></div><button class="azg-btn azg-btn-r" id="azg-refresh">↻ Re-fetch</button></div><div class="azg-ctrl"><select id="azg-az-sel"><option value="1/62AZ">1/62AZ</option><option value="6AZ" selected>6AZ</option><option value="7/55AZ">7/55AZ</option><option value="12AZ">12AZ</option></select><select id="azg-sh-sel" disabled><option value="">-- Select Sheet --</option></select><button class="azg-btn azg-btn-p" id="azg-grab" disabled>Grab Positions</button></div><div id="azg-st"></div><div class="azg-ab azg-ab-m" id="azg-am"><strong>⏫ MIDDAY POSITIONS DETECTED</strong><div id="azg-mi"></div></div><div class="azg-ab azg-ab-l" id="azg-al"><strong>⚠️ LOST RESERVATIONS</strong><ul id="azg-ll"></ul></div><div class="azg-ab azg-ab-r" id="azg-ar"><strong>🚫 REJECTED POSITIONS</strong><ul id="azg-rl2"></ul></div><div id="azg-stats"></div><div id="azg-tc"><table id="azg-tbl"><thead><tr><th>Site</th><th>Position</th><th>Uplink</th><th>Asset</th><th>Rack Type</th><th>Land Date</th><th>Brick</th><th>Midday</th><th>Notes</th></tr></thead><tbody id="azg-tbody"></tbody></table></div><div id="azg-ob"></div><div id="azg-cr"><button class="azg-btn azg-btn-s" id="azg-copy">📋 Copy Cabling (Clean)</button><button class="azg-btn azg-btn-p" id="azg-copya">📋 Copy Cabling (With Alerts)</button><span id="azg-cs" style="font-size:12px;color:#34d399"></span><div class="azg-ch">Cabling only — midday excluded. Paste into Col A. Grouped by site with 2-row gaps.</div></div><div id="azg-ocr"><button class="azg-btn azg-btn-o" id="azg-copyo">🔬 Copy Optics (All)</button><span id="azg-os" style="font-size:12px;color:#7dd3fc"></span><div class="azg-oh">Optics = morning + midday. 6-col TSV. Paste into Col A of optics workload.</div></div></div>`;
    document.body.appendChild(overlay);

    // === REFS ===
    const azSel=document.getElementById('azg-az-sel'),fileInfo=document.getElementById('azg-file-info'),fnameEl=document.getElementById('azg-fname'),fsrcEl=document.getElementById('azg-fsrc'),refreshBtn=document.getElementById('azg-refresh'),shSel=document.getElementById('azg-sh-sel'),grabBtn=document.getElementById('azg-grab'),stEl=document.getElementById('azg-st'),statsEl=document.getElementById('azg-stats'),alEl=document.getElementById('azg-al'),llEl=document.getElementById('azg-ll'),arEl=document.getElementById('azg-ar'),rlEl=document.getElementById('azg-rl2'),amEl=document.getElementById('azg-am'),miEl=document.getElementById('azg-mi'),tcEl=document.getElementById('azg-tc'),tbody=document.getElementById('azg-tbody'),obEl=document.getElementById('azg-ob'),crEl=document.getElementById('azg-cr'),copyBtn=document.getElementById('azg-copy'),copyaBtn=document.getElementById('azg-copya'),csEl=document.getElementById('azg-cs'),closeBtn=document.getElementById('azg-close'),ocrEl=document.getElementById('azg-ocr'),copyoBtn=document.getElementById('azg-copyo'),osEl=document.getElementById('azg-os');

    let workbook=null, lastRows=[];

    // === HELPERS ===
    function isAZ(v){if(!v)return false;const u=v.toUpperCase().trim(),cfg=AZ_MAP[selectedAZ];return cfg?cfg.match.includes(u):false;}
    function isRack(rt){if(!rt)return false;const u=rt.toUpperCase().trim();return VALID_RACK_PREFIXES.some(p=>u.startsWith(p));}
    function isXls(u){return/\.(xlsx|xlsm|xls)/i.test(u);}
    function fname(u){try{const p=decodeURIComponent(u.split('?')[0]).split('/');return p[p.length-1]||null;}catch(e){return null;}}
    function setSt(m,t){stEl.textContent=m;stEl.className='';if(t==='success')stEl.classList.add('azg-ok');else if(t==='error')stEl.classList.add('azg-err');else if(t==='info')stEl.classList.add('azg-inf');}
    function esc(s){const d=document.createElement('div');d.textContent=s;return d.innerHTML;}
    function sortSite(rows){return rows.sort((a,b)=>{const sA=a.site.toUpperCase(),sB=b.site.toUpperCase(),pA=sA.match(/(\d+|\D+)/g)||[],pB=sB.match(/(\d+|\D+)/g)||[];for(let i=0;i<Math.max(pA.length,pB.length);i++){const x=pA[i]||'',y=pB[i]||'',nA=parseInt(x,10),nB=parseInt(y,10);if(!isNaN(nA)&&!isNaN(nB)){if(nA!==nB)return nA-nB;}else{if(x<y)return-1;if(x>y)return 1;}}return a.position.localeCompare(b.position,undefined,{numeric:true});});}
    function grpSite(rows){const gs=[],m=new Map();rows.forEach(r=>{const k=r.site.toUpperCase().trim();if(!m.has(k)){const g={site:r.site,rows:[]};m.set(k,g);gs.push(g);}m.get(k).rows.push(r);});return gs;}

    // === GUID ===
    function cleanGuid(raw){if(!raw)return'';let s=raw;for(let i=0;i<2;i++){try{const d=decodeURIComponent(s);if(d===s)break;s=d;}catch(e){break;}}return s.replace(/[{}\s]/g,'').trim();}
    function fmtGuid(g){if(!g)return'';const c=g.replace(/[^a-fA-F0-9]/g,'');if(c.length!==32)return g;return c.substring(0,8)+'-'+c.substring(8,12)+'-'+c.substring(12,16)+'-'+c.substring(16,20)+'-'+c.substring(20,32);}

    // === FILE URL DETECTION ===
    function detectFile(){
        const url=window.location.href,p=new URLSearchParams(window.location.search),sd=p.get('sourcedoc'),fp=p.get('file');
        if(sd&&url.includes('/_layouts/')){const sp=extSite(url),cg=fmtGuid(cleanGuid(sd)),ws=findWopi(p,url);return{url:apiUrl(sp,cg),type:'sourcedoc',guid:cg,fileName:fp?decodeURIComponent(fp):null,sitePath:sp,wopiSrc:ws};}
        if(url.includes('officeapps.live.com')){const src=p.get('src')||p.get('wdOrigin');if(src&&isXls(src))return{url:src,type:'officeapps',fileName:fname(src)};const ws=p.get('WOPISrc')||p.get('wopisrc');if(ws)return{url:ws+'/contents',type:'wopi',fileName:fp};}
        if(/\.(xlsx|xlsm|xls)(\?|$|#)/i.test(url)){const u=url.split('?')[0];return{url:u,type:'direct',fileName:fname(u)};}
        const id=p.get('id');if(id&&isXls(id)){const fp2=id.startsWith('/')?id:'/'+id;return{url:window.location.origin+fp2,type:'id-param',fileName:fname(id)};}
        return null;
    }
    function findWopi(p,url){const d=p.get('WOPISrc')||p.get('wopisrc');if(d)return d;const m=url.match(/(https?:\/\/[^&\s]+_vti_bin\/wopi\.ashx\/files\/[a-f0-9]+)/i);return m?m[1]:null;}
    function extSite(url){try{let p=new URL(url).pathname;p=p.replace(/^\/:x:\/[rs]\//,'/');const m=p.match(/^(\/(?:sites|teams)\/[^/]+)/i);if(m)return m[1];}catch(e){}return'';}
    function apiUrl(sp,guid){const base=sp?window.location.origin+sp:window.location.origin;return base+`/_api/web/GetFileById('${guid}')/$value`;}

    // === RESPONSE GUARD ===
    function validateBin(buf){
        if(!buf||buf.byteLength===0)return{ok:false,reason:'Empty response.'};
        const u=new Uint8Array(buf);
        if((u[0]===0x50&&u[1]===0x4B&&u[2]===0x03&&u[3]===0x04)||(u[0]===0xD0&&u[1]===0xCF&&u[2]===0x11&&u[3]===0xE0))return{ok:true};
        let txt='';try{txt=new TextDecoder('utf-8',{fatal:false}).decode(u.slice(0,2000)).trim();}catch(e){return{ok:false,reason:'Unknown format.'};}
        const lw=txt.toLowerCase();
        if(lw.includes('<!doctype')||lw.includes('<html')){if(lw.includes('sign in')||lw.includes('login')||lw.includes('microsoftonline'))return{ok:false,reason:'Login page returned. Sign in again.'};if(lw.includes('access denied')||lw.includes('403'))return{ok:false,reason:'Access Denied.'};if(lw.includes('404'))return{ok:false,reason:'404 Not Found.'};return{ok:false,reason:'HTML returned instead of Excel.'};}
        if(lw.startsWith('{')||lw.startsWith('['))return{ok:false,reason:'JSON returned instead of Excel.'};
        return{ok:true};
    }

    // === FETCH ===
    async function fetchXls(fi){
        const strats=[];
        if(fi.sitePath&&fi.guid)strats.push({n:'download.aspx',u:window.location.origin+fi.sitePath+'/_layouts/15/download.aspx?UniqueId='+fi.guid});
        strats.push({n:'GetFileById',u:fi.url});
        if(fi.wopiSrc)strats.push({n:'WOPI',u:fi.wopiSrc+'/contents'});
        for(const s of strats){console.log('[azg] Trying:',s.n);setSt('Fetching ('+s.n+')...','info');try{const buf=await fetchBin(s.u);const chk=validateBin(buf);if(chk.ok){console.log('[azg] ✓',s.n,buf.byteLength,'bytes');return buf;}console.warn('[azg] ✗',s.n,chk.reason);}catch(e){console.warn('[azg] ✗',s.n,e.message);}}
        throw new Error('All fetch strategies failed. Check SharePoint sign-in.');
    }
    function fetchBin(url){return new Promise((res,rej)=>{const nc=!url.includes(window.location.hostname);if(typeof GM_xmlhttpRequest!=='undefined'&&nc){GM_xmlhttpRequest({method:'GET',url,responseType:'arraybuffer',headers:{'Accept':'application/octet-stream'},onload(r){r.status>=200&&r.status<300?res(r.response):rej(new Error('HTTP '+r.status));},onerror(){rej(new Error('Network error'));}});}else{fetch(url,{credentials:'include',headers:{'Accept':'application/octet-stream'}}).then(r=>{if(!r.ok)throw new Error('HTTP '+r.status);return r.arrayBuffer();}).then(res).catch(rej);}});}

    // === MIDDAY TOGGLE ===
    function toggleMD(entry,btn){
        entry.isMidday=!entry.isMidday;entry.isOverride=true;
        entry.notes=entry.notes.filter(n=>n!=='MIDDAY'&&n!=='MIDDAY ✎');
        if(entry.isMidday)entry.notes.push('MIDDAY ✎');
        btn.textContent=entry.isMidday?'☀ Remove':'🌙 Add';
        btn.title=entry.isMidday?'Remove from midday (move to cabling)':'Mark as midday (optics only)';
        btn.classList.toggle('azg-im',entry.isMidday);
        btn.closest('tr').classList.toggle('azg-rm',entry.isMidday);
        refreshStats();
        console.log('[azg] Toggle:',entry.site,entry.position,'→',entry.isMidday?'MIDDAY':'MORNING','(override)');
    }
    function refreshStats(){
        const g=lastRows;if(!g.length)return;
        const md=g.filter(e=>e.isMidday),mo=g.filter(e=>!e.isMidday),cl=g.filter(e=>!e.isLost&&!e.isRejected),ov=g.filter(e=>e.isOverride),sg=grpSite(g);
        if(md.length){amEl.style.display='block';const ovn=ov.length?'<br><span class="azg-oc">('+ov.length+' manual override'+(ov.length>1?'s':'')+' ✎)</span>':'';miEl.innerHTML='<strong>'+md.length+' midday position'+(md.length>1?'s':'')+'</strong>'+ovn+'<br>Included in Optics, excluded from Cabling.<ul>'+md.map(e=>'<li>'+esc(e.site)+' — '+esc(e.position)+(e.rackType?' ('+esc(e.rackType)+')':'')+(e.isOverride?' ✎':'')+'</li>').join('')+'</ul>';}else{amEl.style.display='none';}
        statsEl.textContent=AZ_MAP[selectedAZ].label+': '+g.length+' total | '+mo.length+' morning (cabling) | '+md.length+' midday (optics only) | '+sg.length+' sites | '+g.filter(e=>e.isLost).length+' lost | '+g.filter(e=>e.isRejected).length+' rejected'+(ov.length?' | '+ov.length+' override(s) ✎':'');
        crEl.style.display=mo.length>0?'flex':'none';
        ocrEl.style.display=cl.length>0?'flex':'none';
    }

    // === PROCESS SHEET ===
    function processSheet(){
        if(!workbook){setSt('No workbook loaded.','error');return;}
        const sn=shSel.value;if(!sn){setSt('Select a sheet first.','error');return;}
        const sheet=workbook.Sheets[sn];if(!sheet||!sheet['!ref']){setSt('Sheet is empty.','error');return;}
        setSt('Processing: '+sn+'...','info');
        tbody.innerHTML='';tcEl.style.display='none';obEl.style.display='none';obEl.textContent='';crEl.style.display='none';ocrEl.style.display='none';csEl.textContent='';osEl.textContent='';alEl.style.display='none';llEl.innerHTML='';arEl.style.display='none';rlEl.innerHTML='';amEl.style.display='none';miEl.innerHTML='';statsEl.style.display='none';statsEl.textContent='';lastRows=[];

        const rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:''});
        if(rows.length<2){setSt('No data rows.','error');return;}

        let hIdx=-1;
        for(let i=0;i<Math.min(rows.length,10);i++){const r=rows[i];if(!r)continue;const j=r.map(c=>String(c).toUpperCase().trim()).join('|');if(j.includes('AZ')&&(j.includes('SITE')||j.includes('POSITION')||j.includes('RACK'))){hIdx=i;break;}}
        if(hIdx===-1){setSt('No header row on "'+sn+'". Try a different sheet.','error');return;}

        const hdr=rows[hIdx].map(h=>String(h).toUpperCase().trim());
        console.log('[azg] Header row',hIdx,':',hdr.join(' | '));

        function fC(...a){for(const x of a){const i=hdr.indexOf(x);if(i>=0)return i;}return-1;}
        function fP(...a){for(const x of a){const i=hdr.findIndex(h=>h.includes(x));if(i>=0)return i;}return-1;}

        const cAZ=fC('AZ'),cSite=fC('SITE','SITE CODE','SITE NAME'),cPos=fC('POSITION','POS','RACK POSITION'),cUp=fP('UPLINK','CONFIG'),cAs=fC('ASSET','ASSET ID','ASSET NUMBER'),cRk=fP('RACK TYPE','TYPE'),cLd=fP('LAND','DATE'),cBr=fC('BRICK','BRICK ID'),cMb=fP('MEASURED');
        if(cAZ<0||cSite<0||cPos<0){setSt('Missing columns: AZ='+cAZ+' Site='+cSite+' Pos='+cPos+'. Try another sheet.','error');return;}

        const midSep=findMiddaySeparatorRow(sheet);
        const grabbed=[],lost=[],rejected=[],middays=[];let skip=0;

        for(let i=hIdx+1;i<rows.length;i++){
            const r=rows[i];if(!r||r.every(c=>!c||String(c).trim()===''))continue;
            const az=r[cAZ]?String(r[cAZ]).trim():'';if(!isAZ(az)){skip++;continue;}
            const site=r[cSite]?String(r[cSite]).trim():'',pos=r[cPos]?String(r[cPos]).trim():'';if(!site||!pos){skip++;continue;}
            const rk=cRk>=0&&r[cRk]?String(r[cRk]).trim():'';if(rk&&!isRack(rk)){skip++;continue;}
            const up=cUp>=0&&r[cUp]?String(r[cUp]).trim():'',as=cAs>=0&&r[cAs]?String(r[cAs]).trim():'',ld=cLd>=0&&r[cLd]?String(r[cLd]).trim():'',br=cBr>=0&&r[cBr]?String(r[cBr]).trim():'',mb=cMb>=0&&r[cMb]?String(r[cMb]).trim():'';
            const off=sheet['!ref']?XLSX.utils.decode_range(sheet['!ref']).s.r:0;
            const isMD=midSep>=0&&(i+off)<=midSep;
            const allTxt=r.map(c=>String(c).toLowerCase()).join(' ');
            const isL=ALERT_LOST_RES.some(k=>allTxt.includes(k)),isR=ALERT_REJECTED.some(k=>allTxt.includes(k));
            const e={site,position:pos,uplink:up,asset:as,rackType:rk,landDate:ld,brick:br,measuredBy:mb,isMidday:isMD,isLost:isL,isRejected:isR,isOverride:false,rowIdx:i,notes:[]};
            if(isL){e.notes.push('LOST RES');lost.push(e);}if(isR){e.notes.push('REJECTED');rejected.push(e);}if(isMD){e.notes.push('MIDDAY');middays.push(e);}
            grabbed.push(e);
        }
        if(!grabbed.length){setSt('No '+AZ_MAP[selectedAZ].label+' positions on "'+sn+'". Skipped '+skip+'.','error');return;}

        sortSite(grabbed);lastRows=grabbed;

        if(lost.length){alEl.style.display='block';lost.forEach(e=>{const li=document.createElement('li');li.textContent=e.site+' — '+e.position+(e.rackType?' ('+e.rackType+')':'');llEl.appendChild(li);});}
        if(rejected.length){arEl.style.display='block';rejected.forEach(e=>{const li=document.createElement('li');li.textContent=e.site+' — '+e.position+(e.rackType?' ('+e.rackType+')':'');rlEl.appendChild(li);});}
        if(middays.length){amEl.style.display='block';miEl.innerHTML='<strong>'+middays.length+' midday position'+(middays.length>1?'s':'')+'</strong><br>Included in Optics, excluded from Cabling.<ul>'+middays.map(e=>'<li>'+esc(e.site)+' — '+esc(e.position)+(e.rackType?' ('+esc(e.rackType)+')':'')+'</li>').join('')+'</ul>';}

        const sg=grpSite(grabbed);
        sg.forEach((g,gi)=>{
            g.rows.forEach(e=>{
                const tr=document.createElement('tr');
                if(e.isLost)tr.classList.add('azg-rl');if(e.isRejected)tr.classList.add('azg-rr');if(e.isMidday)tr.classList.add('azg-rm');
                const ns=e.notes.filter(n=>n!=='MIDDAY').map(n=>'<span class="azg-tag '+(n==='LOST RES'?'azg-tag-l':'azg-tag-j')+'">'+n+'</span>').join(' ');
                const btn=document.createElement('button');btn.className='azg-mt'+(e.isMidday?' azg-im':'');btn.textContent=e.isMidday?'☀ Remove':'🌙 Add';btn.title=e.isMidday?'Remove from midday':'Mark as midday';btn.addEventListener('click',function(){toggleMD(e,this);});
                const cells=['site','position','uplink','asset','rackType','landDate','brick'].map(k=>{const td=document.createElement('td');td.textContent=e[k];return td;});
                const tdM=document.createElement('td');tdM.appendChild(btn);
                const tdN=document.createElement('td');tdN.innerHTML=ns;
                cells.forEach(td=>tr.appendChild(td));tr.appendChild(tdM);tr.appendChild(tdN);
                tbody.appendChild(tr);
            });
            if(gi<sg.length-1){for(let s=0;s<SITE_GAP_ROWS;s++){const d=document.createElement('tr');d.classList.add('azg-sd');d.innerHTML='<td colspan="9"></td>';tbody.appendChild(d);}}
        });
        tcEl.style.display='block';

        const mo=grabbed.filter(e=>!e.isMidday).length,cl=grabbed.filter(e=>!e.isLost&&!e.isRejected).length;
        statsEl.textContent=AZ_MAP[selectedAZ].label+': '+grabbed.length+' total | '+mo+' morning (cabling) | '+middays.length+' midday (optics only) | '+sg.length+' sites | '+lost.length+' lost | '+rejected.length+' rejected | '+skip+' skipped';
        statsEl.style.display='block';
        setSt('✓ Found '+grabbed.length+' positions across '+sg.length+' sites.','success');
        obEl.textContent='19-col TSV (A-S) — Cabling excludes midday, Optics includes all.';obEl.style.display='block';
        if(mo>0)crEl.style.display='flex';if(cl>0)ocrEl.style.display='flex';
    }

    // === CLIPBOARD ===
    async function writeClip(text){if(typeof GM_setClipboard!=='undefined'){try{GM_setClipboard(text,'text');return true;}catch(e){}}try{await navigator.clipboard.writeText(text);return true;}catch(e){}try{const ta=document.createElement('textarea');ta.value=text;ta.style.cssText='position:fixed;left:-9999px';document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);return true;}catch(e){return false;}}

    async function copyCabling(grabbed,alerts){
        const cab=grabbed.filter(e=>!e.isMidday),gs=grpSite(cab),lines=[];
        gs.forEach((g,gi)=>{g.rows.forEach(e=>{const ap=alerts?(e.isLost?'⚠ LOST RES | ':e.isRejected?'🚫 REJECTED | ':''):'';lines.push([e.site,e.position,e.uplink,e.asset,e.rackType,e.landDate,e.brick,DEST.brickHandoffDate,DEST.dateWorkloadReceived,e.measuredBy||'',DEST.copper,DEST.copperSig,DEST.fiber,DEST.fiberSig,'',DEST.mnSig,DEST.optics,DEST.brickPatching,ap+DEST.posStatus].join('\t'));});if(gi<gs.length-1){for(let s=0;s<SITE_GAP_ROWS;s++)lines.push('\t'.repeat(18));}});
        return writeClip(lines.join('\n'));
    }
    function buildOptics(grabbed){
        const cl=grabbed.filter(e=>!e.isLost&&!e.isRejected);if(!cl.length)return'';
        const gs=grpSite(cl),lines=[];
        gs.forEach((g,gi)=>{g.rows.forEach(e=>{lines.push([e.site,e.position,e.uplink,e.asset,e.brick,e.landDate].join('\t'));});if(gi<gs.length-1)lines.push('');});
        return lines.join('\n');
    }

    // === AUTO-GRAB ===
    async function autoGrab(){
        overlay.classList.add('azg-visible');fab.classList.add('azg-loading');setSt('Detecting Excel file...','info');
        shSel.innerHTML='<option value="">-- Select Sheet --</option>';shSel.disabled=true;grabBtn.disabled=true;
        fileInfo.style.display='none';tbody.innerHTML='';tcEl.style.display='none';obEl.style.display='none';crEl.style.display='none';ocrEl.style.display='none';alEl.style.display='none';arEl.style.display='none';amEl.style.display='none';statsEl.style.display='none';lastRows=[];
        const fi=detectFile();
        if(!fi){fab.classList.remove('azg-loading');setSt('No Excel file detected. Navigate to a SharePoint Excel file first.','error');return;}
        fnameEl.textContent=fi.fileName||'Unknown file';fsrcEl.textContent=fi.type+' — '+(fi.url||'').substring(0,90);fileInfo.style.display='flex';
        try{const buf=await fetchXls(fi);workbook=XLSX.read(new Uint8Array(buf),{type:'array',cellStyles:true,cellDates:true,raw:false,dense:false});console.log('[azg] Sheets:',workbook.SheetNames);workbook.SheetNames.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=n;shSel.appendChild(o);});shSel.disabled=false;grabBtn.disabled=false;setSt('Workbook loaded — '+workbook.SheetNames.length+' sheet(s). Select a sheet and click Grab Positions.','success');}catch(e){console.error('[azg]',e);setSt('Error: '+e.message,'error');}finally{fab.classList.remove('azg-loading');}
    }

    // === EVENTS ===
    azSel.addEventListener('change',()=>{selectedAZ=azSel.value;fab.textContent=AZ_MAP[selectedAZ]?.label||'AZ';});
    fab.addEventListener('click',()=>{if(overlay.classList.contains('azg-visible'))overlay.classList.remove('azg-visible');else autoGrab();});
    closeBtn.addEventListener('click',()=>overlay.classList.remove('azg-visible'));
    overlay.addEventListener('click',e=>{if(e.target===overlay)overlay.classList.remove('azg-visible');});
    document.addEventListener('keydown',e=>{if(e.key==='Escape'&&overlay.classList.contains('azg-visible'))overlay.classList.remove('azg-visible');});
    refreshBtn.addEventListener('click',()=>autoGrab());
    grabBtn.addEventListener('click',processSheet);
    copyBtn.addEventListener('click',async()=>{if(!lastRows.length)return;csEl.textContent='Copying...';const ok=await copyCabling(lastRows,false);csEl.textContent=ok?'✓ Copied cabling! Paste into Col A.':'⚠ Copy may have failed.';setTimeout(()=>{csEl.textContent='';},4000);});
    copyaBtn.addEventListener('click',async()=>{if(!lastRows.length)return;csEl.textContent='Copying...';const ok=await copyCabling(lastRows,true);csEl.textContent=ok?'✓ Copied cabling with alerts! Paste into Col A.':'⚠ Copy may have failed.';setTimeout(()=>{csEl.textContent='';},4000);});
    copyoBtn.addEventListener('click',async()=>{if(!lastRows.length)return;osEl.textContent='Copying...';const tsv=buildOptics(lastRows);if(!tsv){osEl.textContent='No clean rows.';setTimeout(()=>{osEl.textContent='';},3000);return;}const ok=await writeClip(tsv);osEl.textContent=ok?'✓ Optics copied! Paste into Col A.':'⚠ Copy may have failed.';setTimeout(()=>{osEl.textContent='';},4000);});
})();


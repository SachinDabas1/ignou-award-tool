#!/usr/bin/env python3
import io, re, zipfile, json, os
from collections import OrderedDict
from flask import Flask, request, send_file, render_template_string, Response, stream_with_context
from docgen import generate_award_list, generate_award_list_batched

app = Flask(__name__)

def fix_zoom(docx_bytes):
    buf = io.BytesIO(docx_bytes)
    out = io.BytesIO()
    with zipfile.ZipFile(buf, 'r') as zin,          zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/settings.xml':
                data = re.sub(
                    rb'(<w:zoom\b(?![^>]*w:percent)[^/]*)(/?>)',
                    rb'\1 w:percent="100"\2', data)
            zout.writestr(item, data)
    return out.getvalue()

# ── HTML UI ───────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>IGNOU RC-71 Filter Tool</title>
<style>
:root{
  --bg:#0d1117;--s:#161b22;--s2:#21262d;--bd:#30363d;
  --ac:#58a6ff;--ac2:#7c5fe6;--gr:#3fb950;--re:#f85149;
  --tx:#c9d1d9;--mu:#8b949e;
  --jun:#e6a817;--dec:#58a6ff;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--tx);font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh}
header{background:linear-gradient(135deg,#0d2137,#161b22);border-bottom:2px solid var(--ac);padding:16px 28px}
.logo{font-size:22px;font-weight:800;letter-spacing:2px;background:linear-gradient(90deg,#58a6ff,#7c5fe6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.sub{color:var(--mu);font-size:11px;letter-spacing:1px;text-transform:uppercase;margin-top:3px}
.container{max-width:1300px;margin:0 auto;padding:24px}

/* Upload */
.upload-zone{background:var(--s);border:2px dashed var(--bd);border-radius:14px;padding:40px 28px;text-align:center;cursor:pointer;transition:all .2s;position:relative;margin-bottom:20px}
.upload-zone:hover,.upload-zone.dragover{border-color:var(--ac);background:#1a2030}
.upload-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.upload-zone h2{font-size:18px;margin:10px 0 6px}
.upload-zone p{color:var(--mu);font-size:13px}

/* Stats */
.stats{display:flex;gap:12px;margin-bottom:18px;flex-wrap:wrap}
.stat{background:var(--s);border:1px solid var(--bd);border-radius:10px;padding:12px 18px;flex:1;min-width:120px}
.stat .l{color:var(--mu);font-size:10px;text-transform:uppercase;letter-spacing:1px}
.stat .v{font-size:22px;font-weight:700;color:var(--ac)}

/* Controls */
.controls{background:var(--s);border:1px solid var(--bd);border-radius:12px;padding:18px 20px;margin-bottom:18px}
.ctrl-header{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:14px}
.ctrl-header h3{font-size:12px;color:var(--mu);text-transform:uppercase;letter-spacing:1px}

/* Session toggle */
.session-toggle{display:flex;gap:0;border-radius:8px;overflow:hidden;border:1px solid var(--bd)}
.session-btn{padding:7px 20px;font-size:13px;font-weight:600;cursor:pointer;border:none;background:var(--s2);color:var(--mu);transition:all .15s}
.session-btn.active-jun{background:rgba(230,168,23,.18);color:var(--jun);border-color:var(--jun)}
.session-btn.active-dec{background:rgba(88,166,255,.18);color:var(--dec)}
.session-btn:first-child{border-right:1px solid var(--bd)}

/* Pills */
.pills{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:14px}
.pill{background:var(--s2);border:1px solid var(--bd);border-radius:20px;padding:5px 13px;font-size:12px;cursor:pointer;transition:all .15s;user-select:none;display:flex;align-items:center;gap:6px}
.pill:hover{border-color:var(--ac);color:var(--ac)}
.pill.on{background:rgba(88,166,255,.15);border-color:var(--ac);color:var(--ac)}
.pill.all.on{background:rgba(124,95,230,.15);border-color:var(--ac2);color:var(--ac2)}
.pill .badge{background:rgba(88,166,255,.25);border-radius:10px;padding:1px 7px;font-size:10px;font-weight:700}

/* Search row */
.row{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap}
.cg{flex:1;min-width:180px}
.cg label{display:block;font-size:11px;color:var(--mu);text-transform:uppercase;letter-spacing:1px;margin-bottom:6px}
input{width:100%;background:var(--s2);border:1px solid var(--bd);border-radius:7px;color:var(--tx);padding:9px 12px;font-size:13px;outline:none;transition:border-color .2s}
input:focus{border-color:var(--ac)}

/* Buttons */
.btn{padding:9px 18px;border-radius:7px;border:none;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap}
.btn-p{background:linear-gradient(135deg,#1f6feb,#7c5fe6);color:#fff}
.btn-g{background:linear-gradient(135deg,#1a7f37,#3fb950);color:#fff}
.btn-o{background:transparent;border:1px solid var(--ac);color:var(--ac)}
.btn:hover{opacity:.85;transform:translateY(-1px)}
.btn:disabled{opacity:.35;cursor:not-allowed;transform:none}

/* Table */
.tw{background:var(--s);border:1px solid var(--bd);border-radius:12px;overflow:hidden;margin-bottom:16px}
.th2{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;border-bottom:1px solid var(--bd);flex-wrap:wrap;gap:8px}
.th2 h3{font-size:14px}
.rc{background:rgba(88,166,255,.12);border:1px solid rgba(88,166,255,.25);color:var(--ac);border-radius:16px;padding:3px 12px;font-size:12px;font-weight:600}
.ts{overflow-x:auto;max-height:480px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:9;background:#0d2137}
th{padding:10px 13px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:var(--mu);border-bottom:1px solid var(--bd);white-space:nowrap}
tbody tr{border-bottom:1px solid rgba(48,54,61,.5);transition:background .1s}
tbody tr:hover{background:var(--s2)}
td{padding:9px 13px;white-space:nowrap}
.ec{font-family:monospace;color:var(--ac);font-weight:600}
.cb{background:rgba(124,95,230,.12);border:1px solid rgba(124,95,230,.25);color:#a78bfa;border-radius:5px;padding:2px 8px;font-size:11px;font-weight:600}
.mc{font-weight:700;color:var(--gr)}
.nd{text-align:center;color:var(--mu);padding:50px;font-size:14px}

/* Export row */
.exp{display:flex;gap:10px;justify-content:flex-end;flex-wrap:wrap;margin-bottom:8px}
.export-info{font-size:12px;color:var(--mu);display:flex;align-items:center;gap:6px;flex:1}
.export-info span{background:rgba(63,185,80,.12);border:1px solid rgba(63,185,80,.25);color:var(--gr);border-radius:12px;padding:2px 10px}

/* Toast */
.toast{position:fixed;bottom:20px;right:20px;padding:12px 20px;border-radius:9px;font-weight:600;font-size:13px;transform:translateY(70px);opacity:0;transition:all .3s;z-index:999;box-shadow:0 6px 20px rgba(0,0,0,.4);max-width:320px}
.toast.show{transform:translateY(0);opacity:1}
.toast.ok{background:#1a7f37;color:#fff}
.toast.err{background:#da3633;color:#fff}

::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:var(--s)}
::-webkit-scrollbar-thumb{background:var(--bd);border-radius:3px}
.hidden{display:none!important}
</style>
</head>
<body>
<header>
  <div class="logo">📋 IGNOU RC-71</div>
  <div class="sub">Assignment Filter &amp; Award List Generator · 100% Offline</div>
</header>

<div class="container">

  <!-- Upload -->
  <div class="upload-zone" id="uz">
    <input type="file" id="fi" accept=".xlsx,.xls,.csv">
    <div style="font-size:44px;margin-bottom:8px">📂</div>
    <h2>Upload Excel / CSV File</h2>
    <p>Drag &amp; drop or click · Works 100% offline</p>
  </div>

  <!-- Stats -->
  <div class="stats hidden" id="sb">
    <div class="stat"><div class="l">Total Records</div><div class="v" id="sTot">0</div></div>
    <div class="stat"><div class="l">Unique Students</div><div class="v" id="sStu">0</div></div>
    <div class="stat"><div class="l">Unique Courses</div><div class="v" id="sCrs">0</div></div>
    <div class="stat"><div class="l">Filtered</div><div class="v" id="sFil">0</div></div>
  </div>

  <!-- Controls -->
  <div class="controls hidden" id="ctrl">
    <div class="ctrl-header">
      <h3>🎯 Filter by Course ID</h3>
      <!-- Session selector -->
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <span style="font-size:12px;color:var(--mu);text-transform:uppercase;letter-spacing:1px">TEE Session:</span>
        <div class="session-toggle">
          <button class="session-btn" id="btnJun" onclick="setSession('Jun')">☀️ June</button>
          <button class="session-btn active-dec" id="btnDec" onclick="setSession('Dec')">❄️ December</button>
        </div>
        <span id="sessionYear" style="font-size:12px;color:var(--mu)">Year: <input id="inpYear" type="number" value="2024" min="2020" max="2035" style="width:70px;display:inline-block;padding:4px 8px;font-size:12px"></span>
      </div>
    </div>

    <div class="pills" id="pills"></div>

    <div class="row">
      <div class="cg">
        <label>Search Enrollment / Name</label>
        <input type="text" id="srch" placeholder="Type to search...">
      </div>
      <button class="btn btn-o" onclick="resetF()">↺ Reset</button>
    </div>
  </div>

  <!-- Table -->
  <div class="tw hidden" id="tw">
    <div class="th2">
      <h3>📄 Candidate Records</h3>
      <span class="rc" id="rc">0 records</span>
    </div>
    <div class="ts"><table><thead id="tH"></thead><tbody id="tB"></tbody></table></div>
  </div>

  <!-- Export -->
  <div class="exp hidden" id="ep">
    <div class="export-info">
      <span id="exportSummary">0 courses · 0 records</span>
      <span style="color:var(--mu)" id="batchHint">→ Each course gets its own full page with header</span>
    </div>
    <button class="btn btn-p" onclick="exportXL()">⬇ Excel</button>
    <button class="btn btn-g" id="btnDoc" onclick="exportDoc()">📄 Export IGNOU Award List (.docx)</button>
  </div>

</div>
<div class="toast" id="toast"></div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
let all=[], hdrs=[], enrollIdx=-1, courseIdx=-1, nameIdx=-1, progIdx=-1;
const sel = new Set();
let session = 'Dec';  // 'Jun' or 'Dec'

// ── Column detection ──
function detectCols(h){
  const n = h.map(x => String(x).toLowerCase().replace(/[\s_\-]/g,''));
  enrollIdx = n.findIndex(x => x.includes('enroll') || x.includes('rollno'));
  courseIdx = n.findIndex(x => x.includes('course') || x.includes('subject'));
  nameIdx   = n.findIndex(x => x.includes('name'));
  progIdx   = n.findIndex(x => x.includes('prog') || x.includes('programme') || x.includes('stream'));
  if(enrollIdx<0) enrollIdx=0;
  if(courseIdx<0) courseIdx=3;
  if(nameIdx<0)   nameIdx=2;
  if(progIdx<0)   progIdx=3;
}

// ── File loading ──
document.getElementById('fi').onchange = e => loadFile(e.target.files[0]);
const uz = document.getElementById('uz');
uz.ondragover = e => { e.preventDefault(); uz.classList.add('dragover'); };
uz.ondragleave = () => uz.classList.remove('dragover');
uz.ondrop = e => { e.preventDefault(); uz.classList.remove('dragover'); loadFile(e.dataTransfer.files[0]); };

function loadFile(f){
  if(!f) return;
  const r = new FileReader();
  r.onload = e => {
    try{
      const wb  = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, defval:''});
      let hi = 0;
      for(let i=0; i<Math.min(6,raw.length); i++){
        if(raw[i].filter(c=>String(c).trim()).length >= 3){ hi=i; break; }
      }
      hdrs = raw[hi].map(h => String(h).trim());
      all  = [];
      for(let i=hi+1; i<raw.length; i++){
        const row = raw[i];
        if(!row.some(c => String(c).trim())) continue;
        const obj = {};
        hdrs.forEach((h,j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
        all.push(obj);
      }
      detectCols(hdrs);
      uz.querySelector('h2').textContent = '✅ Loaded: ' + f.name;
      uz.querySelector('p').textContent  = all.length + ' records · Sheet: ' + wb.SheetNames[0];
      buildPills(); updateStats(); render();
      ['sb','ctrl','tw','ep'].forEach(id => show(id));
      toast('✅ Loaded ' + all.length + ' records!');
    } catch(err){ toast('Error: ' + err.message, true); }
  };
  r.readAsArrayBuffer(f);
}

// ── Session ──
function setSession(s){
  session = s;
  document.getElementById('btnJun').className = 'session-btn' + (s==='Jun' ? ' active-jun' : '');
  document.getElementById('btnDec').className = 'session-btn' + (s==='Dec' ? ' active-dec' : '');
}

function getSessionLabel(){
  const yr = document.getElementById('inpYear').value || '2024';
  return (session === 'Jun' ? 'Jun' : 'Dec') + ' ' + yr;
}

// ── Pills ──
function buildPills(){
  const ck = hdrs[courseIdx];
  // Count students per course
  const counts = {};
  all.forEach(r => {
    const c = String(r[ck]||'').trim();
    if(c) counts[c] = (counts[c]||0) + 1;
  });
  const courses = Object.keys(counts).sort();

  const container = document.getElementById('pills');
  container.innerHTML = '';

  // "All" pill
  const ap = document.createElement('span');
  ap.className = 'pill all on';
  ap.innerHTML = '★ All Courses <span class="badge">' + all.length + '</span>';
  ap.dataset.c = '__ALL__';
  ap.onclick = () => toggle('__ALL__', ap);
  container.appendChild(ap);

  courses.forEach(cr => {
    const p = document.createElement('span');
    p.className = 'pill';
    p.innerHTML = cr + ' <span class="badge">' + counts[cr] + '</span>';
    p.dataset.c  = cr;
    p.onclick    = () => toggle(cr, p);
    container.appendChild(p);
  });
  sel.clear();
  document.getElementById('srch').oninput = render;
}

function toggle(c, el){
  const pills = document.querySelectorAll('#pills .pill');
  if(c === '__ALL__'){
    sel.clear();
    pills.forEach(p => p.classList.remove('on'));
    el.classList.add('on');
  } else {
    document.querySelector('.pill.all').classList.remove('on');
    sel.has(c) ? sel.delete(c) : sel.add(c);
    el.classList.toggle('on');
    if(!sel.size) document.querySelector('.pill.all').classList.add('on');
  }
  render();
}

function resetF(){
  sel.clear();
  document.querySelectorAll('#pills .pill').forEach(p => p.classList.remove('on'));
  document.querySelector('.pill.all').classList.add('on');
  document.getElementById('srch').value = '';
  render();
}

// ── Filtering & sorting ──
function sortByEnrollment(arr){
  const ek = hdrs[enrollIdx];
  return arr.slice().sort((a,b) => {
    const ea = String(a[ek]||'').trim(), eb = String(b[ek]||'').trim();
    const na = parseInt(ea.replace(/\D/g,'')), nb = parseInt(eb.replace(/\D/g,''));
    return (!isNaN(na)&&!isNaN(nb)) ? na-nb : ea.localeCompare(eb);
  });
}

function getFiltered(){
  const s  = document.getElementById('srch').value.toLowerCase().trim();
  const ck = hdrs[courseIdx];
  let f = all.filter(r => {
    if(sel.size && !sel.has(String(r[ck]||'').trim())) return false;
    if(s) return Object.values(r).some(v => String(v).toLowerCase().includes(s));
    return true;
  });
  return sortByEnrollment(f);
}

// ── Render table ──
function render(){
  const f  = getFiltered();
  const ek = hdrs[enrollIdx], ck = hdrs[courseIdx];
  document.getElementById('tH').innerHTML = '<tr><th>#</th>' + hdrs.map(h=>'<th>'+h+'</th>').join('') + '</tr>';
  const tb = document.getElementById('tB');
  if(!f.length){
    tb.innerHTML = '<tr><td colspan="'+(hdrs.length+1)+'" class="nd">No records match your filters.</td></tr>';
  } else {
    tb.innerHTML = f.map((r,i) => {
      const cells = hdrs.map(h => {
        const v = r[h] !== undefined ? r[h] : '';
        if(h===ek) return '<td class="ec">'+v+'</td>';
        if(h===ck) return '<td><span class="cb">'+v+'</span></td>';
        if(String(h).toLowerCase().includes('mark')||String(h).toLowerCase().includes('obtain'))
          return '<td class="mc">'+v+'</td>';
        return '<td>'+v+'</td>';
      }).join('');
      return '<tr><td style="color:var(--mu);font-size:11px">'+(i+1)+'</td>'+cells+'</tr>';
    }).join('');
  }
  document.getElementById('rc').textContent = f.length + ' record' + (f.length!==1?'s':'');
  document.getElementById('sFil').textContent = f.length;

  // Update export summary
  const activeCourses = sel.size ? [...sel] : [...new Set(all.map(r=>String(r[hdrs[courseIdx]]||'').trim()))].filter(Boolean);
  const nCourses  = activeCourses.length;
  const nBatches  = Math.ceil(nCourses / BATCH_SIZE);
  document.getElementById('exportSummary').textContent =
    nCourses + ' course' + (nCourses!==1?'s':'') + ' · ' + f.length + ' records';
  const hint = document.getElementById('batchHint');
  if(nBatches > 1){
    hint.textContent = '→ ' + nBatches + ' files will download one after another (10 courses each, each with full heading)';
    hint.style.color = 'var(--ac)';
  } else {
    hint.textContent = '→ Each course gets its own page with full heading · max 10 courses per file';
    hint.style.color = 'var(--mu)';
  }
}

function updateStats(){
  const ek = hdrs[enrollIdx], ck = hdrs[courseIdx];
  document.getElementById('sTot').textContent = all.length;
  document.getElementById('sStu').textContent = new Set(all.map(r=>r[ek])).size;
  document.getElementById('sCrs').textContent = new Set(all.map(r=>r[ck])).size;
  document.getElementById('sFil').textContent = all.length;
}

// ── Build course→candidates map (one entry per course, sorted) ──
function buildCourseMap(){
  const ck = hdrs[courseIdx];
  const ek = hdrs[enrollIdx];
  const nk = hdrs[nameIdx];
  const pk = hdrs[progIdx];
  const f  = getFiltered();

  // Group by course, preserving sorted order per course
  const map = {};
  f.forEach(r => {
    const course = String(r[ck]||'').trim();
    if(!course) return;
    if(!map[course]) map[course] = [];
    map[course].push({
      enrollment: String(r[ek]||'').trim(),
      name:       String(r[nk]||'').trim(),
      programme:  String(r[pk]||'').trim()
    });
  });
  return map;
}

// ── Export to Excel ──
function exportXL(){
  const f = getFiltered();
  if(!f.length){ toast('No data to export!', true); return; }
  const ws = XLSX.utils.aoa_to_sheet([hdrs, ...f.map(r => hdrs.map(h => r[h]||''))]);
  ws['!cols'] = hdrs.map(() => ({wch:20}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filtered');
  const name = sel.size ? [...sel].join('_') : 'ALL';
  XLSX.writeFile(wb, 'IGNOU_' + name + '.xlsx');
  toast('✅ Excel exported!');
}

// ── Export IGNOU Word doc — streams batches from server one at a time ──
const BATCH_SIZE = 10;

async function exportDoc(){
  const courseMap   = buildCourseMap();
  const courseKeys  = Object.keys(courseMap);
  const courseCount = courseKeys.length;
  if(!courseCount){ toast('No data to export!', true); return; }

  const btn          = document.getElementById('btnDoc');
  btn.disabled       = true;
  const sessionLabel = getSessionLabel();
  const safeSession  = sessionLabel.replace(' ','_');

  try{
    if(courseCount <= BATCH_SIZE){
      // ── Fast path: single file, existing endpoint ──
      btn.textContent = '⏳ Generating ' + courseCount + ' page' + (courseCount>1?'s':'') + '...';
      const resp = await fetch('/generate_doc', {
        method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ courseMap, sessionLabel })
      });
      if(!resp.ok) throw new Error(await resp.text());
      triggerDownload(await resp.blob(), 'IGNOU_Award_' + safeSession + '.docx');
      toast('✅ ' + courseCount + ' pages exported!');

    } else {
      // ── Streaming path: server yields one batch at a time ──
      btn.textContent = '⏳ Connecting…';
      const resp = await fetch('/generate_doc_stream', {
        method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ courseMap, sessionLabel })
      });
      if(!resp.ok) throw new Error(await resp.text());

      // Parse the multipart boundary from Content-Type header
      const ct       = resp.headers.get('Content-Type') || '';
      const bmatch   = ct.match(/boundary=(\S+)/);
      if(!bmatch) throw new Error('No multipart boundary in response');
      const boundary = '--' + bmatch[1];

      // Read the entire stream into a Uint8Array
      // (stream is split into small parts — only the current part's bytes
      //  are decoded at once in the loop below)
      const reader  = resp.body.getReader();
      let   chunks  = [];
      let   received = 0;

      while(true){
        const { done, value } = await reader.read();
        if(done) break;
        chunks.push(value);
        received += value.byteLength;
        btn.textContent = '⏳ Receiving… ' + (received/1024).toFixed(0) + ' KB';
      }

      // Concatenate all chunks
      const fullBuf  = new Uint8Array(received);
      let   offset   = 0;
      for(const ch of chunks){ fullBuf.set(ch, offset); offset += ch.byteLength; }
      chunks = null; // free

      // Split on boundary and trigger individual downloads
      const enc      = new TextEncoder();
      const boundBuf = enc.encode(boundary);
      const parts    = splitMultipart(fullBuf, boundBuf);
      let   saved    = 0;

      for(const part of parts){
        const { headers, body } = parseMultipartPart(part);
        const disp     = headers['content-disposition'] || '';
        const fnMatch  = disp.match(/filename="([^"]+)"/);
        const batchHdr = headers['x-batch'] || '';
        if(!fnMatch || !body) continue;
        const filename = fnMatch[1];
        saved++;
        btn.textContent = '💾 Saving file ' + batchHdr + ' — ' + filename;
        triggerDownload(new Blob([body], {type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}), filename);
        // Small delay so browser doesn't block multiple downloads
        await new Promise(r => setTimeout(r, 400));
      }

      toast('✅ ' + saved + ' file' + (saved>1?'s':'') + ' downloaded (' + courseCount + ' courses total)!');
    }

  } catch(err){
    toast('❌ ' + err.message, true);
  } finally {
    btn.disabled    = false;
    btn.textContent = '📄 Export IGNOU Award List (.docx)';
  }
}

// ── Multipart binary splitter ──
function splitMultipart(buf, boundary){
  // Returns array of Uint8Arrays, one per part (excluding the boundary lines)
  const parts   = [];
  let   pos     = 0;
  const crlf    = new Uint8Array([13, 10]);

  function indexOf(haystack, needle, from){
    outer: for(let i = from; i <= haystack.length - needle.length; i++){
      for(let j = 0; j < needle.length; j++){
        if(haystack[i+j] !== needle[j]) continue outer;
      }
      return i;
    }
    return -1;
  }

  while(pos < buf.length){
    const bStart = indexOf(buf, boundary, pos);
    if(bStart === -1) break;
    const afterBoundary = bStart + boundary.length;
    // Skip \r\n after boundary
    const lineEnd = indexOf(buf, crlf, afterBoundary);
    if(lineEnd === -1) break;
    const partStart = lineEnd + 2; // after \r\n
    // Find next boundary
    const nextB = indexOf(buf, boundary, partStart);
    if(nextB === -1) break;
    // Part ends just before \r\n--boundary
    const partEnd = nextB - 2;
    if(partEnd > partStart){
      parts.push(buf.slice(partStart, partEnd));
    }
    pos = nextB;
  }
  return parts;
}

function parseMultipartPart(partBuf){
  // Split headers from body (separated by \r\n\r\n)
  const sep = new Uint8Array([13,10,13,10]);
  function indexOf(h,n,f){
    outer: for(let i=f; i<=h.length-n.length; i++){
      for(let j=0;j<n.length;j++) if(h[i+j]!==n[j]) continue outer;
      return i;
    }
    return -1;
  }
  const sepIdx = indexOf(partBuf, sep, 0);
  if(sepIdx === -1) return { headers: {}, body: null };

  const headerText = new TextDecoder().decode(partBuf.slice(0, sepIdx));
  const body       = partBuf.slice(sepIdx + 4);
  const headers    = {};
  for(const line of headerText.split('\r\n')){
    const ci = line.indexOf(':');
    if(ci > -1){
      headers[line.slice(0,ci).trim().toLowerCase()] = line.slice(ci+1).trim();
    }
  }
  return { headers, body };
}

function triggerDownload(blob, filename){
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href    = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function show(id){ document.getElementById(id).classList.remove('hidden'); }
function toast(m, isErr=false){
  const t = document.getElementById('toast');
  t.textContent = m;
  t.className   = 'toast ' + (isErr ? 'err' : 'ok');
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 4000);
}
</script>
</body>
</html>"""

@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/generate_doc', methods=['POST'])
def generate_doc():
    """
    Single-batch endpoint (≤10 courses).
    Returns one .docx file directly.
    """
    data          = request.json
    course_map    = data.get('courseMap', {})
    session_label = data.get('sessionLabel', 'Dec 2024')

    if not course_map:
        return 'No data provided', 400

    ordered = OrderedDict((k, v) for k, v in course_map.items())
    buf      = generate_award_list(ordered, session_label)
    raw      = buf.read()
    fixed    = fix_zoom(raw)

    safe_s      = session_label.replace(' ', '_')
    keys        = list(ordered.keys())
    courses_str = '_'.join(keys[:3]) + (f'_+{len(keys)-3}more' if len(keys) > 3 else '')
    filename    = f'IGNOU_Award_{safe_s}_{courses_str}.docx'

    return send_file(
        io.BytesIO(fixed),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=filename
    )


@app.route('/generate_doc_stream', methods=['POST'])
def generate_doc_stream():
    """
    Streaming multipart endpoint for large sets (>10 courses).

    Yields each 10-course batch as a separate part of a multipart/mixed response
    immediately after generating it — only one Document object lives in RAM at
    any point.  The browser reads parts one at a time and triggers individual
    file downloads via the filenames embedded in Content-Disposition headers.

    Wire format (each part):
        --BOUNDARY\\r\\n
        Content-Type: application/octet-stream\\r\\n
        Content-Disposition: attachment; filename="...docx"\\r\\n
        X-Batch: 1/3\\r\\n
        \\r\\n
        <raw .docx bytes>
        \\r\\n
    Final:
        --BOUNDARY--\\r\\n
    """
    data          = request.json
    course_map    = data.get('courseMap', {})
    session_label = data.get('sessionLabel', 'Dec 2024')

    if not course_map:
        return 'No data provided', 400

    ordered    = OrderedDict((k, v) for k, v in course_map.items())
    safe_s     = session_label.replace(' ', '_')
    BOUNDARY   = 'IGNOU_BATCH_BOUNDARY'

    def stream():
        for batch_num, total_batches, buf in generate_award_list_batched(
                ordered, session_label, batch_size=10):

            raw   = buf.read()
            fixed = fix_zoom(raw)

            keys        = list(ordered.keys())
            start_idx   = (batch_num - 1) * 10
            end_idx     = min(batch_num * 10, len(keys)) - 1
            part_name   = (
                f'IGNOU_Award_{safe_s}'
                f'_Part{str(batch_num).zfill(2)}'
                f'_{keys[start_idx]}'
                + (f'-{keys[end_idx]}' if end_idx > start_idx else '')
                + '.docx'
            )

            headers = (
                f'--{BOUNDARY}\r\n'
                f'Content-Type: application/octet-stream\r\n'
                f'Content-Disposition: attachment; filename="{part_name}"\r\n'
                f'X-Batch: {batch_num}/{total_batches}\r\n'
                f'Content-Length: {len(fixed)}\r\n'
                f'\r\n'
            )
            yield headers.encode('utf-8')
            yield fixed
            yield b'\r\n'

        yield f'--{BOUNDARY}--\r\n'.encode('utf-8')

    return app.response_class(
        stream(),
        mimetype=f'multipart/mixed; boundary={BOUNDARY}',
        status=200
    )


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5050))
    app.run(host='0.0.0.0', port=port, debug=False)

#!/usr/bin/env python3
"""
IGNOU RC-71 Assignment Filter & Award List Generator
Run: python3 app.py   →  open http://localhost:5050
"""
import io, re, zipfile, json
from collections import OrderedDict
from flask import Flask, request, send_file, render_template_string
from docgen import generate_award_list

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>IGNOU RC-71 · Award List Generator</title>
<style>
:root{
  --bg:#0d1117;--s:#161b22;--s2:#21262d;--bd:#30363d;
  --ac:#58a6ff;--ac2:#7c5fe6;--gr:#3fb950;--re:#f85149;
  --tx:#c9d1d9;--mu:#8b949e;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--tx);font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh}

/* ── Header ── */
header{
  background:linear-gradient(135deg,#0d2137 0%,#161b22 100%);
  border-bottom:2px solid var(--ac);
  padding:14px 28px;display:flex;align-items:center;gap:14px;
}
.logo{font-size:22px;font-weight:800;letter-spacing:2px;
  background:linear-gradient(90deg,#58a6ff,#7c5fe6);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent}
.sub{color:var(--mu);font-size:11px;letter-spacing:1px;text-transform:uppercase;margin-top:3px}

.container{max-width:1320px;margin:0 auto;padding:22px 20px}

/* ── Upload ── */
.upload-zone{
  background:var(--s);border:2px dashed var(--bd);border-radius:14px;
  padding:40px 28px;text-align:center;cursor:pointer;
  transition:all .2s;position:relative;margin-bottom:18px;
}
.upload-zone:hover,.upload-zone.drag{border-color:var(--ac);background:#1a2030}
.upload-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.upload-zone h2{font-size:17px;margin:8px 0 5px}
.upload-zone p{color:var(--mu);font-size:13px}
.uz-icon{font-size:42px;margin-bottom:4px}

/* ── Stats ── */
.stats{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}
.stat{background:var(--s);border:1px solid var(--bd);border-radius:10px;
  padding:11px 16px;flex:1;min-width:110px}
.stat .l{color:var(--mu);font-size:10px;text-transform:uppercase;letter-spacing:1px}
.stat .v{font-size:22px;font-weight:700;color:var(--ac)}

/* ── Controls card ── */
.card{background:var(--s);border:1px solid var(--bd);border-radius:12px;
  padding:16px 20px;margin-bottom:16px}
.card-title{font-size:11px;color:var(--mu);text-transform:uppercase;
  letter-spacing:1px;margin-bottom:12px;font-weight:600}

/* ── Session picker ── */
.session-row{display:flex;gap:10px;align-items:center;margin-bottom:14px;flex-wrap:wrap}
.session-label{font-size:12px;color:var(--mu);text-transform:uppercase;
  letter-spacing:1px;font-weight:600;white-space:nowrap}
.sess-btn{
  padding:7px 18px;border-radius:20px;border:1px solid var(--bd);
  background:var(--s2);color:var(--tx);font-size:13px;font-weight:600;
  cursor:pointer;transition:all .15s;
}
.sess-btn.on{background:rgba(88,166,255,.18);border-color:var(--ac);color:var(--ac)}
.sess-btn:hover{border-color:var(--ac)}

/* ── Course pills ── */
.pills{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px}
.pill{
  background:var(--s2);border:1px solid var(--bd);border-radius:20px;
  padding:5px 13px;font-size:12px;cursor:pointer;
  transition:all .15s;user-select:none;
}
.pill:hover{border-color:var(--ac);color:var(--ac)}
.pill.on{background:rgba(88,166,255,.15);border-color:var(--ac);color:var(--ac)}
.pill.all.on{background:rgba(124,95,230,.15);border-color:var(--ac2);color:var(--ac2)}

/* ── Search + reset row ── */
.row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
.cg{flex:1;min-width:180px}
.cg label{display:block;font-size:11px;color:var(--mu);
  text-transform:uppercase;letter-spacing:1px;margin-bottom:6px}
input[type=text]{
  width:100%;background:var(--s2);border:1px solid var(--bd);
  border-radius:7px;color:var(--tx);padding:9px 12px;
  font-size:13px;outline:none;transition:border-color .2s;
}
input[type=text]:focus{border-color:var(--ac)}

/* ── Buttons ── */
.btn{padding:9px 18px;border-radius:7px;border:none;
  font-size:13px;font-weight:600;cursor:pointer;
  transition:all .15s;white-space:nowrap}
.btn-g{background:linear-gradient(135deg,#1a7f37,#3fb950);color:#fff}
.btn-p{background:linear-gradient(135deg,#1f6feb,#7c5fe6);color:#fff}
.btn-o{background:transparent;border:1px solid var(--ac);color:var(--ac)}
.btn:hover{opacity:.85;transform:translateY(-1px)}
.btn:disabled{opacity:.35;cursor:not-allowed;transform:none}

/* ── Table ── */
.tw{background:var(--s);border:1px solid var(--bd);border-radius:12px;
  overflow:hidden;margin-bottom:14px}
.tw-head{display:flex;justify-content:space-between;align-items:center;
  padding:12px 16px;border-bottom:1px solid var(--bd)}
.tw-head h3{font-size:14px}
.rc{background:rgba(88,166,255,.12);border:1px solid rgba(88,166,255,.25);
  color:var(--ac);border-radius:16px;padding:3px 12px;
  font-size:12px;font-weight:600}
.ts{overflow-x:auto;max-height:460px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:9;background:#0d2137}
th{padding:9px 12px;text-align:left;font-size:10px;text-transform:uppercase;
  letter-spacing:1px;color:var(--mu);border-bottom:1px solid var(--bd);white-space:nowrap}
tbody tr{border-bottom:1px solid rgba(48,54,61,.5);transition:background .1s}
tbody tr:hover{background:var(--s2)}
td{padding:8px 12px;white-space:nowrap}
.ec{font-family:monospace;color:var(--ac);font-weight:600}
.cb{background:rgba(124,95,230,.12);border:1px solid rgba(124,95,230,.25);
  color:#a78bfa;border-radius:5px;padding:2px 8px;font-size:11px;font-weight:600}
.mc{font-weight:700;color:var(--gr)}
.nd{text-align:center;color:var(--mu);padding:50px;font-size:14px}

/* ── Export row ── */
.exp{display:flex;gap:10px;justify-content:flex-end;flex-wrap:wrap;margin-bottom:8px}

/* ── Selected courses summary ── */
.sel-summary{
  background:rgba(88,166,255,.07);border:1px solid rgba(88,166,255,.2);
  border-radius:8px;padding:8px 14px;font-size:12px;color:var(--mu);
  margin-bottom:10px;line-height:1.6;
}
.sel-summary strong{color:var(--ac)}

/* ── Toast ── */
.toast{
  position:fixed;bottom:20px;right:20px;padding:12px 20px;
  border-radius:9px;font-weight:600;font-size:13px;
  transform:translateY(70px);opacity:0;transition:all .3s;
  z-index:999;box-shadow:0 6px 20px rgba(0,0,0,.4);
}
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
  <div>
    <div class="logo">📋 IGNOU RC-71</div>
    <div class="sub">Assignment Filter &amp; Award List Generator · 100% Offline</div>
  </div>
</header>

<div class="container">

  <!-- Upload -->
  <div class="upload-zone" id="uz">
    <input type="file" id="fi" accept=".xlsx,.xls,.csv">
    <div class="uz-icon">📂</div>
    <h2>Upload Excel / CSV File</h2>
    <p>Drag &amp; drop or click to browse · Works 100% offline</p>
  </div>

  <!-- Stats -->
  <div class="stats hidden" id="sb">
    <div class="stat"><div class="l">Total Records</div><div class="v" id="sTot">0</div></div>
    <div class="stat"><div class="l">Unique Students</div><div class="v" id="sStu">0</div></div>
    <div class="stat"><div class="l">Unique Courses</div><div class="v" id="sCrs">0</div></div>
    <div class="stat"><div class="l">Filtered</div><div class="v" id="sFil">0</div></div>
  </div>

  <!-- Controls -->
  <div class="card hidden" id="ctrl">
    <div class="card-title">🗓 Session</div>
    <div class="session-row">
      <span class="session-label">TEE Session:</span>
      <button class="sess-btn on" id="sJun" onclick="setSession('Jun')">☀️ June</button>
      <button class="sess-btn"   id="sDec" onclick="setSession('Dec')">❄️ December</button>
      <span id="sessionYear" style="color:var(--mu);font-size:12px">Year: <strong id="yrDisplay" style="color:var(--ac)">2024</strong></span>
      <input type="number" id="yrInput" value="2024" min="2020" max="2030"
        style="width:80px;background:var(--s2);border:1px solid var(--bd);border-radius:7px;
               color:var(--tx);padding:6px 10px;font-size:13px;outline:none"
        oninput="document.getElementById('yrDisplay').textContent=this.value">
    </div>

    <div class="card-title">🎯 Filter by Course ID</div>
    <div class="pills" id="pills"></div>

    <div class="row">
      <div class="cg">
        <label>Search Enrollment / Name</label>
        <input type="text" id="srch" placeholder="Type to search...">
      </div>
      <button class="btn btn-o" onclick="resetF()">↺ Reset All</button>
    </div>
  </div>

  <!-- Selection summary -->
  <div class="sel-summary hidden" id="selSum"></div>

  <!-- Table -->
  <div class="tw hidden" id="tw">
    <div class="tw-head">
      <h3>📄 Candidate Records</h3>
      <span class="rc" id="rc">0 records</span>
    </div>
    <div class="ts">
      <table><thead id="tH"></thead><tbody id="tB"></tbody></table>
    </div>
  </div>

  <!-- Export -->
  <div class="exp hidden" id="ep">
    <button class="btn btn-g" id="btnDoc" onclick="exportDoc()">
      📄 Export IGNOU Award List (.docx) — 1 page per course
    </button>
    <button class="btn btn-p" onclick="exportXL()">⬇ Export to Excel</button>
  </div>

</div>
<div class="toast" id="toast"></div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
// ── State ──
let all=[], hdrs=[], enrollIdx=-1, courseIdx=-1, nameIdx=-1, progIdx=-1;
const sel = new Set();
let sessionMonth = 'Jun';

// ── Session ──
function setSession(m){
  sessionMonth = m;
  document.getElementById('sJun').classList.toggle('on', m==='Jun');
  document.getElementById('sDec').classList.toggle('on', m==='Dec');
}

function getSessionLabel(){
  const yr = document.getElementById('yrInput').value || '2024';
  return sessionMonth + ' ' + yr;   // e.g. "Jun 2024" or "Dec 2024"
}

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
uz.ondragover = e => { e.preventDefault(); uz.classList.add('drag'); };
uz.ondragleave = () => uz.classList.remove('drag');
uz.ondrop = e => { e.preventDefault(); uz.classList.remove('drag'); loadFile(e.dataTransfer.files[0]); };

function loadFile(f){
  if(!f) return;
  const r = new FileReader();
  r.onload = e => {
    try{
      const wb  = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1,defval:''});
      let hi = 0;
      for(let i=0; i<Math.min(6,raw.length); i++){
        if(raw[i].filter(c=>String(c).trim()).length >= 3){ hi=i; break; }
      }
      hdrs = raw[hi].map(h => String(h).trim());
      all  = [];
      for(let i=hi+1; i<raw.length; i++){
        const row = raw[i];
        if(!row.some(c=>String(c).trim())) continue;
        const obj = {};
        hdrs.forEach((h,j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
        all.push(obj);
      }
      detectCols(hdrs);
      uz.querySelector('h2').textContent = '✅ Loaded: ' + f.name;
      uz.querySelector('p').textContent  = all.length + ' records · ' + wb.SheetNames[0];
      buildPills();
      updateStats();
      render();
      ['sb','ctrl','tw','ep'].forEach(id => show(id));
      toast('✅ Loaded ' + all.length + ' records!');
    } catch(err){ toast('Error: ' + err.message, true); }
  };
  r.readAsArrayBuffer(f);
}

// ── Pills ──
function buildPills(){
  const ck = hdrs[courseIdx];
  const courses = [...new Set(all.map(r => String(r[ck]||'').trim()))].filter(Boolean).sort();
  const c = document.getElementById('pills');
  c.innerHTML = '';
  const ap = document.createElement('span');
  ap.className = 'pill all on'; ap.textContent = '★ All Courses'; ap.dataset.c = '__ALL__';
  ap.onclick = () => toggle('__ALL__', ap);
  c.appendChild(ap);
  courses.forEach(cr => {
    const p = document.createElement('span');
    p.className = 'pill'; p.textContent = cr; p.dataset.c = cr;
    p.onclick = () => toggle(cr, p);
    c.appendChild(p);
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

// ── Filtering ──
function getFiltered(){
  const s  = document.getElementById('srch').value.toLowerCase().trim();
  const ck = hdrs[courseIdx], ek = hdrs[enrollIdx];
  let f = all.filter(r => {
    if(sel.size && !sel.has(String(r[ck]||'').trim())) return false;
    if(s) return Object.values(r).some(v => String(v).toLowerCase().includes(s));
    return true;
  });
  f.sort((a,b) => {
    const ea = String(a[ek]||'').trim(), eb = String(b[ek]||'').trim();
    const na = parseInt(ea.replace(/\D/g,'')), nb = parseInt(eb.replace(/\D/g,''));
    return (!isNaN(na) && !isNaN(nb)) ? na-nb : ea.localeCompare(eb);
  });
  return f;
}

// ── Render table ──
function render(){
  const f  = getFiltered();
  const ek = hdrs[enrollIdx], ck = hdrs[courseIdx];

  document.getElementById('tH').innerHTML =
    '<tr><th>#</th>' + hdrs.map(h=>'<th>'+h+'</th>').join('') + '</tr>';

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

  // Selected courses summary
  const ck2 = hdrs[courseIdx];
  const activeCourses = sel.size > 0 ? [...sel].sort() :
    [...new Set(all.map(r=>String(r[ck2]||'').trim()))].filter(Boolean).sort();
  const sumEl = document.getElementById('selSum');
  if(activeCourses.length > 1){
    const counts = activeCourses.map(c => {
      const n = f.filter(r => String(r[ck2]||'').trim()===c).length;
      return `<strong>${c}</strong> (${n})`;
    });
    sumEl.innerHTML = '📄 Will generate <strong>' + activeCourses.length +
      ' pages</strong> — one per course: ' + counts.join(' · ');
    show('selSum');
  } else if(activeCourses.length === 1){
    const n = f.filter(r=>String(r[ck2]||'').trim()===activeCourses[0]).length;
    sumEl.innerHTML = '📄 1 page for <strong>'+activeCourses[0]+'</strong> — '+n+' candidate'+(n!==1?'s':'');
    show('selSum');
  } else {
    sumEl.classList.add('hidden');
  }
}

function updateStats(){
  const ek = hdrs[enrollIdx], ck = hdrs[courseIdx];
  document.getElementById('sTot').textContent = all.length;
  document.getElementById('sStu').textContent = new Set(all.map(r=>r[ek])).size;
  document.getElementById('sCrs').textContent = new Set(all.map(r=>r[ck])).size;
  document.getElementById('sFil').textContent = all.length;
}

// ── Excel export ──
function exportXL(){
  const f = getFiltered();
  if(!f.length){ toast('No data!', true); return; }
  const ws = XLSX.utils.aoa_to_sheet([hdrs, ...f.map(r=>hdrs.map(h=>r[h]||''))]);
  ws['!cols'] = hdrs.map(()=>({wch:20}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filtered');
  XLSX.writeFile(wb, 'IGNOU_' + ([...sel].join('_')||'ALL') + '.xlsx');
  toast('✅ Excel exported!');
}

// ── Word export (one page per course) ──
async function exportDoc(){
  const f = getFiltered();
  if(!f.length){ toast('No data to export!', true); return; }

  const btn = document.getElementById('btnDoc');
  btn.disabled = true;
  btn.textContent = '⏳ Generating...';

  try {
    const ek = hdrs[enrollIdx], nk = hdrs[nameIdx], pk = hdrs[progIdx], ck = hdrs[courseIdx];

    // Group by course (preserving sorted order)
    const courseMap = {};
    f.forEach(r => {
      const course = String(r[ck]||'').trim();
      if(!courseMap[course]) courseMap[course] = [];
      courseMap[course].push({
        enrollment: String(r[ek]||''),
        name:       String(r[nk]||''),
        programme:  String(r[pk]||''),
      });
    });

    const payload = {
      courseMap,
      sessionLabel: getSessionLabel(),
    };

    const resp = await fetch('/generate_doc', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(payload),
    });
    if(!resp.ok) throw new Error(await resp.text());

    const blob = await resp.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const sessionLabel = getSessionLabel().replace(' ','_');
    const courses = Object.keys(courseMap);
    const fname = courses.length === 1
      ? `IGNOU_Award_${courses[0]}_${sessionLabel}.docx`
      : `IGNOU_Award_${courses.length}_Courses_${sessionLabel}.docx`;
    a.href = url; a.download = fname; a.click();
    URL.revokeObjectURL(url);

    toast(`✅ Generated ${courses.length} page${courses.length!==1?'s':''} — downloaded!`);
  } catch(err){
    toast('Error: ' + err.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = '📄 Export IGNOU Award List (.docx) — 1 page per course';
  }
}

function show(id){ document.getElementById(id).classList.remove('hidden'); }
function toast(m, isErr=false){
  const t = document.getElementById('toast');
  t.textContent = m;
  t.className   = 'toast ' + (isErr ? 'err' : 'ok');
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3500);
}
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────

def fix_zoom(docx_bytes):
    """Fix missing w:percent on w:zoom (python-docx default bug)."""
    buf = io.BytesIO(docx_bytes)
    out = io.BytesIO()
    with zipfile.ZipFile(buf, 'r') as zin, \
         zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/settings.xml':
                data = re.sub(
                    rb'(<w:zoom\b(?![^>]*w:percent)[^/]*)(/?>)',
                    rb'\1 w:percent="100"\2',
                    data
                )
            zout.writestr(item, data)
    return out.getvalue()


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/generate_doc', methods=['POST'])
def generate_doc():
    data          = request.json
    course_map    = data.get('courseMap', {})
    session_label = data.get('sessionLabel', 'Dec 2024')

    # course_map is {courseCode: [candidate dicts]} already grouped by JS
    buf       = generate_award_list(course_map, session_label)
    raw_bytes = buf.read()
    fixed     = fix_zoom(raw_bytes)

    courses   = list(course_map.keys())
    if len(courses) == 1:
        fname = f'IGNOU_Award_{courses[0]}_{session_label.replace(" ","_")}.docx'
    else:
        fname = f'IGNOU_Award_{len(courses)}_Courses_{session_label.replace(" ","_")}.docx'

    return send_file(
        io.BytesIO(fixed),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=fname,
    )


if __name__ == '__main__':
    print('\n' + '='*55)
    print('  IGNOU RC-71 · Award List Generator')
    print('='*55)
    print('  ✅ Starting...')
    print('  🌐 Open:  http://localhost:5050')
    print('  ⛔ Stop:  Ctrl+C')
    print('='*55 + '\n')
    app.run(host='127.0.0.1', port=5050, debug=False)

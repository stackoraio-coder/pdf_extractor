from pathlib import Path
import csv
import hashlib
import hmac
import io
import os
import tempfile
import traceback
import logging
from typing import List

logging.basicConfig(level=logging.INFO)

from fastapi import FastAPI, UploadFile, File, Form, Request, Response
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, StreamingResponse, RedirectResponse
from starlette.middleware.base import BaseHTTPMiddleware

from extractor import extract_fields, fill_excel_template, to_excel_date_display

# ── Auth config ───────────────────────────────────────────────────────────────
APP_PASSWORD = os.environ.get("APP_PASSWORD", "llanoseguros2026*")
_COOKIE_NAME = "pdf_auth"
_SECRET      = hashlib.sha256(APP_PASSWORD.encode()).hexdigest()   # token fijo derivado de la contraseña

def _valid_cookie(request: Request) -> bool:
    return hmac.compare_digest(request.cookies.get(_COOKIE_NAME, ""), _SECRET)

# Rutas que no requieren auth
_PUBLIC = {"/login", "/favicon.ico"}

class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        if request.url.path in _PUBLIC or _valid_cookie(request):
            return await call_next(request)
        return RedirectResponse(url="/login", status_code=302)

app = FastAPI(title="Extractor PDF escaneado a registro Excel")
app.add_middleware(AuthMiddleware)

# ── Login page ────────────────────────────────────────────────────────────────
LOGIN_HTML = """<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Acceso — Extractor PDF</title>
  <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='6' fill='%230f5bd7'/%3E%3Cpath d='M9 7h9l5 5v13a1 1 0 01-1 1H9a1 1 0 01-1-1V8a1 1 0 011-1z' fill='white'/%3E%3Cpath d='M18 7l5 5h-4a1 1 0 01-1-1V7z' fill='%23bfdbfe'/%3E%3Cpath d='M11 16h4M11 19h6M11 22h5' stroke='%230f5bd7' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E"/>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:Inter,Arial,sans-serif;background:#f4f7fb;display:flex;align-items:center;justify-content:center;min-height:100vh;}
    .card{background:#fff;border:1px solid #d9e2ef;border-radius:20px;padding:40px 36px;width:100%;max-width:380px;box-shadow:0 8px 40px rgba(15,23,42,.10);}
    .logo{width:48px;height:48px;background:#0f5bd7;border-radius:12px;display:flex;align-items:center;justify-content:center;margin-bottom:20px;}
    .logo svg{width:28px;height:28px;}
    h2{font-size:22px;font-weight:800;letter-spacing:-.02em;margin-bottom:6px;}
    p{font-size:14px;color:#64748b;margin-bottom:28px;}
    label{display:block;font-size:13px;font-weight:700;margin-bottom:6px;color:#374151;}
    input[type=password]{width:100%;padding:11px 14px;border:1.5px solid #d9e2ef;border-radius:10px;font-size:15px;outline:none;transition:.2s;}
    input[type=password]:focus{border-color:#0f5bd7;box-shadow:0 0 0 3px #dbeafe;}
    button{width:100%;margin-top:18px;padding:12px;background:#0f5bd7;color:#fff;border:0;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;transition:.15s;}
    button:hover{background:#0a42a8;}
    .error{margin-top:14px;padding:10px 14px;background:#fee2e2;color:#991b1b;border-radius:8px;font-size:13px;font-weight:600;display:none;}
    .error.show{display:block;}
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">
      <svg viewBox="0 0 28 28" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M6 4h10l6 6v14a1 1 0 01-1 1H6a1 1 0 01-1-1V5a1 1 0 011-1z" fill="white"/>
        <path d="M16 4l6 6h-5a1 1 0 01-1-1V4z" fill="#bfdbfe"/>
        <path d="M8 14h4M8 17h8M8 20h6" stroke="#0f5bd7" stroke-width="1.5" stroke-linecap="round"/>
      </svg>
    </div>
    <h2>Extractor PDF → Excel</h2>
    <p>Ingresa la contraseña para continuar.</p>
    <form method="POST" action="/login" id="form">
      <label for="pw">Contraseña</label>
      <input type="password" id="pw" name="password" placeholder="••••••••••••" autofocus autocomplete="current-password"/>
      <button type="submit">Entrar</button>
      <div class="error {error_class}" id="err">Contraseña incorrecta. Inténtalo de nuevo.</div>
    </form>
  </div>
</body>
</html>"""

@app.get("/login", response_class=HTMLResponse)
def login_page(error: str = ""):
    html = LOGIN_HTML.replace("{error_class}", "show" if error else "")
    return HTMLResponse(html)

@app.post("/login")
async def login_post(response: Response, password: str = Form(...)):
    if hmac.compare_digest(password, APP_PASSWORD):
        resp = RedirectResponse(url="/", status_code=302)
        resp.set_cookie(_COOKIE_NAME, _SECRET, httponly=True, samesite="lax", max_age=60*60*24*7)
        return resp
    return RedirectResponse(url="/login?error=1", status_code=302)

@app.get("/logout")
def logout():
    resp = RedirectResponse(url="/login", status_code=302)
    resp.delete_cookie(_COOKIE_NAME)
    return resp

HTML = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Extractor PDF → Excel</title>
  <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='6' fill='%230f5bd7'/%3E%3Cpath d='M9 7h9l5 5v13a1 1 0 01-1 1H9a1 1 0 01-1-1V8a1 1 0 011-1z' fill='white'/%3E%3Cpath d='M18 7l5 5h-4a1 1 0 01-1-1V7z' fill='%23bfdbfe'/%3E%3Cpath d='M11 16h4M11 19h6M11 22h5' stroke='%230f5bd7' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E" />
  <style>
    :root { --blue:#0f5bd7; --blue-dark:#0a42a8; --bg:#f4f7fb; --border:#d9e2ef; --text:#111827; }
    * { box-sizing: border-box; }
    body { font-family: Inter, Arial, sans-serif; margin:0; background:var(--bg); color:var(--text); }
    .wrap { max-width: 1200px; margin: 32px auto; padding: 0 18px; }
    h1 { margin:0; font-size:28px; letter-spacing:-.03em; }
    .muted { color:#64748b; margin:6px 0 0; font-size:14px; }
    .card { background:#fff; border:1px solid var(--border); border-radius:16px; padding:20px; box-shadow:0 4px 24px rgba(15,23,42,.06); }
    label { display:block; font-weight:700; margin-bottom:8px; font-size:14px; }
    input, button { font:inherit; }
    input[type=file] { width:100%; padding:10px 12px; border:2px dashed var(--border); border-radius:12px; background:#fafbff; cursor:pointer; font-size:14px; }
    input[type=file]:hover { border-color:var(--blue); }
    button { border:0; border-radius:10px; padding:10px 18px; font-weight:700; cursor:pointer; font-size:14px; transition: opacity .15s, transform .1s; }
    button:active { transform:scale(.97); }
    button:disabled { opacity:.45; cursor:not-allowed; transform:none; }
    .primary { background:var(--blue); color:#fff; }
    .primary:hover:not(:disabled) { background:var(--blue-dark); }
    .ghost { background:#eaf1ff; color:#0b45a4; }
    .ghost:hover:not(:disabled) { background:#d4e4ff; }
    .danger { background:#fee2e2; color:#991b1b; }
    .danger:hover:not(:disabled) { background:#fecaca; }
    .actions { display:flex; gap:10px; flex-wrap:wrap; margin-top:16px; }
    .help { font-size:13px; color:#64748b; margin-top:8px; }

    /* Progress banner */
    .progress-banner { display:none; margin-top:16px; background:#eff6ff; border:1px solid #bfdbfe; border-radius:12px; padding:14px 16px; }
    .progress-banner.active { display:block; }
    .progress-top { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }
    .progress-label { font-size:13px; font-weight:700; color:#1d4ed8; }
    .progress-count { font-size:13px; color:#3b82f6; font-weight:600; }
    .progress-file { font-size:12px; color:#64748b; margin-top:6px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .bar-track { height:8px; background:#dbeafe; border-radius:999px; overflow:hidden; }
    .bar-fill { height:100%; width:0%; background:var(--blue); border-radius:999px; transition:.3s ease; }

    /* KPIs */
    .kpis { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-top:16px; }
    .kpi { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:12px 14px; }
    .kpi-label { font-size:12px; color:#64748b; margin-bottom:4px; }
    .kpi-value { font-size:22px; font-weight:800; }
    .kpi-value.green { color:#16a34a; }
    .kpi-value.amber { color:#d97706; }

    /* Table */
    .table-wrap { margin-top:20px; overflow:auto; border:1px solid var(--border); border-radius:16px; background:#fff; }
    .table-header { display:flex; justify-content:space-between; align-items:center; padding:14px 16px 0; }
    .table-title { font-weight:700; font-size:15px; }
    table { width:100%; border-collapse:collapse; min-width:1050px; }
    thead th { position:sticky; top:0; background:#0f172a; color:#fff; text-align:left; font-size:12px; letter-spacing:.04em; text-transform:uppercase; padding:11px 12px; }
    tbody td { padding:10px 12px; border-bottom:1px solid #f1f5f9; white-space:nowrap; font-size:13px; }
    td[contenteditable="true"] { background:#fffdf5; outline:none; }
    td[contenteditable="true"]:focus { background:#fff8e1; box-shadow:inset 0 0 0 2px #fbbf24; border-radius:4px; }
    tbody tr:hover td { background:#f8fafc; }
    tbody tr:last-child td { border-bottom:none; }

    /* Loading row */
    .loading-row td { background:#f0f7ff !important; }
    .skeleton { display:inline-block; height:12px; border-radius:6px; background: linear-gradient(90deg,#e2e8f0 25%,#f1f5f9 50%,#e2e8f0 75%); background-size:200% 100%; animation:shimmer 1.4s infinite; }
    @keyframes shimmer { 0%{background-position:200% 0} 100%{background-position:-200% 0} }
    .spinner { display:inline-block; width:14px; height:14px; border:2px solid #bfdbfe; border-top-color:var(--blue); border-radius:50%; animation:spin .7s linear infinite; vertical-align:middle; margin-right:6px; }
    @keyframes spin { to{transform:rotate(360deg)} }

    .pill { display:inline-block; padding:3px 8px; border-radius:999px; font-size:11px; font-weight:700; }
    .ok   { background:#dcfce7; color:#166534; }
    .warn { background:#fef3c7; color:#92400e; }
    .err  { background:#fee2e2; color:#991b1b; }

    .empty-state { text-align:center; padding:48px 0; color:#94a3b8; font-size:14px; }
    .empty-icon { font-size:36px; margin-bottom:8px; }

    /* Done toast */
    .toast { display:none; position:fixed; bottom:24px; right:24px; background:#0f172a; color:#fff; padding:12px 20px; border-radius:12px; font-size:14px; font-weight:600; box-shadow:0 8px 32px rgba(0,0,0,.25); z-index:99; }
    .toast.show { display:block; animation:slideUp .3s ease; }
    @keyframes slideUp { from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)} }

    @media (max-width:800px) { .kpis{grid-template-columns:1fr 1fr;} }
  </style>
</head>
<body>
<div class="wrap">

  <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px">
    <div>
      <h1>Extractor PDF → Excel</h1>
      <p class="muted">Sube uno o varios PDFs escaneados. Cada uno genera una fila editable lista para copiar en Excel.</p>
    </div>
    <a href="/logout" style="font-size:13px;color:#64748b;text-decoration:none;padding:8px 14px;border:1px solid #d9e2ef;border-radius:8px;background:#fff;white-space:nowrap;margin-top:4px">Cerrar sesión</a>
  </div>

  <div class="card">
    <label>Seleccionar PDFs</label>
    <input id="pdfs" type="file" accept="application/pdf" multiple />
    <div class="help">Puedes subir 100+ PDFs a la vez. Las fechas se muestran en formato <b>MM-DD-YY</b>.</div>

    <div class="actions">
      <button class="primary" id="btn_extract">⚡ Procesar PDFs</button>
      <button class="ghost" id="btn_copy" disabled>Copiar todo para Excel</button>
      <button class="ghost" id="btn_csv" disabled>Descargar CSV</button>
      <button class="danger" id="btn_clear" disabled>Limpiar tabla</button>
    </div>

    <div class="progress-banner" id="progress_banner">
      <div class="progress-top">
        <span class="progress-label"><span class="spinner"></span>Procesando archivos...</span>
        <span class="progress-count" id="prog_count">0 / 0</span>
      </div>
      <div class="bar-track"><div class="bar-fill" id="bar"></div></div>
      <div class="progress-file" id="prog_file"></div>
    </div>

    <div class="kpis">
      <div class="kpi">
        <div class="kpi-label">Total procesados</div>
        <div class="kpi-value" id="kpi_total">0</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">Completos</div>
        <div class="kpi-value green" id="kpi_ok">0</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">Por revisar</div>
        <div class="kpi-value amber" id="kpi_warn">0</div>
      </div>
    </div>
  </div>

  <div class="table-wrap">
    <div class="table-header">
      <span class="table-title">Resultados</span>
    </div>
    <table id="records">
      <thead>
        <tr>
          <th>Archivo</th>
          <th>Nombre</th>
          <th>Cédula</th>
          <th>Ingreso</th>
          <th>Valor crédito</th>
          <th>Seg. mensual</th>
          <th>Extraprima</th>
          <th>Fecha nacimiento</th>
          <th>Estado</th>
          <th>Notas</th>
        </tr>
      </thead>
      <tbody>
        <tr id="empty_row">
          <td colspan="10">
            <div class="empty-state">
              <div class="empty-icon">📄</div>
              Selecciona PDFs y haz clic en <b>Procesar PDFs</b> para comenzar.
            </div>
          </td>
        </tr>
      </tbody>
    </table>
  </div>

</div>
<div class="toast" id="toast"></div>

<script>
const rows = [];
const colsToCopy = ["nombre","cedula","ingreso_excel","valor_credito","seguro_vida_mensual","extraprima","fecha_nacimiento_excel"];

function money(v){ return v == null ? "" : Number(v).toLocaleString("en-US",{maximumFractionDigits:0}); }

function statusFor(r){
  if(r.error) return {label:"Error", cls:"err", notes: r.error};
  const missing = [];
  if(!r.nombre) missing.push("nombre");
  if(!r.cedula) missing.push("cedula");
  if(!r.ingreso_excel) missing.push("ingreso");
  if(!r.valor_credito) missing.push("valor");
  if(!r.fecha_nacimiento_excel) missing.push("nacimiento");
  return missing.length
    ? {label:"Revisar", cls:"warn", notes:"Falta: "+missing.join(", ")}
    : {label:"OK", cls:"ok", notes: r.raw_confidence_notes || "OK"};
}

function renderRows(){
  const tbody = document.querySelector("#records tbody");
  // remove data rows but keep empty_row and loading_row
  [...tbody.querySelectorAll("tr.data-row")].forEach(r=>r.remove());
  let ok=0, warn=0;
  rows.forEach((r,idx)=>{
    const st = statusFor(r);
    if(st.cls==="ok") ok++; else warn++;
    const tr = document.createElement("tr");
    tr.className = "data-row";
    tr.innerHTML = `
      <td title="${r.file_name||""}">${(r.file_name||"").length>22?(r.file_name||"").slice(0,20)+"…":r.file_name||""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="nombre">${r.nombre||""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="cedula">${r.cedula||""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="ingreso_excel">${r.ingreso_excel||""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="valor_credito">${money(r.valor_credito)}</td>
      <td contenteditable="true" data-i="${idx}" data-k="seguro_vida_mensual">${money(r.seguro_vida_mensual)}</td>
      <td contenteditable="true" data-i="${idx}" data-k="extraprima">${r.extraprima||""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="fecha_nacimiento_excel">${r.fecha_nacimiento_excel||""}</td>
      <td><span class="pill ${st.cls}">${st.label}</span></td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis" title="${st.notes}">${st.notes}</td>`;
    // insert before loading row (if present)
    const loader = document.getElementById("loading_row");
    if(loader) tbody.insertBefore(tr, loader);
    else tbody.appendChild(tr);
  });
  document.getElementById("empty_row").style.display = rows.length ? "none" : "";
  document.getElementById("kpi_total").textContent = rows.length;
  document.getElementById("kpi_ok").textContent = ok;
  document.getElementById("kpi_warn").textContent = warn;
  const hasRows = rows.length > 0;
  document.getElementById("btn_copy").disabled = !hasRows;
  document.getElementById("btn_csv").disabled  = !hasRows;
  document.getElementById("btn_clear").disabled = !hasRows;
}

function addLoadingRow(filename){
  removeLoadingRow();
  const tbody = document.querySelector("#records tbody");
  const tr = document.createElement("tr");
  tr.id = "loading_row";
  tr.className = "loading-row";
  tr.innerHTML = `
    <td><span class="spinner"></span>${filename.length>18?filename.slice(0,16)+"…":filename}</td>
    <td><span class="skeleton" style="width:120px"></span></td>
    <td><span class="skeleton" style="width:80px"></span></td>
    <td><span class="skeleton" style="width:70px"></span></td>
    <td><span class="skeleton" style="width:90px"></span></td>
    <td><span class="skeleton" style="width:60px"></span></td>
    <td><span class="skeleton" style="width:50px"></span></td>
    <td><span class="skeleton" style="width:70px"></span></td>
    <td><span class="skeleton" style="width:50px"></span></td>
    <td></td>`;
  tbody.appendChild(tr);
  tr.scrollIntoView({behavior:"smooth", block:"nearest"});
}

function removeLoadingRow(){
  const el = document.getElementById("loading_row");
  if(el) el.remove();
}

function showToast(msg, duration=3000){
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"), duration);
}

function setProcessing(active, total=0, done=0){
  const banner = document.getElementById("progress_banner");
  banner.classList.toggle("active", active);
  document.getElementById("btn_extract").disabled = active;
  document.getElementById("btn_clear").disabled  = active;
  if(active){
    document.getElementById("prog_count").textContent = done+" / "+total;
  }
}

document.addEventListener("input", e=>{
  const td = e.target.closest("td[contenteditable=true]");
  if(!td) return;
  rows[Number(td.dataset.i)][td.dataset.k] = td.textContent.trim();
});

async function processFiles(){
  const files = [...document.getElementById("pdfs").files];
  if(!files.length){ alert("Selecciona al menos un PDF"); return; }
  setProcessing(true, files.length, 0);
  document.getElementById("empty_row").style.display = "none";
  for(let i=0;i<files.length;i++){
    document.getElementById("prog_count").textContent = (i+1)+" / "+files.length;
    document.getElementById("prog_file").textContent = "Procesando: "+files[i].name;
    document.getElementById("bar").style.width = Math.round((i/files.length)*100)+"%";
    addLoadingRow(files[i].name);
    const form = new FormData();
    form.append("pdf", files[i]);
    const res = await fetch("/extract-json",{method:"POST",body:form});
    const data = await res.json();
    if(data.error){ data.nombre=""; data.raw_confidence_notes=data.error; }
    data.file_name = files[i].name;
    rows.push(data);
    document.getElementById("bar").style.width = Math.round(((i+1)/files.length)*100)+"%";
    renderRows();
  }
  removeLoadingRow();
  setProcessing(false);
  showToast("✅ "+files.length+" PDF"+(files.length>1?"s":"")+" procesado"+(files.length>1?"s":""));
}

function tsv(){
  return rows.map(r=>colsToCopy.map(k=>(r[k]??"").toString().replace(/\\t|\\n/g," ")).join("\\t")).join("\\n");
}

async function copyTSV(){
  if(!rows.length) return;
  await navigator.clipboard.writeText(tsv());
  showToast("📋 Filas copiadas — pega directamente en Excel");
}

function downloadCSV(){
  if(!rows.length) return;
  const header=["NOMBRE","CEDULA","INGRESO","VALOR CREDITO","SEG. MENSUAL","EXTRAPRIMA","FECHA NACIMIENTO"];
  const body=rows.map(r=>colsToCopy.map(k=>`"${(r[k]??"").toString().replaceAll('"','""')}"`).join(","));
  const blob=new Blob([[header.join(","),...body].join("\\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="registros_extraidos.csv"; a.click();
}

function clearRows(){
  if(!rows.length) return;
  if(!confirm("¿Seguro que quieres limpiar todos los registros?")) return;
  rows.length=0;
  renderRows();
  document.getElementById("bar").style.width="0%";
  document.getElementById("empty_row").style.display="";
}

document.getElementById("btn_extract").addEventListener("click", processFiles);
document.getElementById("btn_copy").addEventListener("click", copyTSV);
document.getElementById("btn_csv").addEventListener("click", downloadCSV);
document.getElementById("btn_clear").addEventListener("click", clearRows);
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def home():
    return HTML

@app.get("/preview", response_class=HTMLResponse)
def preview():
    sample = HTML.replace("const rows = [];", "const rows = [{file_name:'ejemplo.pdf',nombre:'LONDOÑO LOZADA JOHN EDWARD',cedula:'9.728.672',ingreso_excel:'06-05-26',valor_credito:5034402,seguro_vida_mensual:2937,extraprima:'',fecha_nacimiento_excel:'08-20-81',raw_confidence_notes:'OK'}]; setTimeout(render, 0);")
    return sample

@app.post("/extract-json")
async def extract_json(pdf: UploadFile = File(...)):
    try:
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = Path(tmp) / pdf.filename
            pdf_path.write_bytes(await pdf.read())
            data = extract_fields(pdf_path)
            data["file_name"] = pdf.filename
        return JSONResponse(data)
    except Exception as e:
        tb = traceback.format_exc()
        logging.error("Error procesando %s:\n%s", pdf.filename, tb)
        return JSONResponse({"error": str(e), "traceback": tb, "file_name": pdf.filename}, status_code=500)

@app.post("/extract-multiple-json")
async def extract_multiple_json(pdfs: List[UploadFile] = File(...)):
    records = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        for pdf in pdfs:
            pdf_path = tmp / pdf.filename
            pdf_path.write_bytes(await pdf.read())
            data = extract_fields(pdf_path)
            data["file_name"] = pdf.filename
            records.append(data)
    return JSONResponse({"records": records})

@app.post("/extract")
async def extract_to_excel(
    pdf: UploadFile = File(...),
    excel: UploadFile = File(...),
    month_sheet: str = Form("JUNIO"),
):
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf_path = tmp / pdf.filename
        xlsx_path = tmp / excel.filename
        out_path = tmp / f"excel_diligenciado_{month_sheet.lower()}.xlsx"
        pdf_path.write_bytes(await pdf.read())
        xlsx_path.write_bytes(await excel.read())
        data = extract_fields(pdf_path)
        fill_excel_template(xlsx_path, out_path, data, month_sheet=month_sheet)
        final = Path(tempfile.gettempdir()) / out_path.name
        final.write_bytes(out_path.read_bytes())
    return FileResponse(final, filename=final.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

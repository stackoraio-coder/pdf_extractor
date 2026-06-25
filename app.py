from pathlib import Path
import csv
import io
import tempfile
import traceback
import logging
from typing import List

logging.basicConfig(level=logging.INFO)

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, StreamingResponse

from extractor import extract_fields, fill_excel_template, to_excel_date_display

app = FastAPI(title="Extractor PDF escaneado a registro Excel")

HTML = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Extractor PDF a registro Excel</title>
  <style>
    :root { --blue:#0f5bd7; --bg:#f4f7fb; --border:#d9e2ef; --text:#111827; }
    * { box-sizing: border-box; }
    body { font-family: Inter, Arial, sans-serif; margin:0; background:var(--bg); color:var(--text); }
    .wrap { max-width: 1180px; margin: 32px auto; padding: 0 18px; }
    .hero { display:flex; justify-content:space-between; gap:20px; align-items:flex-end; margin-bottom:18px; }
    h1 { margin:0; font-size: 30px; letter-spacing:-.03em; }
    .muted { color:#64748b; margin:8px 0 0; }
    .card { background:#fff; border:1px solid var(--border); border-radius:18px; padding:22px; box-shadow:0 14px 40px rgba(15,23,42,.07); }
    .grid { display:grid; grid-template-columns: 1.4fr .6fr; gap:16px; }
    label { display:block; font-weight:800; margin-bottom:8px; }
    input, select, button { font:inherit; }
    input[type=file], select { width:100%; padding:12px; border:1px solid var(--border); border-radius:12px; background:#fff; }
    button { border:0; border-radius:12px; padding:12px 16px; font-weight:800; cursor:pointer; }
    .primary { background:var(--blue); color:#fff; }
    .ghost { background:#eaf1ff; color:#0b45a4; }
    .danger { background:#fee2e2; color:#991b1b; }
    .actions { display:flex; gap:10px; flex-wrap:wrap; margin-top:16px; }
    .table-wrap { margin-top:18px; overflow:auto; border:1px solid var(--border); border-radius:16px; background:#fff; }
    table { width:100%; border-collapse:collapse; min-width:1050px; }
    th { position:sticky; top:0; background:#0f172a; color:#fff; text-align:left; font-size:13px; letter-spacing:.02em; }
    th, td { padding:11px 12px; border-bottom:1px solid #e5e7eb; white-space:nowrap; }
    td[contenteditable="true"] { background:#fffdf5; outline:none; }
    tr:hover td { background:#f8fafc; }
    .pill { display:inline-block; padding:3px 8px; border-radius:999px; font-size:12px; font-weight:800; }
    .ok { background:#dcfce7; color:#166534; }
    .warn { background:#fef3c7; color:#92400e; }
    .progress { height:10px; background:#e2e8f0; border-radius:999px; overflow:hidden; margin-top:14px; display:none; }
    .bar { height:100%; width:0%; background:var(--blue); transition:.25s; }
    .help { font-size:13px; color:#64748b; margin-top:10px; }
    .preview { margin-top:16px; display:grid; grid-template-columns:repeat(4,1fr); gap:10px; }
    .kpi { background:#f8fafc; border:1px solid #e2e8f0; border-radius:14px; padding:12px; }
    .kpi b { font-size:22px; display:block; }
    @media (max-width: 800px) { .preview { grid-template-columns:1fr; } .hero { display:block; } }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <div>
        <h1>Extractor PDF escaneado → filas para Excel</h1>
        <p class="muted">Sube uno o varios PDFs. Cada PDF genera una fila editable para copiar y pegar en la plantilla.</p>
      </div>
    </div>

    <div class="card">
      <div>
        <label>PDFs escaneados</label>
        <input id="pdfs" type="file" accept="application/pdf" multiple />
        <div class="help">Puedes seleccionar varios PDFs a la vez (incluso 100+). Las fechas se muestran como <b>MM-DD-YY</b>.</div>
      </div>
      <div class="actions">
        <button class="primary" id="btn_extract">Extraer PDFs</button>
        <button class="ghost" id="btn_copy">Copiar filas para Excel</button>
        <button class="ghost" id="btn_csv">Descargar CSV</button>
        <button class="danger" id="btn_clear">Limpiar</button>
      </div>
      <div class="progress" id="progress"><div class="bar" id="bar"></div></div>
      <div class="preview">
        <div class="kpi"><span>PDFs procesados</span><b id="kpi_total">0</b></div>
        <div class="kpi"><span>OK</span><b id="kpi_ok">0</b></div>
        <div class="kpi"><span>Por revisar</span><b id="kpi_warn">0</b></div>
        <div class="kpi"><span>Última acción</span><b id="last_action" style="font-size:15px">Listo</b></div>
      </div>
    </div>

    <div class="table-wrap">
      <table id="records">
        <thead>
          <tr>
            <th>Archivo</th>
            <th>NOMBRE</th>
            <th>CEDULA</th>
            <th>INGRESO</th>
            <th>VALOR CREDITO</th>
            <th>MES / SEGURO</th>
            <th>EXTRAPRIMA</th>
            <th>FECHA NACIMIENTO</th>
            <th>Estado</th>
            <th>Notas</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>

    <div class="card" style="margin-top:18px">
      <b>Opcional:</b> si quieres seguir generando el Excel directamente, usa el endpoint <code>POST /extract</code> con un PDF y la plantilla. La tabla anterior está pensada para revisar, corregir y copiar/pegar varias filas.
    </div>
  </div>
<script>
const rows = [];
const colsToCopy = ["nombre", "cedula", "ingreso_excel", "valor_credito", "seguro_vida_mensual", "extraprima", "fecha_nacimiento_excel"];
function money(v){ return v == null ? "" : Number(v).toLocaleString("en-US", {maximumFractionDigits:0}); }
function statusFor(r){
  const missing = [];
  if(!r.nombre) missing.push("nombre");
  if(!r.cedula) missing.push("cedula");
  if(!r.ingreso_excel) missing.push("ingreso");
  if(!r.valor_credito) missing.push("valor");
  if(!r.fecha_nacimiento_excel) missing.push("nacimiento");
  return missing.length ? {label:"Revisar", cls:"warn", notes:"Falta: " + missing.join(", ")} : {label:"OK", cls:"ok", notes:r.raw_confidence_notes || "OK"};
}
function render(){
  const tbody = document.querySelector("#records tbody");
  tbody.innerHTML = "";
  let ok = 0, warn = 0;
  rows.forEach((r, idx) => {
    const st = statusFor(r); if(st.cls === "ok") ok++; else warn++;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.file_name || ""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="nombre">${r.nombre || ""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="cedula">${r.cedula || ""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="ingreso_excel">${r.ingreso_excel || ""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="valor_credito">${money(r.valor_credito)}</td>
      <td contenteditable="true" data-i="${idx}" data-k="seguro_vida_mensual">${money(r.seguro_vida_mensual)}</td>
      <td contenteditable="true" data-i="${idx}" data-k="extraprima">${r.extraprima || ""}</td>
      <td contenteditable="true" data-i="${idx}" data-k="fecha_nacimiento_excel">${r.fecha_nacimiento_excel || ""}</td>
      <td><span class="pill ${st.cls}">${st.label}</span></td>
      <td>${st.notes}</td>`;
    tbody.appendChild(tr);
  });
  document.getElementById("kpi_total").textContent = rows.length;
  document.getElementById("kpi_ok").textContent = ok;
  document.getElementById("kpi_warn").textContent = warn;
}
document.addEventListener("input", e => {
  const td = e.target.closest("td[contenteditable=true]");
  if(!td) return;
  rows[Number(td.dataset.i)][td.dataset.k] = td.textContent.trim();
});
async function processFiles(){
  const files = [...document.getElementById("pdfs").files];
  if(!files.length){ alert("Selecciona uno o varios PDFs"); return; }
  document.getElementById("progress").style.display = "block";
  for(let i=0;i<files.length;i++){
    const form = new FormData(); form.append("pdf", files[i]);
    document.getElementById("last_action").textContent = `Procesando ${files[i].name}`;
    const res = await fetch("/extract-json", {method:"POST", body: form});
    const data = await res.json();
    if(data.error){ data.nombre = "ERROR: " + data.error; data.raw_confidence_notes = data.traceback || data.error; }
    data.file_name = files[i].name;
    rows.push(data);
    document.getElementById("bar").style.width = `${Math.round(((i+1)/files.length)*100)}%`;
    render();
  }
  document.getElementById("last_action").textContent = "Extracción terminada";
}
function tsv(){
  return rows.map(r => colsToCopy.map(k => (r[k] ?? "").toString().replace(/\\t|\\n/g," ")).join("\\t")).join("\\n");
}
async function copyTSV(){
  if(!rows.length){ alert("No hay filas para copiar"); return; }
  await navigator.clipboard.writeText(tsv());
  document.getElementById("last_action").textContent = "Filas copiadas";
}
function downloadCSV(){
  if(!rows.length){ alert("No hay filas para descargar"); return; }
  const header = ["NOMBRE","CEDULA","INGRESO","VALOR CREDITO","MES / SEGURO","EXTRAPRIMA","FECHA NACIMIENTO"];
  const body = rows.map(r => colsToCopy.map(k => `"${(r[k] ?? "").toString().replaceAll('"','""')}"`).join(","));
  const blob = new Blob([[header.join(","), ...body].join("\\n")], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "registros_extraidos.csv"; a.click();
}
function clearRows(){ rows.length = 0; render(); document.getElementById("bar").style.width = "0%"; document.getElementById("last_action").textContent = "Limpio"; }
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

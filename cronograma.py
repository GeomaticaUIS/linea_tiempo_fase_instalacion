#!/usr/bin/env python3
"""
Genera el HTML "Cronograma de instalación de embarcaderos" (Portal de Interventoría)
a partir de la hoja "Reporte semanal" del Excel de control de instalación.

Uso:
    python cronograma.py [ruta_excel] [ruta_html_salida]

Por defecto lee Tablero_control_General_Instalación_180926.xlsx y escribe
instalacion.html en la misma carpeta del script.
"""
import sys, json, datetime as dt
from pathlib import Path
import openpyxl

HOJA = "Reporte semanal"
FILA_ENCABEZADO = 9
TOLERANCIA_DIAS = 30  # una fecha de obra/transporte/armado a más de 30 días de la instalación se marca como inconsistente

# columna del Excel -> clave interna
COLS = {
    "num": "No.EMB Instalado", "cod": "CÓDIGO DIPOLA MUN", "emb": "# EMB Fabricación",
    "mun": "MUNICIPIO", "dep": "DEPARTAMENTO", "sup": "SUPERVISOR UIS",
    "soc": "FECHA SOCIALIZACIÓN INICIO",
    "obra_a": "FECHA INICIO OBRA CIVIL", "obra_b": "FECHA FINALIZACIÓN OBRA CIVIL",
    "tra_a": "FECHA INICIO TRANSPORTE EMB", "tra_b": "FECHA FIN TRANSPORTE EMB",
    "arm_a": "FECHA INICIO ARMADO EMB", "arm_b": "FECHA FINALIZACIÓN ARMADO EMB",
    "ins": "FECHA INSTALACIÓN EMB", "cap": "FECHA CAPACITACIÓN Y ENTREGA EMB",
    "dias": "DIAS DE EJECUCIÓN", "habiles": "DIAS HABILES DE EJECUCIÓN",
}
COLS_OPCIONALES = {"habiles"}
MINUSCULAS = {"de", "del", "la", "las", "los", "el", "y"}
MUNICIPIOS_EXCLUIDOS = {"SAN FERNANDO"}


def nombre(txt):
    txt = " ".join(str(txt).replace("_", " - ").split())
    partes = []
    for i, w in enumerate(txt.lower().split(" ")):
        partes.append(w if (w in MINUSCULAS and i) else w[:1].upper() + w[1:])
    return " ".join(partes).replace(" - ", " – ")


def iso(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    return None


def leer(xlsx):
    ws = openpyxl.load_workbook(xlsx, data_only=True)[HOJA]
    head = {}
    for c in ws[FILA_ENCABEZADO]:
        if c.value:
            head[str(c.value).strip().upper()] = c.column
    def col(nombre_col, opcional=False):
        for k, v in head.items():
            if k.startswith(nombre_col.upper()[:28]):
                return v
        if opcional:
            return None
        raise KeyError(f"No encontré la columna '{nombre_col}' en la fila {FILA_ENCABEZADO}")
    idx = {k: col(v, opcional=k in COLS_OPCIONALES) for k, v in COLS.items()}

    sitios, avisos = [], []
    for fila in range(FILA_ENCABEZADO + 1, ws.max_row + 1):
        g = lambda k: ws.cell(fila, idx[k]).value if idx[k] is not None else None
        if not g("mun"):
            continue
        f = {k: iso(g(k)) for k in ("soc", "obra_a", "obra_b", "tra_a", "tra_b", "arm_a", "arm_b", "ins", "cap")}
        ref = f["ins"] or f["cap"]
        refd = dt.date.fromisoformat(ref) if ref else None

        def ok(clave):
            v = f[clave]
            if not v:
                return False
            if clave in ("obra_a", "obra_b", "tra_a", "tra_b", "arm_a", "arm_b") and refd:
                return abs((dt.date.fromisoformat(v) - refd).days) <= TOLERANCIA_DIAS
            return True

        nom = nombre(g("mun"))
        if nom.upper() in MUNICIPIOS_EXCLUIDOS:
            continue
        etapas = []
        for k, etiqueta, a, b in (
            ("soc", "Socialización", "soc", None), ("obra", "Obra civil", "obra_a", "obra_b"),
            ("tra", "Transporte EMB", "tra_a", "tra_b"), ("arm", "Armado EMB", "arm_a", "arm_b"),
            ("ins", "Instalación", "ins", None), ("cap", "Capacitación y entrega", "cap", None)):
            va, vb = f[a], f[b] if b else f[a]
            aok, bok = ok(a), ok(b) if b else ok(a)
            for clave, v, o in ((a, va, aok), (b, vb, bok)):
                if clave and v and not o:
                    avisos.append(f"{nom}: {etiqueta} ({clave}) = {v} no es coherente con la fecha de instalación; se excluye del gráfico (fila {fila})")
            if va or vb:
                etapas.append({"k": k, "n": etiqueta, "a": va, "b": vb, "aok": aok, "bok": bok})
        validas = sorted(d for e in etapas for d, o in ((e["a"], e["aok"]), (e["b"], e["bok"])) if d and o)
        cod = g("cod")
        sitios.append({
            "id": fila, "num": g("num"), "cod": str(cod) if cod else "", "emb": g("emb"),
            "mun": nom, "dep": nombre(g("dep") or ""), "sup": (g("sup") or "").strip() if isinstance(g("sup"), str) else "",
            "inst": bool(f["cap"]), "etapas": etapas,
            "desde": validas[0] if validas else None, "hasta": validas[-1] if validas else None,
            "dias": g("dias"), "habiles": g("habiles"),
        })
    return sitios, avisos


def main():
    aqui = Path(__file__).resolve().parent
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else aqui / "Tablero_control_General_Instalación_180926.xlsx"
    salida = Path(sys.argv[2]) if len(sys.argv) > 2 else aqui / "instalacion.html"
    sitios, avisos = leer(xlsx)
    datos = {"sitios": sitios, "avisos": avisos, "fuente": xlsx.name,
             "generado": dt.date.today().isoformat()}
    html = PLANTILLA.replace("/*__DATA__*/null", json.dumps(datos, ensure_ascii=False))
    salida.write_text(html, encoding="utf-8")
    inst = sum(s["inst"] for s in sitios)
    print(f"{len(sitios)} sitios · {inst} instalados · {len(sitios) - inst} pendientes · {len(avisos)} avisos de fechas")
    for a in avisos:
        print("  ⚠", a)
    print("HTML:", salida)


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Cronograma de instalación de embarcaderos · Interventoría FONTUR</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#F2F4F3; --bg-panel:#FFFFFF; --ink:#16232B; --muted:#4B5B63; --faint:#8A9AA1;
  --line:#D9E0DF; --soft:#F7F8F7; --accent:#0E7490; --accent-ink:#063542; --accent-soft:#E1EFF1;
  --shadow:0 1px 2px rgba(15,35,40,.06); --radius:10px;
  --ok:#0f9d8f; --pend:#d97706;
  --soc:#64748b; --obra:#d97706; --tra:#7c3aed; --arm:#0891b2; --ins:#0f9d8f; --cap:#db2777;
}
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]){
    --bg:#10171A; --bg-panel:#172024; --ink:#E8EDEC; --muted:#A9B7B9; --faint:#6C7A7C;
    --line:#253134; --soft:#131C1F; --accent:#4FB3C9; --accent-ink:#DFF4F8; --accent-soft:#1B3439;
    --shadow:0 1px 2px rgba(0,0,0,.4);
  }
}
:root[data-theme="dark"]{
  --bg:#10171A; --bg-panel:#172024; --ink:#E8EDEC; --muted:#A9B7B9; --faint:#6C7A7C;
  --line:#253134; --soft:#131C1F; --accent:#4FB3C9; --accent-ink:#DFF4F8; --accent-soft:#1B3439;
  --shadow:0 1px 2px rgba(0,0,0,.4);
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:"IBM Plex Sans",-apple-system,"Segoe UI",Roboto,sans-serif;font-size:15px;line-height:1.45}
.wrap{max-width:1180px;margin:0 auto;padding:28px 20px 60px}
h1.page{font-size:clamp(22px,3.2vw,30px);font-weight:700;margin:0 0 16px;letter-spacing:-.01em}
.page-nav{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 18px}
.page-nav a{display:inline-flex;align-items:center;min-height:40px;padding:8px 14px;border:1px solid var(--line);border-radius:7px;background:var(--bg-panel);color:var(--muted);font:inherit;font-weight:500;text-decoration:none;transition:background .15s ease,border-color .15s ease,color .15s ease}
.page-nav a:hover,.page-nav a[aria-current="page"]{border-color:var(--accent);background:var(--accent-soft);color:var(--accent-ink)}
.page-nav a[aria-current="page"]{font-weight:600}
.page-nav a:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.portal{background:var(--bg-panel);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
.band{background:var(--accent-soft);color:var(--accent-ink);padding:18px 20px;display:flex;gap:14px;align-items:center}
.band svg{flex:none}
.band b{display:block;font-size:22px;letter-spacing:.2px}
.band span{font-size:14px;opacity:.92}
.body{padding:20px}
.docs h2{font-size:20px;margin:0 0 4px;font-weight:600}
.docs p{margin:0;color:var(--muted);font-size:15px}
.tabs{display:flex;gap:10px;margin:20px 0 18px;flex-wrap:wrap}
.pill{border:1px solid var(--line);background:var(--bg-panel);color:var(--ink);border-radius:999px;padding:10px 18px;font:inherit;font-size:15px;cursor:default;display:inline-flex;gap:8px;align-items:center}
.pill.on{background:var(--accent-soft);color:var(--accent-ink);border-color:var(--accent)}
button.pill{cursor:pointer}
h2.sec{font-size:22px;margin:0 0 4px;font-weight:600}
.sub{color:var(--muted);margin:0 0 18px;font-size:15px}
.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px}
.kpi b{display:block;font-size:28px;font-weight:600;line-height:1.1}
.kpi small{color:var(--muted);font-size:12px}
.seg{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px}
.seg .pill{justify-content:center;padding:13px 18px;font-size:16px;width:100%}
.filters{display:grid;grid-template-columns:1fr 1fr;gap:12px 12px;margin-bottom:22px}
label.f{display:block;color:var(--muted);font-size:15px;font-weight:500}
label.f select,label.f input{display:block;width:100%;margin-top:6px;border:1px solid var(--line);border-radius:999px;padding:11px 16px;font:inherit;font-size:16px;color:var(--ink);background:var(--bg-panel);appearance:none}
label.f select{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%23444' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 14px center;padding-right:38px}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap}
.toggle{display:flex;gap:4px;background:var(--bg-panel);border-radius:999px}
.toggle button{border:0;background:transparent;font:inherit;font-size:15px;color:var(--muted);padding:9px 18px;border-radius:999px;cursor:pointer}
.toggle button.on{background:var(--accent-soft);color:var(--accent-ink);font-weight:600}
.meta{display:flex;justify-content:space-between;align-items:center;margin:14px 0 14px;color:var(--muted);font-size:14px;gap:10px;flex-wrap:wrap}
.chips-help{color:var(--muted);font-size:13px;margin:0 0 8px}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px}
.chip{border:1px solid var(--line);background:#fff;border-radius:999px;padding:7px 14px;font:inherit;font-size:14px;cursor:pointer;display:inline-flex;gap:8px;align-items:center}
.chip i{width:10px;height:10px;border-radius:50%;display:inline-block}
.chip.on{background:var(--ink);color:#fff;border-color:var(--ink)}
.scroll{overflow-x:auto}
.gantt{position:relative;min-width:680px}
.axis-year{text-align:center;font-weight:600;font-size:14px;margin:0 28px 6px 190px}
.axis,.row{display:grid;grid-template-columns:190px 1fr 28px;align-items:center}
.axis{font-size:13px;color:#374151;height:30px}
.axis .l{font-size:12px;color:var(--muted)}
.axis .t{position:relative;height:100%}
.axis .t span{position:absolute;top:50%;transform:translate(-50%,-50%);white-space:nowrap}
.gl{position:absolute;left:190px;right:28px;top:30px;bottom:0;pointer-events:none}
.gl i{position:absolute;top:0;bottom:0;width:1px;background:var(--line)}
.row{position:relative;min-height:44px;cursor:pointer;border-radius:8px}
.row:hover{background:var(--accent-soft)}
.row.sel{background:var(--accent-soft)}
.row .n{font-size:14px;padding:6px 10px 6px 0}
.row.sel .n,.row:hover .n{font-weight:600}
.row .t{position:relative;height:44px}
.bar{position:absolute;top:50%;height:16px;transform:translateY(-50%);border-radius:8px;background:var(--ok);min-width:12px}
.bar.pend{background:var(--pend)}
.row .c{justify-self:center;color:#4b5563;transition:transform .15s}
.row.sel .c{transform:rotate(180deg)}
.legend{display:flex;gap:18px;margin:10px 0 18px;color:#374151;font-size:15px;flex-wrap:wrap}
.legend i{display:inline-block;width:12px;height:12px;border-radius:50%;margin-right:6px;vertical-align:-1px}
.detail{border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-top:6px;background:var(--bg-panel)}
.detail header{display:flex;justify-content:space-between;align-items:flex-start}
.detail h3{margin:0;font-size:20px;font-weight:600}
.detail header p{margin:2px 0 0;color:var(--muted);font-size:15px}
.x{border:0;background:none;font-size:22px;line-height:1;cursor:pointer;color:#374151;padding:4px 8px}
.stage{display:grid;grid-template-columns:150px 1fr;gap:12px;align-items:center;margin:14px 0}
.stage .s b{display:block;font-size:12px;font-weight:600}
.stage .s small{font-size:11px;color:#374151}
.stage .s small.bad{color:#b91c1c}
.track{position:relative;height:18px;background:var(--soft);border-radius:9px}
.track i{position:absolute;top:0;bottom:0;border-radius:9px;min-width:8px}
.note{color:#374151;font-size:15px;margin:10px 0 0}
.facts{display:flex;gap:22px;flex-wrap:wrap;margin:14px 0 0;color:#374151;font-size:14px}
.facts b{font-weight:600}
table{border-collapse:collapse;width:100%;min-width:900px;font-size:14px}
th,td{text-align:left;padding:10px 10px;border-bottom:1px solid var(--line);white-space:nowrap}
th{color:var(--muted);font-weight:600;font-size:13px;background:#fff;position:sticky;top:0}
td.m{font-weight:600}
.st{display:inline-flex;gap:6px;align-items:center}
.st i{width:9px;height:9px;border-radius:50%}
.bad{color:#b91c1c}
.bad s{opacity:.8}
.empty{padding:34px 10px;text-align:center;color:var(--muted)}
.foot{color:var(--muted);font-size:14px;margin:22px 0 0}
details.av{margin-top:10px;font-size:13px;color:#7f1d1d;background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:8px 12px}
details.av summary{cursor:pointer;font-weight:600}
details.av ul{margin:8px 0 0;padding-left:18px}
@media (max-width:640px){
  .wrap{padding:18px 14px 40px}
  .body{padding:14px}
  .kpis{grid-template-columns:repeat(3,1fr)}
  .axis,.row{grid-template-columns:130px 1fr 24px}
  .axis-year{margin-left:130px;margin-right:24px}
  .gl{left:130px;right:24px}
  .stage{grid-template-columns:110px 1fr}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style>
</head>
<body>
<div class="wrap">
  <h1 class="page">Cronograma de instalación</h1>
  <nav class="page-nav" aria-label="Secciones del proyecto">
    <a href="index.html">Soportes documentales</a>
    <a href="instalacion.html" aria-current="page">Cronograma de instalación</a>
  </nav>
  <div class="portal">
    <div class="band">
      <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 21V8l8-4v17"/><path d="M12 9l8 3v9"/><path d="M2 21h20"/><path d="M8 11v.01M8 15v.01"/></svg>
      <div><b>INTERVENTORÍA</b><span>Proyecto FONTUR · Fase de instalación</span></div>
    </div>
    <div class="body">
      
      <h2 class="sec">Cronograma de instalación de embarcaderos</h2>
      <p class="sub" id="subt"></p>

      <div class="kpis" id="kpis"></div>

      <div class="seg">
        <button class="pill on" id="vGeneral" type="button">Vista general</button>
        <button class="pill" id="vAct" type="button">Por actividad</button>
      </div>

      <div class="filters">
        <label class="f">Año<select id="fAnio"></select></label>
        <label class="f">Mes<select id="fMes"></select></label>
        <label class="f">Departamento<select id="fDep"></select></label>
        <label class="f">Buscar sitio<input id="fQ" type="search" placeholder="Nombre del municipio" autocomplete="off"></label>
      </div>

      <div class="head">
        <div>
          <h2 class="sec" style="font-size:20px" id="tTitulo">Cronograma general</h2>
          <p class="sub" id="tSub" style="margin:0"></p>
        </div>
        <div class="toggle" role="group" aria-label="Tipo de vista">
          <button type="button" id="mGraf" class="on">Gráfico</button>
          <button type="button" id="mTab">Tabla de fechas</button>
        </div>
      </div>

      <div class="meta"><span>Distribución cronológica por municipio</span></div>
      <p class="chips-help" id="chips-help" hidden>Selecciona una fase para filtrar los sitios; el punto de color identifica la fase.</p>
      <div class="chips" id="chips" role="group" aria-label="Filtrar cronograma por actividad" hidden></div>
      <div class="legend activity-legend" id="activity-legend" aria-label="Leyenda de actividades"></div>
      <div id="chart"></div>
      <div class="legend" id="legend"></div>
      <div id="detail"></div>
      <div id="avisos"></div>
      <p class="foot" id="foot"></p>
    </div>
  </div>
</div>

<script>
const DATA = /*__DATA__*/null;
const DAY = 86400000;
const MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const MESES_L = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const ACT = [
  {k:"soc",n:"Socialización",c:"var(--soc)"},{k:"obra",n:"Obra civil",c:"var(--obra)"},
  {k:"tra",n:"Transporte EMB",c:"var(--tra)"},{k:"arm",n:"Armado EMB",c:"var(--arm)"},
  {k:"ins",n:"Instalación",c:"var(--ins)"},{k:"cap",n:"Capacitación y entrega",c:"var(--cap)"}
];
const T = s => s ? Date.parse(s + "T00:00:00Z") : null;
const dm = t => { const d = new Date(t); return String(d.getUTCDate()).padStart(2,"0") + "/" + String(d.getUTCMonth()+1).padStart(2,"0") + "/" + d.getUTCFullYear(); };
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const norm = s => s.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase();
const $ = id => document.getElementById(id);

const S = DATA.sitios.filter(s => s.desde).sort((a,b) => a.desde.localeCompare(b.desde) || a.mun.localeCompare(b.mun));
const years = [...new Set(S.flatMap(s => [s.desde.slice(0,4), s.hasta.slice(0,4)]))].sort();
const deps = [...new Set(S.map(s => s.dep).filter(Boolean))].sort((a,b) => a.localeCompare(b,"es"));

const st = {vista:"general", modo:"graf", anio:years.includes("2026") ? "2026" : years[years.length-1], mes:0, dep:"", q:"", act:"obra", sel:null};

// ---- cabecera y filtros
const total = DATA.sitios.length, inst = DATA.sitios.filter(s => s.inst).length;
$("kpis").innerHTML = [[total,"Sitios registrados"],[inst,"Instalados"],[total-inst,"Pendiente" + (total-inst===1?"":"s")]]
  .map(([n,l]) => `<div class="kpi"><b>${n}</b><small>${l}</small></div>`).join("");
$("subt").textContent = `Consulta histórica de actividades · ${years[0]}–${years[years.length-1]}`;
$("fAnio").innerHTML = years.map(y => `<option>${y}</option>`).join("");
$("fAnio").value = st.anio;
$("fMes").innerHTML = `<option value="0">Todos los meses</option>` + MESES_L.map((m,i) => `<option value="${i+1}">${m}</option>`).join("");
$("fDep").innerHTML = `<option value="">Todos</option>` + deps.map(d => `<option>${esc(d)}</option>`).join("");
$("chips").innerHTML = ACT.map(a => `<button type="button" class="chip" data-k="${a.k}" aria-pressed="false" title="Filtrar por ${a.n}"><i style="background:${a.c}"></i>${a.n}</button>`).join("");


function periodo() {
  const y = +st.anio;
  return st.mes ? [Date.UTC(y, st.mes-1, 1), Date.UTC(y, st.mes, 1)] : [Date.UTC(y,0,1), Date.UTC(y+1,0,1)];
}
function visibles() {
  const [p0, p1] = periodo();
  return S.filter(s => {
    if (st.dep && s.dep !== st.dep) return false;
    if (st.q && !norm(s.mun).includes(norm(st.q.trim()))) return false;
    if (!(T(s.desde) < p1 && T(s.hasta) + DAY > p0)) return false;
    if (st.vista === "act") return rangoAct(s, st.act) !== null;
    return true;
  });
}
function rangoAct(s, k) {
  const e = s.etapas.find(x => x.k === k); if (!e) return null;
  const v = [e.aok ? e.a : null, e.bok ? e.b : null].filter(Boolean).map(T);
  if (!v.length) return null;
  return [Math.min(...v), Math.max(...v)];
}
function rangoSitio(s) { return [T(s.desde), T(s.hasta)]; }

// ---- gráfico
function renderGantt(rows) {
  const [p0, p1] = periodo(), span = p1 - p0;
  const pct = t => (t - p0) / span * 100;
  let ticks = "", lines = "";
  if (!st.mes) {
    for (let m = 0; m < 12; m++) {
      const a = Date.UTC(+st.anio, m, 1), b = Date.UTC(+st.anio, m+1, 1);
      ticks += `<span style="left:${pct((a+b)/2)}%">${MESES[m]}</span>`;
      lines += `<i style="left:${pct(a)}%"></i>`;
    }
    lines += `<i style="left:100%"></i>`;
  } else {
    const n = (p1 - p0) / DAY;
    for (let d = 0; d < n; d++) {
      ticks += `<span style="left:${(d+.5)/n*100}%;font-size:11px">${d+1}</span>`;
      lines += `<i style="left:${d/n*100}%"></i>`;
    }
  }
  const act = ACT.find(a => a.k === st.act);
  const html = rows.map(s => {
    const [a, b] = st.vista === "act" ? rangoAct(s, st.act) : rangoSitio(s);
    const l = Math.max(pct(a), 0), r = Math.min(pct(b + DAY), 100);
    const color = st.vista === "act" ? act.c : (s.inst ? "var(--ok)" : "var(--pend)");
    const tip = `${s.mun} · ${dm(a)}${b !== a ? " – " + dm(b) : ""} · ${s.inst ? "Instalado" : "Pendiente"}`;
    return `<div class="row${st.sel === s.id ? " sel" : ""}" data-id="${s.id}" role="button" tabindex="0" aria-expanded="${st.sel === s.id}">
      <div class="n">${esc(s.mun)}</div>
      <div class="t"><div class="bar${st.vista!=="act" && !s.inst ? " pend" : ""}" style="left:${l}%;width:${Math.max(r-l,.9)}%;background:${color}" title="${esc(tip)}"></div></div>
      <svg class="c" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg></div>`;
  }).join("");
  $("chart").innerHTML = `<div class="scroll"><div class="gantt">
    <div class="axis-year">${st.mes ? MESES_L[st.mes-1] + " " : ""}${st.anio}</div>
    <div class="axis"><div class="l">Municipio / Sitio</div><div class="t">${ticks}</div><div></div></div>
    <div class="gl">${lines}</div>${html}</div></div>`;
}

// ---- tabla
function celda(e, par) {
  if (!e) return "—";
  const f = (v, ok) => v ? (ok ? dm(T(v)) : `<span class="bad" title="Fecha inconsistente: ${v}">⚠ <s>${v.split("-").reverse().join("/")}</s></span>`) : "";
  if (par) { const a = f(e.a, e.aok), b = f(e.b, e.bok); return a === b || !b ? a : `${a} – ${b}`; }
  return f(e.a, e.aok);
}
function renderTabla(rows) {
  const g = (s, k) => s.etapas.find(e => e.k === k);
  $("chart").innerHTML = `<div class="scroll"><table><thead><tr>
    <th>Municipio</th><th>Departamento</th><th>Supervisor UIS</th><th>Socialización</th><th>Obra civil</th><th>Transporte EMB</th><th>Armado EMB</th><th>Instalación</th><th>Capacitación y entrega</th><th>Días ejec.</th><th>Estado</th></tr></thead><tbody>` +
    rows.map(s => `<tr><td class="m">${esc(s.mun)}</td><td>${esc(s.dep)}</td><td>${esc(s.sup) || "—"}</td>
      <td>${celda(g(s,"soc"))}</td><td>${celda(g(s,"obra"),1)}</td><td>${celda(g(s,"tra"),1)}</td><td>${celda(g(s,"arm"),1)}</td>
      <td>${celda(g(s,"ins"))}</td><td>${celda(g(s,"cap"))}</td><td>${s.dias ?? "—"}</td>
      <td><span class="st"><i style="background:${s.inst ? "var(--ok)" : "var(--pend)"}"></i>${s.inst ? "Instalado" : "Pendiente"}</span></td></tr>`).join("") +
    `</tbody></table></div>`;
}

// ---- detalle del sitio
function renderDetalle() {
  const s = S.find(x => x.id === st.sel);
  if (!s) { $("detail").innerHTML = ""; return; }
  const [a, b] = rangoSitio(s), span = b + DAY - a;
  const pos = t => (t - a) / span * 100;
  const filas = s.etapas.map(e => {
    const col = ACT.find(x => x.k === e.k).c;
    const v = [e.aok ? e.a : null, e.bok ? e.b : null].filter(Boolean).map(T);
    const bad = [!e.aok && e.a ? e.a : null, !e.bok && e.b ? e.b : null].filter(Boolean);
    const txt = v.length ? (v[0] === v[v.length-1] ? dm(v[0]) : dm(Math.min(...v)) + " – " + dm(Math.max(...v))) : "sin fecha válida";
    const barra = v.length ? (() => { const x0 = Math.min(...v), x1 = Math.max(...v); return `<i style="left:${pos(x0)}%;width:${Math.max(pos(x1 + DAY) - pos(x0), 2)}%;background:${col}"></i>`; })() : "";
    return `<div class="stage"><div class="s"><b>${e.n}</b><small>${txt}</small>${bad.length ? `<small class="bad"> · ⚠ revisar ${bad.map(x => x.split("-").reverse().join("/")).join(", ")}</small>` : ""}</div><div class="track">${barra}</div></div>`;
  }).join("");
  const sub = [s.dep, s.cod ? "DIVIPOLA " + s.cod : "", s.sup ? "Supervisor UIS: " + s.sup : ""].filter(Boolean).join(" · ");
  $("detail").innerHTML = `<section class="detail" aria-label="Detalle de ${esc(s.mun)}">
    <header><div><h3>${esc(s.mun)}</h3><p>Línea de tiempo detallada · ${esc(sub)}</p></div><button class="x" type="button" aria-label="Cerrar detalle" id="xDet">×</button></header>
    ${filas}
    <p class="note">Las barras representan la duración de cada etapa y permiten identificar actividades simultáneas.</p>
    <div class="facts">${s.emb ? `<span>Embarcadero Número: <b>${s.emb}</b></span>` : ""}${s.dias ? `<span>Días de ejecución: <b>${s.dias}</b></span>` : ""}${s.habiles ? `<span>Días hábiles: <b>${s.habiles}</b></span>` : ""}<span>Estado: <b>${s.inst ? "Instalado" : "Pendiente"}</b></span></div>
  </section>`;
}

// ---- render general
function render() {
  const rows = visibles();
  if (st.sel && !rows.some(r => r.id === st.sel)) st.sel = null;
  $("vGeneral").classList.toggle("on", st.vista === "general");
  $("vAct").classList.toggle("on", st.vista === "act");
  $("mGraf").classList.toggle("on", st.modo === "graf");
  $("mTab").classList.toggle("on", st.modo === "tab");
  $("chips").hidden = st.vista !== "act";
  $("chips-help").hidden = st.vista !== "act";
  $("activity-legend").hidden = st.vista !== "general";
  document.querySelectorAll(".chip").forEach(c => {
    const activo = c.dataset.k === st.act;
    c.classList.toggle("on", activo);
    c.setAttribute("aria-pressed", String(activo));
  });
  const actN = ACT.find(a => a.k === st.act).n;
  $("tTitulo").textContent = st.vista === "act" ? "Cronograma por actividad" : "Cronograma general";
  $("tSub").textContent = (st.vista === "act" ? actN + " · " : "") + (st.mes ? MESES_L[st.mes-1] + " " : "Año ") + st.anio;
  $("activity-legend").innerHTML = `<b>Leyenda:</b>` + ACT
    .map(a => `<span><i style="background:${a.c}"></i>${a.n}</span>`).join("");
  $("legend").innerHTML = st.vista === "act"
    ? `<span><i style="background:${ACT.find(a => a.k === st.act).c}"></i>${actN}</span>`
    : `<span><i style="background:var(--ok)"></i>Instalado</span><span><i style="background:var(--pend)"></i>Pendiente</span>`;
  if (!rows.length) $("chart").innerHTML = `<div class="empty">No hay sitios para este filtro. Cambia el año, el mes o el departamento.</div>`;
  else if (st.modo === "tab") renderTabla(rows); else renderGantt(rows);
  renderDetalle();
}

// ---- eventos
$("fAnio").onchange = e => { st.anio = e.target.value; render(); };
$("fMes").onchange = e => { st.mes = +e.target.value; render(); };
$("fDep").onchange = e => { st.dep = e.target.value; render(); };
$("fQ").oninput = e => { st.q = e.target.value; render(); };
$("vGeneral").onclick = () => { st.vista = "general"; render(); };
$("vAct").onclick = () => { st.vista = "act"; render(); };
$("mGraf").onclick = () => { st.modo = "graf"; render(); };
$("mTab").onclick = () => { st.modo = "tab"; render(); };
$("chips").onclick = e => { const c = e.target.closest(".chip"); if (c) { st.act = c.dataset.k; render(); } };
const abrir = el => { const r = el.closest(".row"); if (!r) return; const id = +r.dataset.id; st.sel = st.sel === id ? null : id; render(); if (st.sel) $("detail").scrollIntoView({behavior:"smooth", block:"nearest"}); };
$("chart").onclick = e => abrir(e.target);
$("chart").onkeydown = e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); abrir(e.target); } };
$("detail").onclick = e => { if (e.target.id === "xDet") { st.sel = null; render(); } };
render();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
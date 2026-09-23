#!/usr/bin/env python3
"""
generar_tablero_fontur.py
--------------------------
Lee el archivo "Tablero_control_General_Instalación_*.xlsx" (hoja "Línea de
Tiempo") y genera un HTML interactivo, filtrable por MUNICIPIO y
DEPARTAMENTO, con una fila por cada tipo de soporte documental de cada
municipio.

USO:
    python generar_tablero_fontur.py [ruta_excel.xlsx] [salida.html]

Si no se indican rutas, usa los valores por defecto de abajo (pensados para
ejecutarse desde la carpeta del proyecto en tu equipo, donde vive el Excel:
C:\\Users\\carlos.garcia\\Documents\\DOCUMENTACION REPASO\\PROYECTO FONTUR\\).

Cada vez que actualices el Excel maestro, vuelve a correr este script para
regenerar el HTML con los datos más recientes.
"""

import json
import re
import sys
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

# Rutas por defecto (ajusta si tu Excel vive en otra carpeta / nombre)
EXCEL_POR_DEFECTO = r"C:\Users\carlos.garcia\Desktop\github\linea_tiempo_fase_instalacion\Tablero_control_General_Instalación_180926.xlsx"
HTML_POR_DEFECTO = r"C:\Users\carlos.garcia\Desktop\github\linea_tiempo_fase_instalacion\index.html"

 
HOJA = "Línea de Tiempo"
FILA_ENCABEZADOS = 10
PRIMERA_FILA_DATOS = 11
 
COL_CODIGO = 2  # CÓDIGO DIPOLA MUNICIPIO
COL_MUNICIPIO = 6
COL_DEPARTAMENTO = 7
 
# Columna (índice openpyxl, 1-based) -> etiqueta legible.
# Los 16 tipos de soporte documental pedidos (se excluye a propósito la
# columna AF "RUTA SOPORTE REGISTRO FOTOGRÁFICO OBRA CIVIL2", que no forma
# parte de la lista original).
COLUMNAS_SOPORTES = [
    (17, "Póliza de Cumplimiento"),
    (18, "Póliza de Responsabilidad Civil"),
    (19, "Actas de Inicio Autoridad de Transporte"),
    (20, "Cronograma de Instalación"),
    (21, "Reunión Preparación Instalación Municipio"),
    (22, "Socialización Técnica Municipio"),
    (23, "Socialización Inicio Comunidad"),
    (24, "Autorización Ocupación Espacio Público"),
    (25, "Actas de Vecindad"),
    (26, "Bitácoras"),
    (27, "Certificaciones Calidad Materiales"),
    (28, "Control Calidad Interventoría"),
    (29, "Seguimiento Armado EMB"),
    (30, "Capacitación EMB"),
    (31, "Entrega EMB"),
    (33, "Recepción EMB Cotecmar"),
]
 
MARCADOR_ARCHIVO = re.compile(r"NOMBRE\s*ARCHIVO\s*:", re.IGNORECASE)
 
 
# ---------------------------------------------------------------------------
# Parseo
# ---------------------------------------------------------------------------
 
def parsear_celda(valor):
    """Convierte el contenido crudo de una celda RUTA SOPORTE DOCUMENTAL en
    {estado, ruta, archivos}.
 
    estado es uno de: 'no_requiere', 'pendiente', 'con_soporte'.
    """
    if valor is None or str(valor).strip() == "":
        return {"estado": "pendiente", "ruta": "", "archivos": []}
 
    texto = str(valor).strip()
 
    if texto.upper() == "NO REQUIERE":
        return {"estado": "no_requiere", "ruta": "", "archivos": []}
 
    partes = MARCADOR_ARCHIVO.split(texto, maxsplit=1)
    ruta = partes[0].strip()
    archivos = []
    if len(partes) > 1:
        for linea in partes[1].split("\n"):
            linea = linea.strip()
            if linea:
                archivos.append(linea)
 
    return {"estado": "con_soporte", "ruta": ruta, "archivos": archivos}
 
 
def cargar_registros(ruta_excel):
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    if HOJA not in wb.sheetnames:
        sys.exit(f"No encontré la hoja '{HOJA}' en {ruta_excel}. Hojas disponibles: {wb.sheetnames}")
    ws = wb[HOJA]
 
    registros = []
    for fila in range(PRIMERA_FILA_DATOS, ws.max_row + 1):
        municipio = ws.cell(row=fila, column=COL_MUNICIPIO).value
        if municipio is None or str(municipio).strip() == "":
            continue
        municipio = str(municipio).strip()
        departamento = str(ws.cell(row=fila, column=COL_DEPARTAMENTO).value or "").strip()
        codigo = ws.cell(row=fila, column=COL_CODIGO).value
 
        for col_idx, etiqueta in COLUMNAS_SOPORTES:
            valor = ws.cell(row=fila, column=col_idx).value
            parseado = parsear_celda(valor)
            registros.append({
                "municipio": municipio,
                "departamento": departamento,
                "codigo": str(codigo) if codigo is not None else "",
                "tipo": etiqueta,
                **parseado,
            })
 
    return registros
 
 
# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
 
PLANTILLA_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Tablero de Soportes Documentales — Instalación de Embarcaderos</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root {
  --bg: #F2F4F3;
  --bg-panel: #FFFFFF;
  --bg-row-alt: #F7F8F7;
  --ink: #16232B;
  --ink-soft: #4B5B63;
  --ink-faint: #8A9AA1;
  --line: #D9E0DF;
  --accent: #0E7490;
  --accent-ink: #063542;
  --accent-soft: #E1EFF1;
  --tag-noreq-bg: #E7E4DD;
  --tag-noreq-ink: #5A5347;
  --shadow: 0 1px 2px rgba(15, 35, 40, 0.06);
  --radius: 10px;
  box-sizing: border-box;
  padding-top: env(safe-area-inset-top, 0px);
  padding-bottom: env(safe-area-inset-bottom, 0px);
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --bg: #10171A;
    --bg-panel: #172024;
    --bg-row-alt: #131C1F;
    --ink: #E8EDEC;
    --ink-soft: #A9B7B9;
    --ink-faint: #6C7A7C;
    --line: #253134;
    --accent: #4FB3C9;
    --accent-ink: #DFF4F8;
    --accent-soft: #1B3439;
    --tag-noreq-bg: #2A2A24;
    --tag-noreq-ink: #C9C0AC;
    --shadow: 0 1px 2px rgba(0, 0, 0, 0.4);
  }
}
:root[data-theme="dark"] {
  --bg: #10171A;
  --bg-panel: #172024;
  --bg-row-alt: #131C1F;
  --ink: #E8EDEC;
  --ink-soft: #A9B7B9;
  --ink-faint: #6C7A7C;
  --line: #253134;
  --accent: #4FB3C9;
  --accent-ink: #DFF4F8;
  --accent-soft: #1B3439;
  --tag-noreq-bg: #2A2A24;
  --tag-noreq-ink: #C9C0AC;
  --shadow: 0 1px 2px rgba(0, 0, 0, 0.4);
}
html { scroll-padding-top: env(safe-area-inset-top, 0px); }
* { box-sizing: border-box; }
html, body {
  margin: 0;
  background: var(--bg);
  color: var(--ink);
  font-family: "IBM Plex Sans", -apple-system, "Segoe UI", Roboto, sans-serif;
  font-size: 15px;
  line-height: 1.45;
}
.wrap {
  max-width: 1180px;
  margin: 0 auto;
  padding: 28px 20px 60px;
}
header.page-head {
  margin-bottom: 22px;
}
.eyebrow {
  color: var(--accent);
  font-weight: 600;
  font-size: 13px;
  letter-spacing: 0.02em;
  margin: 0 0 6px;
}
h1 {
  font-size: clamp(22px, 3.2vw, 30px);
  font-weight: 700;
  margin: 0 0 8px;
  letter-spacing: -0.01em;
}
.subtitle {
  color: var(--ink-soft);
  margin: 0;
  max-width: 62ch;
}
.panel {
  background: var(--bg-panel);
  border: 1px solid var(--line);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
}
.filters {
  padding: 16px;
  margin: 22px 0 16px;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 12px;
  align-items: end;
}
.field label {
  display: block;
  font-size: 12.5px;
  color: var(--ink-soft);
  margin-bottom: 5px;
}
.field select,
.field input {
  width: 100%;
  padding: 9px 10px;
  border: 1px solid var(--line);
  border-radius: 7px;
  background: var(--bg);
  color: var(--ink);
  font: inherit;
}
.field select:focus,
.field input:focus {
  outline: 2px solid var(--accent);
  outline-offset: 1px;
}
.filters .reset {
  align-self: end;
}
button.reset {
  padding: 9px 14px;
  border: 1px solid var(--line);
  background: var(--bg-panel);
  color: var(--ink-soft);
  border-radius: 7px;
  cursor: pointer;
  font: inherit;
}
button.reset:hover { color: var(--ink); border-color: var(--ink-faint); }
.meta-row {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  color: var(--ink-soft);
  font-size: 13px;
  margin: 0 2px 10px;
  flex-wrap: wrap;
  gap: 6px;
}
.meta-row strong { color: var(--ink); }
.table-scroll {
  overflow-x: auto;
  border: 1px solid var(--line);
  border-radius: var(--radius);
  background: var(--bg-panel);
  box-shadow: var(--shadow);
}
table {
  border-collapse: collapse;
  width: 100%;
  min-width: 680px;
  table-layout: fixed;
}
thead th {
  text-align: left;
  font-size: 12.5px;
  font-weight: 600;
  color: var(--ink-soft);
  padding: 10px 14px;
  border-bottom: 1px solid var(--line);
  position: sticky;
  top: 0;
  background: var(--bg-panel);
  white-space: nowrap;
}
th.col-tipo, td.col-tipo { width: 24%; }
th.col-ruta, td.col-ruta { width: 42%; }
th.col-archivo, td.col-archivo { width: 34%; }
tbody td {
  padding: 10px 14px;
  border-bottom: 1px solid var(--line);
  vertical-align: top;
}
tbody tr:nth-child(even) { background: var(--bg-row-alt); }
tbody tr:hover { background: var(--accent-soft); }
tr.group-row td {
  background: var(--accent-soft);
  padding: 8px 14px;
  border-bottom: 1px solid var(--line);
}
tr.group-row:hover td { background: var(--accent-soft); }
.group-municipio { font-weight: 600; color: var(--accent-ink); }
.group-departamento { color: var(--ink-soft); font-size: 13px; margin-left: 8px; }
td.col-tipo { color: var(--ink); }
td.col-ruta { font-family: "IBM Plex Mono", ui-monospace, Consolas, monospace; font-size: 12.5px; color: var(--ink-soft); overflow-wrap: break-word; }
td.col-archivo { font-family: "IBM Plex Mono", ui-monospace, Consolas, monospace; font-size: 12.5px; overflow-wrap: break-word; }
td.col-archivo ul { margin: 0; padding-left: 16px; }
td.col-archivo li { margin: 2px 0; }
.tag-noreq {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 999px;
  background: var(--tag-noreq-bg);
  color: var(--tag-noreq-ink);
  font-family: "IBM Plex Sans", sans-serif;
  font-size: 12px;
  font-weight: 500;
}
.dash { color: var(--ink-faint); }
.empty-state {
  padding: 40px 14px;
  text-align: center;
  color: var(--ink-soft);
}
footer.note {
  margin-top: 18px;
  color: var(--ink-faint);
  font-size: 12.5px;
}
@media (max-width: 640px) {
  .wrap { padding: 18px 14px 40px; }
  table { min-width: 720px; }
}
</style>
</head>
<body>
<div class="wrap">
  <header class="page-head">
    <p class="eyebrow">PROYECTO FONTUR · Instalación de embarcaderos</p>
    <h1>Tablero de soportes documentales fase instalación</h1>
    <p class="subtitle">Puede encontrar los siguientes documentos:</p>
    <p class="subtitle">Bitácoras, Pólizas, Cronogramas, Actas, Permisos</p>
  </header>
 
  <div class="panel filters">
    <div class="field">
      <label for="f-municipio">Municipio</label>
      <select id="f-municipio">
        <option value="">Todos los municipios</option>
      </select>
    </div>
    <div class="field">
      <label for="f-departamento">Departamento</label>
      <select id="f-departamento">
        <option value="">Todos los departamentos</option>
      </select>
    </div>
    <div class="field">
      <label for="f-buscar">Buscar (tipo, ruta o archivo)</label>
      <input id="f-buscar" type="text" placeholder="p. ej. bitácora, cronograma, poliza...">
    </div>
    <div class="reset">
      <button class="reset" id="btn-reset" type="button">Limpiar filtros</button>
    </div>
  </div>
 
  <div class="meta-row">
    <span id="conteo">—</span>
    <span id="pendientes-nota"></span>
  </div>
 
  <div class="table-scroll">
    <table>
      <thead>
        <tr>
          <th class="col-tipo">Tipo de documento</th>
          <th class="col-archivo">Archivo(s)</th>
          <th class="col-ruta">Ruta</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="empty-state" id="empty-state" style="display:none;">No hay filas que coincidan con el filtro actual.</div>
  </div>
 
  <footer class="note">Generado automáticamente el· __FECHA_GENERACION__</footer>
</div>
 
<script>
const DATA = __DATA_JSON__;
 
const elMunicipio = document.getElementById('f-municipio');
const elDepartamento = document.getElementById('f-departamento');
const elBuscar = document.getElementById('f-buscar');
const elTbody = document.getElementById('tbody');
const elConteo = document.getElementById('conteo');
const elPendientes = document.getElementById('pendientes-nota');
const elEmpty = document.getElementById('empty-state');
 
function unico(campo) {
  return [...new Set(DATA.map(r => r[campo]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
}
 
function poblarSelect(el, valores) {
  for (const v of valores) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    el.appendChild(opt);
  }
}
 
poblarSelect(elMunicipio, unico('municipio'));
poblarSelect(elDepartamento, unico('departamento'));
 
function escapeHtml(s) {
  return s.replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
 
function renderRuta(r) {
  if (r.estado === 'no_requiere') return '<span class="tag-noreq">No requiere</span>';
  if (r.estado === 'pendiente') return '<span class="dash">—</span>';
  // permite que el navegador corte la ruta después de cada backslash,
  // en vez de partir palabras a la mitad
  return escapeHtml(r.ruta || '').split('\\').join('\\<wbr>');
}
 
function renderArchivo(r) {
  if (r.estado === 'no_requiere' || r.estado === 'pendiente') return '<span class="dash">—</span>';
  if (!r.archivos || r.archivos.length === 0) return '<span class="dash">—</span>';
  if (r.archivos.length === 1) return escapeHtml(r.archivos[0]);
  return '<ul>' + r.archivos.map(a => '<li>' + escapeHtml(a) + '</li>').join('') + '</ul>';
}
 
function aplicarFiltros() {
  const municipio = elMunicipio.value;
  const departamento = elDepartamento.value;
  const texto = elBuscar.value.trim().toLowerCase();
 
  const filtradas = DATA.filter(r => {
    if (municipio && r.municipio !== municipio) return false;
    if (departamento && r.departamento !== departamento) return false;
    if (texto) {
      const hay = (r.tipo + ' ' + r.ruta + ' ' + (r.archivos || []).join(' ')).toLowerCase();
      if (!hay.includes(texto)) return false;
    }
    return true;
  });
 
  let filas = '';
  let grupoActual = null;
  for (const r of filtradas) {
    const clave = r.municipio + '||' + r.departamento;
    if (clave !== grupoActual) {
      grupoActual = clave;
      filas += `
        <tr class="group-row">
          <td colspan="3">
            <span class="group-municipio">${escapeHtml(r.municipio)}</span>
            <span class="group-departamento">${escapeHtml(r.departamento)}</span>
          </td>
        </tr>`;
    }
    filas += `
      <tr>
        <td class="col-tipo">${escapeHtml(r.tipo)}</td>
        <td class="col-archivo">${renderArchivo(r)}</td>
        <td class="col-ruta">${renderRuta(r)}</td>
      </tr>`;
  }
  elTbody.innerHTML = filas;
 
  elEmpty.style.display = filtradas.length === 0 ? 'block' : 'none';
 
  const pendientes = filtradas.filter(r => r.estado === 'pendiente').length;
  elConteo.innerHTML = `Mostrando <strong>${filtradas.length}</strong> de ${DATA.length} filas`;
  elPendientes.textContent = filtradas.length ? `${pendientes} pendientes sin soporte cargado` : '';
}
 
// Si se elige un departamento, restringe la lista de municipios a ese departamento
elDepartamento.addEventListener('change', () => {
  const dep = elDepartamento.value;
  const actual = elMunicipio.value;
  elMunicipio.innerHTML = '<option value="">Todos los municipios</option>';
  const municipios = dep ? unico('municipio').filter(m => DATA.some(r => r.municipio === m && r.departamento === dep)) : unico('municipio');
  poblarSelect(elMunicipio, municipios);
  if (municipios.includes(actual)) elMunicipio.value = actual;
  aplicarFiltros();
});
 
elMunicipio.addEventListener('change', aplicarFiltros);
elBuscar.addEventListener('input', aplicarFiltros);
document.getElementById('btn-reset').addEventListener('click', () => {
  elMunicipio.value = '';
  elDepartamento.value = '';
  elBuscar.value = '';
  elMunicipio.innerHTML = '<option value="">Todos los municipios</option>';
  poblarSelect(elMunicipio, unico('municipio'));
  aplicarFiltros();
});
 
aplicarFiltros();
</script>
</body>
</html>
"""
 
 
def construir_html(registros, ruta_salida):
    from datetime import datetime
    data_json = json.dumps(registros, ensure_ascii=False).replace("</script", "<\\/script")
    html = PLANTILLA_HTML.replace("__DATA_JSON__", data_json)
    html = html.replace("__FECHA_GENERACION__", datetime.now().strftime("%Y-%m-%d %H:%M"))
    Path(ruta_salida).write_text(html, encoding="utf-8")
 
 
def main():
    ruta_excel = sys.argv[1] if len(sys.argv) > 1 else EXCEL_POR_DEFECTO
    ruta_salida = sys.argv[2] if len(sys.argv) > 2 else HTML_POR_DEFECTO
 
    if not Path(ruta_excel).exists():
        sys.exit(f"No encontré el archivo: {ruta_excel}")
 
    registros = cargar_registros(ruta_excel)
    construir_html(registros, ruta_salida)
 
    con_soporte = sum(1 for r in registros if r["estado"] == "con_soporte")
    no_requiere = sum(1 for r in registros if r["estado"] == "no_requiere")
    pendientes = sum(1 for r in registros if r["estado"] == "pendiente")
    municipios = len(set(r["municipio"] for r in registros))
 
    print(f"Municipios procesados: {municipios}")
    print(f"Filas generadas: {len(registros)}")
    print(f"  con soporte cargado: {con_soporte}")
    print(f"  no requiere:         {no_requiere}")
    print(f"  pendientes:          {pendientes}")
    print(f"HTML generado en: {ruta_salida}")
 
 
if __name__ == "__main__":
    main()
 
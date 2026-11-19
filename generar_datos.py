#!/usr/bin/env python3
"""
Genera index.html (aplicación de consulta de causales) a partir del Excel.

Uso:
    python generar_datos.py TABLA_CAUSALES_HFC_2026.xlsx index.html

Requiere: pip install openpyxl
Cada vez que cambie el Excel, vuelva a ejecutar este script.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

# Valores que la columna A usa como "producto" (aplicabilidad). Cualquier otro
# texto de la columna A, sin datos en la fila, se toma como encabezado de sección.
PRODUCTOS = {
    "TV,LINEA", "INTERNET", "TV,INTERNET,LINEA", "TV,INTERNNET,LINEA",
    "TV,INTERNNET", "TV,INTERNET", "TV", "IPTV", "LINEA",
}

HOJAS_ORDEN = ["HFC", "CANOPY", "GPON", "Hoja2"]

# Mapeo de columnas por hoja.
#   grupos:  bloques de columnas que se muestran juntos en la ficha
#   codigos: columna de código -> columna de descripción que le corresponde
SPECS = {
    "HFC": {
        "inicio": 5,        # primera fila de datos
        "cabecera": 4,      # fila con los encabezados originales
        "grupos": [
            ("Ubicación", ["B", "C", "D"]),
            ("Tipo", ["E", "F", "G", "H"]),
            ("Causa", ["I", "J", "K", "L"]),
            ("Combinación", ["M"]),
        ],
        "etiquetas_extra": {"M": "Código combinado (causa - tipo - ubicación)"},
        "codigos": {"B": "C", "E": "F", "G": "H", "I": "J", "K": "L"},
    },
    "GPON": {
        "inicio": 3,
        "cabecera": 2,
        "grupos": [("Registro", ["B", "C", "D", "E", "F", "G"])],
        "seccion": "UBICACIÓN OPEN",   # la columna A se hereda hacia abajo
        "codigos": {"B": "C", "D": "E", "F": "G"},
    },
    "CANOPY": {
        "inicio": 4,
        # La hoja no tiene encabezados: se usan etiquetas genéricas
        "etiquetas": {"B": "Código", "C": "Descripción", "D": "Código",
                      "E": "Descripción", "F": "Código", "G": "Tipo"},
        "grupos": [("Registro", ["B", "C", "D", "E", "F", "G"])],
        "contexto": [("Sección", "CANOPY")],
        "codigos": {"B": "C", "D": "E", "F": "G"},
    },
    "Hoja2": {
        "inicio": 1,
        "etiquetas": {"A": "Código", "B": "Descripción"},
        "grupos": [("Registro", ["A", "B"])],
        "codigos": {"A": "B"},
    },
}


def limpiar(v):
    """Convierte la celda a texto tal como aparece, sin cambiar las palabras."""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()
    return s or None


def leer_hoja(ws, nombre, spec, registros, indice):
    base = [c for _, cols in spec["grupos"] for c in cols]
    cols = base + ([] if "A" in base else ["A"])

    etiq = {}
    if spec.get("cabecera"):
        etiq = {c: limpiar(ws[f"{c}{spec['cabecera']}"].value) for c in base}
    etiq.update(spec.get("etiquetas", {}))
    etiq.update(spec.get("etiquetas_extra", {}))

    seccion = None
    carry = None
    for r in range(spec["inicio"], ws.max_row + 1):
        v = {c: limpiar(ws[f"{c}{r}"].value) for c in cols}
        a = v.get("A")
        hay_datos = any(v[c] for c in base)

        if nombre == "HFC":
            if not hay_datos:
                # Fila solo con texto en A: encabezado de sección (o producto sin datos)
                if a and a not in PRODUCTOS:
                    seccion = a
                continue
            ctx = [("Sección", seccion)] if seccion else []
            if a:
                ctx.append(("Producto", a))
        elif nombre == "GPON":
            if a:
                carry = a
            if not hay_datos:
                continue
            ctx = [(spec["seccion"], carry)] if carry else []
        else:
            if not hay_datos:
                continue
            ctx = list(spec.get("contexto", []))

        grupos = []
        for titulo, gcols in spec["grupos"]:
            campos = [
                {"etiqueta": etiq.get(c) or c, "valor": v[c], "codigo": c in spec["codigos"]}
                for c in gcols if v[c]
            ]
            if campos:
                grupos.append({"titulo": titulo, "campos": campos})

        idx = len(registros)
        registros.append({"hoja": nombre, "contexto": [list(x) for x in ctx], "grupos": grupos})

        # Índice de códigos: cada código apunta a sus registros
        for cc, dc in spec["codigos"].items():
            cod = v.get(cc)
            if not cod:
                continue
            entrada = indice.setdefault(
                (nombre, cod), {"hoja": nombre, "codigo": cod, "descripciones": [], "registros": []}
            )
            d = v.get(dc)
            if d and d not in entrada["descripciones"]:
                entrada["descripciones"].append(d)
            if idx not in entrada["registros"]:
                entrada["registros"].append(idx)


def main():
    xlsx = Path(sys.argv[1] if len(sys.argv) > 1 else "TABLA_CAUSALES_HFC_2026.xlsx")
    salida = Path(sys.argv[2] if len(sys.argv) > 2 else "index.html")

    # data_only=True: toma los valores calculados de las fórmulas (p. ej. =+J5)
    wb = openpyxl.load_workbook(xlsx, data_only=True)

    registros, indice, leidas = [], {}, []
    for nombre, spec in SPECS.items():
        if nombre not in wb.sheetnames:
            print(f"Aviso: no existe la hoja '{nombre}', se omite.")
            continue
        leer_hoja(wb[nombre], nombre, spec, registros, indice)
        leidas.append(nombre)

    entradas = sorted(
        indice.values(),
        key=lambda e: (
            HOJAS_ORDEN.index(e["hoja"]),
            (0, int(e["codigo"])) if e["codigo"].isdigit() else (1, e["codigo"]),
        ),
    )
    for i, e in enumerate(entradas):
        e["id"] = i
        e["busqueda"] = (e["codigo"] + " " + " ".join(e["descripciones"])).lower()

    datos = {"hojas": leidas, "registros": registros, "indice": entradas}
    json_txt = json.dumps(datos, ensure_ascii=False).replace("</", "<\\/")
    salida.write_text(PLANTILLA.replace("__DATOS__", json_txt), encoding="utf-8")
    print(f"Listo: {len(entradas)} códigos, {len(registros)} registros -> {salida}")


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Consulta de Causales HFC 2026</title>
<style>
:root{--azul:#1f4e79;--fondo:#f4f6f9;--borde:#d9dee7;--texto:#1d2330;--gris:#667085;--resalta:#fff3c4}
*{box-sizing:border-box}
body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--fondo);color:var(--texto)}
header{background:var(--azul);color:#fff;padding:16px 24px}
header h1{margin:0;font-size:20px}
header p{margin:4px 0 0;font-size:13px;opacity:.9}
.layout{display:grid;grid-template-columns:380px 1fr;gap:16px;padding:16px;max-width:1280px;margin:0 auto;align-items:start}
.panel{background:#fff;border:1px solid var(--borde);border-radius:8px;padding:14px}
#buscar{width:100%;padding:10px;font-size:15px;border:1px solid var(--borde);border-radius:6px;font-family:inherit}
.filtros{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}
.filtros button{border:1px solid var(--borde);background:#fff;border-radius:16px;padding:4px 10px;cursor:pointer;font:13px Arial,sans-serif}
.filtros button.activo{background:var(--azul);color:#fff;border-color:var(--azul)}
#contador{font-size:12px;color:var(--gris);margin-bottom:6px}
#lista{max-height:68vh;overflow:auto;display:flex;flex-direction:column;gap:4px}
.item{display:grid;grid-template-columns:70px 1fr auto;gap:8px;text-align:left;border:1px solid transparent;background:#fff;padding:8px;border-radius:6px;cursor:pointer;font:inherit;color:inherit}
.item:hover{background:#eef3f9}
.item.activo{border-color:var(--azul);background:#e8f0fa}
.cod{font-weight:bold;color:var(--azul);word-break:break-all}
.desc{font-size:14px}
.tag{font-size:11px;color:var(--gris);align-self:center;white-space:nowrap}
.vacio{color:var(--gris);text-align:center;padding:40px 10px}
#detalle h2{margin:0 0 4px;color:var(--azul)}
.sub{color:var(--gris);font-size:13px;margin-bottom:12px}
.descs{margin:0 0 14px;padding-left:18px}
.descs li{margin-bottom:4px}
.card{border:1px solid var(--borde);border-radius:8px;padding:12px;margin-bottom:12px}
.ctx{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:6px}
.ctx span{background:#eef3f9;border-radius:4px;padding:3px 8px;font-size:12px}
h3{font-size:12px;text-transform:uppercase;color:var(--gris);margin:12px 0 4px;letter-spacing:.03em}
table{width:100%;border-collapse:collapse;font-size:14px}
th,td{text-align:left;padding:6px 8px;border-bottom:1px solid #eef0f4;vertical-align:top}
th{width:36%;color:#344054;font-weight:600}
tr.match td,tr.match th{background:var(--resalta)}
@media(max-width:800px){.layout{grid-template-columns:1fr}#lista{max-height:40vh}}
</style>
</head>
<body>
<header>
  <h1>Consulta de causales HFC 2026</h1>
  <p>Seleccione un código para ver su descripción completa, tal como aparece en la tabla.</p>
</header>
<main class="layout">
  <section class="panel">
    <input id="buscar" type="search" placeholder="Buscar por código o descripción…" autocomplete="off">
    <div class="filtros" id="filtros"></div>
    <div id="contador"></div>
    <div id="lista"></div>
  </section>
  <section class="panel" id="detalle"><p class="vacio">Seleccione un código de la lista.</p></section>
</main>
<script>
const DATOS = __DATOS__;
const estado = {hoja: 'Todas', q: '', sel: null};
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

function pintarFiltros() {
  $('filtros').innerHTML = ['Todas', ...DATOS.hojas].map(h =>
    `<button data-hoja="${esc(h)}" class="${h === estado.hoja ? 'activo' : ''}">${esc(h)}</button>`).join('');
}

function pintarLista() {
  const q = estado.q.trim().toLowerCase();
  const items = DATOS.indice.filter(e =>
    (estado.hoja === 'Todas' || e.hoja === estado.hoja) && (!q || e.busqueda.includes(q)));
  $('contador').textContent = `${items.length} código${items.length === 1 ? '' : 's'}`;
  $('lista').innerHTML = items.map(e => `
    <button class="item${e.id === estado.sel ? ' activo' : ''}" data-id="${e.id}">
      <span class="cod">${esc(e.codigo)}</span>
      <span class="desc">${esc(e.descripciones[0] || '(sin descripción)')}</span>
      <span class="tag">${esc(e.hoja)}</span>
    </button>`).join('') || '<p class="vacio">Sin resultados</p>';
}

function tarjeta(r, codigo) {
  return `<article class="card">
    <div class="ctx">${r.contexto.map(([k, v]) => `<span>${esc(k)}: <b>${esc(v)}</b></span>`).join('')}</div>
    ${r.grupos.map(g => `
      <h3>${esc(g.titulo)}</h3>
      <table>${g.campos.map(c => `
        <tr class="${c.codigo && c.valor === codigo ? 'match' : ''}">
          <th>${esc(c.etiqueta)}</th><td>${esc(c.valor)}</td>
        </tr>`).join('')}</table>`).join('')}
  </article>`;
}

function pintarDetalle() {
  if (estado.sel === null) {
    $('detalle').innerHTML = '<p class="vacio">Seleccione un código de la lista.</p>';
    return;
  }
  const e = DATOS.indice[estado.sel];
  const regs = e.registros.map(i => DATOS.registros[i]);
  const n = e.descripciones.length;
  $('detalle').innerHTML = `
    <h2>Código ${esc(e.codigo)}</h2>
    <div class="sub">Hoja: ${esc(e.hoja)} · aparece en ${regs.length} registro${regs.length === 1 ? '' : 's'}</div>
    <h3>${n > 1 ? 'Descripciones de este código' : 'Descripción'}</h3>
    <ul class="descs">${e.descripciones.map(d => `<li>${esc(d)}</li>`).join('')}</ul>
    ${regs.map(r => tarjeta(r, e.codigo)).join('')}`;
}

$('buscar').addEventListener('input', ev => { estado.q = ev.target.value; pintarLista(); });
$('filtros').addEventListener('click', ev => {
  const b = ev.target.closest('button'); if (!b) return;
  estado.hoja = b.dataset.hoja; pintarFiltros(); pintarLista();
});
$('lista').addEventListener('click', ev => {
  const b = ev.target.closest('button[data-id]'); if (!b) return;
  estado.sel = Number(b.dataset.id); pintarLista(); pintarDetalle();
});

pintarFiltros();
pintarLista();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()

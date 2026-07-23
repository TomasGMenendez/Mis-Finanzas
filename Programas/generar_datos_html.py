"""
Finanzas_Toto.xlsx -> Mi Dashboard de Finanzas.html
Uso:  python generar_datos_html.py

Lee Finanzas_Toto.xlsx (hojas Gastos, Ingresos, Inversiones, Config) y
reemplaza la variable EMBEDDED_DATA dentro de "Mi Dashboard de Finanzas.html"
con los datos reales actuales. Así, al abrir ese archivo haciendo doble clic
(sin server, sin conexión en vivo), el dashboard ya muestra tus datos reales
tal como estaban en el Excel al momento de correr este script.

Corré este script cada vez que quieras que el Dashboard refleje los últimos
cambios del Excel (después de sincronizar MercadoPago/Bybit o editar a mano).
"""
import sys, os, subprocess, json, re, shutil, tempfile
from pathlib import Path
from datetime import datetime, date

BASE = Path(__file__).parent
RAIZ = BASE.parent  # donde viven Finanzas_Toto.xlsx y el Dashboard

def ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])
ensure("openpyxl")

from openpyxl import load_workbook

xlsx_path = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "Finanzas_Toto.xlsx"
if not xlsx_path.exists():
    print(f"No encuentro {xlsx_path}")
    sys.exit(1)

# El Excel puede estar abierto en la app de Excel o sincronizando con OneDrive,
# lo que bloquea la lectura directa (incluso para copiar el archivo). Se
# reintenta con una pausa breve porque casi siempre se libera en segundos.
tmp_path = Path(tempfile.gettempdir()) / "_Finanzas_Toto_read.xlsx"
last_err = None
for attempt in range(5):
    try:
        shutil.copyfile(xlsx_path, tmp_path)
        last_err = None
        break
    except PermissionError as e:
        last_err = e
        if attempt < 4:
            import time
            time.sleep(2)
if last_err:
    print(f"No pude leer {xlsx_path} — está abierto en Excel o sincronizando con OneDrive.")
    print("Guardalo y cerralo (o esperá a que OneDrive termine de sincronizar) y volvé a correr el script.")
    sys.exit(1)
wb = load_workbook(tmp_path, data_only=True)


def cell(ws, r, c):
    return ws.cell(r, c).value


def num(v):
    if isinstance(v, (int, float)):
        return v
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


def fecha_str(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return None


gastos = []
ws = wb["Gastos"]
for r in range(5, ws.max_row + 1):
    f = fecha_str(cell(ws, r, 2))
    monto = num(cell(ws, r, 9))
    if not f or not monto:
        continue
    gastos.append({
        "fecha": f,
        "cat": (cell(ws, r, 4) or "").strip() if isinstance(cell(ws, r, 4), str) else (cell(ws, r, 4) or ""),
        "desc": cell(ws, r, 5) or "",
        "medio": cell(ws, r, 6) or "",
        "tipo": cell(ws, r, 7) or "",
        "moneda": cell(ws, r, 8) or "ARS",
        "monto": monto,
    })

ingresos = []
ws = wb["Ingresos"]
for r in range(5, ws.max_row + 1):
    f = fecha_str(cell(ws, r, 2))
    monto = num(cell(ws, r, 7))
    if not f or not monto:
        continue
    ingresos.append({
        "fecha": f,
        "fuente": cell(ws, r, 4) or "",
        "desc": cell(ws, r, 5) or "",
        "moneda": cell(ws, r, 6) or "ARS",
        "monto": monto,
    })

inversiones = []
ws = wb["Inversiones"]
skip_re = re.compile(r"^──|^Subtotal|^RESUMEN|^Secci[oó]n|^Bybit|^Bull Market|^TOTAL", re.I)
for r in range(5, ws.max_row + 1):
    activo = cell(ws, r, 2)
    if not activo or skip_re.match(str(activo)):
        continue
    cantidad = num(cell(ws, r, 5))
    if not cantidad:
        continue
    inversiones.append({
        "activo": str(activo).strip(),
        "tipo": cell(ws, r, 3) or "",
        "cantidad": cantidad,
        "precioCompra": num(cell(ws, r, 6)),
        "precioActual": num(cell(ws, r, 7)),
    })

ws = wb["Config"]
usdt_ars = num(cell(ws, 5, 3)) or 1000
blue_rate = num(cell(ws, 6, 3)) or usdt_ars

data = {
    "gastos": gastos,
    "ingresos": ingresos,
    "inversiones": inversiones,
    "config": {"usdtArs": usdt_ars, "blueRate": blue_rate},
}

wb.close()
tmp_path.unlink(missing_ok=True)

html_path = RAIZ / "Mi Dashboard de Finanzas.html"
html = html_path.read_text(encoding="utf-8")
new_line = "var EMBEDDED_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";"
new_html, n = re.subn(r"var EMBEDDED_DATA = .*?;", lambda m: new_line, html, count=1)
if n == 0:
    print("No encontré 'var EMBEDDED_DATA' en el Dashboard — no se modificó nada.")
    sys.exit(1)

html_path.write_text(new_html, encoding="utf-8")
print(f"Dashboard actualizado con datos reales: {len(gastos)} gastos, {len(ingresos)} ingresos, {len(inversiones)} inversiones.")

# Copia para publicar (docs/index.html, versión iPhone/PWA). Si todavía no
# existe esa carpeta no pasa nada — este paso es opcional y no afecta al
# Dashboard principal.
docs_path = RAIZ / "docs" / "index.html"
if docs_path.exists():
    docs_html = docs_path.read_text(encoding="utf-8")
    new_docs_html, n2 = re.subn(r"var EMBEDDED_DATA = .*?;", lambda m: new_line, docs_html, count=1)
    if n2:
        docs_path.write_text(new_docs_html, encoding="utf-8")
        print("Copia para el iPhone (docs/index.html) también actualizada.")
print('Abrí "Mi Dashboard de Finanzas.html" haciendo doble clic para verlo con estos datos.')

"""
CRM de Finanzas Toto — servidor local.

Levanta una web en la propia compu para cargar gastos, ingresos, aportes,
posiciones de Bull Market, metas y ventas. Cada vez que se guarda algo:

  1. Se escribe en Finanzas_Toto.xlsx (con backup previo, ver crm_excel.py).
  2. Se regenera el Dashboard de la compu y la copia del iPhone.
  3. Se publica a Internet (git push) para que el iPhone lo vea, agrupando
     varios guardados seguidos en un solo commit.

Se sirve en 0.0.0.0 para poder entrar también desde el celular estando en la
misma WiFi. Por eso pide PIN.

Uso:  python crm_server.py        (o doble clic en "Abrir mi CRM.bat")
"""
import json
import os
import socket
import subprocess
import sys
import threading
from datetime import datetime
from functools import wraps
from pathlib import Path

BASE = Path(__file__).parent
RAIZ = BASE.parent


def ensure(pkg, import_name=None):
    try:
        __import__(import_name or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])


ensure("flask")
ensure("openpyxl")

from flask import Flask, jsonify, request, session, send_file  # noqa: E402

import crm_excel as XL  # noqa: E402
import crm_mercadopago as MP  # noqa: E402

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
CONFIG_PATH = BASE / "crm_config.json"
CONFIG_DEFECTO = {"pin": "1930", "puerto": 8765, "publicar_automatico": True,
                  "abrir_navegador": True}


def cargar_config():
    if CONFIG_PATH.exists():
        try:
            return {**CONFIG_DEFECTO, **json.loads(CONFIG_PATH.read_text(encoding="utf-8"))}
        except json.JSONDecodeError:
            pass
    CONFIG_PATH.write_text(json.dumps(CONFIG_DEFECTO, indent=2), encoding="utf-8")
    return dict(CONFIG_DEFECTO)


CONFIG = cargar_config()

app = Flask(__name__, static_folder=None)
# La clave de sesión se regenera en cada arranque: si se reinicia el servidor
# hay que volver a poner el PIN, que para un uso doméstico está bien.
app.secret_key = os.urandom(24)

# Todas las escrituras al Excel pasan por acá. El servidor es multi-hilo y dos
# guardados simultáneos (compu + celular) corromperían el archivo.
CANDADO_EXCEL = threading.Lock()


# ---------------------------------------------------------------------------
# Publicación automática
#
# Regenerar el Dashboard es instantáneo y se hace en cada guardado. Subir a
# Internet tarda unos segundos, así que se agenda con retraso: si se cargan
# cinco cosas seguidas, se sube una sola vez al final en un único commit.
# ---------------------------------------------------------------------------
RETRASO_PUBLICACION = 12  # segundos desde el último guardado


class Publicador:
    def __init__(self):
        self.estado = "al-dia"     # al-dia | pendiente | publicando | error
        self.mensaje = ""
        self.ultima = None
        self._timer = None
        self._lock = threading.Lock()

    def regenerar_dashboard(self):
        """Actualiza el Dashboard de la compu y la copia del iPhone."""
        r = subprocess.run(
            [sys.executable, str(BASE / "generar_datos_html.py")],
            capture_output=True, text=True, cwd=str(BASE),
        )
        if r.returncode != 0:
            raise RuntimeError((r.stderr or r.stdout or "").strip()[:300])

    def agendar(self):
        if not CONFIG.get("publicar_automatico", True):
            return
        with self._lock:
            if self._timer:
                self._timer.cancel()
            self.estado = "pendiente"
            self.mensaje = "Se sube a Internet en unos segundos…"
            self._timer = threading.Timer(RETRASO_PUBLICACION, self._publicar)
            self._timer.daemon = True
            self._timer.start()

    def publicar_ya(self):
        with self._lock:
            if self._timer:
                self._timer.cancel()
                self._timer = None
        self._publicar()

    def _git(self, *args):
        return subprocess.run(["git", *args], capture_output=True, text=True,
                              cwd=str(RAIZ), timeout=120)

    def _publicar(self):
        self.estado = "publicando"
        self.mensaje = "Subiendo a Internet…"
        try:
            self._git("add", "Mi Dashboard de Finanzas.html", "docs/index.html")
            commit = self._git("commit", "-m", "Actualizar datos del dashboard")
            if commit.returncode != 0 and "nothing to commit" not in (
                    commit.stdout + commit.stderr).lower():
                raise RuntimeError((commit.stderr or commit.stdout).strip()[:300])
            push = self._git("push")
            if push.returncode != 0:
                raise RuntimeError((push.stderr or push.stdout).strip()[:300])
            self.estado = "al-dia"
            self.ultima = datetime.now().strftime("%H:%M")
            self.mensaje = f"iPhone al día ({self.ultima})"
        except Exception as e:  # noqa: BLE001 — se muestra tal cual en la UI
            self.estado = "error"
            self.mensaje = f"No pude subir a Internet: {e}"


PUB = Publicador()


def tras_guardar():
    """Regenera el Dashboard y agenda la subida. Devuelve un aviso si falla."""
    try:
        PUB.regenerar_dashboard()
    except Exception as e:  # noqa: BLE001
        return f"Guardado en el Excel, pero no pude actualizar el Dashboard: {e}"
    PUB.agendar()
    return None


# ---------------------------------------------------------------------------
# Autenticación por PIN
# ---------------------------------------------------------------------------
def requiere_pin(fn):
    @wraps(fn)
    def envoltorio(*a, **kw):
        if not session.get("ok"):
            return jsonify({"error": "Necesitás poner el PIN."}), 401
        return fn(*a, **kw)
    return envoltorio


@app.post("/api/entrar")
def entrar():
    if str((request.json or {}).get("pin", "")).strip() == str(CONFIG["pin"]):
        session["ok"] = True
        session.permanent = True
        return jsonify({"ok": True})
    return jsonify({"error": "PIN incorrecto."}), 401


@app.post("/api/salir")
def salir():
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/sesion")
def sesion():
    return jsonify({"autenticado": bool(session.get("ok"))})


# ---------------------------------------------------------------------------
# Lectura
# ---------------------------------------------------------------------------
@app.get("/")
def inicio():
    return send_file(BASE / "crm_ui.html")


@app.get("/api/catalogos")
@requiere_pin
def catalogos():
    with CANDADO_EXCEL:
        return jsonify(XL.leer_catalogos())


@app.get("/api/estado")
@requiere_pin
def estado():
    with CANDADO_EXCEL:
        datos = XL.leer_estado()
    datos["publicacion"] = {"estado": PUB.estado, "mensaje": PUB.mensaje}
    return jsonify(datos)


@app.get("/api/publicacion")
@requiere_pin
def publicacion():
    return jsonify({"estado": PUB.estado, "mensaje": PUB.mensaje, "ultima": PUB.ultima})


@app.post("/api/publicar")
@requiere_pin
def publicar():
    threading.Thread(target=PUB.publicar_ya, daemon=True).start()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Escritura
#
# Todas las altas comparten la misma forma: escribir con el candado tomado,
# traducir los errores previsibles a un mensaje claro, y disparar la
# actualización del Dashboard.
# ---------------------------------------------------------------------------
def _guardar(fn, datos):
    try:
        with CANDADO_EXCEL:
            mensaje = fn(datos)
    except XL.ExcelAbierto as e:
        return jsonify({"error": str(e), "excelAbierto": True}), 409
    except XL.ErrorCRM as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"Error inesperado: {e}"}), 500
    aviso = tras_guardar()
    return jsonify({"ok": True, "mensaje": mensaje, "aviso": aviso})


@app.post("/api/gasto")
@requiere_pin
def api_gasto():
    return _guardar(XL.agregar_gasto, request.json or {})


@app.post("/api/ingreso")
@requiere_pin
def api_ingreso():
    return _guardar(XL.agregar_ingreso, request.json or {})


@app.post("/api/aporte")
@requiere_pin
def api_aporte():
    return _guardar(XL.agregar_aporte, request.json or {})


@app.post("/api/venta")
@requiere_pin
def api_venta():
    return _guardar(XL.agregar_venta, request.json or {})


@app.post("/api/meta")
@requiere_pin
def api_meta():
    return _guardar(XL.agregar_meta, request.json or {})


@app.post("/api/posiciones")
@requiere_pin
def api_posiciones():
    return _guardar(lambda d: XL.guardar_posiciones(d.get("posiciones", [])),
                    request.json or {})


@app.post("/api/cotizacion")
@requiere_pin
def api_cotizacion():
    return _guardar(lambda d: XL.guardar_cotizacion(d.get("usdtArs"), d.get("dolarBlue")),
                    request.json or {})


@app.post("/api/borrar")
@requiere_pin
def api_borrar():
    return _guardar(lambda d: XL.borrar_movimiento(d.get("hoja"), d.get("fila")),
                    request.json or {})


# ---------------------------------------------------------------------------
# MercadoPago: se sube el CSV, se muestra qué se va a cargar y recién cuando
# se confirma (asignando categoría a las transferencias desconocidas) se
# escribe al Excel.
# ---------------------------------------------------------------------------
@app.post("/api/mercadopago/analizar")
@requiere_pin
def mp_analizar():
    archivo = request.files.get("archivo")
    if not archivo:
        return jsonify({"error": "No llegó ningún archivo."}), 400
    try:
        with CANDADO_EXCEL:
            return jsonify(MP.analizar(archivo.stream.read()))
    except MP.ErrorMP as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"No pude leer el CSV: {e}"}), 400


@app.post("/api/mercadopago/confirmar")
@requiere_pin
def mp_confirmar():
    datos = request.json or {}
    try:
        with CANDADO_EXCEL:
            resultado = MP.confirmar(datos.get("filas", []),
                                     datos.get("recordarReglas", True))
    except XL.ExcelAbierto as e:
        return jsonify({"error": str(e), "excelAbierto": True}), 409
    except (XL.ErrorCRM, MP.ErrorMP) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"Error inesperado: {e}"}), 500
    aviso = tras_guardar()
    resultado["aviso"] = aviso
    resultado["ok"] = True
    return jsonify(resultado)


# ---------------------------------------------------------------------------
# Arranque
# ---------------------------------------------------------------------------
def puerto_ocupado(puerto):
    """True si ya hay algo escuchando ahí (típicamente, otro CRM ya abierto)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(0.6)
    try:
        return s.connect_ex(("127.0.0.1", puerto)) == 0
    finally:
        s.close()


def ip_local():
    """IP de esta compu en la red de casa, para entrar desde el celular."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))  # no manda tráfico, solo elige la interfaz
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


def main():
    # La consola de Windows no viene en UTF-8: sin esto los acentos de los
    # mensajes de abajo salen como símbolos raros.
    for flujo in (sys.stdout, sys.stderr):
        if hasattr(flujo, "reconfigure"):
            try:
                flujo.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass

    puerto = int(CONFIG.get("puerto", 8765))

    if puerto_ocupado(puerto):
        print("=" * 62)
        print("  El CRM ya está abierto")
        print("=" * 62)
        print(f"  Ya hay un CRM andando en el puerto {puerto}.")
        print(f"  Abrilo en el navegador:  http://localhost:{puerto}")
        print()
        print("  (Si creés que no debería estar abierto, cerrá la otra ventana")
        print("   negra del CRM y volvé a intentar.)")
        print("=" * 62)
        return

    ip = ip_local()
    print("=" * 62)
    print("  CRM de Finanzas — andando")
    print("=" * 62)
    print(f"  En esta compu:   http://localhost:{puerto}")
    print(f"  Desde el iPhone: http://{ip}:{puerto}   (misma WiFi)")
    print(f"  PIN: {CONFIG['pin']}    (se cambia en Programas/crm_config.json)")
    print()
    print("  Dejá esta ventana abierta mientras uses el CRM.")
    print("  Para cerrarlo: cerrá esta ventana o apretá Ctrl+C.")
    print("=" * 62)
    if not XL.XLSX.exists():
        print(f"\n  OJO: no encuentro {XL.XLSX}")

    # Se abre el navegador con un respiro para que el servidor ya esté
    # escuchando y no muestre un error de conexión.
    if CONFIG.get("abrir_navegador", True):
        import webbrowser
        threading.Timer(
            1.5, lambda: webbrowser.open(f"http://localhost:{puerto}")
        ).start()

    app.run(host="0.0.0.0", port=puerto, threaded=True, debug=False)


if __name__ == "__main__":
    main()

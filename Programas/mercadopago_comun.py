"""
Lógica compartida para procesar los CSV de MercadoPago.

La usan los dos caminos de carga:
  - mercadopago_processor.py  (el .bat de siempre, por consola)
  - crm_mercadopago.py        (el CRM, por el navegador)

Vive acá para que las reglas de clasificación sean las mismas en los dos
lados: si se agrega una regla nueva, la toman ambos.
"""
import json
import re
from pathlib import Path

BASE = Path(__file__).parent
REGLAS_PATH = BASE / "mis_reglas.json"
PRIVADO_PATH = BASE / "mis_reglas_privado.json"


def cargar_reglas():
    """Combina las reglas públicas con las privadas.

    Las reglas con nombres reales de personas viven aparte en
    mis_reglas_privado.json (fuera de git, ver .gitignore) porque el repo puede
    ser público. Se devuelven las dos cosas: el combinado para clasificar, y el
    privado suelto porque las categorías que se cargan a mano se guardan
    únicamente ahí, nunca en el público.
    """
    with open(REGLAS_PATH, encoding="utf-8") as f:
        reglas = json.load(f)

    if PRIVADO_PATH.exists():
        with open(PRIVADO_PATH, encoding="utf-8") as f:
            privadas = json.load(f)
    else:
        privadas = {"reglas_gastos": [], "reglas_ingresos": [], "excluir": []}

    reglas["reglas_gastos"] = privadas.get("reglas_gastos", []) + reglas.get("reglas_gastos", [])
    reglas["reglas_ingresos"] = privadas.get("reglas_ingresos", []) + reglas.get("reglas_ingresos", [])
    reglas["excluir"] = privadas.get("excluir", []) + reglas.get("excluir", [])
    return reglas, privadas


def clasificar(desc, amt, reglas):
    """Decide qué es cada movimiento del CSV.

    Devuelve (tipo, fuente, categoria, nota) donde tipo es
    INGRESO / GASTO / REVISAR, o None cuando hay que excluirlo.
    """
    d = desc.lower()
    for ex in reglas["excluir"]:
        if ex["match"] in d:
            return None, None, None, f"EXCLUIDO: {ex['razon']}"
    for r in reglas["reglas_ingresos"]:
        if r["match"] in d:
            return "INGRESO", r["fuente"], None, None
    if "transferencia recibida" in d and amt > 0:
        nombre = desc.replace("Transferencia recibida", "").strip()
        return "INGRESO", f"Transf: {nombre}", None, None
    for r in reglas["reglas_gastos"]:
        if r["match"] in d:
            return "GASTO", r["descripcion"], r["categoria"], None
    if "transferencia enviada" in d:
        return "REVISAR", desc, "?", "Transferencia a persona nueva — necesita categorización manual"
    return "GASTO", desc, reglas["categoria_default_gasto"], None


def normalizar_nombre(desc):
    """Deja el nombre de la persona/comercio sin el prefijo de transferencia."""
    return re.sub(r"^transferencia (enviada|recibida)\s*", "", desc, flags=re.I).strip().lower()


def guardar_regla_privada(privadas, descripcion, categoria):
    """Aprende una categoría nueva para que la próxima vez se cargue sola.

    Se escribe sólo en mis_reglas_privado.json para no filtrar nombres de
    personas al repositorio.
    """
    nombre = normalizar_nombre(descripcion)
    privadas.setdefault("reglas_gastos", []).insert(0, {
        "match": nombre, "categoria": categoria, "descripcion": descripcion,
    })
    with open(PRIVADO_PATH, "w", encoding="utf-8") as f:
        json.dump(privadas, f, ensure_ascii=False, indent=2)
    return nombre


def es_duplicado(existentes, fecha_str, descripcion):
    """True cuando ya hay un gasto de la misma fecha y misma persona/descripción.

    El monto NO participa de la comparación, para evitar falsos duplicados
    cuando una misma transferencia cambia de importe o se corrige en un
    reporte nuevo.
    """
    if not fecha_str or not descripcion:
        return False
    return (str(fecha_str)[:10], (descripcion or "").strip().lower()) in existentes


def leer_csv(origen):
    """Lee el CSV de MercadoPago (acepta una ruta o los bytes del archivo)."""
    import io

    import pandas as pd

    if isinstance(origen, (bytes, bytearray)):
        origen = io.BytesIO(origen)
    df = pd.read_csv(origen, sep=";", skiprows=3, decimal=",", thousands=".")
    faltantes = {"TRANSACTION_TYPE", "RELEASE_DATE", "TRANSACTION_NET_AMOUNT"} - set(df.columns)
    if faltantes:
        raise ValueError(
            "El archivo no parece un reporte de MercadoPago "
            f"(le faltan las columnas: {', '.join(sorted(faltantes))})."
        )
    df["desc"] = df["TRANSACTION_TYPE"].str.strip()
    df["fecha"] = pd.to_datetime(df["RELEASE_DATE"], format="%d-%m-%Y")
    df["amt"] = df["TRANSACTION_NET_AMOUNT"]
    return df

"""
Importación de MercadoPago desde el CRM (navegador).

A diferencia del .bat de consola, acá el proceso tiene dos pasos: primero se
analiza el CSV y se muestra en pantalla qué se va a cargar, y recién cuando el
usuario confirma (asignando categoría a las transferencias que ninguna regla
reconoce) se escribe al Excel.
"""
from datetime import date, datetime

import crm_excel as XL
import mercadopago_comun as MPC


class ErrorMP(Exception):
    """Error previsible al importar, para mostrar tal cual en pantalla."""


def _iso(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v)[:10]


def analizar(csv_bytes):
    """Lee el CSV y arma la vista previa, sin tocar el Excel."""
    try:
        df = MPC.leer_csv(csv_bytes)
    except ValueError as e:
        raise ErrorMP(str(e))
    except Exception as e:  # noqa: BLE001
        raise ErrorMP(f"No pude leer el CSV: {e}")

    reglas, _privadas = MPC.cargar_reglas()

    # Para detectar duplicados hace falta lo que ya está cargado en el Excel.
    estado = XL.leer_estado()
    existentes = {(g["fecha"], (g["descripcion"] or "").strip().lower())
                  for g in estado["gastos"]}
    por_fecha_monto = {}
    for g in estado["gastos"]:
        por_fecha_monto.setdefault((g["fecha"], abs(g["monto"] or 0)), []).append(
            g["descripcion"])

    filas, ingresos_ignorados, excluidos = [], 0, 0
    for _, r in df.iterrows():
        tipo, fuente, categoria, _nota = MPC.clasificar(r["desc"], r["amt"], reglas)
        if tipo is None:
            excluidos += 1
            continue
        if tipo == "INGRESO":
            # Pedido explícito de Toto: la importación de MercadoPago nunca
            # escribe en la hoja Ingresos. Se cuentan sólo para informar.
            ingresos_ignorados += 1
            continue

        fecha = _iso(r["fecha"])
        descripcion = str(fuente)[:100].strip()
        monto = abs(float(r["amt"]))
        duplicado = MPC.es_duplicado(existentes, fecha, descripcion)
        parecidos = por_fecha_monto.get((fecha, monto), [])

        filas.append({
            "fecha": fecha,
            "descripcion": descripcion,
            "monto": monto,
            "categoria": None if tipo == "REVISAR" else categoria,
            "necesitaCategoria": tipo == "REVISAR",
            "duplicado": duplicado,
            "parecidos": [p for p in parecidos if p and p.strip().lower() != descripcion.lower()],
            # Los duplicados vienen destildados: se cargan sólo si el usuario
            # los tilda a mano.
            "cargar": not duplicado,
        })

    return {
        "filas": filas,
        "resumen": {
            "total": len(filas),
            "nuevos": sum(1 for f in filas if not f["duplicado"]),
            "duplicados": sum(1 for f in filas if f["duplicado"]),
            "aRevisar": sum(1 for f in filas if f["necesitaCategoria"]),
            "ingresosIgnorados": ingresos_ignorados,
            "excluidos": excluidos,
        },
    }


def confirmar(filas, recordar_reglas=True):
    """Escribe al Excel las filas tildadas y aprende las categorías nuevas."""
    elegidas = [f for f in filas if f.get("cargar")]
    if not elegidas:
        raise ErrorMP("No tildaste ninguna fila para cargar.")

    sin_categoria = [f for f in elegidas if not (f.get("categoria") or "").strip()]
    if sin_categoria:
        raise ErrorMP(
            f"Faltan categorías: {len(sin_categoria)} movimiento(s) sin categoría asignada."
        )

    _reglas, privadas = MPC.cargar_reglas()
    aprendidas = []
    if recordar_reglas:
        vistos = set()
        for f in elegidas:
            if not f.get("necesitaCategoria"):
                continue
            nombre = MPC.normalizar_nombre(f["descripcion"])
            categoria = f["categoria"].strip()
            if nombre in vistos or categoria == "?":
                continue
            MPC.guardar_regla_privada(privadas, f["descripcion"], categoria)
            vistos.add(nombre)
            aprendidas.append({"nombre": f["descripcion"], "categoria": categoria})

    mensaje = XL.agregar_gastos([{
        "fecha": f["fecha"],
        "categoria": f["categoria"],
        "descripcion": f["descripcion"],
        "medio": "MercadoPago",
        "tipo": "Variable",
        "moneda": "ARS",
        "monto": f["monto"],
    } for f in elegidas])

    return {"mensaje": mensaje, "cargados": len(elegidas), "aprendidas": aprendidas}

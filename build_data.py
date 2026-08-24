#!/usr/bin/env python3
"""
build_data.py — Convierte la base de datos de encuestas de Auxilio24 (Excel)
en el archivo data.json que consume el panel web (index.html).

Uso:
    python3 build_data.py BBDD_AA24.xlsx

Requiere: openpyxl (pip install openpyxl --break-system-packages)

Columnas esperadas en la hoja "Reporte Metricas Encuestas" (o la primera hoja):
    A  Mes
    B  IdOrdenDeServicio
    C  CIA
    D  Campaña
    E  FechaDeEncuesta
    F  TipoDePoliza
    G  FechaDeServicioPrestado
    H  NroPoliza
    I  Pregunta: Atención (1-5)
    J  Pregunta: Servicio de la grúa (1-5)
    K  Pregunta: Calidad general (1-5)
    L  Tiempo de arribo de la grúa
    M  Comentario libre (texto)
    N  RESPONDIO? (Responde / No Responde)
    O  Tipo de comentario (Comentarios Positivos / Aspectos de Mejora / Comentario Mixto)
    P  Servicio
    Q  Problema
    R  Provincia
    S  Tipo de Vehículo

Si en el futuro cambia el orden de columnas, ajustar el diccionario COL más abajo.
"""
import sys
import json
import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Falta openpyxl. Instalalo con: pip install openpyxl --break-system-packages")

# Índices de columna (base 0)
COL = {
    "mes": 0,
    "id_orden": 1,
    "cia": 2,
    "campana": 3,
    "atencion": 8,
    "grua": 9,
    "calidad": 10,
    "tiempo_arribo": 11,
    "comentario": 12,
    "respondio": 13,
    "tipo_comentario": 14,
    "servicio": 15,
    "problema": 16,
    "provincia": 17,
    "tipo_vehiculo": 18,
}

MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

CATEGORIAS_COMENTARIO = {
    "Comentarios Positivos": "Comentarios positivos",
    "Aspectos de Mejora": "Aspectos de mejora",
    "Comentario Mixto": "Comentarios mixtos",
}


def clean(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    return v


def mes_label(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return f"{MESES_ES[v.month - 1]} {v.year}"
    return clean(v)


def main():
    if len(sys.argv) < 2:
        sys.exit("Uso: python3 build_data.py <archivo.xlsx>")

    path = Path(sys.argv[1])
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]

    casos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[COL["id_orden"]] is None:
            continue

        atencion = row[COL["atencion"]]
        grua = row[COL["grua"]]
        calidad = row[COL["calidad"]]
        tiempo_arribo = clean(row[COL["tiempo_arribo"]])
        comentario_txt = clean(row[COL["comentario"]])
        respondio = clean(row[COL["respondio"]])
        tipo_comentario_raw = clean(row[COL["tipo_comentario"]])

        caso = {
            "mes": mes_label(row[COL["mes"]]),
            "id": row[COL["id_orden"]],
            "atencion": atencion if isinstance(atencion, (int, float)) else None,
            "grua": grua if isinstance(grua, (int, float)) else None,
            "calidad": calidad if isinstance(calidad, (int, float)) else None,
            "tiempoArribo": tiempo_arribo or None,
            "servicio": clean(row[COL["servicio"]]) or None,
            "problema": clean(row[COL["problema"]]) or None,
            "provincia": clean(row[COL["provincia"]]) or None,
            "tipoVehiculo": clean(row[COL["tipo_vehiculo"]]) or None,
            "respondio": respondio,
        }

        # Comentario: solo lo incluimos si respondió, hay categoría reconocida
        # y el texto no es un placeholder vacío tipo "Sin comentarios"
        if (
            respondio == "Responde"
            and tipo_comentario_raw in CATEGORIAS_COMENTARIO
            and comentario_txt
            and comentario_txt.lower() not in ("sin comentarios", "sin comentario", "-", "n/a")
        ):
            caso["tipoComentario"] = CATEGORIAS_COMENTARIO[tipo_comentario_raw]
            caso["comentario"] = comentario_txt

        casos.append(caso)

    data = {
        "actualizado": datetime.date.today().isoformat(),
        "cliente": "Auxilio24",
        "casos": casos,
    }

    out_path = path.parent / "data.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK -> {out_path} ({len(casos)} casos)")


if __name__ == "__main__":
    main()

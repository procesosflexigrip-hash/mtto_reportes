"""
Arma un Excel con los indicadores del supervisor, una fila por semana.

Solo entran los CSV de captura digital: nombre MANT_AAAA-Wss.csv
desde 2026-W41. Antes de esa semana seguían aceptándose formatos en
papel, así que los exportados por fecha no se leen.

Uso:
    python indicadores_por_semana.py
"""
import re
from pathlib import Path

import pandas as pd

CARPETA = Path(__file__).resolve().parent
DESDE = (2026, 41)
PATRON = re.compile(r"^MANT_(\d{4})-W(\d{2})\.csv$", re.IGNORECASE)
SALIDA = CARPETA / "indicadores_por_semana.xlsx"
EXCLUIR_MAQUINA = {"SERVICIOS", "NO APLICA"}

COLUMNAS = [
    "Semana",
    "Total",
    "Esperando asignar",
    "Por confirmar",
    "En reparación",
    "Pausados",
    "Cerrados",
    "Liberados",
    "Espera promedio (h)",
    "Reparación promedio (h)",
]


def semana_de_nombre(path):
    coincidencia = PATRON.match(path.name)
    if not coincidencia:
        return None
    year = int(coincidencia.group(1))
    week = int(coincidencia.group(2))
    if (year, week) < DESDE:
        return None
    return f"{year}-W{week:02d}", year, week


def texto(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    s = str(valor).strip()
    if s.lower() == "nan":
        return ""
    return s


def maquina_normalizada(valor):
    s = texto(valor).upper()
    return s.translate(str.maketrans("ÁÉÍÓÚÜ", "AEIOUU"))


def numero(valor):
    s = texto(valor).replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def columna_que_empieza(df, prefijo):
    for col in df.columns:
        if str(col).startswith(prefijo):
            return col
    return None


def promedio(valores):
    if not valores:
        return 0.0
    return round(sum(valores) / len(valores), 1)


def indicadores(df):
    estados = [texto(v) for v in df["ESTADO"]] if "ESTADO" in df.columns else [""] * len(df)
    maquinas = [maquina_normalizada(v) for v in df["MAQUINA"]] if "MAQUINA" in df.columns else [""] * len(df)

    def contar(nombre):
        return sum(1 for estado in estados if estado == nombre)

    col_espera = columna_que_empieza(df, "TIEMPO ESPERA (h)")
    col_reparacion = columna_que_empieza(df, "TIEMPO REPARACION (h)")
    espera = []
    reparacion = []
    for i, maquina in enumerate(maquinas):
        if maquina in EXCLUIR_MAQUINA:
            continue
        if col_espera:
            horas = numero(df.iloc[i][col_espera])
            if horas is not None:
                espera.append(horas)
        if col_reparacion and estados[i] in ("Cerrado", "Liberado"):
            horas = numero(df.iloc[i][col_reparacion])
            if horas is not None:
                reparacion.append(horas)

    return {
        "Total": len(df),
        "Esperando asignar": contar("Pendiente"),
        "Por confirmar": contar("Asignado"),
        "En reparación": contar("En reparación"),
        "Pausados": contar("Pausado"),
        "Cerrados": contar("Cerrado"),
        "Liberados": contar("Liberado"),
        "Espera promedio (h)": promedio(espera),
        "Reparación promedio (h)": promedio(reparacion),
    }


def csv_digitales():
    encontrados = []
    for path in CARPETA.glob("MANT_*.csv"):
        semana = semana_de_nombre(path)
        if semana:
            encontrados.append((semana, path))
    encontrados.sort(key=lambda item: (item[0][1], item[0][2]))
    return encontrados


def main():
    archivos = csv_digitales()
    if not archivos:
        raise SystemExit("No hay CSV MANT_AAAA-Wss.csv desde 2026-W41 en " + str(CARPETA))

    filas = []
    for (semana, _year, _week), path in archivos:
        df = pd.read_csv(path, encoding="utf-8-sig", dtype=str, keep_default_na=False)
        fila = {"Semana": semana}
        fila.update(indicadores(df))
        filas.append(fila)
        print(f"{path.name}: {len(df)} servicios -> {semana}")

    pd.DataFrame(filas, columns=COLUMNAS).to_excel(SALIDA, index=False, sheet_name="Indicadores")
    print(f"Escrito {SALIDA.name}: {len(filas)} semana(s)")


if __name__ == "__main__":
    main()

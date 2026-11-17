"""Motor de la Revisión de Abastecimiento CD Coquimbo.

Cruza el modelo de abastecimiento (Excel ABASTECIMIENTO_x.xlsx: stock CD, stock CL, forecast
del stock 7 hrs y masivos) con la venta real del CD (exporte SAP 951) para decidir qué pallets
del masivo viajan primero, cuáles se postergan y qué sobra en el CD.

No depende de Streamlit: se puede probar o automatizar por separado.
"""
import io
import math

import numpy as np
import pandas as pd

# ------------------------------------------------------------------ parámetros por defecto
PARAMS = {
    "semanas_venta": 4,        # ventana de venta real usada para la demanda
    "usar_tendencia": True,    # demanda = mayor entre promedio del período y promedio última semana
    "dias_objetivo": 5,        # días venta que se quieren tener en el CD (igual al modelo)
    "umbral_urgente": 2.0,     # cobertura (días) bajo la cual el producto es urgente
    "umbral_exceso": 15.0,     # cobertura (días) sobre la cual sobra stock en el CD
    "dias_sin_venta": [6],     # domingo (0 = lunes ... 6 = domingo)
}

PRIORIDADES = ["1-URGENTE", "2-NORMAL", "3-POSTERGAR", "4-NO TRAER"]
HOJAS_MASIVO = ["MASIVO GM", "MASIVO CVP", "MASIVO TRANSITO GM", "MASIVO TRANSITO CVP"]
COLS_MASIVO = ["GRUPO", "ORIGEN", "DESTINO", "FECHA", "CODIGO SAP", "CANTIDAD CAJAS", "PALLETS"]
COLS_VENTA_951 = ["fecha contabiliz.", "material", "cantidad"]


# ------------------------------------------------------------------ utilidades
def _cod(s: pd.Series) -> pd.Series:
    """Código como texto limpio (sin .0 de Excel ni espacios)."""
    return (s.astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .replace({"nan": "", "None": ""}))


def _num(s: pd.Series) -> pd.Series:
    """Número desde texto en formato chileno o normal; 'No encontrado' -> 0."""
    txt = s.astype(str).str.strip()
    chileno = txt.str.contains(",", regex=False) | txt.str.fullmatch(r"-?\d{1,3}(\.\d{3})+")
    txt = txt.where(~chileno, txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(txt, errors="coerce").fillna(0)


def leer_tabla(contenido: bytes, nombre: str, columnas=None) -> pd.DataFrame:
    """CSV (separador ; o ,) o Excel, columnas en minúsculas.
    columnas: si se indica, solo se cargan esas (ahorra memoria con archivos de venta grandes)."""
    usar = (lambda c: str(c).strip().lower() in columnas) if columnas else None
    if nombre.lower().endswith(".csv"):
        primera = contenido[:2000].split(b"\n", 1)[0]
        sep = ";" if primera.count(b";") > primera.count(b",") else ","
        df = None
        for cod in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(contenido), sep=sep, dtype=str, encoding=cod, usecols=usar)
                break
            except UnicodeDecodeError:
                continue
    else:
        df = pd.read_excel(io.BytesIO(contenido), dtype=str, usecols=usar)
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def _fecha(txt: pd.Series) -> pd.Series:
    """Fechas de SAP/Excel: prueba formatos fijos (rápido) y al final deja que pandas adivine."""
    txt = txt.astype(str).str.strip().str.slice(0, 19)
    out = pd.Series(pd.NaT, index=txt.index)
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y%m%d"):
        falta = out.isna()
        if not falta.any():
            break
        out[falta] = pd.to_datetime(txt[falta], format=fmt, errors="coerce")
    falta = out.isna()
    if falta.any():
        out[falta] = pd.to_datetime(txt[falta], dayfirst=True, errors="coerce")
    return out


# ------------------------------------------------------------------ modelo de abastecimiento
def leer_modelo(contenido: bytes) -> dict:
    """Lee el Excel del modelo (hojas ABASTECIMIENTO, Stock a Mantener, DICCIONARIO y masivos)."""
    xls = pd.ExcelFile(io.BytesIO(contenido))
    hojas = {h.strip().upper(): h for h in xls.sheet_names}
    if "ABASTECIMIENTO" not in hojas:
        raise ValueError("El archivo del modelo no tiene la hoja ABASTECIMIENTO.")

    a = pd.read_excel(xls, hojas["ABASTECIMIENTO"], dtype=str)
    a.columns = [str(c).strip().upper() for c in a.columns]
    faltan = [c for c in ("PRODUCTO", "SAP", "DESCRIPCION", "CAJAS_X_PALLET", "STOCK CD",
                          "STOCK CL CAJAS") if c not in a.columns]
    if faltan:
        raise ValueError(f"A la hoja ABASTECIMIENTO le faltan columnas: {', '.join(faltan)}.")
    m = pd.DataFrame({
        "producto": _cod(a["PRODUCTO"]),
        "sap": _cod(a["SAP"]),
        "descripcion": a["DESCRIPCION"].astype(str).str.strip(),
        "cajas_x_pallet": _num(a["CAJAS_X_PALLET"]),
        "stock_cd": _num(a["STOCK CD"]),
        "stock_cl": _num(a["STOCK CL CAJAS"]),
        "transito": _num(a["TRANSITO"]) if "TRANSITO" in a.columns else 0.0,
        "pallets_exceso_modelo": _num(a["TRASLADAR CD - CL (PALLET)"])
        if "TRASLADAR CD - CL (PALLET)" in a.columns else 0.0,
    })
    m = m[m["sap"] != ""].drop_duplicates("sap").set_index("sap")

    # forecast y revisión manual
    m["forecast_dia"] = 0.0
    m["revision_pallets"] = 0.0
    if "STOCK A MANTENER" in hojas:
        s = pd.read_excel(xls, hojas["STOCK A MANTENER"], dtype=str)
        s.columns = [str(c).strip().upper() for c in s.columns]
        s["SAP"] = _cod(s["SAP"])
        s = s.drop_duplicates("SAP").set_index("SAP")
        if "FORECAST POR DIA" in s.columns:
            m["forecast_dia"] = _num(s["FORECAST POR DIA"]).reindex(m.index).fillna(0)
    if "REVISION" in hojas:
        r = pd.read_excel(xls, hojas["REVISION"], dtype=str)
        r.columns = [str(c).strip().upper() for c in r.columns]
        if {"SAP", "REVISION PALLETS"} <= set(r.columns):
            r["SAP"] = _cod(r["SAP"])
            rv = _num(r.drop_duplicates("SAP").set_index("SAP")["REVISION PALLETS"])
            m["revision_pallets"] = rv.reindex(m.index).fillna(0)

    # pasillo
    m["pasillo"] = ""
    if "DICCIONARIO" in hojas:
        d = pd.read_excel(xls, hojas["DICCIONARIO"], dtype=str)
        d.columns = [str(c).strip().upper() for c in d.columns]
        if {"CODIGO", "PASILLO ACTUAL"} <= set(d.columns):
            d["CODIGO"] = _cod(d["CODIGO"])
            pas = d.drop_duplicates("CODIGO").set_index("CODIGO")["PASILLO ACTUAL"]
            m["pasillo"] = m["producto"].map(pas).fillna("")
    m.loc[m["cajas_x_pallet"] <= 0, "cajas_x_pallet"] = np.nan

    # masivos
    filas = []
    for h in HOJAS_MASIVO:
        if h not in hojas:
            continue
        x = pd.read_excel(xls, hojas[h], header=None, dtype=str)
        if x.empty:
            continue
        if str(x.iat[0, 0]).strip().upper() == "GRUPO":
            x = x.iloc[1:]
        x = x.iloc[:, :8]
        x.columns = COLS_MASIVO + (["PASILLO ACTUAL"] if x.shape[1] == 8 else [])
        x["CODIGO SAP"] = _cod(x["CODIGO SAP"])
        x = x[x["CODIGO SAP"] != ""]
        x["HOJA"] = h
        filas.append(x)
    # plantilla de la hoja de traslado CD - CL (GRUPO, ORIGEN, DESTINO tal como los usa el modelo)
    plantilla = {"GRUPO": 1, "ORIGEN": "COQUIMEXT", "DESTINO": "COQUIMBO"}
    if "CD - CL" in hojas:
        t = pd.read_excel(xls, hojas["CD - CL"], header=None, dtype=str, nrows=3)
        if len(t) and str(t.iat[0, 0]).strip().upper() == "GRUPO":
            t = t.iloc[1:]
        if len(t) and t.shape[1] >= 3 and str(t.iat[0, 1]) not in ("nan", ""):
            plantilla = {"GRUPO": int(_num(t.iloc[:1, 0]).iat[0]) or 1,
                         "ORIGEN": str(t.iat[0, 1]).strip(), "DESTINO": str(t.iat[0, 2]).strip()}
    if not filas:
        raise ValueError("No encontré líneas en las hojas MASIVO GM / MASIVO CVP.")
    mas = pd.concat(filas, ignore_index=True)
    mas["CANTIDAD CAJAS"] = _num(mas["CANTIDAD CAJAS"])
    mas["PALLETS"] = _num(mas["PALLETS"]).astype(int)
    mas["GRUPO"] = _num(mas["GRUPO"]).astype(int)
    mas["FECHA"] = _cod(mas["FECHA"])
    return {"modelo": m, "masivo": mas, "plantilla_cd_cl": plantilla}


# ------------------------------------------------------------------ venta real
def leer_venta(archivos: list[tuple[str, bytes]]) -> pd.DataFrame:
    """Une uno o más exportes de venta. Formato SAP 951 (Fecha contabiliz., Material, Cantidad
    negativa = salida) o formato simple (FECHA, CODIGO, CAJAS)."""
    partes = []
    for nombre, contenido in archivos:
        v = leer_tabla(contenido, nombre, columnas=set(COLS_VENTA_951) | {
            "documento material", "fecha", "codigo", "cajas"})
        if all(c in v.columns for c in COLS_VENTA_951):
            cajas = -_num(v["cantidad"])
            fecha_txt, cod = v["fecha contabiliz."], v["material"]
        elif {"fecha", "codigo", "cajas"} <= set(v.columns):
            cajas = _num(v["cajas"])
            fecha_txt, cod = v["fecha"], v["codigo"]
        else:
            raise ValueError(f"{nombre}: no reconozco el formato. Debe traer Fecha contabiliz., "
                             "Material y Cantidad (exporte 951) o FECHA, CODIGO y CAJAS.")
        fecha = _fecha(fecha_txt)
        d = pd.DataFrame({"fecha": fecha.dt.normalize(), "material": _cod(cod), "cajas": cajas})
        if "documento material" in v.columns:
            d["doc"] = v["documento material"].astype(str)
        d["archivo"] = nombre
        partes.append(d)
    v = pd.concat(partes, ignore_index=True)
    v = v[v["fecha"].notna() & (v["cajas"] > 0) & (v["material"] != "")]
    # si dos archivos traen el mismo período, no contar dos veces
    claves = [c for c in ("fecha", "material", "cajas", "doc") if c in v.columns]
    if v["archivo"].nunique() > 1 and "doc" in v.columns:
        v = v.drop_duplicates(claves)
    return v.drop(columns="archivo")


def mapear_codigos(venta: pd.DataFrame, modelo: pd.DataFrame) -> pd.DataFrame:
    """Lleva el código de la venta al SAP del modelo (acepta SAP, código producto o con ceros)."""
    mapa = {}
    for sap, prod in zip(modelo.index, modelo["producto"]):
        for k in (prod, prod.lstrip("0"), sap, sap.lstrip("0")):
            if k:
                mapa.setdefault(k, sap)
    for sap in modelo.index:  # el SAP exacto manda sobre el código producto
        mapa[sap] = sap
    v = venta.copy()
    v["sap"] = v["material"].map(mapa)
    falta = v["sap"].isna()
    v.loc[falta, "sap"] = v.loc[falta, "material"].str.lstrip("0").map(mapa)
    v["en_modelo"] = v["sap"].notna()
    v["sap"] = v["sap"].fillna(v["material"])
    return v


def demanda(venta: pd.DataFrame, p: dict) -> tuple[pd.DataFrame, dict]:
    """Venta diaria por SKU en la ventana: promedio del período y de la última semana operativa."""
    fin = venta["fecha"].max()
    ini = fin - pd.Timedelta(weeks=p["semanas_venta"]) + pd.Timedelta(days=1)
    v = venta[(venta["fecha"] >= ini) & (~venta["fecha"].dt.weekday.isin(p["dias_sin_venta"]))]
    dias = sorted(v["fecha"].unique())
    n = max(1, len(dias))
    ult = dias[-6:] if len(dias) >= 6 else dias
    tot = v.groupby("sap")["cajas"].sum()
    sem = v[v["fecha"].isin(ult)].groupby("sap")["cajas"].sum() / max(1, len(ult))
    dias_con = v.groupby("sap")["fecha"].nunique()
    d = pd.DataFrame({"venta_periodo": tot, "venta_dia_prom": tot / n,
                      "venta_dia_ult_semana": sem, "dias_con_venta": dias_con}).fillna(0)
    d["pct_dias_con_venta"] = d["dias_con_venta"] / n
    d["demanda_dia"] = (np.maximum(d["venta_dia_prom"], d["venta_dia_ult_semana"])
                        if p["usar_tendencia"] else d["venta_dia_prom"])
    en_modelo = venta.groupby("sap")["en_modelo"].first()
    d["en_modelo"] = en_modelo.reindex(d.index).fillna(False)
    info = {"desde": pd.Timestamp(dias[0]) if dias else ini, "hasta": fin, "dias_operativos": n,
            "skus_con_venta": int((d["venta_periodo"] > 0).sum()),
            "skus_fuera_modelo": int((~d["en_modelo"].astype(bool)).sum())}
    return d, info


# ------------------------------------------------------------------ cálculo
def _base(modelo: pd.DataFrame, dem: pd.DataFrame, p: dict) -> pd.DataFrame:
    b = modelo.join(dem, how="left")
    for c in ("venta_periodo", "venta_dia_prom", "venta_dia_ult_semana", "demanda_dia",
              "dias_con_venta", "pct_dias_con_venta"):
        b[c] = b[c].fillna(0)
    b["cobertura_dias"] = np.where(b["demanda_dia"] > 0, b["stock_cd"] / b["demanda_dia"],
                                   np.where(b["stock_cd"] > 0, 999, 0))
    b["real_vs_forecast"] = np.where(b["forecast_dia"] > 0, b["demanda_dia"] / b["forecast_dia"], np.nan)
    nec = (b["demanda_dia"] * p["dias_objetivo"] - b["stock_cd"]).clip(lower=0)
    b["pallets_necesarios"] = np.ceil(nec / b["cajas_x_pallet"]).fillna(0)
    b["pallets_sobrantes"] = np.floor(((b["stock_cd"] - b["demanda_dia"] * p["dias_objetivo"])
                                       .clip(lower=0)) / b["cajas_x_pallet"]).fillna(0)
    return b


def priorizar_masivo(m: dict, dem: pd.DataFrame, p: dict) -> pd.DataFrame:
    """Cada línea del masivo con su prioridad y los pallets que conviene traer hoy."""
    b = _base(m["modelo"], dem, p)
    mas = m["masivo"].copy()
    # orden de asignación: primero masivo normal, después tránsito
    mas["orden_hoja"] = mas["HOJA"].map({h: i for i, h in enumerate(HOJAS_MASIVO)})
    mas = mas.sort_values(["CODIGO SAP", "orden_hoja"])
    cols = ["descripcion", "pasillo", "cajas_x_pallet", "stock_cd", "stock_cl", "forecast_dia",
            "revision_pallets", "venta_dia_prom", "venta_dia_ult_semana", "demanda_dia",
            "real_vs_forecast", "cobertura_dias", "pallets_necesarios"]
    mas = mas.join(b[cols], on="CODIGO SAP")
    mas["descripcion"] = mas["descripcion"].fillna("NO ESTÁ EN ABASTECIMIENTO")
    if "PASILLO ACTUAL" in mas.columns:
        mas["pasillo"] = mas["PASILLO ACTUAL"].fillna(mas["pasillo"])
    mas["cajas_x_pallet"] = mas["cajas_x_pallet"].fillna(mas["CANTIDAD CAJAS"] / mas["PALLETS"].clip(lower=1))
    for c in ("stock_cd", "demanda_dia", "pallets_necesarios", "forecast_dia"):
        mas[c] = mas[c].fillna(0)
    mas["cobertura_dias"] = mas["cobertura_dias"].fillna(999)

    # reparte los pallets necesarios del SKU entre sus líneas del masivo
    acumulado = mas.groupby("CODIGO SAP")["PALLETS"].cumsum() - mas["PALLETS"]
    mas["pallets_sugeridos"] = (mas["pallets_necesarios"] - acumulado).clip(lower=0)
    mas["pallets_sugeridos"] = np.minimum(mas["pallets_sugeridos"], mas["PALLETS"]).astype(int)
    mas["pallets_no_viajan"] = mas["PALLETS"] - mas["pallets_sugeridos"]
    mas["cobertura_si_llega"] = np.where(
        mas["demanda_dia"] > 0,
        (mas["stock_cd"] + mas["pallets_sugeridos"] * mas["cajas_x_pallet"]) / mas["demanda_dia"], 999)
    sin_dem = mas["demanda_dia"] <= 0
    sug = mas["pallets_sugeridos"] > 0
    urg = mas["cobertura_dias"] < p["umbral_urgente"]
    mas["prioridad"] = np.select([sin_dem, ~sug, urg], ["4-NO TRAER", "3-POSTERGAR", "1-URGENTE"], "2-NORMAL")
    mas["motivo"] = np.select(
        [sin_dem,
         ~sug & (mas["pallets_necesarios"] > 0),
         ~sug,
         urg,
         mas["pallets_no_viajan"] > 0],
        ["Sin venta en el período",
         "Lo necesario ya viene en otra línea del masivo",
         "Ya tiene los días objetivo en el CD",
         "Cobertura bajo el umbral urgente",
         "Con menos pallets ya llega a los días objetivo"],
        "Le falta para llegar a los días objetivo")
    mas = mas.sort_values(["prioridad", "cobertura_dias", "demanda_dia"], ascending=[True, True, False])
    return mas.drop(columns=["orden_hoja"]).reset_index(drop=True)


def cruce_completo(m: dict, dem: pd.DataFrame, p: dict) -> pd.DataFrame:
    """Todos los SKU del modelo (y los que venden y no están en él) con la acción sugerida."""
    b = _base(m["modelo"], dem, p)
    en_mas = m["masivo"].groupby("CODIGO SAP")["PALLETS"].sum()
    b["pallets_en_masivo"] = en_mas.reindex(b.index).fillna(0)
    pallets_cl = np.floor(b["stock_cl"] / b["cajas_x_pallet"]).fillna(0)
    b["pallets_a_agregar"] = np.where(b["pallets_en_masivo"] == 0,
                                      np.minimum(b["pallets_necesarios"], pallets_cl), 0)
    b["accion"] = np.select(
        [(b["stock_cd"] <= 0) & (b["demanda_dia"] <= 0),
         (b["cobertura_dias"] >= p["umbral_exceso"]) & (b["pallets_sobrantes"] >= 1),
         (b["cobertura_dias"] < p["umbral_urgente"]) & (b["pallets_en_masivo"] == 0) & (b["pallets_a_agregar"] > 0),
         (b["cobertura_dias"] < p["umbral_urgente"]) & (b["pallets_en_masivo"] == 0) & (b["demanda_dia"] > 0),
         (b["cobertura_dias"] < p["umbral_urgente"])],
        ["-", "SOBRA EN CD", "AGREGAR AL MASIVO", "QUIEBRE (sin pallet en CL)", "BAJO (viene en masivo)"],
        "OK")
    fuera = dem[~dem["en_modelo"].astype(bool) & (dem["venta_periodo"] > 0)]
    if len(fuera):
        f = pd.DataFrame(index=fuera.index)
        f["descripcion"] = "NO ESTÁ EN EL MODELO"
        for c in ("venta_periodo", "venta_dia_prom", "venta_dia_ult_semana", "demanda_dia"):
            f[c] = fuera[c]
        f["accion"] = "VENDE Y NO ESTÁ EN EL MODELO"
        b = pd.concat([b, f])
    b.index.name = "sap"
    orden = {"AGREGAR AL MASIVO": 0, "SOBRA EN CD": 1, "QUIEBRE (sin pallet en CL)": 2,
             "VENDE Y NO ESTÁ EN EL MODELO": 3, "BAJO (viene en masivo)": 4, "OK": 5, "-": 6}
    b["_o"] = b["accion"].map(orden)
    b = b.sort_values(["_o", "pallets_sobrantes", "cobertura_dias"], ascending=[True, False, True])
    return b.drop(columns="_o").reset_index()


def resumen(mas: pd.DataFrame, cruce: pd.DataFrame) -> dict:
    r = {"pallets_masivo": int(mas["PALLETS"].sum()),
         "pallets_sugeridos": int(mas["pallets_sugeridos"].sum())}
    r["pallets_no_viajan"] = r["pallets_masivo"] - r["pallets_sugeridos"]
    r["por_prioridad"] = (mas.groupby("prioridad")
                          .agg(lineas=("CODIGO SAP", "size"), pallets_masivo=("PALLETS", "sum"),
                               pallets_sugeridos=("pallets_sugeridos", "sum"))
                          .reindex(PRIORIDADES).fillna(0).astype(int))
    r["urgentes"] = int((mas["prioridad"] == "1-URGENTE").sum())
    s = cruce[cruce["accion"] == "SOBRA EN CD"]
    r["skus_sobra"] = len(s)
    r["pallets_sobra"] = int(s["pallets_sobrantes"].sum())
    a = cruce[cruce["accion"] == "AGREGAR AL MASIVO"]
    r["skus_agregar"] = len(a)
    r["pallets_agregar"] = int(a["pallets_a_agregar"].sum())
    r["skus_quiebre"] = int((cruce["accion"] == "QUIEBRE (sin pallet en CL)").sum())
    return r


# ------------------------------------------------------------------ sobra en CD -> CL
def sobra_cd(cruce: pd.DataFrame, p: dict) -> pd.DataFrame:
    """Productos con stock de más en el CD, de mayor a menor pallets sobrantes."""
    s = cruce[cruce["accion"] == "SOBRA EN CD"].copy()
    s["pallets_en_cd"] = s["stock_cd"] / s["cajas_x_pallet"]
    s["pallets_a_dejar"] = np.ceil(s["demanda_dia"] * p["dias_objetivo"] / s["cajas_x_pallet"])
    s["cajas_a_enviar"] = s["pallets_sobrantes"] * s["cajas_x_pallet"]
    s["dias_que_quedan"] = np.where(
        s["demanda_dia"] > 0, (s["stock_cd"] - s["cajas_a_enviar"]) / s["demanda_dia"], 999)
    s["observacion"] = np.select(
        [s["demanda_dia"] <= 0, s["pallets_en_masivo"] > 0],
        ["Sin venta en el período: enviar completo salvo que sea estacional",
         "Ojo: además viene en el masivo de hoy"], "")
    return s.sort_values(["pallets_sobrantes", "cobertura_dias"], ascending=[False, False]).reset_index(drop=True)


def traslado_cd_cl(sobra: pd.DataFrame, plantilla: dict, fecha: str) -> pd.DataFrame:
    """Sobra en formato de la hoja CD - CL del modelo, listo para cargar."""
    return pd.DataFrame({
        "GRUPO": plantilla["GRUPO"], "ORIGEN": plantilla["ORIGEN"], "DESTINO": plantilla["DESTINO"],
        "FECHA": fecha, "CODIGO SAP": sobra["sap"].values,
        "CANTIDAD CAJAS": sobra["cajas_a_enviar"].round(0).values,
        "PALLETS": sobra["pallets_sobrantes"].astype(int).values,
        "PASILLO ACTUAL": sobra["pasillo"].values,
    })


COLS_VISTA_SOBRA = ["sap", "descripcion", "pasillo", "cajas_x_pallet", "stock_cd", "pallets_en_cd",
                    "demanda_dia", "forecast_dia", "cobertura_dias", "pallets_a_dejar", "pallets_sobrantes",
                    "cajas_a_enviar", "dias_que_quedan", "stock_cl", "pallets_exceso_modelo", "observacion"]


# ------------------------------------------------------------------ salidas
def masivo_listo(mas: pd.DataFrame, cruce: pd.DataFrame, fecha: str) -> pd.DataFrame:
    """Masivo con el mismo formato del modelo, solo con los pallets sugeridos, en orden de prioridad,
    más las líneas a agregar (SKU urgentes que el modelo no pidió)."""
    x = mas[mas["pallets_sugeridos"] > 0].copy()
    x["CANTIDAD CAJAS"] = np.where(x["pallets_sugeridos"] == x["PALLETS"], x["CANTIDAD CAJAS"],
                                   x["pallets_sugeridos"] * x["cajas_x_pallet"]).round(0)
    x["PALLETS"] = x["pallets_sugeridos"]
    x["PASILLO ACTUAL"] = x["pasillo"]
    out = x[COLS_MASIVO + ["PASILLO ACTUAL", "HOJA", "prioridad"]]
    a = cruce[cruce["accion"] == "AGREGAR AL MASIVO"]
    if len(a):
        g = mas[["GRUPO", "ORIGEN", "DESTINO"]].iloc[0] if len(mas) else pd.Series(
            {"GRUPO": 1, "ORIGEN": "COQUIMEXT", "DESTINO": "COQUIMBO"})
        add = pd.DataFrame({
            "GRUPO": g["GRUPO"], "ORIGEN": g["ORIGEN"], "DESTINO": g["DESTINO"], "FECHA": fecha,
            "CODIGO SAP": a["sap"].values,
            "CANTIDAD CAJAS": (a["pallets_a_agregar"] * a["cajas_x_pallet"]).round(0).values,
            "PALLETS": a["pallets_a_agregar"].astype(int).values,
            "PASILLO ACTUAL": a["pasillo"].values, "HOJA": "AGREGADO (venta real)",
            "prioridad": "1-URGENTE"})
        out = pd.concat([add, out], ignore_index=True)
    orden = {pr: i for i, pr in enumerate(PRIORIDADES)}
    out = out.assign(_o=out["prioridad"].map(orden)).sort_values("_o", kind="stable")
    return out.drop(columns="_o").rename(columns={"prioridad": "PRIORIDAD", "HOJA": "HOJA ORIGEN"})


NOMBRES = {
    "HOJA": "HOJA", "CODIGO SAP": "CODIGO SAP", "sap": "SAP", "producto": "PRODUCTO",
    "descripcion": "DESCRIPCION", "pasillo": "PASILLO", "PALLETS": "PALLETS MASIVO",
    "CANTIDAD CAJAS": "CAJAS MASIVO", "cajas_x_pallet": "CAJAS X PALLET", "stock_cd": "STOCK CD",
    "stock_cl": "STOCK CL", "transito": "TRANSITO", "forecast_dia": "FORECAST DIA (7 HRS)",
    "revision_pallets": "REVISION MANUAL (PALLETS)", "venta_periodo": "VENTA PERIODO",
    "venta_dia_prom": "VENTA DIA PROM", "venta_dia_ult_semana": "VENTA DIA ULT. SEMANA",
    "demanda_dia": "DEMANDA DIA USADA", "pct_dias_con_venta": "% DIAS CON VENTA",
    "real_vs_forecast": "REAL / FORECAST", "cobertura_dias": "COBERTURA CD (DIAS)",
    "cobertura_si_llega": "COBERTURA SI LLEGA (DIAS)", "pallets_necesarios": "PALLETS NECESARIOS",
    "pallets_sugeridos": "PALLETS SUGERIDOS", "pallets_no_viajan": "PALLETS QUE NO VIAJAN",
    "prioridad": "PRIORIDAD", "motivo": "MOTIVO", "pallets_en_masivo": "PALLETS EN MASIVO",
    "pallets_sobrantes": "PALLETS SOBRANTES", "pallets_a_agregar": "PALLETS A AGREGAR",
    "pallets_exceso_modelo": "EXCESO SEGUN MODELO (PALLETS)", "accion": "ACCION",
    "pallets_en_cd": "PALLETS EN CD", "pallets_a_dejar": "PALLETS A DEJAR (DIAS OBJETIVO)",
    "cajas_a_enviar": "CAJAS A ENVIAR", "dias_que_quedan": "DIAS QUE QUEDAN EN CD",
    "observacion": "OBSERVACION",
}
COLS_VISTA_MASIVO = ["prioridad", "HOJA", "CODIGO SAP", "descripcion", "pasillo", "PALLETS",
                     "pallets_sugeridos", "pallets_no_viajan", "stock_cd", "demanda_dia",
                     "forecast_dia", "real_vs_forecast", "cobertura_dias", "cobertura_si_llega",
                     "revision_pallets", "motivo"]
COLS_VISTA_CRUCE = ["accion", "sap", "descripcion", "pasillo", "cajas_x_pallet", "stock_cd", "stock_cl",
                    "forecast_dia", "venta_dia_prom", "venta_dia_ult_semana", "demanda_dia",
                    "real_vs_forecast", "cobertura_dias", "pallets_en_masivo", "pallets_a_agregar",
                    "pallets_sobrantes", "pallets_exceso_modelo", "revision_pallets"]


def vista(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    return df[[c for c in cols if c in df.columns]].rename(columns=NOMBRES)


def a_excel(listo, mas, cruce, res, info, p, sobra=None, traslado=None) -> bytes:
    """Excel con el masivo priorizado listo para cargar, el detalle y el cruce completo."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    colores = {"1-URGENTE": "A6093D", "2-NORMAL": "FFB81C", "3-POSTERGAR": "64A70B", "4-NO TRAER": "BFBFBF",
               "AGREGAR AL MASIVO": "A6093D", "QUIEBRE (sin pallet en CL)": "7F7F7F",
               "SOBRA EN CD": "64A70B", "VENDE Y NO ESTÁ EN EL MODELO": "007DBA",
               "BAJO (viene en masivo)": "FFB81C"}
    blanco = {"A6093D", "64A70B", "7F7F7F", "007DBA"}
    salida = io.BytesIO()
    hojas = {
        "MASIVO PRIORIZADO": listo,
        "DETALLE MASIVO": vista(mas, COLS_VISTA_MASIVO),
        "CRUCE COMPLETO": vista(cruce[cruce["accion"] != "-"], COLS_VISTA_CRUCE),
    }
    if sobra is not None:
        hojas["SOBRA EN CD"] = vista(sobra, COLS_VISTA_SOBRA)
        hojas["TRASLADO CD - CL"] = traslado
    rp = res["por_prioridad"].reset_index().rename(columns={
        "prioridad": "PRIORIDAD", "lineas": "LINEAS", "pallets_masivo": "PALLETS MASIVO",
        "pallets_sugeridos": "PALLETS SUGERIDOS"})
    with pd.ExcelWriter(salida, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
        rp.to_excel(w, sheet_name="RESUMEN", index=False, startrow=9)
        ws = w.book["RESUMEN"]
        filas = [("Revisión de abastecimiento CD Coquimbo", ""),
                 ("Venta usada", f"{info['desde']:%d-%m-%Y} a {info['hasta']:%d-%m-%Y} "
                                 f"({info['dias_operativos']} días operativos)"),
                 ("Días objetivo", p["dias_objetivo"]),
                 ("Umbral urgente / exceso (días)", f"{p['umbral_urgente']} / {p['umbral_exceso']}"),
                 ("Demanda", "mayor entre promedio del período y última semana" if p["usar_tendencia"]
                  else "promedio del período"),
                 ("Pallets masivo / sugeridos / no viajan",
                  f"{res['pallets_masivo']} / {res['pallets_sugeridos']} / {res['pallets_no_viajan']}"),
                 ("Sobra en CD (SKU / pallets)", f"{res['skus_sobra']} / {res['pallets_sobra']}"),
                 ("Agregar al masivo (SKU / pallets)", f"{res['skus_agregar']} / {res['pallets_agregar']}")]
        for i, (k, v) in enumerate(filas, 1):
            ws.cell(i, 1, k).font = Font(name="Arial", bold=True, size=13 if i == 1 else 10)
            ws.cell(i, 2, v).font = Font(name="Arial")
        w.book.move_sheet("RESUMEN", offset=-(len(hojas)))
        for ws in w.book.worksheets:
            hdr_row = 10 if ws.title == "RESUMEN" else 1
            for c in ws[hdr_row]:
                if c.value is None:
                    continue
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="205C40")
                c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            ws.row_dimensions[hdr_row].height = 32
            for col in range(1, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(col)].width = 13
            ws.column_dimensions["A"].width = 34 if ws.title == "RESUMEN" else 16
            if ws.title != "RESUMEN":
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
            for fila in ws.iter_rows(min_row=hdr_row + 1):
                for c in fila:
                    c.font = Font(name="Arial")
                    if isinstance(c.value, float):
                        c.number_format = "0%" if ws.cell(hdr_row, c.column).value == "REAL / FORECAST" else "#,##0.0"
                    if isinstance(c.value, str) and c.value in colores:
                        col = colores[c.value]
                        c.fill = PatternFill("solid", fgColor=col)
                        c.font = Font(name="Arial", bold=True, color="FFFFFF" if col in blanco else "000000")
            for c in ws[hdr_row]:
                if c.value in ("DESCRIPCION",):
                    ws.column_dimensions[c.column_letter].width = 36
                if c.value in ("MOTIVO", "ACCION", "PRIORIDAD", "HOJA ORIGEN"):
                    ws.column_dimensions[c.column_letter].width = 26
                if c.value == "OBSERVACION":
                    ws.column_dimensions[c.column_letter].width = 50
    return salida.getvalue()


def a_excel_sobra(sobra, traslado, info, p) -> bytes:
    """Excel solo con la sobra del CD (ordenada) y la hoja de traslado CD - CL."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as w:
        vista(sobra, COLS_VISTA_SOBRA).to_excel(w, sheet_name="SOBRA EN CD", index=False, startrow=2)
        traslado.to_excel(w, sheet_name="TRASLADO CD - CL", index=False)
        ws = w.book["SOBRA EN CD"]
        ws["A1"] = (f"Sobra en CD sobre {p['umbral_exceso']:g} días de cobertura · deja {p['dias_objetivo']} días "
                    f"venta · venta {info['desde']:%d-%m-%Y} a {info['hasta']:%d-%m-%Y} · "
                    f"{int(sobra['pallets_sobrantes'].sum())} pallets en {len(sobra)} SKU")
        ws["A1"].font = Font(name="Arial", bold=True, size=12)
        for hoja, fila in (("SOBRA EN CD", 3), ("TRASLADO CD - CL", 1)):
            ws = w.book[hoja]
            for c in ws[fila]:
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="205C40")
                c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
            ws.row_dimensions[fila].height = 32
            for col in range(1, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(col)].width = 13
            for c in ws[fila]:
                if c.value == "DESCRIPCION":
                    ws.column_dimensions[c.column_letter].width = 36
                if c.value == "OBSERVACION":
                    ws.column_dimensions[c.column_letter].width = 50
            ws.freeze_panes = ws.cell(fila + 1, 3)
            ws.auto_filter.ref = f"A{fila}:{get_column_letter(ws.max_column)}{ws.max_row}"
            for r in ws.iter_rows(min_row=fila + 1):
                for c in r:
                    c.font = Font(name="Arial")
                    if isinstance(c.value, float):
                        c.number_format = "#,##0.0"
    return salida.getvalue()


def procesar(modelo_bytes: bytes, ventas: list[tuple[str, bytes]], p: dict | None = None) -> dict:
    p = {**PARAMS, **(p or {})}
    m = leer_modelo(modelo_bytes)
    v = mapear_codigos(leer_venta(ventas), m["modelo"])
    dem, info = demanda(v, p)
    mas = priorizar_masivo(m, dem, p)
    cru = cruce_completo(m, dem, p)
    res = resumen(mas, cru)
    fecha = str(m["masivo"]["FECHA"].iloc[0]) if len(m["masivo"]) else ""
    listo = masivo_listo(mas, cru, fecha)
    sob = sobra_cd(cru, p)
    tras = traslado_cd_cl(sob, m["plantilla_cd_cl"], fecha)
    return {"masivo": mas, "cruce": cru, "resumen": res, "info": info, "listo": listo, "params": p,
            "sobra": sob, "traslado": tras,
            "excel": a_excel(listo, mas, cru, res, info, p, sob, tras),
            "excel_sobra": a_excel_sobra(sob, tras, info, p)}

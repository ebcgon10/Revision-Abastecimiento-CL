import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sistema de Revisión Abastecimiento", page_icon="🚚", layout="wide")
import tema  # noqa: E402
import motor  # noqa: E402

tema.aplicar()

# Paleta CCU (misma de Revisión Venta)
VERDE, VERDE_OSC, AMBAR, AZUL, VINO, GRIS = "#64A70B", "#205C40", "#FFB81C", "#007DBA", "#A6093D", "#7F7F7F"
COLOR = {"1-URGENTE": VINO, "2-NORMAL": AMBAR, "3-POSTERGAR": VERDE, "4-NO TRAER": GRIS,
         "AGREGAR AL MASIVO": VINO, "SOBRA EN CD": VERDE, "QUIEBRE (sin pallet en CL)": GRIS,
         "VENDE Y NO ESTÁ EN EL MODELO": AZUL, "BAJO (viene en masivo)": AMBAR}
TEXTO_OSCURO = {AMBAR}

# Carpeta de venta por defecto (VENTAS en el Drive personal). Se puede cambiar en los secrets.
CARPETA_VENTA = "18xG-E755XZcz9f-eCd1DpboJg8gKwxsN"
PATRON_VENTA = "VENTA"

st.markdown("""<style>
.ccu-card {background:#fff;border-radius:10px;padding:0.9rem 1rem;height:100%;}
.ccu-lab {font-size:0.8rem;font-weight:700;text-transform:uppercase;letter-spacing:0.02em;}
.ccu-val {font-size:2rem;font-weight:800;line-height:1.1;}
.ccu-sub {font-size:0.8rem;color:#4b5b50;}
</style>""", unsafe_allow_html=True)


# ---------------------------------------------------------------- Drive (opcional)
def secretos_drive():
    """Credenciales de la cuenta de servicio desde los secrets. Acepta el JSON pegado como texto
    (gcp_json = '''{...}''') o como tabla TOML ([gcp_json] / [gcp_service_account])."""
    import json
    cred = None
    for clave in ("gcp_json", "gcp_service_account"):
        try:
            cred = st.secrets[clave]
            break
        except (KeyError, FileNotFoundError):
            continue
    if cred is None:
        return None
    if not isinstance(cred, str):
        cred = json.dumps(dict(cred))
    cred = cred.strip()
    try:
        info = json.loads(cred)
        if "client_email" not in info or "private_key" not in info:
            raise ValueError("faltan client_email o private_key")
    except Exception as e:  # noqa: BLE001
        st.warning("El secret gcp_json no es un JSON válido de cuenta de servicio "
                   f"({e}). Pega el contenido completo del archivo .json, desde la primera {{ "
                   "hasta la última }, entre triples comillas. Mientras tanto, sube la venta a mano.")
        return None
    try:
        carpeta = st.secrets.get("drive_venta_folder_id", CARPETA_VENTA)
    except Exception:
        carpeta = CARPETA_VENTA
    return cred, carpeta


@st.cache_resource
def cliente_drive(cred):
    import drive
    return drive.conectar(cred)


@st.cache_data(ttl=1800, show_spinner="Buscando venta en Drive...")
def listar_drive(cred, carpeta):
    import drive
    return drive.listar_archivos(cliente_drive(cred), carpeta)


@st.cache_data(show_spinner=False, max_entries=50)
def bajar_drive(cred, archivo_id, modificado):
    import drive
    return drive.descargar(cliente_drive(cred), archivo_id)


# ---------------------------------------------------------------- tarjetas y tablas
def tarjeta(col, etiqueta, valor, color, sub=""):
    col.markdown(f"""<div class="ccu-card" style="border-left:6px solid {color};">
        <div class="ccu-lab" style="color:{color};">{etiqueta}</div>
        <div class="ccu-val" style="color:{color};">{valor}</div>
        <div class="ccu-sub">{sub}</div></div>""", unsafe_allow_html=True)


def pintar(df, columna):
    def f(fila):
        c = COLOR.get(fila.get(columna))
        if not c:
            return [""] * len(fila)
        txt = "#3a2a00" if c in TEXTO_OSCURO else "white"
        return [f"background-color:{c};color:{txt};font-weight:700" if k == columna else "" for k in fila.index]
    fmt = {c: "{:,.0f}" for c in df.columns if df[c].dtype.kind in "fi"}
    for c in ("DEMANDA DIA USADA", "VENTA DIA PROM", "VENTA DIA ULT. SEMANA", "COBERTURA CD (DIAS)",
              "COBERTURA SI LLEGA (DIAS)", "FORECAST DIA (7 HRS)"):
        if c in df.columns:
            fmt[c] = "{:,.1f}"
    if "REAL / FORECAST" in df.columns:
        fmt["REAL / FORECAST"] = "{:.0%}"
    return df.style.apply(f, axis=1).format(fmt, na_rep="")


def miles(x):
    return f"{int(round(x)):,}".replace(",", ".")


# ---------------------------------------------------------------- interfaz
st.markdown("## 🚚 Sistema de Revisión Abastecimiento")
st.markdown(
    "**Masivo CL → CD priorizado con la venta real:** cruza el modelo de abastecimiento "
    "(stock CD, stock CL y forecast del stock 7 hrs) con lo que el CD realmente vende, para "
    "traer primero lo que se va a quebrar, postergar lo que ya tiene días suficientes y detectar "
    "lo que sobra ocupando strech.")

with st.sidebar:
    st.markdown("### ⚙️ Parámetros")
    p = dict(motor.PARAMS)
    p["semanas_venta"] = st.slider("Semanas de venta a considerar", 1, 8, p["semanas_venta"])
    p["usar_tendencia"] = st.toggle("Considerar alza de la última semana", p["usar_tendencia"],
                                    help="Demanda = la mayor entre el promedio del período y el promedio "
                                         "de la última semana. Sirve en temporada (fiestas, calor).")
    p["dias_objetivo"] = st.number_input("Días venta objetivo en CD", 1, 15, p["dias_objetivo"])
    p["umbral_urgente"] = st.number_input("Urgente bajo (días de cobertura)", 0.5, 5.0,
                                          p["umbral_urgente"], 0.5)
    p["umbral_exceso"] = st.number_input("Sobra en CD sobre (días de cobertura)", 6.0, 60.0,
                                         p["umbral_exceso"], 1.0)
    st.caption("Domingo no se cuenta como día de venta.")

st.markdown("---")
c1, c2 = st.columns(2)
with c1:
    st.markdown("### 1. Modelo de abastecimiento")
    st.caption("El Excel ABASTECIMIENTO del día (hojas ABASTECIMIENTO, Stock a Mantener, "
               "REVISION, DICCIONARIO y masivos GM / CVP), tal como lo genera el modelo.")
    f_modelo = st.file_uploader("Sube el modelo:", type=["xlsx", "xlsm"], key="modelo")

with c2:
    st.markdown("### 2. Venta real del CD")
    conf = secretos_drive()
    ventas_drive = []
    if conf:
        cred, carpeta = conf
        try:
            arch = [f for f in listar_drive(cred, carpeta)
                    if PATRON_VENTA in f["name"].upper()]
        except Exception as e:  # noqa: BLE001
            arch = []
            st.warning(f"No pude leer la carpeta de venta en Drive: {e}")
        if arch:
            arch = sorted(arch, key=lambda f: f["modifiedTime"], reverse=True)
            nombres = [f"{f['carpeta']}{f['name']}" for f in arch]
            elegidos = st.multiselect("Archivos de venta en Drive", nombres, default=nombres[:1],
                                      help="Por defecto el más reciente. Si eliges varios que se traslapan, "
                                           "las filas repetidas se cuentan una vez (pero usa más memoria).")
            ventas_drive = [f for f, n in zip(arch, nombres) if n in elegidos]
            if st.button("🔄 Actualizar lista de Drive"):
                listar_drive.clear()
                st.rerun()
    else:
        st.caption("Exporte de venta SAP (951) en CSV o Excel. Puedes subir varios meses.")
    f_venta = st.file_uploader("…o sube la venta:", type=["csv", "xlsx", "xls"], key="venta",
                               accept_multiple_files=True)

st.markdown("---")
if st.button("🚀 Procesar Revisión de Abastecimiento", use_container_width=True):
    ventas = [(f.name, f.getvalue()) for f in (f_venta or [])]
    if not ventas and ventas_drive:
        with st.spinner("Bajando venta desde Drive..."):
            ventas = [(f["name"], bajar_drive(conf[0], f["id"], f["modifiedTime"])) for f in ventas_drive]
    if not f_modelo or not ventas:
        st.error("⚠️ Necesito el modelo de abastecimiento y al menos un archivo de venta.")
    else:
        with st.spinner("Cruzando modelo y venta..."):
            try:
                st.session_state["res"] = motor.procesar(f_modelo.getvalue(), ventas, p)
            except ValueError as e:
                st.error(f"⚠️ {e}")
            except Exception as e:  # noqa: BLE001
                st.error(f"⚠️ Ocurrió un error procesando los archivos: {e}")

if "res" not in st.session_state:
    st.stop()

r = st.session_state["res"]
res, info, mas, cru = r["resumen"], r["info"], r["masivo"], r["cruce"]
st.caption(f"Venta del {info['desde']:%d-%m-%Y} al {info['hasta']:%d-%m-%Y} · "
           f"{info['dias_operativos']} días operativos · {info['skus_con_venta']} SKU con venta")
if info["skus_fuera_modelo"]:
    st.warning(f"{info['skus_fuera_modelo']} código(s) de la venta no están en el modelo. "
               "Revísalos en la pestaña Cruce completo (acción 'VENDE Y NO ESTÁ EN EL MODELO').")

k = st.columns(6)
tarjeta(k[0], "Pallets masivo", miles(res["pallets_masivo"]), VERDE_OSC, "lo que pide el modelo")
tarjeta(k[1], "Pallets sugeridos", miles(res["pallets_sugeridos"]), AZUL, "lo que conviene traer")
tarjeta(k[2], "Pallets que no viajan", miles(res["pallets_no_viajan"]), VERDE,
        f"{res['pallets_no_viajan'] / max(1, res['pallets_masivo']):.0%} del masivo")
tarjeta(k[3], "Líneas urgentes", miles(res["urgentes"]), VINO, f"cobertura < {p['umbral_urgente']:g} días")
tarjeta(k[4], "Sobra en CD", f"{miles(res['pallets_sobra'])} pall.", AMBAR, f"{res['skus_sobra']} SKU")
tarjeta(k[5], "Agregar al masivo", f"{miles(res['pallets_agregar'])} pall.", VINO,
        f"{res['skus_agregar']} SKU · {res['skus_quiebre']} sin CL")
st.markdown("")

hoy = datetime.date.today().strftime("%Y%m%d")
d1, _ = st.columns(2)
d1.download_button("⬇️ Descargar masivo priorizado + detalle (Excel)", r["excel"],
                   f"Revision_Abastecimiento_{hoy}.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)

t1, t2, t3 = st.tabs(["📦 Masivo priorizado", "🔎 Cruce completo", "📉 Forecast vs venta real"])

with t1:
    pr = res["por_prioridad"].reset_index()
    pr.columns = ["PRIORIDAD", "LÍNEAS", "PALLETS MASIVO", "PALLETS SUGERIDOS"]
    st.dataframe(pintar(pr, "PRIORIDAD"), hide_index=True, use_container_width=True)
    st.caption("Carga en este orden. 3-POSTERGAR y 4-NO TRAER no suben hoy: el CD ya tiene los días "
               "objetivo o el producto no se vendió en el período.")
    filtro = st.multiselect("Prioridad", motor.PRIORIDADES, default=motor.PRIORIDADES[:2])
    vm = motor.vista(mas[mas["prioridad"].isin(filtro)], motor.COLS_VISTA_MASIVO)
    st.dataframe(pintar(vm, "PRIORIDAD"), hide_index=True, use_container_width=True, height=520)

with t2:
    acciones = [a for a in ["AGREGAR AL MASIVO", "SOBRA EN CD", "QUIEBRE (sin pallet en CL)",
                            "VENDE Y NO ESTÁ EN EL MODELO", "BAJO (viene en masivo)", "OK"]
                if a in set(cru["accion"])]
    elegidas = st.multiselect("Acción", acciones, default=[a for a in acciones if a != "OK"])
    pas = sorted(x for x in cru["pasillo"].dropna().unique() if x)
    f_pas = st.multiselect("Pasillo", pas)
    vc = cru[cru["accion"].isin(elegidas)]
    if f_pas:
        vc = vc[vc["pasillo"].isin(f_pas)]
    st.caption("SOBRA EN CD = pallets por sobre los días objetivo, en productos con cobertura sobre "
               f"{p['umbral_exceso']:g} días: candidatos a devolver al CL o a no reponer y liberar strech.")
    st.dataframe(pintar(motor.vista(vc, motor.COLS_VISTA_CRUCE), "ACCION"), hide_index=True,
                 use_container_width=True, height=560)

with t3:
    st.caption("Productos donde el forecast del stock 7 hrs se aleja de la venta real. "
               "Son los que más pallets sin rotación (forecast alto) o quiebres (forecast bajo) generan; "
               "sirven para corregir la hoja REVISION.")
    f = cru[(cru["forecast_dia"] > 0) & (cru["descripcion"] != "NO ESTÁ EN EL MODELO")].copy()
    f["dif_cajas_dia"] = f["forecast_dia"] - f["demanda_dia"]
    f["dif_pallets_5d"] = f["dif_cajas_dia"] * p["dias_objetivo"] / f["cajas_x_pallet"]
    cols = ["sap", "descripcion", "pasillo", "forecast_dia", "demanda_dia", "real_vs_forecast",
            "dif_pallets_5d", "revision_pallets", "stock_cd", "cobertura_dias"]
    ren = {**motor.NOMBRES, "dif_pallets_5d": "DIF. PALLETS EN 5 DÍAS"}
    a, b = st.columns(2)
    with a:
        st.markdown("**Forecast sobre la venta real** (llegan pallets de más)")
        st.dataframe(f.sort_values("dif_pallets_5d", ascending=False).head(40)[cols].rename(columns=ren)
                     .style.format({"REAL / FORECAST": "{:.0%}", "DIF. PALLETS EN 5 DÍAS": "{:.1f}",
                                    "DEMANDA DIA USADA": "{:.1f}", "COBERTURA CD (DIAS)": "{:.1f}"}),
                     hide_index=True, use_container_width=True, height=500)
    with b:
        st.markdown("**Forecast bajo la venta real** (riesgo de quiebre)")
        st.dataframe(f.sort_values("dif_pallets_5d").head(40)[cols].rename(columns=ren)
                     .style.format({"REAL / FORECAST": "{:.0%}", "DIF. PALLETS EN 5 DÍAS": "{:.1f}",
                                    "DEMANDA DIA USADA": "{:.1f}", "COBERTURA CD (DIAS)": "{:.1f}"}),
                     hide_index=True, use_container_width=True, height=500)

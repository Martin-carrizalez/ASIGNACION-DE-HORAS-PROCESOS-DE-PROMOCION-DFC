# -*- coding: utf-8 -*-
"""
Asignación de horas — Concurso Cerrado de Incremento de Horas
Dirección de Formación Continua · Comisión Dictaminadora Interna

El código no contiene ningún dato personal: los dos Excel se cargan en el momento,
se procesan en memoria y no se guardan en disco ni en el repositorio.
"""

import io
import json
import re
from datetime import date

import pandas as pd
import streamlit as st

import docx_merge

st.set_page_config(page_title="Asignación de horas · DFC", layout="wide")

# ---------------------------------------------------------------------------
# Reglas del proceso
# ---------------------------------------------------------------------------

TOPE_PROCESO = 15      # máximo de horas que un asesor puede tomar en este concurso
TOPE_TOTAL = 48        # máximo de horas que puede tener un empleado en total

# Orden de prelación: primero la categoría, y dentro de ella el puntaje.
ORDEN_CATEGORIA = ["HORAS DE TITULAR C", "HORAS DE TITULAR B", "HORAS DE TITULAR A",
                   "HORAS DE ASOCIADO C", "HORAS DE ASOCIADO B", "HORAS DE ASOCIADO A"]

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

LEYENDA_RENUNCIA = ("El interesado renuncia a su plaza de otras áreas por exceder las "
                    "48 horas permitidas, a fin de incrementar horas como asesor.")

LEYENDA_DESIERTA = ("La presente convocatoria se declara DESIERTA al no haberse asignado "
                    "horas a ningún participante.")

VACIOS = ("", "nan", "none", "n/a", "na", "#n/a", "-", "s/d", "sd")


def limpiar(v):
    if v is None:
        return ""
    t = re.sub(r"\s+", " ", str(v)).strip()
    return "" if t.lower() in VACIOS else t


def numero(v, por_defecto=None):
    """Devuelve el número o por_defecto; 's/d' y vacíos cuentan como sin dato."""
    try:
        if limpiar(v) == "":
            return por_defecto
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return por_defecto


def enlistar(xs):
    xs = [x for x in xs if x]
    if not xs:
        return ""
    if len(xs) == 1:
        return xs[0]
    return ", ".join(xs[:-1]) + " y " + xs[-1]


# ---------------------------------------------------------------------------
# Carga de los dos Excel
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def cargar_vacancia(contenido, anio, formato_folio):
    df = pd.read_excel(io.BytesIO(contenido)).fillna("")
    for c in df.columns:
        df[c] = df[c].map(limpiar)
    faltan = [c for c in ("CONVOCATORIA", "Clave", "hrs", "Categoria") if c not in df.columns]
    if faltan:
        raise ValueError("Al Excel de vacancia le faltan columnas: " + ", ".join(faltan))

    # El autorrelleno de Excel arrastró el año de algunos folios (/2027, /2028…).
    df["folio_origen"] = df["CONVOCATORIA"]
    df["CONVOCATORIA"] = df["CONVOCATORIA"].str.replace(r"/\d{4}$", "/" + str(anio), regex=True)
    if formato_folio:
        df["CONVOCATORIA"] = df["CONVOCATORIA"].map(
            lambda f: (formato_folio.format(n=re.search(r"[A-Za-z]+-(\d+)", f).group(1), anio=anio)
                       if re.search(r"[A-Za-z]+-(\d+)", f) else f))
    df["hrs"] = df["hrs"].map(lambda v: int(numero(v, 0)))
    df["id_clave"] = df.index.map(lambda i: "C%03d" % i)   # cada renglón es una clave
    return df


@st.cache_data(show_spinner=False)
def cargar_prelacion(contenido):
    df = pd.read_excel(io.BytesIO(contenido)).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in ("rfc", "nombre", "puesto") if c not in df.columns]
    if faltan:
        raise ValueError("Al Excel de prelación le faltan columnas: " + ", ".join(faltan))

    col_pun = next((c for c in df.columns if c.lower() == "puntaje"), None)
    col_ase = next((c for c in df.columns if c.lower() == "horas asesor"), None)
    col_otr = next((c for c in df.columns if c.lower() == "horas otras áreas"), None)
    col_bas = next((c for c in df.columns if c.lower() == "horas de base"), None)

    gente = []
    for _, r in df.iterrows():
        rfc = limpiar(r.get("rfc"))
        if not rfc:
            continue
        gente.append({
            "rfc": rfc,
            "nombre": limpiar(r.get("nombre")),
            "puesto": limpiar(r.get("puesto")),
            "plaza": limpiar(r.get("plaza")),
            "puntaje": numero(r.get(col_pun) if col_pun else None),
            "horas_asesor": numero(r.get(col_ase) if col_ase else None),
            "horas_otras": numero(r.get(col_otr) if col_otr else None),
            "horas_base": numero(r.get(col_bas) if col_bas else None),
        })
    p = pd.DataFrame(gente)
    # Orden de prelación: categoría primero, puntaje después. Quien no participó
    # (sin puntaje) queda al final, visible pero sin lugar en la lista.
    p["rango"] = p["puesto"].map(lambda v: ORDEN_CATEGORIA.index(v)
                                 if v in ORDEN_CATEGORIA else len(ORDEN_CATEGORIA))
    p["sin_puntaje"] = p["puntaje"].isna()
    p = p.sort_values(["sin_puntaje", "rango", "puntaje"],
                      ascending=[True, True, False]).reset_index(drop=True)
    p["prelacion"] = p.index + 1
    return p


def cupo_de(fila, base_manual, renuncias, asignado):
    """Horas que todavía puede tomar un asesor. Devuelve (cupo, motivo)."""
    asesor = fila["horas_asesor"]
    if asesor is not None and asesor >= TOPE_TOTAL:
        return 0, "Tiempo completo: no concursa"

    base = base_manual.get(fila["rfc"], fila["horas_base"])
    if fila["rfc"] in renuncias:
        # Al renunciar a otras áreas conserva únicamente sus horas de asesor.
        base = asesor if asesor is not None else base
    if base is None:
        return 0, "Falta capturar sus horas de base"

    cupo = max(0, min(TOPE_PROCESO, TOPE_TOTAL - base)) - asignado
    if cupo <= 0 and base >= TOPE_TOTAL:
        return 0, "Está en 48 horas: puede renunciar a otras áreas para incrementar"
    return max(0, cupo), ""


# ---------------------------------------------------------------------------
# Estado de la sesión
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def _almacen():
    """El avance vive en el proceso del servidor, no en la sesión del navegador ni
    en disco. Así sobrevive a recargar la página, a cerrar la pestaña y a que se
    caiga la conexión, que es lo que borra el trabajo a media asignación.
    Lo que NO sobrevive es un reinicio del servidor: para eso está el respaldo
    en JSON de la pestaña Salidas."""
    return {}


def estado():
    clave = st.session_state.get("clave_sesion") or "comision"
    return _almacen().setdefault(
        clave, {"asignaciones": [], "renuncias": set(), "base_manual": {}})


def asignado_por_rfc(rfc):
    return sum(a["hrs"] for a in estado()["asignaciones"] if a["rfc"] == rfc)


def asignado_por_clave(id_clave):
    return sum(a["hrs"] for a in estado()["asignaciones"] if a["id_clave"] == id_clave)


# ---------------------------------------------------------------------------
# Barra lateral: archivos y parámetros
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Archivos")
    f_vac = st.file_uploader("Excel de vacancia", type="xlsx")
    f_pre = st.file_uploader("Excel de prelación", type="xlsx")
    st.caption("Plantillas de dictamen (solo para generar los Word)")
    f_est = st.file_uploader("Dictamen estatal", type="docx")
    f_int = st.file_uploader("Dictamen interno", type="docx")
    f_ind = st.file_uploader("Dictamen individual", type="docx")

    st.header("Sesión")
    st.text_input("Nombre de la sesión de trabajo", "comision", key="clave_sesion",
                  help="El avance se guarda bajo este nombre. Si se recarga la "
                       "página o se cae la conexión, se recupera solo al volver "
                       "a entrar con el mismo nombre.")

    st.header("Parámetros")
    anio = st.text_input("Año del proceso", "2026")
    formato_folio = st.text_input("Formato del folio",
                                  "IN-{n}-Dirección de Formación Continua/{anio}")
    fecha_conv = st.text_input("Fecha de la convocatoria", "7 de octubre de 2026")
    fecha_dict = st.date_input("Fecha de los dictámenes", date.today())
    prefijo = st.text_input("Prefijo del dictamen individual", "D")
    sufijo = st.text_input("Sufijo del dictamen individual", "-DFC/2026")

if not (f_vac and f_pre):
    st.info("Carga el Excel de vacancia y el de prelación para comenzar.")
    st.stop()

try:
    vac = cargar_vacancia(f_vac.getvalue(), anio, formato_folio)
    pre = cargar_prelacion(f_pre.getvalue())
except ValueError as e:
    st.error(str(e))
    st.stop()

t_datos, t_prel, t_asig, t_sal = st.tabs(
    ["Datos", "Prelación", "Asignación", "Salidas"])

# ---------------------------------------------------------------------------
# Datos
# ---------------------------------------------------------------------------

with t_datos:
    convocatorias = list(dict.fromkeys(vac["CONVOCATORIA"]))
    c1, c2, c3 = st.columns(3)
    c1.metric("Convocatorias", len(convocatorias))
    c2.metric("Horas ofertadas", int(vac["hrs"].sum()))
    c3.metric("Asesores en prelación", len(pre))

    corregidos = vac[vac["folio_origen"] != vac["CONVOCATORIA"]]["folio_origen"].nunique()
    if corregidos:
        st.warning("Se emparejó el año de %d folio(s) con %s (arrastre del autorrelleno "
                   "de Excel)." % (corregidos, anio))

    st.subheader("Horas de base sin capturar")
    sin_base = pre[pre["horas_base"].isna() &
                   ~(pre["horas_asesor"].fillna(0) >= TOPE_TOTAL)]
    if sin_base.empty:
        st.success("Todos los asesores que pueden concursar traen sus horas de base.")
    else:
        st.caption("Sin este dato no se puede calcular el cupo. Captúralo aquí; "
                   "no se guarda en ningún lado al cerrar la sesión.")
        for _, r in sin_base.iterrows():
            c1, c2 = st.columns([3, 1])
            c1.write("**%s** · %s · horas asesor: %s"
                     % (r["nombre"], r["puesto"].replace("HORAS DE ", ""),
                        "s/d" if r["horas_asesor"] is None else int(r["horas_asesor"])))
            valor = c2.number_input("Horas de base", 0, 60,
                                    int(estado()["base_manual"].get(r["rfc"], TOPE_TOTAL)),
                                    key="base_" + r["rfc"], label_visibility="collapsed")
            estado()["base_manual"][r["rfc"]] = valor

# ---------------------------------------------------------------------------
# Prelación
# ---------------------------------------------------------------------------

with t_prel:
    st.caption("Orden calculado con la regla del concurso: primero la categoría "
               "(Titular C hacia abajo) y dentro de ella el puntaje.")
    filas = []
    for _, r in pre.iterrows():
        ya = asignado_por_rfc(r["rfc"])
        cupo, motivo = cupo_de(r, estado()["base_manual"], estado()["renuncias"], ya)
        filas.append({
            "#": r["prelacion"],
            "Asesor": r["nombre"],
            "Categoría": r["puesto"].replace("HORAS DE ", ""),
            "Puntaje": r["puntaje"],
            "Base": estado()["base_manual"].get(r["rfc"], r["horas_base"]),
            "Asignadas": ya,
            "Puede tomar": cupo,
            "Situación": motivo,
        })
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    st.subheader("Renuncia a horas de otras áreas")
    st.caption("Quien está en 48 horas puede renunciar a lo que tiene en otras áreas "
               "para crecer como asesor. Al marcarlo, su cupo se recalcula sobre sus "
               "horas de asesor y el dictamen individual lleva la leyenda.")
    candidatos = pre[pre["horas_otras"].fillna(0) > 0]
    for _, r in candidatos.iterrows():
        marcado = st.checkbox(
            "%s — asesor %s / otras áreas %s"
            % (r["nombre"],
               "s/d" if r["horas_asesor"] is None else int(r["horas_asesor"]),
               int(r["horas_otras"])),
            value=r["rfc"] in estado()["renuncias"], key="ren_" + r["rfc"])
        if marcado:
            estado()["renuncias"].add(r["rfc"])
        else:
            estado()["renuncias"].discard(r["rfc"])

# ---------------------------------------------------------------------------
# Asignación
# ---------------------------------------------------------------------------

with t_asig:
    convocatorias = list(dict.fromkeys(vac["CONVOCATORIA"]))
    etiquetas = []
    for c in convocatorias:
        sub = vac[vac["CONVOCATORIA"] == c]
        libre = sum(int(r["hrs"]) - asignado_por_clave(r["id_clave"]) for _, r in sub.iterrows())
        etiquetas.append("%s — %d hrs libres de %d" % (c, libre, int(sub["hrs"].sum())))

    elegida = st.selectbox("Convocatoria", range(len(convocatorias)),
                           format_func=lambda i: etiquetas[i])
    conv = convocatorias[elegida]
    sub = vac[vac["CONVOCATORIA"] == conv]

    st.write("**%s** · %s" % (sub.iloc[0].get("Sustituido", ""),
                              sub.iloc[0].get("Motivo Baja", "")))

    opciones = []
    for _, r in sub.iterrows():
        libre = int(r["hrs"]) - asignado_por_clave(r["id_clave"])
        opciones.append((r["id_clave"], "%s · %s · %d de %d hrs libres"
                         % (r["Clave"], r["Categoria"], libre, int(r["hrs"])), libre, r))

    c1, c2, c3 = st.columns([3, 2, 1])
    idx = c1.selectbox("Clave presupuestal", range(len(opciones)),
                       format_func=lambda i: opciones[i][1])
    id_clave, _, libre_clave, fila_clave = opciones[idx]

    # Los asesores se ofrecen en orden de prelación, con su cupo ya calculado.
    disponibles = []
    for _, r in pre.iterrows():
        cupo, motivo = cupo_de(r, estado()["base_manual"], estado()["renuncias"],
                               asignado_por_rfc(r["rfc"]))
        disponibles.append((r, cupo, motivo))

    quien = c2.selectbox(
        "Asesor (en orden de prelación)", range(len(disponibles)),
        format_func=lambda i: "%d. %s — puede tomar %d%s"
        % (disponibles[i][0]["prelacion"], disponibles[i][0]["nombre"],
           disponibles[i][1], " · " + disponibles[i][2] if disponibles[i][2] else ""))
    asesor, cupo_asesor, motivo = disponibles[quien]

    tope = int(min(libre_clave, cupo_asesor))
    horas = c3.number_input("Horas", 0, max(tope, 0), max(tope, 0) if tope > 0 else 0)

    if tope <= 0:
        st.warning("No hay horas que asignar aquí: la clave tiene %d libres y el asesor "
                   "puede tomar %d. %s" % (libre_clave, cupo_asesor, motivo))
    elif st.button("Asignar", type="primary"):
        estado()["asignaciones"].append({
            "convocatoria": conv, "id_clave": id_clave,
            "clave": fila_clave["Clave"], "categoria": fila_clave["Categoria"],
            "hrs": int(horas), "rfc": asesor["rfc"], "nombre": asesor["nombre"],
            "categoria_actual": asesor["puesto"], "puntaje": asesor["puntaje"],
            "clave_actual": asesor["plaza"], "prelacion": int(asesor["prelacion"]),
            "horas_actuales": asesor["horas_asesor"],
        })
        st.rerun()

    st.subheader("Asignaciones registradas")
    if not estado()["asignaciones"]:
        st.caption("Todavía no hay ninguna.")
    else:
        for i, a in enumerate(estado()["asignaciones"]):
            c1, c2 = st.columns([8, 1])
            c1.write("%s · %s · **%d hrs** · %s" % (a["convocatoria"], a["clave"],
                                                    a["hrs"], a["nombre"]))
            if c2.button("Quitar", key="del_%d" % i):
                estado()["asignaciones"].pop(i)
                st.rerun()

# ---------------------------------------------------------------------------
# Salidas
# ---------------------------------------------------------------------------

def excel_de(dataframes):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for hoja, df in dataframes.items():
            df.to_excel(w, sheet_name=hoja, index=False)
    return buf.getvalue()


def bloques_por_convocatoria():
    """Un bloque por convocatoria para los dictámenes estatal e interno."""
    bloques = []
    for conv in dict.fromkeys(vac["CONVOCATORIA"]):
        sub = vac[vac["CONVOCATORIA"] == conv]
        base = sub.iloc[0]
        ganadores = [a for a in estado()["asignaciones"] if a["convocatoria"] == conv]
        motivo = base.get("Motivo Baja", "")
        sustituido = base.get("Sustituido", "")
        valores = {
            "CONVOCATORIA": conv,
            "Fecha convocatoria": fecha_conv,
            "Motivo": ("la %s de %s" % (motivo, sustituido)) if motivo else ("el %s" % sustituido),
            "RFC sustituido": base.get("RFC", ""),
            "Plaza": "%d %s" % (int(sub["hrs"].sum()), base.get("Categoria", "")),
            "Clave": ", ".join(sub["Clave"]),
            "Escuela": base.get("Escuela", ""),
            "CCT": base.get("CCT", ""),
            "Leyenda": "" if ganadores else LEYENDA_DESIERTA,
        }
        filas = [{
            "N": str(i), "Nombre": g["nombre"], "RFC": g["rfc"],
            "Categoría actual": g["categoria_actual"].replace("HORAS DE ", ""),
            "Clave actual": g["clave_actual"], "Hrs actuales": str(g["hrs"]),
            "Categoría asignada": g["categoria"].replace("HORAS DE ", ""),
            "Clave asignada": g["clave"], "Hrs asignadas": str(g["hrs"]),
            "Puntaje": "" if g["puntaje"] is None else g["puntaje"],
        } for i, g in enumerate(ganadores, 1)]
        bloques.append((valores, {"N": filas}))
    return bloques


def bloques_por_asesor():
    """Un bloque por asesor que ganó horas, numerado D01, D02… en orden de prelación."""
    por_rfc = {}
    for a in estado()["asignaciones"]:
        por_rfc.setdefault(a["rfc"], []).append(a)
    orden = sorted(por_rfc, key=lambda r: min(a["prelacion"] for a in por_rfc[r]))

    bloques = []
    for n, rfc in enumerate(orden, 1):
        asigs = por_rfc[rfc]
        primera = asigs[0]
        # Una línea por categoría asignada, sumando las horas de esa categoría.
        suma = {}
        for a in asigs:
            suma[a["categoria"]] = suma.get(a["categoria"], 0) + a["hrs"]
        actuales = ("" if primera["horas_actuales"] is None
                    else str(int(primera["horas_actuales"])))
        filas = [{"Categoría actual": primera["categoria_actual"],
                  "Horas actuales": actuales,
                  "Categoría asignada": cat, "Horas asignadas": str(hrs)}
                 for cat, hrs in suma.items()]
        valores = {
            "Dictamen": "%s%02d%s" % (prefijo, n, sufijo),
            "Asesor": primera["nombre"],
            "Convocatorias": enlistar(sorted({a["convocatoria"] for a in asigs})),
            "Puntaje": "" if primera["puntaje"] is None else "%.2f" % primera["puntaje"],
            "Día": str(fecha_dict.day),
            "Mes": MESES[fecha_dict.month - 1],
            "Año": str(fecha_dict.year),
            "Fecha convocatoria": fecha_conv,
            "Nota": LEYENDA_RENUNCIA if rfc in estado()["renuncias"] else "",
        }
        bloques.append((valores, {"Categoría actual": filas}))
    return bloques


with t_sal:
    libres = []
    for _, r in vac.iterrows():
        queda = int(r["hrs"]) - asignado_por_clave(r["id_clave"])
        if queda > 0:
            fila = r.drop(labels=["id_clave", "folio_origen"]).to_dict()
            fila["hrs"] = queda
            libres.append(fila)

    c1, c2, c3 = st.columns(3)
    c1.metric("Horas asignadas", sum(a["hrs"] for a in estado()["asignaciones"]))
    c2.metric("Horas libres", sum(f["hrs"] for f in libres))
    c3.metric("Asesores con horas", len({a["rfc"] for a in estado()["asignaciones"]}))

    st.download_button(
        "Excel · vacancia que quedó libre",
        excel_de({"vacancia libre": pd.DataFrame(libres)}),
        "vacancia_libre_%s.xlsx" % anio,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.download_button(
        "Excel · asignaciones",
        excel_de({"asignaciones": pd.DataFrame(estado()["asignaciones"])}),
        "asignaciones_%s.xlsx" % anio,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.subheader("Respaldo del avance")
    st.caption("El avance se conserva solo mientras el servidor siga encendido. "
               "Descarga este archivo de vez en cuando: devuelve la asignación "
               "tal como iba, aunque la app se haya reiniciado.")
    c1, c2 = st.columns(2)
    c1.download_button(
        "Descargar avance (JSON)",
        json.dumps({"asignaciones": estado()["asignaciones"],
                    "renuncias": sorted(estado()["renuncias"]),
                    "base_manual": estado()["base_manual"]},
                   ensure_ascii=False, indent=1).encode("utf-8"),
        "avance_%s.json" % anio, "application/json")
    respaldo = c2.file_uploader("Restaurar avance", type="json",
                                label_visibility="collapsed")
    if respaldo is not None and c2.button("Restaurar este archivo"):
        datos = json.loads(respaldo.getvalue().decode("utf-8"))
        estado()["asignaciones"] = datos.get("asignaciones", [])
        estado()["renuncias"] = set(datos.get("renuncias", []))
        estado()["base_manual"] = datos.get("base_manual", {})
        st.rerun()

    st.divider()
    MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    for etiqueta, archivo, generador, nombre in (
            ("Word · dictamen estatal", f_est, bloques_por_convocatoria, "dictamen_estatal"),
            ("Word · dictamen interno", f_int, bloques_por_convocatoria, "dictamen_interno"),
            ("Word · dictámenes individuales", f_ind, bloques_por_asesor, "dictamenes_individuales")):
        if archivo is None:
            st.caption("%s — falta cargar su plantilla." % etiqueta)
            continue
        bloques = generador()
        if not bloques:
            st.caption("%s — todavía no hay nada que dictaminar." % etiqueta)
            continue
        datos, faltantes = docx_merge.generar(archivo.getvalue(), bloques)
        st.download_button(etiqueta, datos, "%s_%s.docx" % (nombre, anio), MIME)
        if faltantes:
            st.caption("Marcadores sin dato: " + ", ".join(sorted(faltantes)))

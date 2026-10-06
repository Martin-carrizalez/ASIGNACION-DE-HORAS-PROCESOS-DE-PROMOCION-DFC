"""
Registro del Concurso Cerrado de Incremento de Horas 2026
Comisión Dictaminadora Interna — Dirección de Formación Continua / CEPD

Tres etapas, en el mismo orden del calendario de la Base SÉPTIMA:
  1. RECEPCIÓN   (7 al 21 de octubre) — alta del participante, convocatorias
                 que solicita, requisitos y documentos entregados.
  2. VALORACIÓN  (22 al 28 de octubre) — captura del tabulador por expediente.
  3. PRELACIÓN   — orden final y exportación que alimenta el sistema de asignación.

Persistencia: Google Sheets (gspread + cuenta de servicio en st.secrets), mismo
patrón que el resto de las apps del área. No se guarda nada en el repositorio.

Configuración requerida en .streamlit/secrets.toml:

    [gcp_service_account]
    ... (el bloque de la cuenta de servicio que ya usas)

    [registro_incremento]
    sheet_id = "ID_DE_LA_HOJA_NUEVA"

La hoja debe estar compartida como Editor con el correo de la cuenta de servicio.
Las pestañas se crean solas la primera vez que corre la app.
"""

import io
import json
from datetime import datetime

import gspread
import pandas as pd
import pytz
import streamlit as st
from google.oauth2.service_account import Credentials

import tabulador as tb

TZ = pytz.timezone("America/Mexico_City")

# --- Catálogos de la convocatoria -----------------------------------------
# Centros de trabajo que convoca (Base PRIMERA.6 de la Convocatoria Marco).
CCTS = ["14ADG1075P — Dirección de Formación Continua",
        "14FMP0001B — Centro Estatal de Profesionalización Docente"]

CATEGORIAS = ["Asociado", "Titular"]
NIVELES = ["A", "B", "C"]

# Requisitos de participación. El texto se conserva idéntico al del checklist
# impreso para que ventanilla y sistema digan exactamente lo mismo.
REQUISITOS = [
    ("REQ_01", "Adscrito a la DFC (14ADG1075P) o al CEPD (14FMP0001B)", "Base PRIMERA.6"),
    ("REQ_02", "Plaza con nombramiento Código 10, alta definitiva", "Art. 4; Base PRIMERA.6"),
    ("REQ_03", "Desempeña funciones académicas, directivas o técnico pedagógicas acreditadas con constancia", "Art. 9; Base PRIMERA.7"),
    ("REQ_04", "Cumple los requisitos de categoría y nivel del Art. 13 o del Adendum", "Art. 13 y Adendum"),
    ("REQ_05", "Si invoca formación equivalente: se verificó que no hubiere participantes con la formación requerida", "Art. 9 reformado"),
    ("REQ_06", "Sin demanda legal pendiente contra la SEJ por asuntos laborales", "Art. 12; Base PRIMERA.9"),
    ("REQ_07", "Si goza de licencia sin goce de sueldo: acredita haber reactivado su(s) clave(s) de base", "Art. 8 reformado"),
    ("REQ_08", "Si obtuvo promoción previa: transcurrió un año, salvo que no haya participantes", "Art. 17 reformado"),
    ("REQ_09", "Las horas solicitadas están entre 1 y 15", "Art. 14 Primera b); Base PRIMERA.10"),
    ("REQ_10", "Constancias posteriores a la última promoción y no presentadas en recategorización", "Base PRIMERA.5"),
]

# Documentos del expediente (Base SEGUNDA + Etapa 2 del calendario).
DOCUMENTOS = [
    ("DOC_01", "Solicitud de participación por duplicado, firmada"),
    ("DOC_02", "Identificación oficial"),
    ("DOC_03", "Formación académica (grado o título)"),
    ("DOC_04", "Experiencia docente o profesional"),
    ("DOC_05", "Talón de cheque reciente"),
    ("DOC_06", "Dictamen de homologación de plaza de origen (si aplica)"),
    ("DOC_07", "Constancia de antigüedad del Depto. de Registros y Archivo"),
    ("DOC_08", "Constancia de funciones de áreas sustantivas firmada por la Dirección (Art. 9)"),
    ("DOC_09", "Relación escrita de los documentos entregados"),
]

ESTATUS = ["RECIBIDA", "RECIBIDA CON FALTANTES", "SUBSANADA", "IMPROCEDENTE"]

# Encabezados de cada pestaña de la hoja. El orden es el contrato con el Sheet:
# si se agrega una columna, va al final, nunca en medio.
COLS_PARTICIPANTES = (
    ["folio", "fecha_recepcion", "nombre", "rfc", "curp", "cct", "correo", "telefono",
     "categoria_actual", "nivel_actual", "horas_actuales", "fecha_ultima_promocion",
     "horas_solicitadas", "estatus"]
    + [k for k, _, _ in REQUISITOS]
    + [k for k, _ in DOCUMENTOS]
    + ["observaciones", "capturo", "actualizado"]
)
COLS_CONVOCATORIAS = ["folio", "no_convocatoria", "categoria", "nivel",
                      "clave_presupuestal", "horas_vacante"]
COLS_PUNTAJES = ["folio", "concepto", "cantidad", "capturo", "actualizado"]

HOJAS = {
    "PARTICIPANTES": COLS_PARTICIPANTES,
    "CONVOCATORIAS": COLS_CONVOCATORIAS,
    "PUNTAJES": COLS_PUNTAJES,
}


# =========================================================================
# Capa de datos
# =========================================================================

@st.cache_resource(show_spinner=False)
def abrir_libro():
    """Abre el Sheet y crea las pestañas que falten con sus encabezados."""
    alcances = ["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"]
    cred = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=alcances)
    libro = gspread.authorize(cred).open_by_key(
        st.secrets["registro_incremento"]["sheet_id"])

    existentes = {h.title for h in libro.worksheets()}
    for nombre, cols in HOJAS.items():
        if nombre not in existentes:
            hoja = libro.add_worksheet(title=nombre, rows=500, cols=len(cols) + 5)
            hoja.append_row(cols)
        else:
            hoja = libro.worksheet(nombre)
            # Si la pestaña existe pero está vacía, se le ponen los encabezados.
            if not hoja.row_values(1):
                hoja.append_row(cols)
    return libro


def leer(nombre):
    """Lee una pestaña completa como DataFrame, respetando sus columnas."""
    hoja = abrir_libro().worksheet(nombre)
    datos = hoja.get_all_records()
    df = pd.DataFrame(datos, columns=HOJAS[nombre]) if datos else pd.DataFrame(columns=HOJAS[nombre])
    return df


def agregar(nombre, registro):
    """Agrega un renglón al final, en el orden de las columnas de la pestaña."""
    hoja = abrir_libro().worksheet(nombre)
    hoja.append_row([str(registro.get(c, "")) for c in HOJAS[nombre]],
                    value_input_option="USER_ENTERED")


def reemplazar_por_folio(nombre, folio, registros):
    """
    Borra los renglones de un folio y escribe los nuevos.
    Se usa para convocatorias y puntajes, que se recapturan en bloque.
    """
    hoja = abrir_libro().worksheet(nombre)
    valores = hoja.get_all_values()
    if valores:
        encabezado, cuerpo = valores[0], valores[1:]
        i_folio = encabezado.index("folio")
        conservados = [r for r in cuerpo if (r[i_folio] if len(r) > i_folio else "") != folio]
    else:
        encabezado, conservados = HOJAS[nombre], []
    nuevos = [[str(r.get(c, "")) for c in HOJAS[nombre]] for r in registros]
    hoja.clear()
    hoja.update([encabezado] + conservados + nuevos, value_input_option="USER_ENTERED")


def actualizar_participante(folio, registro):
    """Reescribe el renglón de un participante conservando su posición."""
    hoja = abrir_libro().worksheet("PARTICIPANTES")
    valores = hoja.get_all_values()
    i_folio = valores[0].index("folio")
    for n, fila in enumerate(valores[1:], start=2):
        if len(fila) > i_folio and fila[i_folio] == folio:
            hoja.update(
                [[str(registro.get(c, "")) for c in COLS_PARTICIPANTES]],
                f"A{n}", value_input_option="USER_ENTERED")
            return True
    return False


def siguiente_folio(df):
    """Folio consecutivo INC-2026-001. Es por docente, no por convocatoria."""
    usados = [int(f.split("-")[-1]) for f in df["folio"].astype(str)
              if f.startswith("INC-2026-") and f.split("-")[-1].isdigit()]
    return f"INC-2026-{(max(usados) + 1 if usados else 1):03d}"


def ahora():
    return datetime.now(TZ).strftime("%d/%m/%Y %H:%M")


# Contrato con el sistema de asignación (repo SISTEMA-PARA-PROPUESTAS-DFC):
# cargar_prelacion() de esa app lee exactamente estas columnas, con estos
# encabezados. No cambiar el texto sin cambiarlo allá también.
COLS_ASIGNACION = ["rfc", "nombre", "puesto", "plaza", "puntaje",
                   "horas asesor", "horas otras áreas", "horas de base"]


def excel_para_asignacion(orden):
    """
    Arma el Excel de prelación en el esquema que consume el sistema de
    asignación. El ORDEN lo decide esa app (categoría primero, puntaje
    después); aquí sólo se entregan los datos.

    Lo que se deja en blanco a propósito:
      - puntaje: si el expediente no está valorado. La app de asignación
        manda al final a quien no trae puntaje, que es lo correcto.
      - plaza: la clave presupuestal actual no se captura en recepción.
      - horas asesor / horas otras áreas / horas de base: salen de nómina,
        no del expediente. Las horas actuales que declara el participante
        NO sirven como horas de base: el tope de 48 se calcula sumando todas
        sus plazas, incluidas las de otros centros, niveles o sostenimientos.
        Se capturan en la pantalla de la app de asignación o se llenan desde
        el historial de nóminas.
    """
    filas = []
    for e in orden:
        filas.append({
            "rfc": e["rfc"],
            "nombre": e["nombre"],
            # ORDEN_CATEGORIA de la otra app compara esta cadena exacta.
            "puesto": f"HORAS DE {str(e['categoria_actual']).upper()} {str(e['nivel_actual']).upper()}",
            "plaza": "",
            "puntaje": e["total"] if e["valorado"] == "Sí" else None,
            "horas asesor": None,
            "horas otras áreas": None,
            "horas de base": None,
        })
    buffer = io.BytesIO()
    pd.DataFrame(filas, columns=COLS_ASIGNACION).to_excel(buffer, index=False)
    return buffer.getvalue()


# =========================================================================
# Interfaz
# =========================================================================

st.set_page_config(page_title="Registro Incremento de Horas 2026",
                   page_icon="📋", layout="wide")

st.title("Registro del Concurso Cerrado de Incremento de Horas 2026")
st.caption("Comisión Dictaminadora Interna · Dirección de Formación Continua y CEPD · "
           "Recepción del 7 al 21 de octubre de 2026, cierre a las 15:00 h")

# Quién captura: queda asentado en cada renglón para efectos de auditoría,
# porque el oficio CDESE/017/2026 pide un concentrado del proceso.
capturo = st.sidebar.text_input("Captura (nombre de quien registra)", key="capturo")
if st.sidebar.button("Recargar datos"):
    st.cache_data.clear()
    st.rerun()

try:
    df_part = leer("PARTICIPANTES")
    df_conv = leer("CONVOCATORIAS")
    df_punt = leer("PUNTAJES")
except Exception as e:
    st.error(f"No se pudo abrir la hoja de cálculo: {e}")
    st.stop()

st.sidebar.metric("Solicitudes registradas", len(df_part))

tab_rec, tab_exp, tab_val, tab_pre = st.tabs(
    ["1 · Recepción", "2 · Expedientes", "3 · Valoración", "4 · Prelación"])


# --- 1. RECEPCIÓN ---------------------------------------------------------
with tab_rec:
    st.subheader("Alta de solicitud")
    st.info("Una solicitud por docente. En el apartado de convocatorias se anotan "
            "todas aquellas por las que concursa (Art. 10 inciso b).")

    folio_nuevo = siguiente_folio(df_part)
    st.write(f"**Folio asignado:** {folio_nuevo}")

    with st.form("alta", clear_on_submit=False):
        c1, c2, c3 = st.columns(3)
        nombre = c1.text_input("Nombre completo *")
        rfc = c2.text_input("R.F.C. *")
        curp = c3.text_input("C.U.R.P.")

        c1, c2, c3 = st.columns(3)
        cct = c1.selectbox("Centro de trabajo *", CCTS)
        correo = c2.text_input("Correo electrónico")
        telefono = c3.text_input("Teléfono")

        c1, c2, c3, c4 = st.columns(4)
        categoria = c1.selectbox("Categoría actual *", CATEGORIAS)
        nivel = c2.selectbox("Nivel actual *", NIVELES)
        horas_act = c3.number_input("Horas actuales *", min_value=0, max_value=48, step=1)
        # El tope de 15 horas es del Art. 14 Primera b) reformado y de la Base PRIMERA.10.
        horas_sol = c4.number_input("Horas solicitadas *", min_value=1, max_value=15, step=1)

        fecha_prom = st.text_input("Fecha de última promoción (dd/mm/aaaa)")

        st.markdown("**Requisitos de participación**")
        req_vals = {}
        for clave, texto, fund in REQUISITOS:
            req_vals[clave] = st.radio(
                f"{texto}  ·  _{fund}_",
                ["Sí", "No", "No aplica"], horizontal=True, index=0, key=f"n_{clave}")

        st.markdown("**Documentos entregados**")
        st.caption("Original para cotejo, dos copias legibles y respaldo digital (Base SEGUNDA).")
        doc_vals = {}
        for clave, texto in DOCUMENTOS:
            cols = st.columns([5, 1, 1, 1])
            cols[0].write(texto)
            o = cols[1].checkbox("Original", key=f"n_{clave}_o")
            cp = cols[2].checkbox("2 copias", key=f"n_{clave}_c")
            dg = cols[3].checkbox("Digital", key=f"n_{clave}_d")
            doc_vals[clave] = "".join([("O" if o else ""), ("C" if cp else ""), ("D" if dg else "")])

        estatus = st.selectbox("Estatus de la solicitud", ESTATUS)
        observaciones = st.text_area("Observaciones y documentos faltantes")

        enviar = st.form_submit_button("Registrar solicitud", type="primary")

    if enviar:
        faltan = [n for n, v in [("nombre", nombre), ("R.F.C.", rfc)] if not v.strip()]
        if faltan:
            st.error("Falta capturar: " + ", ".join(faltan))
        elif not capturo.strip():
            st.error("Anota en la barra lateral quién está capturando.")
        elif rfc.strip().upper() in df_part["rfc"].astype(str).str.upper().values:
            st.error(f"Ese R.F.C. ya tiene folio registrado. "
                     "La solicitud es por docente: edítala en la pestaña de Expedientes.")
        else:
            registro = {
                "folio": folio_nuevo, "fecha_recepcion": ahora(),
                "nombre": nombre.strip().upper(), "rfc": rfc.strip().upper(),
                "curp": curp.strip().upper(), "cct": cct, "correo": correo.strip(),
                "telefono": telefono.strip(), "categoria_actual": categoria,
                "nivel_actual": nivel, "horas_actuales": horas_act,
                "fecha_ultima_promocion": fecha_prom.strip(),
                "horas_solicitadas": horas_sol, "estatus": estatus,
                "observaciones": observaciones.strip(),
                "capturo": capturo.strip(), "actualizado": ahora(),
                **req_vals, **doc_vals,
            }
            agregar("PARTICIPANTES", registro)
            st.success(f"Solicitud {folio_nuevo} registrada. "
                       "Captura sus convocatorias en la pestaña de Expedientes.")
            st.rerun()


# --- 2. EXPEDIENTES -------------------------------------------------------
with tab_exp:
    st.subheader("Expedientes registrados")
    if df_part.empty:
        st.info("Todavía no hay solicitudes registradas.")
    else:
        st.dataframe(
            df_part[["folio", "fecha_recepcion", "nombre", "cct", "categoria_actual",
                     "nivel_actual", "horas_actuales", "horas_solicitadas", "estatus"]],
            width='stretch', hide_index=True)

        folio = st.selectbox(
            "Expediente a editar",
            df_part["folio"].tolist(),
            format_func=lambda f: f"{f} — {df_part.loc[df_part.folio == f, 'nombre'].iloc[0]}")
        fila = df_part[df_part.folio == folio].iloc[0].to_dict()

        st.markdown("#### Convocatorias por las que participa")
        previas = df_conv[df_conv.folio == folio][
            ["no_convocatoria", "categoria", "nivel", "clave_presupuestal", "horas_vacante"]]
        if previas.empty:
            previas = pd.DataFrame([{"no_convocatoria": "", "categoria": "", "nivel": "",
                                     "clave_presupuestal": "", "horas_vacante": ""}])
        editadas = st.data_editor(previas, num_rows="dynamic", width='stretch',
                                  key=f"conv_{folio}")
        if st.button("Guardar convocatorias", key=f"bconv_{folio}"):
            registros = [{"folio": folio, **r} for _, r in editadas.iterrows()
                         if str(r.get("no_convocatoria", "")).strip()]
            reemplazar_por_folio("CONVOCATORIAS", folio, registros)
            st.success(f"{len(registros)} convocatoria(s) guardada(s) para {folio}.")
            st.rerun()

        st.markdown("#### Estatus y observaciones")
        c1, c2 = st.columns([1, 3])
        nuevo_est = c1.selectbox("Estatus", ESTATUS,
                                 index=ESTATUS.index(fila["estatus"])
                                 if fila.get("estatus") in ESTATUS else 0,
                                 key=f"est_{folio}")
        nueva_obs = c2.text_area("Observaciones", value=str(fila.get("observaciones", "")),
                                 key=f"obs_{folio}")
        st.markdown("##### Documentos")
        doc_edit = {}
        for clave, texto in DOCUMENTOS:
            actual = str(fila.get(clave, ""))
            cols = st.columns([5, 1, 1, 1])
            cols[0].write(texto)
            o = cols[1].checkbox("Original", value="O" in actual, key=f"e_{folio}_{clave}_o")
            cp = cols[2].checkbox("2 copias", value="C" in actual, key=f"e_{folio}_{clave}_c")
            dg = cols[3].checkbox("Digital", value="D" in actual, key=f"e_{folio}_{clave}_d")
            doc_edit[clave] = "".join([("O" if o else ""), ("C" if cp else ""), ("D" if dg else "")])

        if st.button("Guardar cambios del expediente", key=f"bexp_{folio}", type="primary"):
            fila.update(doc_edit)
            fila["estatus"] = nuevo_est
            fila["observaciones"] = nueva_obs
            fila["actualizado"] = ahora()
            actualizar_participante(folio, fila)
            st.success("Expediente actualizado.")
            st.rerun()


# --- 3. VALORACIÓN --------------------------------------------------------
with tab_val:
    st.subheader("Valoración conforme al Tabulador de Factores y Puntajes")
    st.caption("Título Octavo, Art. 43 del Reglamento. Sólo se valoran actividades "
               "posteriores a la última promoción (Base PRIMERA.5).")

    if df_part.empty:
        st.info("Todavía no hay solicitudes registradas.")
    else:
        folio_v = st.selectbox(
            "Expediente a valorar",
            df_part["folio"].tolist(),
            format_func=lambda f: f"{f} — {df_part.loc[df_part.folio == f, 'nombre'].iloc[0]}",
            key="folio_val")

        # Cantidades ya capturadas para este folio.
        previos = df_punt[df_punt.folio == folio_v]
        cant = {r["concepto"]: float(r["cantidad"]) for _, r in previos.iterrows()
                if str(r["cantidad"]).strip() not in ("", "0")}

        nuevas = {}
        for factor in tb.FACTORES:
            with st.expander(factor.etiqueta, expanded=False):
                if factor.nota:
                    st.caption(factor.nota)
                for c in factor.conceptos:
                    etiqueta = f"{c.etiqueta} — {c.puntos} pts"
                    if c.tope_cantidad:
                        etiqueta += f" (máx. {c.tope_cantidad})"
                    nuevas[c.clave] = st.number_input(
                        etiqueta, min_value=0.0, step=1.0,
                        value=float(cant.get(c.clave, 0)),
                        key=f"pt_{folio_v}_{c.clave}")

        desg = tb.desglose(nuevas)
        st.markdown("#### Puntaje")
        cols = st.columns(len(tb.FACTORES) + 1)
        for i, factor in enumerate(tb.FACTORES):
            cols[i].metric(factor.clave, desg[factor.clave])
        cols[-1].metric("TOTAL", desg["TOTAL"])

        # Avisos sobre los dos puntos que el Reglamento no resuelve solo.
        capturados_otros = [c.clave for c in tb.F_OTROS.conceptos if nuevas.get(c.clave, 0) > 0]
        if len(capturados_otros) > 1:
            st.warning("Se capturó más de un grado en 'Otros estudios'. El Reglamento no "
                       "dice si son acumulables entre sí; asienta el criterio en el acta.")
        if tb.puntos_factor(tb.F_ACTUALIZACION, nuevas) == tb.F_ACTUALIZACION.tope_puntos:
            st.info("El inciso b) llegó a su tope de 10 puntos.")

        if st.button("Guardar valoración", type="primary", key=f"bval_{folio_v}"):
            if not capturo.strip():
                st.error("Anota en la barra lateral quién está capturando.")
            else:
                registros = [{"folio": folio_v, "concepto": k, "cantidad": v,
                              "capturo": capturo.strip(), "actualizado": ahora()}
                             for k, v in nuevas.items() if v > 0]
                reemplazar_por_folio("PUNTAJES", folio_v, registros)
                st.success(f"Valoración guardada: {desg['TOTAL']} puntos.")
                st.rerun()


# --- 4. PRELACIÓN ---------------------------------------------------------
with tab_pre:
    st.subheader("Lista de prelación")
    st.caption("Orden por puntaje total; desempate por formación académica y, en segundo "
               "término, por antigüedad en el nivel superior de la SEJ (Art. 24 inciso i).")

    # Sólo entran a prelación los expedientes que no quedaron improcedentes.
    elegibles = df_part[df_part["estatus"] != "IMPROCEDENTE"] if not df_part.empty else df_part

    if elegibles.empty:
        st.info("No hay expedientes elegibles todavía.")
    else:
        participantes = []
        for _, r in elegibles.iterrows():
            pts = df_punt[df_punt.folio == r["folio"]]
            cantidades = {p["concepto"]: float(p["cantidad"]) for _, p in pts.iterrows()
                          if str(p["cantidad"]).strip() not in ("", "0")}
            convs = df_conv[df_conv.folio == r["folio"]]["no_convocatoria"].astype(str).tolist()
            participantes.append({
                "folio": r["folio"], "nombre": r["nombre"], "rfc": r["rfc"],
                "cct": r["cct"], "categoria_actual": r["categoria_actual"],
                "nivel_actual": r["nivel_actual"], "horas_actuales": r["horas_actuales"],
                "horas_solicitadas": r["horas_solicitadas"],
                "convocatorias": " | ".join(convs),
                "valorado": "Sí" if cantidades else "No",
                "cantidades": cantidades,
            })

        orden = tb.ordenar_prelacion(participantes)
        tabla = pd.DataFrame([{
            "lugar": e["lugar"], "folio": e["folio"], "nombre": e["nombre"],
            "rfc": e["rfc"], "cct": e["cct"],
            "categoria_actual": e["categoria_actual"], "nivel_actual": e["nivel_actual"],
            "horas_actuales": e["horas_actuales"], "horas_solicitadas": e["horas_solicitadas"],
            "puntos_total": e["total"], "puntos_formacion": e["formacion"],
            "antiguedad_anios": e["antiguedad"], "convocatorias": e["convocatorias"],
            "valorado": e["valorado"], "empate": "Sí" if e["empate_sin_resolver"] else "",
        } for e in orden])

        sin_valorar = tabla[tabla.valorado == "No"]
        if not sin_valorar.empty:
            st.warning(f"{len(sin_valorar)} expediente(s) sin valorar aparecen con 0 puntos.")
        if (tabla.empate == "Sí").any():
            st.warning("Hay empates que los tres criterios no resuelven. "
                       "La Comisión debe resolverlos y asentarlo en el acta.")

        st.dataframe(tabla, width='stretch', hide_index=True)

        # Exportación. Es el insumo del sistema de asignación: verifica los nombres
        # de columna contra la pestaña que ese sistema espera antes de cargarlo.
        st.download_button(
            "Descargar prelación (CSV)",
            tabla.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"prelacion_incremento_2026_{datetime.now(TZ).strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv", type="primary")

        st.download_button(
            "Descargar Excel para el sistema de asignación",
            excel_para_asignacion(orden),
            file_name=f"prelacion_para_asignacion_{datetime.now(TZ).strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Se carga tal cual en el sistema de asignación. Las horas de "
                   "asesor, de otras áreas y de base van en blanco: salen de nómina "
                   "y se capturan allá.")

        # Concentrado del proceso que pide el oficio CDESE/017/2026, inciso d).
        concentrado = {
            "proceso": "Concurso Cerrado de Incremento de Horas 2026",
            "generado": ahora(),
            "solicitudes_recibidas": int(len(df_part)),
            "elegibles": int(len(elegibles)),
            "improcedentes": int(len(df_part) - len(elegibles)),
            "convocatorias_solicitadas": int(len(df_conv)),
            "expedientes_valorados": int((tabla.valorado == "Sí").sum()),
        }
        st.download_button(
            "Descargar concentrado del proceso (JSON)",
            json.dumps(concentrado, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name="concentrado_incremento_2026.json", mime="application/json")

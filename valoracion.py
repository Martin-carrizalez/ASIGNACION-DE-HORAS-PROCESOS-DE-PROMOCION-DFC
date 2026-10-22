# -*- coding: utf-8 -*-
"""
Registro y valoración — Concurso Cerrado de Incremento de Horas
Dirección de Formación Continua · Comisión Dictaminadora Interna

Captura los documentos de cada participante y calcula su puntaje con el
Tabulador de Factores y Puntajes (Reglamento, Título Octavo, artículo 43).
La salida alimenta la app de asignación.

El código no contiene datos personales: el padrón se carga en el momento y
todo se procesa en memoria.
"""

import base64
import io
import json
import pathlib
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Valoración · DFC", layout="wide")

# ---------------------------------------------------------------------------
# Tabulador de Factores y Puntajes — Reglamento, Título Octavo, artículo 43
#
# Cada concepto es: (etiqueta, puntos por unidad, tope de unidades, grupo)
#   tope  = máximo de unidades que el Reglamento permite contar (None = sin tope)
#   grupo = etiqueta interna para las reglas especiales de más abajo
# ---------------------------------------------------------------------------

TABULADOR = {
    "Formación académica": [
        # "El puntaje no es acumulable entre sí": de cada grupo cuenta solo el mayor.
        ("Licenciatura (estudios en educación)", 20, 1, "educacion"),
        ("Maestría en Educación", 30, 1, "educacion"),
        ("Doctorado en Educación con antecedente de Maestría", 40, 1, "educacion"),
        # "Este puntaje se adiciona a Estudios en Educación".
        ("Licenciatura (estudios no relacionados con educación)", 15, 1, "otros"),
        ("Maestría con antecedente de Licenciatura (otros estudios)", 25, 1, "otros"),
        ("Doctorado con antecedente de Maestría (otros estudios)", 35, 1, "otros"),
    ],
    "Actualización profesional": [
        ("Diplomado acreditado y con reconocimiento oficial", 3, None, None),
        ("Especialidad acreditada y con reconocimiento oficial", 5, None, None),
        ("Seminario, curso o taller como asistente — 40 o más horas", 1.25, None, "asistente"),
        ("Seminario, curso o taller como asistente — 30 a 39 horas", 1, None, "asistente"),
        ("Seminario, curso o taller como asistente — 20 a 29 horas", 0.5, None, "asistente"),
        ("Seminario, curso o taller como asistente — menos de 20 horas", 0.25, None, "asistente"),
    ],
    "Antigüedad": [
        ("Por cada año de servicio en el Nivel Superior", 1, None, None),
    ],
    "Experiencia profesional y/o académica": [
        ("Curso de actualización impartido — menos de 30 horas", 2.5, None, None),
        ("Curso de actualización impartido — de 30 horas o más", 5, None, None),
        ("Asignatura en el Nivel Superior en instituciones que regula el Reglamento", 5, None, None),
    ],
    "Productividad profesional": [
        ("Libro científico o pedagógico publicado — autor", 5, None, None),
        ("Libro científico o pedagógico publicado — coautor", 2.5, None, None),
        ("Libro de texto publicado — autor", 5, None, None),
        ("Libro de texto publicado — coautor", 2.5, None, None),
        ("Proyecto interinstitucional de la SEJ — autor", 2.0, None, None),
        ("Proyecto interinstitucional de la SEJ — coautor", 1, None, None),
        ("Proyecto institucional — autor", 1, None, None),
        ("Proyecto institucional — coautor", 0.5, None, None),
        ("Artículo educativo publicado en revista o periódico", 0.15, None, None),
        ("Grupo de tutoría atendido", 0.5, None, None),
        ("Participar en Cuerpo Académico sin perfil PROMEP", 2, None, None),
        ("Contar con Perfil PROMEP", 3, None, None),
        ("Cuerpo Académico con Línea de Generación de Conocimiento y perfil PROMEP", 2, None, None),
        ("Coordinar un CA con Línea de Generación de Conocimiento", 3, None, None),
        ("Diseño y ejecución del Plan de Desarrollo Institucional", 2, None, None),
        ("Asesoría de tesis a nivel Maestría", 0.75, 5, None),
        ("Asesoría de tesis a nivel Licenciatura", 0.5, 10, None),
        ("Sinodal a nivel Licenciatura y Maestría", 0.25, 10, None),
        ("Diseño de curso o taller — de 40 horas a más", 2, None, None),
        ("Diseño de curso o taller — de 30 a 39 horas", 1.5, None, None),
        ("Diseño de curso o taller — de 20 a 29 horas", 1, None, None),
        ("Diseño de curso o taller — de 10 a 19 horas", 0.5, None, None),
    ],
    "Congresos, simposios, coloquios y foros": [
        ("Constancia — Internacional", 1.5, None, None),
        ("Constancia — Nacional", 1.2, None, None),
        ("Constancia — Regional", 1.0, None, None),
        ("Constancia — Estatal", 0.8, None, None),
        ("Constancia — Municipal", 0.5, None, None),
        ("Constancia — Institucional", 0.2, None, None),
    ],
    "Diplomados": [
        ("Autoría en el diseño de diplomados autorizados por la DG de Formación Continua",
         5, None, None),
        ("Autoría de especialidad autorizada por la Dirección de Posgrado / UPN",
         7, None, None),
    ],
}

# Tope de puntos por grupo. La nota "Sólo se considera máximo diez puntos" va
# justo debajo de la tabla de seminarios, cursos y talleres como asistente.
TOPES_DE_GRUPO = {"asistente": 10}

# Grupos donde el puntaje no se acumula: solo cuenta el concepto de mayor valor.
GRUPOS_NO_ACUMULABLES = ("educacion", "otros")

APARTADOS = list(TABULADOR)
VACIOS = ("", "nan", "none", "n/a", "na", "#n/a", "-", "s/d", "sd")


def limpiar(v):
    if v is None:
        return ""
    t = re.sub(r"\s+", " ", str(v)).strip()
    return "" if t.lower() in VACIOS else t


def numero(v, por_defecto=None):
    try:
        if limpiar(v) == "":
            return por_defecto
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return por_defecto


# ---------------------------------------------------------------------------
# Cálculo del puntaje
# ---------------------------------------------------------------------------

def conceptos_de(apartado):
    return {c[0]: c for c in TABULADOR[apartado]}


def puntos_apartado(apartado, renglones):
    """renglones = [{'Concepto':…, 'Detalle':…, 'Cantidad':…}, …] -> puntos del apartado."""
    catalogo = conceptos_de(apartado)
    por_grupo = {}          # grupo -> puntos acumulados
    sueltos = 0.0
    mejor_no_acumulable = {}   # grupo no acumulable -> mayor puntaje unitario

    for r in renglones:
        concepto = limpiar(r.get("Concepto"))
        if concepto not in catalogo:
            continue
        _, valor, tope, grupo = catalogo[concepto]
        cantidad = numero(r.get("Cantidad"), 0) or 0
        if cantidad <= 0:
            continue
        if tope is not None:
            cantidad = min(cantidad, tope)

        if grupo in GRUPOS_NO_ACUMULABLES:
            # Solo cuenta el grado más alto que acredite, una vez.
            mejor_no_acumulable[grupo] = max(mejor_no_acumulable.get(grupo, 0), valor)
        elif grupo:
            por_grupo[grupo] = por_grupo.get(grupo, 0) + valor * cantidad
        else:
            sueltos += valor * cantidad

    total = sueltos + sum(mejor_no_acumulable.values())
    for grupo, puntos in por_grupo.items():
        total += min(puntos, TOPES_DE_GRUPO.get(grupo, puntos))
    return round(total, 2)


def puntajes_de(expediente):
    """expediente = {apartado: [renglones]} -> {apartado: puntos} y el total."""
    por_apartado = {a: puntos_apartado(a, expediente.get(a, [])) for a in APARTADOS}
    return por_apartado, round(sum(por_apartado.values()), 2)


# ---------------------------------------------------------------------------
# Identidad visual (la misma de la app de asignación)
# ---------------------------------------------------------------------------

RECURSOS = pathlib.Path(__file__).parent / "recursos"


@st.cache_data(show_spinner=False)
def _imagen(nombre):
    ruta = RECURSOS / nombre
    return base64.b64encode(ruta.read_bytes()).decode() if ruta.exists() else ""


def identidad():
    st.markdown("""
<style>
  .stApp { background:
      radial-gradient(1100px 380px at 10% -6%, #FFF1E4 0%, rgba(255,241,228,0) 58%),
      radial-gradient(900px 340px at 92% -4%, #E9F8FD 0%, rgba(233,248,253,0) 52%), #FBFAF8; }
  .block-container { padding-top: 3.2rem; max-width: 1280px; }
  .marco-logo { position: relative; width: 100%; margin: 0 0 6px; }
  .capa { width: 100%; display: block; }
  .capa-leds { position: absolute; top: 0; left: 0;
               animation: destello 2.6s ease-in-out infinite; }
  @keyframes destello {
      0%, 100% { opacity: 1;   filter: brightness(1.5) drop-shadow(0 0 6px #22C7F0); }
      45%, 55% { opacity: .3;  filter: brightness(.85); } }
  @media (prefers-reduced-motion: reduce) { .capa-leds { animation: none; } }
  .pie-marca { color:#6B7688; font-size:.76rem; text-align:center; margin:0 0 14px; }
  .titulo { font-size:1.7rem; font-weight:700; color:#1B2432; margin:0; }
  .bajada { color:#6B7688; font-size:.95rem; margin:4px 0 0; }
  .proceso { display:inline-block; margin:10px 0 4px; padding:4px 12px; border-radius:999px;
             font-size:.72rem; font-weight:700; letter-spacing:.6px; text-transform:uppercase;
             color:#8A3B00; background:linear-gradient(90deg,#FFE2C7,#FFF0DF);
             border:1px solid #FFD2AC; }
  .stTabs [data-baseweb="tab-list"] { gap:4px; border-bottom:1px solid #E3E6EC; }
  .stTabs [data-baseweb="tab"] { height:44px; padding:0 18px; font-weight:600; color:#6B7688; }
  .stTabs [aria-selected="true"] { color:#F2720C; }
  .stTabs [data-baseweb="tab-highlight"] { background:#F2720C; height:3px; }
  div[data-testid="stMetric"] { background:#fff; border:1px solid #E3E6EC; border-radius:12px;
      padding:14px 18px; box-shadow:0 1px 2px rgba(27,36,50,.05); }
  div[data-testid="stMetric"] [data-testid="stMetricValue"] { color:#F2720C; font-weight:700; }
  .stButton > button, .stDownloadButton > button { border-radius:9px; font-weight:600; }
  .stButton > button[kind="primary"] {
      background:linear-gradient(135deg,#F2720C,#FF9A3C); border:0; color:#fff; }
  .stDownloadButton > button { width:100%; text-align:left; }
  section[data-testid="stSidebar"] { background:#fff; border-right:1px solid #E3E6EC; }
  section[data-testid="stSidebar"] h2 { font-size:.76rem; text-transform:uppercase;
      letter-spacing:.8px; color:#F2720C; margin-top:1.1rem; }
  .tarjeta { background:#fff; border:1px solid #E3E6EC; border-left:3px solid #F2720C;
             border-radius:10px; padding:12px 16px; margin-bottom:10px; }
  .tarjeta span { color:#6B7688; font-size:.88rem; }
</style>
""", unsafe_allow_html=True)


def logotipo():
    base, leds = _imagen("dfc_logo.png"), _imagen("dfc_leds.png")
    if not base:
        st.sidebar.warning("No encuentro recursos/dfc_logo.png junto a valoracion.py.")
        return
    capas = '<img src="data:image/png;base64,%s" class="capa">' % base
    if leds:
        capas += '<img src="data:image/png;base64,%s" class="capa capa-leds">' % leds
    st.sidebar.markdown('<div class="marco-logo">%s</div>'
                        '<p class="pie-marca">Comisión Dictaminadora Interna</p>' % capas,
                        unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Padrón de participantes y estado de la sesión
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def cargar_padron(contenido):
    df = pd.read_excel(io.BytesIO(contenido)).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in ("rfc", "nombre") if c not in df.columns]
    if faltan:
        raise ValueError("Al Excel le faltan columnas: " + ", ".join(faltan))
    df["rfc"] = df["rfc"].map(limpiar)
    return df[df["rfc"] != ""].reset_index(drop=True)


@st.cache_resource(show_spinner=False)
def _almacen():
    """El avance vive en el proceso del servidor: sobrevive a recargar la página
    y a que se caiga la conexión. No sobrevive a un reinicio: para eso está el
    respaldo en JSON."""
    return {}


def estado():
    clave = st.session_state.get("clave_sesion") or "valoracion"
    return _almacen().setdefault(clave, {"expedientes": {}})


def expediente_de(rfc):
    return estado()["expedientes"].setdefault(rfc, {a: [] for a in APARTADOS})


identidad()

with st.sidebar:
    logotipo()
    st.header("Archivos")
    f_pad = st.file_uploader("Excel del padrón de participantes", type="xlsx")

    st.header("Sesión")
    st.text_input("Nombre de la sesión de trabajo", "valoracion", key="clave_sesion",
                  help="El avance se guarda bajo este nombre; si se recarga la página "
                       "se recupera al volver a entrar con el mismo nombre.")

    st.header("Parámetros")
    anio = st.text_input("Año del proceso", "2026")

if not f_pad:
    st.info("Carga el Excel del padrón para comenzar. Puede ser el mismo archivo de "
            "prelación: necesita al menos las columnas rfc y nombre.")
    st.stop()

try:
    padron = cargar_padron(f_pad.getvalue())
except ValueError as e:
    st.error(str(e))
    st.stop()

st.markdown('<p class="titulo">Registro y valoración</p>'
            '<p class="bajada">Captura de documentos y cálculo de puntaje con el '
            'Tabulador de Factores y Puntajes (Reglamento, art. 43).</p>'
            '<div class="proceso">Incremento de horas %s</div>' % anio,
            unsafe_allow_html=True)

t_val, t_prel = st.tabs(["Valoración", "Prelación"])

# ---------------------------------------------------------------------------
# Valoración
# ---------------------------------------------------------------------------

with t_val:
    etiquetas = ["%s — %s" % (r["nombre"], r["rfc"]) for _, r in padron.iterrows()]
    i = st.selectbox("Participante", range(len(etiquetas)),
                     format_func=lambda k: etiquetas[k])
    persona = padron.iloc[i]
    rfc = persona["rfc"]
    exp = expediente_de(rfc)

    por_apartado, total = puntajes_de(exp)
    st.markdown('<div class="tarjeta"><b>%s</b> · %s<br><span>%s</span></div>'
                % (persona["nombre"], rfc, limpiar(persona.get("puesto", ""))),
                unsafe_allow_html=True)

    c = st.columns(4)
    c[0].metric("Puntaje total", "%.2f" % total)
    c[1].metric("Documentos capturados",
                sum(len([r for r in exp[a] if limpiar(r.get("Concepto"))]) for a in APARTADOS))
    c[2].metric("Apartados con puntaje", sum(1 for a in APARTADOS if por_apartado[a] > 0))
    c[3].metric("Participantes valorados",
                sum(1 for e in estado()["expedientes"].values()
                    if puntajes_de(e)[1] > 0))

    st.caption("Agrega un renglón por documento. En Detalle escribe el nombre del "
               "diplomado, curso o publicación; en Cantidad, cuántos acreditas de ese "
               "mismo concepto.")

    for apartado in APARTADOS:
        opciones = [c[0] for c in TABULADOR[apartado]]
        with st.expander("%s — %.2f puntos" % (apartado, por_apartado[apartado]),
                         expanded=por_apartado[apartado] > 0):
            base = pd.DataFrame(exp[apartado] or [],
                                columns=["Concepto", "Detalle", "Cantidad"])
            editado = st.data_editor(
                base, num_rows="dynamic", use_container_width=True, hide_index=True,
                key="ed_%s_%s" % (rfc, apartado),
                column_config={
                    "Concepto": st.column_config.SelectboxColumn(
                        "Concepto", options=opciones, width="large", required=True),
                    "Detalle": st.column_config.TextColumn(
                        "Detalle (nombre del documento)", width="medium"),
                    "Cantidad": st.column_config.NumberColumn(
                        "Cantidad", min_value=0, step=1, default=1, width="small"),
                })
            renglones = editado.to_dict("records")
            if renglones != exp[apartado]:
                exp[apartado] = renglones
                st.rerun()

            # Desglose de lo que aporta cada renglón, con los topes ya aplicados.
            if renglones:
                catalogo = conceptos_de(apartado)
                detalle = []
                for r in renglones:
                    con = limpiar(r.get("Concepto"))
                    if con not in catalogo:
                        continue
                    _, valor, tope, _ = catalogo[con]
                    cant = numero(r.get("Cantidad"), 0) or 0
                    usada = min(cant, tope) if tope is not None else cant
                    detalle.append({"Concepto": con, "Detalle": limpiar(r.get("Detalle")),
                                    "Unitario": valor, "Cantidad": cant,
                                    "Cuenta": usada, "Puntos": round(valor * usada, 2)})
                if detalle:
                    st.dataframe(pd.DataFrame(detalle), hide_index=True,
                                 use_container_width=True)

# ---------------------------------------------------------------------------
# Prelación
# ---------------------------------------------------------------------------

def tabla_prelacion():
    """Una fila por participante, una columna por apartado y el total al final."""
    filas = []
    for _, r in padron.iterrows():
        exp = estado()["expedientes"].get(r["rfc"])
        if not exp:
            por_apartado, total = {a: 0.0 for a in APARTADOS}, 0.0
        else:
            por_apartado, total = puntajes_de(exp)
        fila = {"RFC": r["rfc"], "Nombre": r["nombre"],
                "Categoría": limpiar(r.get("puesto", "")).replace("HORAS DE ", "")}
        fila.update({a: por_apartado[a] for a in APARTADOS})
        fila["TOTAL"] = total
        filas.append(fila)
    t = pd.DataFrame(filas).sort_values("TOTAL", ascending=False).reset_index(drop=True)
    t.insert(0, "Folio", t.index + 1)
    return t


def excel_de(hojas):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

with t_prel:
    tabla = tabla_prelacion()
    st.caption("Orden por puntaje del Tabulador, de mayor a menor. La categoría no "
               "interviene en el orden: el Adendum (art. 14) permite incrementar "
               "horas sin importar la categoría y nivel que se ostente.")
    st.dataframe(tabla, hide_index=True, use_container_width=True)

    valorados = int((tabla["TOTAL"] > 0).sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("Participantes", len(tabla))
    c2.metric("Con puntaje", valorados)
    c3.metric("Sin valorar", len(tabla) - valorados)

    empates = tabla[tabla["TOTAL"] > 0].groupby("TOTAL").size()
    empatados = {k: int(v) for k, v in empates[empates > 1].items()}
    if empatados:
        st.warning("Hay empates de puntaje en: %s. La convocatoria manda desempatar "
                   "por formación académica y después por antigüedad en el nivel "
                   "superior; la app todavía no lo resuelve."
                   % ", ".join("%.2f (%d personas)" % (k, v) for k, v in empatados.items()))

    st.divider()
    st.download_button("Excel · prelación con el desglose por apartado",
                       excel_de({"prelación": tabla}),
                       "prelacion_%s.xlsx" % anio, MIME_XLSX)

    # Insumo para la app de asignación: el padrón original con el Puntaje calculado.
    insumo = padron.copy()
    puntajes = dict(zip(tabla["RFC"], tabla["TOTAL"]))
    col_pun = next((c for c in insumo.columns if c.lower() == "puntaje"), "Puntaje")
    insumo[col_pun] = insumo["rfc"].map(puntajes)
    st.download_button("Excel · insumo para la app de asignación",
                       excel_de({"prelación": insumo}),
                       "insumo_asignacion_%s.xlsx" % anio, MIME_XLSX)
    st.caption("Es el mismo padrón que cargaste, con la columna Puntaje recalculada. "
               "Ese archivo es el que lee la app de asignación.")

    st.divider()
    st.subheader("Respaldo del avance")
    st.caption("La captura se conserva mientras el servidor siga encendido. Descarga "
               "este archivo de vez en cuando: devuelve los expedientes tal como iban.")
    c1, c2 = st.columns(2)
    c1.download_button(
        "Descargar captura (JSON)",
        json.dumps(estado()["expedientes"], ensure_ascii=False, indent=1).encode("utf-8"),
        "valoracion_%s.json" % anio, "application/json")
    respaldo = c2.file_uploader("Restaurar captura", type="json",
                                label_visibility="collapsed")
    if respaldo is not None and c2.button("Restaurar este archivo"):
        estado()["expedientes"] = json.loads(respaldo.getvalue().decode("utf-8"))
        st.rerun()

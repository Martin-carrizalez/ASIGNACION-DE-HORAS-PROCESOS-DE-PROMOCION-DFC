"""
Valoración y prelación del Concurso Cerrado de Incremento de Horas 2026
Comisión Dictaminadora Interna — Dirección de Formación Continua / CEPD

Dos etapas, las que siguen vigentes del calendario de la Base SÉPTIMA:
  1. VALORACIÓN  (22 al 28 de octubre) — captura del tabulador por expediente,
                 documento por documento.
  2. PRELACIÓN   — orden final y exportación que alimenta el sistema de asignación.

La recepción (7 al 21 de octubre) cerró, así que sus dos pestañas se retiraron.
Los participantes se leen de la pestaña PARTICIPANTES tal como quedó; esta
página ya no la escribe.

Persistencia: Google Sheets (gspread + cuenta de servicio en st.secrets), mismo
patrón que el resto de las apps del área. No se guarda nada en el repositorio.

Configuración requerida en .streamlit/secrets.toml:

    [gcp_service_account]
    ... (el bloque de la cuenta de servicio que ya usas)

    [registro_incremento]
    sheet_id = "ID_DE_LA_HOJA_NUEVA"

La hoja debe estar compartida como Editor con el correo de la cuenta de servicio.
"""

import io
import json
import os
import re
import sys
import unicodedata
from datetime import datetime
from difflib import get_close_matches

import gspread
import pandas as pd
import pytz
import streamlit as st
from google.oauth2.service_account import Credentials

# tabulador.py vive en la raíz del repo. Cuando esta página corre como página
# del multipágina, la raíz ya está en la ruta de búsqueda; cuando se despliega
# apuntando directo a este archivo, no. Esto cubre los dos casos.
_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

import tabulador as tb  # noqa: E402

TZ = pytz.timezone("America/Mexico_City")

# --- Catálogos de la convocatoria -----------------------------------------
# Se conservan porque definen las columnas de la pestaña PARTICIPANTES, que
# esta página sigue leyendo. El contrato con el Sheet no cambia.
ESTATUS = ["RECIBIDA", "RECIBIDA CON FALTANTES", "SUBSANADA", "IMPROCEDENTE"]

# Encabezados de cada pestaña. A la izquierda el nombre interno que usa el
# código; a la derecha el que se lee al abrir la hoja de cálculo. Nadie que
# abra el Sheet tiene por qué descifrar abreviaturas.
#
# Las columnas de requisitos y documentos de la recepción se retiraron: esa
# etapa ya no se captura desde aquí y sólo dejaban columnas vacías.
HOJAS = {
    "PARTICIPANTES": [
        ("folio", "Folio"),
        ("no_expediente", "No. de expediente"),
        ("nombre", "Nombre del participante"),
        ("rfc", "R.F.C."),
        ("curp", "C.U.R.P."),
        ("cct", "Centro de trabajo"),
        ("correo", "Correo electrónico"),
        ("telefono", "Teléfono"),
        ("categoria_actual", "Categoría actual"),
        ("nivel_actual", "Nivel actual"),
        ("horas_actuales", "Horas actuales"),
        ("fecha_ultima_promocion", "Fecha de última promoción"),
        ("horas_solicitadas", "Horas solicitadas"),
        ("estatus", "Estatus de la solicitud"),
        ("observaciones", "Observaciones"),
        ("fecha_recepcion", "Fecha de recepción"),
        ("capturo", "Capturó"),
        ("actualizado", "Última actualización"),
    ],
    "CONVOCATORIAS": [
        ("folio", "Folio"),
        ("no_convocatoria", "No. de convocatoria"),
        ("categoria", "Categoría"),
        ("nivel", "Nivel"),
        ("clave_presupuestal", "Clave presupuestal"),
        ("horas_vacante", "Horas de la vacante"),
    ],
    "PUNTAJES": [
        ("folio", "Folio"),
        ("apartado", "Apartado"),
        ("concepto_texto", "Concepto del tabulador"),
        ("concepto", "Clave del concepto"),
        ("cantidad", "Cantidad"),
        ("detalle", "Detalle del documento"),
        ("no_documento", "No. de documento"),
        ("notas", "Notas y observaciones"),
        ("sin_valor", "Documento sin valor"),
        ("capturo", "Capturó"),
        ("actualizado", "Última actualización"),
    ],
}

# Nombres con los que una columna pudo haberse guardado antes. Sirven para
# reacomodar una pestaña vieja sin perder lo capturado.
ANTERIORES = {"capturo": ("capturo",), "actualizado": ("actualizado",),
              "no_documento": ("no_documento",), "notas": ("notas",),
              "sin_valor": ("sin_valor",), "detalle": ("detalle",),
              "concepto": ("concepto",), "no_expediente": ("no_expediente",)}


def internos(nombre):
    return [a for a, _ in HOJAS[nombre]]


def visibles(nombre):
    return [b for _, b in HOJAS[nombre]]


def columna_interna(nombre, encabezado):
    """Del texto de un encabezado al nombre interno que le corresponde."""
    for interno, visible in HOJAS[nombre]:
        if encabezado == visible or encabezado == interno:
            return interno
        if encabezado in ANTERIORES.get(interno, ()):
            return interno
    return None


# =========================================================================
# Capa de datos
# =========================================================================

@st.cache_resource(show_spinner=False)
def abrir_libro():
    """
    Abre el Sheet y deja cada pestaña con los encabezados legibles.

    Si una pestaña viene de una versión anterior —con los nombres internos o
    sin las columnas nuevas— se reacomoda una sola vez: se leen sus renglones
    por nombre de columna y se vuelven a escribir en el orden actual. Nada se
    pierde y ninguna columna se adivina por posición.
    """
    alcances = ["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"]
    cred = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=alcances)
    libro = gspread.authorize(cred).open_by_key(
        st.secrets["registro_incremento"]["sheet_id"])

    existentes = {h.title for h in libro.worksheets()}
    for nombre in HOJAS:
        cabecera = visibles(nombre)
        if nombre not in existentes:
            hoja = libro.add_worksheet(title=nombre, rows=500,
                                       cols=len(cabecera) + 5)
            hoja.append_row(cabecera)
            continue

        hoja = libro.worksheet(nombre)
        valores = hoja.get_all_values()
        if not valores or not any(c.strip() for c in valores[0]):
            hoja.update([cabecera], "A1", value_input_option="USER_ENTERED")
            continue
        if valores[0] == cabecera:
            continue

        # Reacomodo: cada renglón se reconstruye por nombre de columna.
        viejos = [columna_interna(nombre, c.strip()) for c in valores[0]]
        filas = []
        for fila in valores[1:]:
            if not any(c.strip() for c in fila):
                continue
            por_interno = {k: (fila[j] if j < len(fila) else "")
                           for j, k in enumerate(viejos) if k}
            filas.append([por_interno.get(k, "") for k in internos(nombre)])
        if hoja.col_count < len(cabecera):
            hoja.add_cols(len(cabecera) - hoja.col_count)
        hoja.clear()
        hoja.update([cabecera] + filas, value_input_option="USER_ENTERED")
    return libro


def _parece_encabezado(nombre, fila):
    """¿El renglón 1 es un encabezado y no un registro?"""
    llenas = [c.strip() for c in fila if c.strip()]
    if not llenas:
        return True                      # vacío: se puede reponer sin riesgo
    return all(columna_interna(nombre, c) for c in llenas)


def _hoja(nombre):
    """
    Devuelve la pestaña y sus valores con el encabezado garantizado.

    Todas las escrituras dan por hecho que el renglón 1 es el encabezado; si
    se borró o quedó a medias, se repone aquí. Si el renglón 1 trae datos de
    verdad, se inserta uno nuevo arriba en vez de pisarlo.
    """
    hoja = abrir_libro().worksheet(nombre)
    valores = hoja.get_all_values()
    cabecera = visibles(nombre)
    if valores and valores[0] == cabecera:
        return hoja, valores
    if not valores:
        hoja.update([cabecera], "A1", value_input_option="USER_ENTERED")
        return hoja, [cabecera]
    if _parece_encabezado(nombre, valores[0]):
        hoja.update([cabecera], "A1", value_input_option="USER_ENTERED")
        valores[0] = cabecera
    else:
        hoja.insert_row(cabecera, 1)
        valores = [cabecera] + valores
    return hoja, valores


def leer(nombre):
    """Lee una pestaña completa con los nombres internos como columnas."""
    hoja, valores = _hoja(nombre)
    claves = [columna_interna(nombre, c.strip()) for c in valores[0]] if valores else []
    filas = []
    for fila in valores[1:]:
        if not any(c.strip() for c in fila):
            continue
        filas.append({k: (fila[j] if j < len(fila) else "")
                      for j, k in enumerate(claves) if k})
    df = pd.DataFrame(filas, columns=internos(nombre)) if filas \
        else pd.DataFrame(columns=internos(nombre))
    return df.fillna("")


def _renglon(nombre, registro):
    """Arma el renglón en el orden de la pestaña, llenando el texto legible
    del concepto a partir de su clave."""
    r = dict(registro)
    if nombre == "PUNTAJES":
        clave = str(r.get("concepto", "")).strip()
        # El nombre del concepto va tal cual lo escribe el Reglamento; el
        # apartado al que pertenece va en su propia columna, no pegado.
        r.setdefault("concepto_texto", "")
        r.setdefault("apartado", "")
        if not r["concepto_texto"]:
            r["concepto_texto"] = ETIQUETA.get(clave, "")
        if not r["apartado"]:
            r["apartado"] = APARTADO_DE.get(clave, "")
    return [str(r.get(k, "")) for k in internos(nombre)]


def agregar_varios(nombre, registros):
    """Agrega muchos renglones en una sola petición."""
    if not registros:
        return
    hoja, _ = _hoja(nombre)
    hoja.append_rows([_renglon(nombre, r) for r in registros],
                     value_input_option="USER_ENTERED")


def reemplazar_varios_folios(nombre, folios, registros, conceptos=None):
    """
    Reescribe la pestaña completa en UNA petición.

    De los folios indicados se quitan los renglones viejos —todos, o sólo los
    de los conceptos que se listen— y se agregan los nuevos. Lo de los demás
    folios no se toca.

    Va en bloque a propósito: una llamada por participante son más de cien
    peticiones seguidas y Google corta a las 60 escrituras por minuto.
    """
    hoja, valores = _hoja(nombre)
    cabecera = visibles(nombre)
    claves = internos(nombre)
    i_folio = claves.index("folio")
    i_concepto = claves.index("concepto") if "concepto" in claves else None

    objetivo = set(folios)
    conservados = []
    for fila in valores[1:]:
        if not any(c.strip() for c in fila):
            continue
        suyo = fila[i_folio] if len(fila) > i_folio else ""
        if suyo not in objetivo:
            conservados.append(fila)
            continue
        if conceptos is not None and i_concepto is not None:
            clave = fila[i_concepto].strip() if len(fila) > i_concepto else ""
            if clave not in conceptos:
                conservados.append(fila)        # de otro apartado: se respeta

    hoja.clear()
    hoja.update([cabecera] + conservados + [_renglon(nombre, r) for r in registros],
                value_input_option="USER_ENTERED")


def reemplazar_por_folio(nombre, folio, registros):
    """Borra los renglones de un folio y escribe los nuevos."""
    reemplazar_varios_folios(nombre, [folio], registros)


def actualizar_participante(folio, registro):
    """Reescribe el renglón de un participante conservando su posición."""
    hoja, valores = _hoja("PARTICIPANTES")
    i_folio = internos("PARTICIPANTES").index("folio")
    for n, fila in enumerate(valores[1:], start=2):
        if len(fila) > i_folio and fila[i_folio] == folio:
            hoja.update([_renglon("PARTICIPANTES", registro)], f"A{n}",
                        value_input_option="USER_ENTERED")
            return True
    return False


def siguiente_folio(df, usados_ya=()):
    """Llave interna consecutiva. No es el número de expediente."""
    todos = list(df["folio"].astype(str)) + list(usados_ya)
    n = [int(f.split("-")[-1]) for f in todos
         if f.startswith("INC-2026-") and f.split("-")[-1].isdigit()]
    return "INC-2026-%03d" % ((max(n) + 1) if n else 1)


def publicar_concentrado(tabla):
    """
    Escribe el concentrado horizontal en la pestaña RESULTADOS: una fila por
    participante, una columna por apartado y el Total al final.

    PUNTAJES guarda un renglón por documento porque cada uno lleva su número,
    su detalle y sus notas. Esta pestaña es la vista horizontal de eso mismo.
    """
    libro = abrir_libro()
    columnas = [str(c) for c in tabla.columns]
    try:
        hoja = libro.worksheet("RESULTADOS")
    except gspread.WorksheetNotFound:
        hoja = libro.add_worksheet(title="RESULTADOS", rows=max(len(tabla) + 20, 100),
                                   cols=len(columnas) + 5)
    if hoja.col_count < len(columnas):
        hoja.add_cols(len(columnas) - hoja.col_count)
    hoja.clear()
    hoja.update([columnas] + tabla.astype(str).values.tolist(),
                value_input_option="USER_ENTERED")
    return len(tabla)


def ahora():
    return datetime.now(TZ).strftime("%d/%m/%Y %H:%M")


def normalizar(texto):
    """Minúsculas, sin acentos y sin espacios de más, para poder comparar."""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


def es_numero(v):
    try:
        float(str(v).replace(",", ""))
        return True
    except (TypeError, ValueError):
        return False


# Traducción exacta de los textos de la hoja "Opciones" del formato 2025 a los
# conceptos del tabulador. No se deja al parecido: las etiquetas del Reglamento
# ("Doctorado en Educación con antecedente de Maestría (titulado)") y las del
# formato ("Doctorado en Educación") no se parecen lo suficiente, y un grado mal
# emparejado le mueve los puntos a alguien.
EQUIVALENCIAS_2025 = {
    "licenciatura en educacion": "LIC_EDU",
    "maestria en educacion": "MAE_EDU",
    "doctorado en educacion": "DOC_EDU",
    "licenciatura no relacionada con educacion": "LIC_OTR",
    "maestria no relacionada con educacion": "MAE_OTR",
    "doctorado no relacionado con educacion": "DOC_OTR",
}
SIN_GRADO = {"", "ninguno", "documento sin valor", "sin valor", "nan", "0"}


def emparejar_concepto(texto, nombres_por_clave, corte=0.80):
    """
    Traduce el texto del archivo de precarga al concepto del tabulador.
    Devuelve None si no hay una coincidencia clara: más vale dejarlo en blanco
    y capturarlo a mano que precargar un grado equivocado.
    """
    objetivo = normalizar(texto)
    if objetivo in SIN_GRADO:
        return None
    if objetivo in EQUIVALENCIAS_2025:
        return EQUIVALENCIAS_2025[objetivo]
    normalizados = {normalizar(v): k for k, v in nombres_por_clave.items()}
    if objetivo in normalizados:
        return normalizados[objetivo]
    parecidos = get_close_matches(objetivo, list(normalizados), n=1, cutoff=corte)
    return normalizados[parecidos[0]] if parecidos else None


# =========================================================================
# Puente con el tabulador
#
# El tabulador razona con cantidades por concepto: {"CURSO_40": 3.0, ...}.
# La captura de esta pantalla es documento por documento, un renglón por
# constancia. Estas funciones traducen de una forma a la otra, para no
# tocar tabulador.py ni romper lo ya capturado en el Sheet.
# =========================================================================

def _nombres_legibles():
    """
    Nombre visible de cada concepto: su propia etiqueta del artículo 43, tal
    cual se lee. Las claves internas (LIC_OTR, CURSO_40) no salen a pantalla
    ni a los encabezados de la base: quien descarga el archivo no tiene por
    qué saberlas.

    Si dos conceptos se llaman igual —el 43 repite "Autor" y "Coautor" en
    varios incisos— se desambigua anteponiendo el factor y, si todavía
    chocan, con el puntaje.
    """
    veces = {}
    for f in tb.FACTORES:
        for c in f.conceptos:
            veces[c.etiqueta.strip()] = veces.get(c.etiqueta.strip(), 0) + 1

    nombres, usados = {}, set()
    for f in tb.FACTORES:
        for c in f.conceptos:
            etiqueta = c.etiqueta.strip()
            nombre = etiqueta if veces[etiqueta] == 1 else "%s — %s" % (f.etiqueta, etiqueta)
            if nombre in usados:                       # mismo texto y mismo factor
                nombre = "%s (%g pts)" % (nombre, c.puntos)
            raiz, n = nombre, 2
            while nombre in usados:                    # último recurso
                nombre = "%s (%d)" % (raiz, n)
                n += 1
            usados.add(nombre)
            nombres[c.clave] = nombre
    return nombres


NOMBRE = _nombres_legibles()                 # clave interna -> nombre legible
CLAVES = {v: k for k, v in NOMBRE.items()}   # nombre legible -> clave interna
CONCEPTO = {c.clave: c for f in tb.FACTORES for c in f.conceptos}


# Columnas del cuadro general, con la redacción exacta del concentrado de
# Recategorización 2025. Son siete, no ocho: "Otros estudios" no es columna
# propia, se suma dentro de a) Formación académica (el Reglamento dice que
# ese puntaje "se adiciona a Estudios en Educación").
CUADRO = [
    ("a) Formación académica", ("FORM_EDU", "FORM_OTROS")),
    # b1 (diplomados y especialidades, sin tope) + b2 (cursos como asistente,
    # tope de 10 puntos). El concentrado de 2025 los suma en una sola columna.
    ("b) Actualización profesional", ("ACTUALIZACION", "CURSOS_TOMADOS")),
    ("c) Antigüedad", ("ANTIGUEDAD",)),
    ("d) Experiencia Profesional y/o Académica", ("EXPERIENCIA",)),
    ("e) Productividad profesional", ("PRODUCTIVIDAD",)),
    ("f)Congresos, simposios, coloquios y foros", ("CONGRESOS",)),
    ("g) Diplomados diseñados", ("DIPLOMADOS",)),
]

# Si tabulador.py nombra un factor distinto, su puntaje se perdería en silencio
# y el concentrado saldría mal sin avisar. Mejor que truene aquí.
_MAPEADAS = {clave for _, claves in CUADRO for clave in claves}
_SUELTAS = [f.clave for f in tb.FACTORES if f.clave not in _MAPEADAS]


def cuadro_general(cantidades):
    """Puntos por apartado, agrupados como el concentrado oficial."""
    desg = tb.desglose(cantidades)
    fila = {col: round(sum(desg.get(k, 0) for k in claves), 2)
            for col, claves in CUADRO}
    fila["Total"] = desg["TOTAL"]
    return fila


def _opciones_legibles():
    """
    Lo que se ve en el selector: el nombre del concepto y lo que vale.

    Sin el nombre del apartado delante — el selector ya está dentro de su
    apartado y el prefijo sólo estorba. Sólo cuando dos conceptos quedarían
    con el mismo texto se les antepone el factor para poder distinguirlos.
    """
    bruto = {}
    for f in tb.FACTORES:
        for c in f.conceptos:
            texto = "%s — %g pts" % (c.etiqueta.strip(), c.puntos)
            if c.tope_cantidad:
                texto += " (máx. %g)" % c.tope_cantidad
            bruto.setdefault(texto, []).append((f, c))

    opciones = {}
    for texto, grupo in bruto.items():
        for f, c in grupo:
            opciones[c.clave] = texto if len(grupo) == 1 \
                else "%s · %s" % (f.etiqueta, texto)
    return opciones


OPCIONES = _opciones_legibles()
DE_OPCION = {v: k for k, v in OPCIONES.items()}
NADA = "— no precargar —"

# Para el Sheet: el nombre del concepto y el apartado al que pertenece.
ETIQUETA = {c.clave: c.etiqueta.strip() for f in tb.FACTORES for c in f.conceptos}
APARTADO_DE = {c.clave: f.etiqueta for f in tb.FACTORES for c in f.conceptos}

# Lo único que se arrastra de Recategorización 2025: el grado académico y los
# años de servicio. Todo lo demás vuelve a cero por la Base PRIMERA.5.
_FACTORES_A = CUADRO[0][1]            # los que componen a) Formación académica
CONCEPTOS_FORMACION = [c for f in tb.FACTORES if f.clave in _FACTORES_A
                       for c in f.conceptos]
_FACTOR_ANTIG = next((f for f in tb.FACTORES if f.clave == "ANTIGUEDAD"), None)
CONCEPTO_ANIO = (_FACTOR_ANTIG.conceptos[0]
                 if _FACTOR_ANTIG and _FACTOR_ANTIG.conceptos else None)
CLAVES_PRECARGA = {c.clave for c in CONCEPTOS_FORMACION} | (
    {CONCEPTO_ANIO.clave} if CONCEPTO_ANIO else set())


def cantidades_de(renglones):
    """Suma los renglones por concepto: 3 cursos de 40 h capturados por
    separado son lo mismo que un renglón con cantidad 3."""
    total = {}
    for r in renglones:
        if r.get("Sin valor"):          # documento recibido que no puntúa
            continue
        clave = DE_OPCION.get(str(r.get("Concepto", "")).strip())
        if not clave:
            continue
        try:
            cantidad = float(r.get("Cantidad") or 0)
        except (TypeError, ValueError):
            continue
        if cantidad > 0:
            total[clave] = total.get(clave, 0) + cantidad
    return total


def puntos_sueltos(cantidades):
    """Puntos de cada concepto antes de topes y de las reglas de no
    acumulación. Sirve para el desglose horizontal; el número oficial de cada
    factor lo sigue dando tabulador.desglose()."""
    sueltos = {}
    for clave, cantidad in cantidades.items():
        c = CONCEPTO.get(clave)
        if c is None:
            continue
        efectiva = min(cantidad, c.tope_cantidad) if c.tope_cantidad else cantidad
        sueltos[clave] = round(c.puntos * efectiva, 2)
    return sueltos


# Contrato con el sistema de asignación (repo SISTEMA-PARA-PROPUESTAS-DFC):
# cargar_prelacion() de esa app lee exactamente estas columnas, con estos
# encabezados. No cambiar el texto sin cambiarlo allá también.
COLS_ASIGNACION = ["rfc", "nombre", "puesto", "plaza", "puntaje",
                   "horas asesor", "horas otras áreas", "horas de base"]


def excel_para_asignacion(orden):
    """
    Arma el Excel de prelación en el esquema que consume el sistema de
    asignación. El ORDEN lo decide esa app, que ya ordena por puntaje;
    aquí sólo se entregan los datos.

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


def excel_de(hojas):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
    return buffer.getvalue()


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# =========================================================================
# Interfaz
# =========================================================================

st.set_page_config(page_title="Valoración y prelación · Incremento 2026",
                   page_icon="📋", layout="wide")

st.title("Valoración y prelación · Incremento de Horas 2026")
st.caption("Comisión Dictaminadora Interna · Dirección de Formación Continua y CEPD · "
           "Valoración del 22 al 28 de octubre de 2026")

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

if _SUELTAS:
    st.error("Estos factores de tabulador.py no están mapeados al concentrado "
             "oficial y su puntaje no aparecería en ninguna columna: "
             + ", ".join(_SUELTAS) + ". Hay que agregarlos a CUADRO antes de usar "
             "la prelación.")
    st.stop()

st.sidebar.metric("Expedientes registrados", len(df_part))
st.sidebar.metric("Expedientes valorados", df_punt["folio"].nunique() if not df_punt.empty else 0)

tab_val, tab_pre = st.tabs(["1 · Valoración", "2 · Prelación"])


# --- 1. VALORACIÓN --------------------------------------------------------
with tab_val:
    st.subheader("Valoración conforme al Tabulador de Factores y Puntajes")
    st.caption("Título Octavo, Art. 43 del Reglamento. Sólo se valoran actividades "
               "posteriores a la última promoción (Base PRIMERA.5).")

    # --- Precarga desde el archivo Precarga_Incremento_2026.xlsx -----------
    # Ese archivo sale de las hojas individuales de Recategorización 2025 y lo
    # revisa la Comisión antes de subirlo. Aquí sólo se escribe lo que trae.
    #
    # Se arrastran a) Formación académica y c) Antigüedad, y nada más: la Base
    # PRIMERA.5 excluye las constancias de actualización, experiencia,
    # producción académica y laboral ya presentadas en recategorización.
    #
    # Siempre abierto: es la única puerta de alta de participantes, porque la
    # pestaña de Recepción ya no existe. Plegado no se ve y no se encuentra.
    with st.container(border=True):
        st.markdown("#### Cargar el archivo de precarga")
        st.caption("Sube aquí **Precarga_Incremento_2026.xlsx** ya revisado. Da de "
                   "alta a los participantes que falten (empareja por R.F.C.) y "
                   "escribe los incisos a) Formación académica y c) Antigüedad. "
                   "Los apartados b) a g) no se tocan.")
        archivo = st.file_uploader("Archivo de precarga (.xlsx)", type="xlsx",
                                   key="archivo_precarga")

        if CONCEPTO_ANIO is None:
            st.error("tabulador.py no expone un concepto de antigüedad; no puedo "
                     "precargar el inciso c).")
        elif archivo is not None:
            try:
                hoja = pd.read_excel(archivo, sheet_name="Precarga").fillna("")
            except Exception as err:
                st.error("No pude leer la hoja «Precarga»: %s" % err)
                hoja = pd.DataFrame()

            faltan = [c for c in ("Nombre completo", "R.F.C.") if c not in hoja.columns]
            if faltan:
                st.error("Al archivo le faltan columnas: " + ", ".join(faltan))
            elif not hoja.empty:
                por_rfc = {normalizar(r): f for f, r
                           in zip(df_part["folio"], df_part["rfc"])}
                propuesta = []
                for _, r in hoja.iterrows():
                    rfc = str(r["R.F.C."]).strip().upper()
                    if not rfc:
                        continue
                    grado = emparejar_concepto(r.get("Estudios en educación"), NOMBRE)
                    otros = emparejar_concepto(r.get("Otros estudios"), NOMBRE)
                    anios = r.get("Años de servicio 2026")
                    propuesta.append({
                        "Cargar": True,
                        "Nombre completo": str(r["Nombre completo"]).strip(),
                        "R.F.C.": rfc,
                        "Horas solicitadas": r.get("Horas solicitadas", ""),
                        "Ya está": "Sí" if normalizar(rfc) in por_rfc else "",
                        "Estudios en educación": OPCIONES[grado] if grado else NADA,
                        "Otros estudios": OPCIONES[otros] if otros else NADA,
                        "Años de servicio": float(anios) if es_numero(anios) else 0.0,
                        "Del archivo": " / ".join(
                            str(x) for x in [r.get("Estudios en educación"),
                                             r.get("Otros estudios"), anios] if str(x).strip()),
                    })

                sin_leer = [p["Nombre completo"] for p in propuesta
                            if p["Estudios en educación"] == NADA
                            and p["Otros estudios"] == NADA]
                if sin_leer:
                    st.warning("%d participante(s) sin grado identificado: %s. Se dan "
                               "de alta igual y su formación se captura a mano."
                               % (len(sin_leer), ", ".join(sin_leer[:5])
                                  + (", …" if len(sin_leer) > 5 else "")))

                opciones_grado = [NADA] + [OPCIONES[c.clave] for c in CONCEPTOS_FORMACION]
                revisado = st.data_editor(
                    pd.DataFrame(propuesta), width="stretch", hide_index=True,
                    key="rev_precarga",
                    column_config={
                        "Cargar": st.column_config.CheckboxColumn("Cargar"),
                        "Ya está": st.column_config.TextColumn(
                            "Ya registrado", disabled=True, width="small"),
                        "Estudios en educación": st.column_config.SelectboxColumn(
                            "Estudios en educación", options=opciones_grado, width="large"),
                        "Otros estudios": st.column_config.SelectboxColumn(
                            "Otros estudios", options=opciones_grado, width="large"),
                        "Años de servicio": st.column_config.NumberColumn(
                            "Años de servicio", min_value=0.0, step=1.0),
                        "Del archivo": st.column_config.TextColumn(
                            "Leído del archivo", disabled=True, width="large"),
                    })

                if st.button("Cargar en el Sheet", type="primary", key="bprecarga"):
                    if not capturo.strip():
                        st.error("Anota en la barra lateral quién está capturando.")
                    else:
                        # Todo se arma en memoria y se escribe de un solo viaje.
                        # Una llamada por participante son más de cien peticiones
                        # seguidas y Google corta a las 60 escrituras por minuto.
                        altas_participantes = []
                        renglones_puntajes = []
                        folios_tocados = []
                        nuevos_folios = []

                        for r in revisado.to_dict("records"):
                            if not r["Cargar"]:
                                continue
                            rfc = str(r["R.F.C."]).strip().upper()
                            folio = por_rfc.get(normalizar(rfc))
                            if folio is None:
                                # Alta mínima. El número de expediente queda en
                                # blanco a propósito: lo pone la Comisión al
                                # revisar, en el orden que ella decida.
                                folio = siguiente_folio(df_part, nuevos_folios)
                                nuevos_folios.append(folio)
                                altas_participantes.append({
                                    "folio": folio, "fecha_recepcion": ahora(),
                                    "nombre": str(r["Nombre completo"]).strip().upper(),
                                    "rfc": rfc, "estatus": ESTATUS[0],
                                    "horas_solicitadas": r.get("Horas solicitadas", ""),
                                    "no_expediente": "",
                                    "capturo": capturo.strip(), "actualizado": ahora()})

                            folios_tocados.append(folio)
                            nuevos = []
                            for columna in ("Estudios en educación", "Otros estudios"):
                                clave = DE_OPCION.get(str(r[columna]).strip())
                                if clave:
                                    nuevos.append({
                                        "folio": folio, "concepto": clave, "cantidad": 1,
                                        "detalle": ""})
                            anios = float(r["Años de servicio"] or 0)
                            if anios > 0:
                                nuevos.append({
                                    "folio": folio, "concepto": CONCEPTO_ANIO.clave,
                                    "cantidad": anios,
                                    "detalle": ""})
                            for n in nuevos:
                                n.update({"capturo": capturo.strip(),
                                          "actualizado": ahora(), "sin_valor": "",
                                          "no_documento": "", "notas": ""})
                            renglones_puntajes.extend(nuevos)

                        try:
                            if altas_participantes:
                                agregar_varios("PARTICIPANTES", altas_participantes)
                            if folios_tocados:
                                # Conserva lo capturado de los demás apartados y
                                # reemplaza sólo los renglones de a) y c).
                                reemplazar_varios_folios(
                                    "PUNTAJES", folios_tocados, renglones_puntajes,
                                    conceptos=CLAVES_PRECARGA)
                        except Exception as err:
                            st.error("No se pudo escribir en el Sheet: %s" % err)
                        else:
                            st.success(
                                "%d participante(s) procesados, %d dados de alta, "
                                "%d renglón(es) de puntaje escritos. El número de "
                                "expediente queda en blanco: se asigna al revisar."
                                % (len(folios_tocados), len(altas_participantes),
                                   len(renglones_puntajes)))
                            st.rerun()


    if df_part.empty:
        st.info("Todavía no hay participantes registrados. Sube el archivo de "
                "precarga aquí arriba para darlos de alta; de ahí en adelante "
                "aparecen en esta lista.")
        st.stop()

    # Se elige por nombre, en orden alfabético: todavía puede no haber número
    # de expediente, porque lo asigna la Comisión conforme revisa.
    orden_alfabetico = df_part.sort_values("nombre")["folio"].tolist()

    def _rotulo_participante(f):
        fila = df_part.loc[df_part.folio == f].iloc[0]
        exp = str(fila.get("no_expediente", "")).strip()
        return ("%s · exp. %s" % (fila["nombre"], exp)) if exp else fila["nombre"]

    folio_v = st.selectbox("Participante a valorar", orden_alfabetico,
                           format_func=_rotulo_participante, key="folio_val")

    # Número de expediente: lo escribe quien revisa, en el orden que decida.
    fila_v = df_part.loc[df_part.folio == folio_v].iloc[0].to_dict()
    c_exp, c_bot = st.columns([1, 3])
    exp_nuevo = c_exp.text_input("No. de expediente",
                                 value=str(fila_v.get("no_expediente", "")),
                                 key="exp_%s" % folio_v,
                                 help="Lo asigna la Comisión al revisar. "
                                      "Puede quedar en blanco mientras tanto.")
    if exp_nuevo.strip() != str(fila_v.get("no_expediente", "")).strip():
        if c_bot.button("Guardar el número de expediente", key="bexp_%s" % folio_v):
            fila_v["no_expediente"] = exp_nuevo.strip()
            fila_v["actualizado"] = ahora()
            actualizar_participante(folio_v, fila_v)
            st.success("Expediente %s asignado." % exp_nuevo.strip())
            st.rerun()

    st.info("Un renglón por documento. En **Detalle** escribe el nombre de la "
            "constancia — «Diplomado en Gestión Escolar», «Curso de Evaluación "
            "Formativa» — y en **Cantidad** cuántos documentos iguales amparas "
            "con ese mismo renglón.")

    # Lo ya capturado para este folio, repartido entre los factores.
    previos = df_punt[df_punt.folio == folio_v]
    renglones_previos = {f.clave: [] for f in tb.FACTORES}
    for _, r in previos.iterrows():
        clave = str(r["concepto"]).strip()
        if clave not in CONCEPTO:
            continue
        factor = next(f for f in tb.FACTORES if any(c.clave == clave for c in f.conceptos))
        renglones_previos[factor.clave].append({
            "No.": str(r.get("no_documento") or ""),
            "Concepto": OPCIONES[clave],
            "Detalle": str(r.get("detalle") or ""),
            "Cantidad": float(r["cantidad"] or 0),
            "Sin valor": str(r.get("sin_valor") or "").strip().lower() in ("sí", "si", "true", "1"),
            "Notas y observaciones": str(r.get("notas") or ""),
        })

    capturado = {}
    for factor in tb.FACTORES:
        opciones = [OPCIONES[c.clave] for c in factor.conceptos]
        base = pd.DataFrame(renglones_previos[factor.clave] or [],
                            columns=["No.", "Concepto", "Detalle", "Cantidad",
                                     "Sin valor", "Notas y observaciones"])
        parciales = cantidades_de(base.to_dict("records"))
        subtotal = tb.puntos_factor(factor, parciales)

        with st.expander("%s — %s pts" % (factor.etiqueta, subtotal),
                         expanded=bool(renglones_previos[factor.clave])):
            if factor.nota:
                st.caption(factor.nota)
            editado = st.data_editor(
                base, num_rows="dynamic", width="stretch", hide_index=True,
                key=f"doc_{folio_v}_{factor.clave}",
                column_config={
                    "No.": st.column_config.TextColumn(
                        "No.", width="small",
                        help="Número del documento en la relación que entregó "
                             "el participante. Es el que se cita en las notas."),
                    "Concepto": st.column_config.SelectboxColumn(
                        "Concepto del tabulador", options=opciones,
                        width="large", required=True),
                    "Detalle": st.column_config.TextColumn(
                        "Detalle (nombre del documento)", width="medium"),
                    "Cantidad": st.column_config.NumberColumn(
                        "Cantidad", min_value=0.0, step=1.0, default=1.0, width="small"),
                    "Sin valor": st.column_config.CheckboxColumn(
                        "Sin valor", default=False,
                        help="Documento recibido que no puntúa. Se conserva en "
                             "el expediente con cero puntos."),
                    "Notas y observaciones": st.column_config.TextColumn(
                        "Notas y observaciones", width="medium"),
                })
            capturado[factor.clave] = editado.to_dict("records")

    # Cantidades de todos los factores juntos: es lo que entiende el tabulador.
    cantidades = {}
    for renglones in capturado.values():
        for clave, cantidad in cantidades_de(renglones).items():
            cantidades[clave] = cantidades.get(clave, 0) + cantidad

    desg = tb.desglose(cantidades)
    st.markdown("#### Puntaje")
    cols = st.columns(len(tb.FACTORES) + 1)
    for i, factor in enumerate(tb.FACTORES):
        cols[i].metric(factor.etiqueta, desg[factor.clave])
    cols[-1].metric("TOTAL DE PUNTOS", desg["TOTAL"])

    # Avisos sobre los dos puntos que el Reglamento no resuelve solo.
    # El Reglamento no dice si los grados de "Otros estudios" se acumulan entre
    # sí, pero la Comisión ya lo resolvió en Recategorización 2025: los
    # expedientes 34 y 58 traen Licenciatura 15 más Maestría 25 con subtotal 40.
    # Por eso F_OTROS va con excluyente=False. El aviso queda sólo como registro.
    capturados_otros = [c.clave for c in tb.F_OTROS.conceptos if cantidades.get(c.clave, 0) > 0]
    if len(capturados_otros) > 1:
        st.info("Se capturó más de un grado en 'Otros estudios' y se están sumando, "
                "como en Recategorización 2025 (expedientes 34 y 58).")
    # El tope de 10 puntos es del bloque b2 (cursos como asistente), no de b1.
    tope_b2 = tb.F_CURSOS_TOMADOS.tope_puntos
    if tope_b2 and tb.puntos_factor(tb.F_CURSOS_TOMADOS, cantidades) >= tope_b2:
        st.info("El bloque b2 llegó a su tope de %g puntos; los cursos que "
                "sigas capturando ya no suman." % tope_b2)

    if st.button("Guardar valoración", type="primary", key=f"bval_{folio_v}"):
        if not capturo.strip():
            st.error("Anota en la barra lateral quién está capturando.")
        else:
            registros = []
            # Números ya escritos a mano: no se reasignan.
            ocupados = set()
            for renglones in capturado.values():
                for r in renglones:
                    n = str(r.get("No.") or "").strip()
                    if n.isdigit():
                        ocupados.add(int(n))
            siguiente = 1

            for renglones in capturado.values():
                for r in renglones:
                    clave = DE_OPCION.get(str(r.get("Concepto", "")).strip())
                    try:
                        cantidad = float(r.get("Cantidad") or 0)
                    except (TypeError, ValueError):
                        cantidad = 0
                    sin_valor = bool(r.get("Sin valor"))
                    # Un documento sin valor se guarda aunque no sume puntos:
                    # desaparecerlo deja el expediente incompleto.
                    if clave and (cantidad > 0 or sin_valor):
                        numero = str(r.get("No.") or "").strip()
                        if not numero.isdigit():
                            while siguiente in ocupados:
                                siguiente += 1
                            numero = str(siguiente)
                            ocupados.add(siguiente)
                        registros.append({
                            "folio": folio_v, "concepto": clave, "cantidad": cantidad,
                            "detalle": str(r.get("Detalle") or "").strip(),
                            "no_documento": numero,
                            "notas": str(r.get("Notas y observaciones") or "").strip(),
                            "sin_valor": "Sí" if sin_valor else "",
                            "capturo": capturo.strip(), "actualizado": ahora()})
            reemplazar_por_folio("PUNTAJES", folio_v, registros)
            st.success(f"Valoración guardada: {len(registros)} documento(s), "
                       f"{desg['TOTAL']} puntos.")
            st.rerun()


# --- 2. PRELACIÓN ---------------------------------------------------------
with tab_pre:
    st.subheader("Lista de prelación")
    st.caption("Orden por puntaje total; desempate por formación académica y, en segundo "
               "término, por antigüedad en el nivel superior de la SEJ (Art. 24 inciso i).")

    # Sólo entran a prelación los expedientes que no quedaron improcedentes.
    elegibles = df_part[df_part["estatus"] != "IMPROCEDENTE"]

    if elegibles.empty:
        st.info("No hay expedientes elegibles todavía.")
    else:
        participantes = []
        for _, r in elegibles.iterrows():
            pts = df_punt[df_punt.folio == r["folio"]]
            cantidades = {}
            for _, p in pts.iterrows():
                if str(p.get("sin_valor", "")).strip():     # no puntúa
                    continue
                clave = str(p["concepto"]).strip()
                try:
                    cantidad = float(p["cantidad"] or 0)
                except (TypeError, ValueError):
                    continue
                if clave and cantidad > 0:
                    cantidades[clave] = cantidades.get(clave, 0) + cantidad
            convs = df_conv[df_conv.folio == r["folio"]]["no_convocatoria"].astype(str).tolist()
            participantes.append({
                "folio": r["folio"], "nombre": r["nombre"], "rfc": r["rfc"],
                "no_expediente": str(r.get("no_expediente", "")).strip(),
                "cct": r["cct"], "categoria_actual": r["categoria_actual"],
                "nivel_actual": r["nivel_actual"], "horas_actuales": r["horas_actuales"],
                "horas_solicitadas": r["horas_solicitadas"],
                "convocatorias": " | ".join(convs),
                "valorado": "Sí" if cantidades else "No",
                "cantidades": cantidades,
            })

        orden = tb.ordenar_prelacion(participantes)

        # --- tabla horizontal: una fila por expediente, una columna por factor ---
        # Los encabezados son los del artículo 43, escritos completos. Quien
        # descargue el archivo debe poder leer de corrido cuántos puntos sacó
        # cada quien en antigüedad sin preguntarle a nadie qué significa una
        # abreviatura.
        filas = []
        for e in orden:
            fila = {
                "Lugar": e["lugar"],
                "Nombre": e["nombre"],
                "No. Exp.": e["no_expediente"],
                "Folio": e["folio"], "R.F.C.": e["rfc"],
                "Categoría": e["categoria_actual"], "Nivel": e["nivel_actual"],
                "Horas actuales": e["horas_actuales"],
                "Horas solicitadas": e["horas_solicitadas"],
            }
            fila.update(cuadro_general(e["cantidades"]))
            fila["Convocatorias en que participa"] = e["convocatorias"]
            fila["Expediente valorado"] = e["valorado"]
            fila["Empate sin resolver"] = "Sí" if e["empate_sin_resolver"] else ""
            filas.append(fila)
        tabla = pd.DataFrame(filas)

        sin_valorar = tabla[tabla["Expediente valorado"] == "No"]
        if not sin_valorar.empty:
            st.warning(f"{len(sin_valorar)} expediente(s) sin valorar aparecen con 0 puntos.")
        if (tabla["Empate sin resolver"] == "Sí").any():
            st.warning("Hay empates que los tres criterios no resuelven. "
                       "La Comisión debe resolverlos y asentarlo en el acta.")

        st.dataframe(tabla, width="stretch", hide_index=True)
        st.caption("Los siete apartados vienen con la misma redacción y el mismo orden "
                   "del concentrado de Recategorización 2025. «Otros estudios» no es "
                   "columna aparte: se suma dentro de a) Formación académica.")

        # --- desglose horizontal documento por documento ---
        # Una columna por concepto que alguien haya acreditado; se omiten los
        # conceptos que nadie capturó para que la hoja siga siendo legible.
        con_captura = [c.clave for f in tb.FACTORES for c in f.conceptos
                       if any(e["cantidades"].get(c.clave, 0) > 0 for e in orden)]
        detalle = []
        for e in orden:
            sueltos = puntos_sueltos(e["cantidades"])
            desg = tb.desglose(e["cantidades"])
            fila = {"Lugar": e["lugar"], "Nombre": e["nombre"], "Folio": e["folio"],
                    "R.F.C.": e["rfc"]}
            fila.update({NOMBRE[k]: sueltos.get(k, "") for k in con_captura})
            fila.update(cuadro_general(e["cantidades"]))
            detalle.append(fila)
        tabla_detalle = pd.DataFrame(detalle)

        st.markdown("#### Desglose documento por documento")
        st.dataframe(tabla_detalle, width="stretch", hide_index=True)
        st.caption("Cada columna es un concepto del tabulador, con su nombre completo. "
                   "Trae los puntos en bruto, antes de los topes y de la regla de no "
                   "acumulación; las columnas de apartado y el Total son las que valen.")

        st.divider()

        # El concentrado horizontal, en el Sheet y no sólo en pantalla.
        c_pub, c_msg = st.columns([1, 2])
        if c_pub.button("Publicar el concentrado en el Sheet", type="primary",
                        key="bpublicar"):
            try:
                n = publicar_concentrado(tabla)
            except Exception as err:
                st.error("No se pudo escribir la pestaña RESULTADOS: %s" % err)
            else:
                st.success("Pestaña RESULTADOS actualizada con %d participante(s)." % n)
        c_msg.caption("Deja la pestaña **RESULTADOS** del Sheet igual a esta tabla: "
                      "una fila por participante, una columna por apartado y el "
                      "Total al final. Se reescribe completa cada vez.")

        st.divider()
        sello = datetime.now(TZ).strftime("%Y%m%d_%H%M")

        st.download_button(
            "Excel · prelación con desglose por factor y por concepto",
            excel_de({"prelación": tabla, "por concepto": tabla_detalle}),
            file_name=f"prelacion_incremento_2026_{sello}.xlsx",
            mime=MIME_XLSX, type="primary")

        st.download_button(
            "Descargar prelación (CSV)",
            tabla.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"prelacion_incremento_2026_{sello}.csv",
            mime="text/csv")

        st.download_button(
            "Excel para el sistema de asignación",
            excel_para_asignacion(orden),
            file_name=f"prelacion_para_asignacion_{sello}.xlsx",
            mime=MIME_XLSX)
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
            "expedientes_valorados": int((tabla["Expediente valorado"] == "Sí").sum()),
            "documentos_capturados": int(len(df_punt)),
        }
        st.download_button(
            "Descargar concentrado del proceso (JSON)",
            json.dumps(concentrado, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name="concentrado_incremento_2026.json", mime="application/json")

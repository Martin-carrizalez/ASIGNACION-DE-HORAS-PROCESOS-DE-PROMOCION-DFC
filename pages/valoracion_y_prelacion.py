"""
Valoración y prelación del Concurso Cerrado de Incremento de Horas 2026
Comisión Dictaminadora Interna — Dirección de Formación Continua / CEPD

Dos etapas, las que siguen vigentes del calendario de la Base SÉPTIMA:
  1. VALORACIÓN  (22 al 28 de octubre) — captura del tabulador por expediente,
                 documento por documento.
  2. PRELACIÓN   — orden final y exportación que alimenta el sistema de asignación.

Los participantes se dan de alta a mano desde la barra lateral, con nombre y
R.F.C.; con eso se arma la lista. No se carga nada de procesos anteriores.

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
from datetime import date, datetime

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

# Las seis categorías del nivel superior. Se guardan partidas en categoría y
# nivel porque así las lee el sistema de asignación, que arma el puesto como
# "HORAS DE TITULAR A".
# De menor a mayor: el orden decide cuál es "la mayor que ostenta" cuando
# alguien tiene horas en varias.
CATEGORIAS = ["ASOCIADO A", "ASOCIADO B", "ASOCIADO C",
              "TITULAR A", "TITULAR B", "TITULAR C"]


def columna_horas(categoria):
    """Nombre interno de la columna de horas de una categoría."""
    return "horas_" + categoria.lower().replace(" ", "_")


def reparto_de(fila):
    """«Asociado A 10 · Asociado B 5» a partir de las seis columnas de horas."""
    partes = []
    for categoria in CATEGORIAS:
        valor = str(fila.get(columna_horas(categoria), "")).strip()
        try:
            horas = float(valor)
        except (TypeError, ValueError):
            continue
        if horas > 0:
            partes.append("%s %g" % (categoria.title(), horas))
    return " · ".join(partes)

# Art. 14, Primera, inciso b del Adendum y Base PRIMERA.10: hasta quince horas.
TOPE_SOLICITUD = 15
# Tope de horas por persona sumando todas sus plazas.
TOPE_TOTAL = 48

# Fecha a la que se cuentan los años de titulación. Sirve para acreditar el
# perfil del Adendum, Título Tercero, artículo 13, al momento de asignar la
# plaza. Es independiente del puntaje: un título que no suma puntos puede, por
# su antigüedad, cubrir el perfil de la categoría.
CORTE_PERFIL = date(2026, 10, 7)

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
        ("horas_actuales", "Horas actuales (total)"),
        ("fecha_ultima_promocion", "Fecha de última promoción"),
        ("horas_solicitadas", "Horas solicitadas"),
        ("estatus", "Estatus de la solicitud"),
        ("observaciones", "Observaciones"),
        ("fecha_recepcion", "Fecha de recepción"),
        ("capturo", "Capturó"),
        ("actualizado", "Última actualización"),
    ] + [(columna_horas(c), "Horas de %s" % c.title()) for c in CATEGORIAS],
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
        ("rfc", "R.F.C."),
        ("nombre", "Nombre del participante"),
        ("apartado", "Apartado"),
        ("concepto_texto", "Concepto del tabulador"),
        ("concepto", "Clave del concepto"),
        ("cantidad", "Cantidad"),
        ("detalle", "Detalle del documento"),
        ("no_documento", "No. de documento"),
        ("notas", "Notas y observaciones"),
        ("sin_valor", "Documento sin valor"),
        ("fecha_titulo", "Fecha de expedición del título"),
        ("capturo", "Capturó"),
        ("actualizado", "Última actualización"),
    ],
}

def internos(nombre):
    return [a for a, _ in HOJAS[nombre]]


def visibles(nombre):
    return [b for _, b in HOJAS[nombre]]


def columna_interna(nombre, encabezado):
    """Del texto de un encabezado al nombre interno que le corresponde."""
    # Se acepta el nombre visible de hoy y el nombre interno, que es como se
    # guardaban las pestañas de la versión anterior.
    for interno, visible in HOJAS[nombre]:
        if encabezado == visible or encabezado == interno:
            return interno
    return None


# =========================================================================
# Capa de datos
# =========================================================================

@st.cache_resource(show_spinner=False)
def abrir_libro():
    """Abre el Sheet y crea las pestañas que falten. El acomodo de columnas
    lo hace _hoja(), que corre en cada lectura y en cada escritura."""
    alcances = ["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"]
    cred = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=alcances)
    libro = gspread.authorize(cred).open_by_key(
        st.secrets["registro_incremento"]["sheet_id"])
    existentes = {h.title for h in libro.worksheets()}
    for nombre in HOJAS:
        if nombre not in existentes:
            cabecera = visibles(nombre)
            libro.add_worksheet(title=nombre, rows=500,
                                cols=len(cabecera) + 5).append_row(cabecera)
    return libro


def _escribir_encima(hoja, filas, renglones_antes, ancho):
    """
    Deja la pestaña con exactamente estos renglones, sin vaciarla primero.

    Vaciar y volver a escribir deja una ventana en la que la pestaña está en
    blanco: si la conexión se cae justo ahí, se pierde lo capturado. Aquí se
    escribe encima y después se recortan los renglones que sobraron, así que
    en ningún momento hay menos datos de los que debe haber.

    Cada renglón se rellena hasta el ancho del encabezado para que no queden
    celdas viejas asomando a la derecha.
    """
    parejas = [list(f) + [""] * (ancho - len(f)) for f in filas]
    hoja.update(parejas, "A1", value_input_option="USER_ENTERED")
    sobrantes = renglones_antes - len(parejas)
    if sobrantes > 0:
        hoja.delete_rows(len(parejas) + 1, renglones_antes)


def _parece_encabezado(nombre, fila):
    """¿El renglón 1 es un encabezado y no un registro?"""
    llenas = [c.strip() for c in fila if c.strip()]
    if not llenas:
        return False
    return all(columna_interna(nombre, c) for c in llenas)


def _hoja(nombre):
    """
    Devuelve la pestaña y sus valores con las columnas en el orden actual.

    Si la pestaña viene de una versión anterior —le falta una columna, o están
    en otro orden— cada renglón se reconstruye por NOMBRE de columna y se
    vuelve a escribir. Nunca por posición: sobrescribir el encabezado sin
    recorrer los datos recorre el significado de todas las columnas.
    """
    hoja = abrir_libro().worksheet(nombre)
    valores = hoja.get_all_values()
    cabecera = visibles(nombre)
    claves = internos(nombre)

    if valores and valores[0] == cabecera:
        return hoja, valores

    if not valores:
        hoja.update([cabecera], "A1", value_input_option="USER_ENTERED")
        return hoja, [cabecera]

    if _parece_encabezado(nombre, valores[0]):
        # Encabezado reconocible: se reacomoda columna por columna.
        viejas = [columna_interna(nombre, c.strip()) for c in valores[0]]
        filas = []
        for fila in valores[1:]:
            if not any(c.strip() for c in fila):
                continue
            por_clave = {k: (fila[j] if j < len(fila) else "")
                         for j, k in enumerate(viejas) if k}
            filas.append([por_clave.get(k, "") for k in claves])
    else:
        # Sin encabezado reconocible: los datos ya vienen en el orden actual.
        filas = [list(f) for f in valores if any(c.strip() for c in f)]

    if hoja.col_count < len(cabecera):
        hoja.add_cols(len(cabecera) - hoja.col_count)
    nuevas = [cabecera] + filas
    _escribir_encima(hoja, nuevas, len(valores), len(cabecera))
    return hoja, nuevas


@st.cache_data(show_spinner=False)
def leer(nombre):
    """
    Lee una pestaña completa con los nombres internos como columnas.

    Va en caché porque Streamlit reejecuta el script entero en cada clic y sin
    esto cada casilla del editor costaba tres lecturas al Sheet. Toda escritura
    llama a olvidar_lo_leido() para que lo siguiente que se lea sea lo nuevo.
    """
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


def olvidar_lo_leido():
    """Tira la caché de lectura. Se llama después de cada escritura."""
    leer.clear()


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

    nuevas = [cabecera] + conservados + [_renglon(nombre, r) for r in registros]
    _escribir_encima(hoja, nuevas, len(valores), len(cabecera))


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


def completar_identidad(df_participantes):
    """
    Deja el R.F.C. y el Nombre de PUNTAJES iguales a los de PARTICIPANTES.

    PARTICIPANTES manda: es donde se corrigen los nombres. PUNTAJES guarda una
    copia sólo para que al abrir el Sheet se vea de quién es cada renglón, así
    que esa copia se refresca —no sólo se rellena cuando está vacía—. Si no,
    un nombre corregido dejaba con el viejo a los renglones capturados a mano.

    Se escriben únicamente esas dos columnas; los puntajes no se tocan.
    """
    hoja, valores = _hoja("PUNTAJES")
    claves = internos("PUNTAJES")
    i_folio, i_rfc, i_nom = (claves.index("folio"), claves.index("rfc"),
                             claves.index("nombre"))
    datos = {f: (r, n) for f, r, n in zip(df_participantes["folio"],
                                          df_participantes["rfc"],
                                          df_participantes["nombre"])}
    columnas, corregidos = [], 0
    for fila in valores[1:]:
        folio = fila[i_folio] if len(fila) > i_folio else ""
        rfc = fila[i_rfc] if len(fila) > i_rfc else ""
        nombre = fila[i_nom] if len(fila) > i_nom else ""
        if folio in datos and (rfc, nombre) != datos[folio]:
            rfc, nombre = datos[folio]
            corregidos += 1
        columnas.append([rfc, nombre])

    if corregidos:
        rango = "%s2:%s%d" % (chr(ord("A") + i_rfc), chr(ord("A") + i_nom),
                              len(columnas) + 1)
        hoja.update(columnas, rango, value_input_option="USER_ENTERED")
        olvidar_lo_leido()
    return corregidos


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
    antes = len(hoja.get_all_values())
    _escribir_encima(hoja, [columnas] + tabla.astype(str).values.tolist(),
                     antes, len(columnas))
    return len(tabla)


def avisar(mensaje):
    """
    Deja el aviso para DESPUÉS de recargar la pantalla.

    st.success() seguido de st.rerun() no se alcanza a dibujar: el rerun
    descarta la pantalla a medio pintar. Por eso el mensaje se guarda y se
    muestra al principio de la pasada siguiente.
    """
    st.session_state["aviso"] = mensaje


def ahora():
    return datetime.now(TZ).strftime("%d/%m/%Y %H:%M")


def normalizar(texto):
    """Minúsculas, sin acentos y sin espacios de más, para poder comparar."""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


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

# Para el Sheet: el nombre del concepto y el apartado al que pertenece.
ETIQUETA = {c.clave: c.etiqueta.strip() for f in tb.FACTORES for c in f.conceptos}
APARTADO_DE = {c.clave: f.etiqueta for f in tb.FACTORES for c in f.conceptos}


def a_fecha(valor):
    """Lee una fecha venga como venga: del editor, del Sheet o vacía."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor).strip()
    if not texto or texto.lower() in ("nat", "nan", "none"):
        return None
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto[:10], formato).date()
        except ValueError:
            continue
    return None


def anios_cumplidos(fecha, corte=CORTE_PERFIL):
    """
    Años enteros entre la expedición del título y la fecha de corte.

    Cuenta años cumplidos, no fracciones: quien se tituló el 8 de octubre de
    2017 tiene nueve al 7 de octubre de 2027, no diez.
    """
    f = a_fecha(fecha)
    if f is None or f > corte:
        return None
    anios = corte.year - f.year
    if (corte.month, corte.day) < (f.month, f.day):
        anios -= 1
    return anios


# A qué grado corresponde cada concepto de formación académica. El artículo 13
# del Adendum pide años por GRADO —doctor, maestro, licenciatura— y no
# distingue si el grado es en educación o no: dice "el título en alguna carrera
# profesional a nivel de licenciatura", "el grado de maestro", "el grado de
# doctor". Por eso las dos áreas caen en el mismo renglón.
GRADO_DE = {
    "LIC_EDU": "licenciatura", "LIC_OTR": "licenciatura",
    "MAE_EDU": "maestría", "MAE_OTR": "maestría",
    "DOC_EDU": "doctorado", "DOC_OTR": "doctorado",
}
GRADOS = ("licenciatura", "maestría", "doctorado")

# Factores cuyos conceptos tienen grado: son los únicos donde se pide la fecha
# del título. Sale del propio GRADO_DE, no de una lista aparte que se
# desincronice.
GRADOS_DE_FACTOR = {f.clave for f in tb.FACTORES
                    for c in f.conceptos if c.clave in GRADO_DE}


def titulos_de(renglones):
    """
    Años de titulación por grado, para el perfil del Adendum, art. 13.

    Se cuentan aunque el renglón esté marcado sin valor: el perfil mira la
    antigüedad del título, no su puntaje. De cada grado se toma el título más
    antiguo, que es el que más ayuda a cubrir el perfil.

    -> ({grado: años}, detalle legible)
    """
    por_grado = {}
    detalle = []
    for r in renglones:
        clave = str(r.get("concepto", "")).strip()
        grado = GRADO_DE.get(clave)
        if grado is None:
            continue
        n = anios_cumplidos(a_fecha(r.get("fecha_titulo")))
        if n is None:
            continue
        if n > por_grado.get(grado, -1):
            por_grado[grado] = n
        detalle.append((ETIQUETA.get(clave, grado), n))
    detalle.sort(key=lambda x: -x[1])
    return por_grado, " · ".join("%s %d año(s)" % (e, n) for e, n in detalle)


def cantidad_de(valor, por_defecto=1.0):
    """
    Cuántos documentos ampara el renglón.

    Un renglón recién agregado llega con la celda vacía. Si ya se eligió el
    concepto, lo que quiso decir quien captura es «uno»; descartarlo en
    silencio hacía que la valoración se guardara sin ese documento.
    """
    try:
        n = float(valor)
    except (TypeError, ValueError):
        return por_defecto
    if n != n:                 # NaN: la celda quedó vacía
        return por_defecto
    return n


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
        cantidad = cantidad_de(r.get("Cantidad"))
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
# La columna de años de titulación va AL FINAL para no mover las ocho que esa
# app ya leía. "plaza" se queda vacía a propósito: esta etapa no maneja claves
# presupuestales.
COLS_ASIGNACION = ["rfc", "nombre", "puesto", "plaza", "puntaje",
                   "horas asesor", "horas otras áreas", "horas de base",
                   "años licenciatura", "años maestría", "años doctorado"]


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
        categoria = str(e["categoria_actual"]).strip().upper()
        nivel = str(e["nivel_actual"]).strip().upper()
        filas.append({
            "rfc": e["rfc"],
            "nombre": e["nombre"],
            # Sin categoría no se arma el puesto: vale más dejarlo vacío que
            # entregar "HORAS DE" pelón y que la otra app lo dé por bueno.
            "puesto": ("HORAS DE %s %s" % (categoria, nivel)).strip()
                      if categoria else "",
            "plaza": "",
            "puntaje": e["total"] if e["valorado"] == "Sí" else None,
            "horas asesor": None,
            "horas otras áreas": None,
            "horas de base": None,
            # Para acreditar el perfil del Adendum, Título Tercero, art. 13,
            # al momento de asignar la plaza. Vacío = no acreditó ese grado.
            **{"años %s" % g: e.get("titulos_por_grado", {}).get(g)
               for g in GRADOS},
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

# Aviso de la operación anterior, que se guardó antes de recargar la pantalla.
if st.session_state.get("aviso"):
    st.success(st.session_state.pop("aviso"))

# Quién captura: queda asentado en cada renglón para efectos de auditoría,
# porque el oficio CDESE/017/2026 pide un concentrado del proceso.
capturo = st.sidebar.text_input("Captura (nombre de quien registra)", key="capturo")
if st.sidebar.button("Recargar datos del Sheet"):
    olvidar_lo_leido()
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
# Valorado = tiene al menos un documento que puntúa. Los que sólo traen
# documentos sin valor están capturados, pero no valorados.
_con_puntos = df_punt[df_punt["sin_valor"].astype(str).str.strip() == ""] \
    if not df_punt.empty else df_punt
st.sidebar.metric("Expedientes valorados",
                  _con_puntos["folio"].nunique() if not _con_puntos.empty else 0)

# El alta vive en la barra lateral para no estorbar en la captura: se usa
# conforme van llegando los expedientes, no de una sola vez.
# PUNTAJES copia el R.F.C. y el nombre de cada participante para que el Sheet
# se lea solo. Si en PARTICIPANTES se corrigió un nombre, aquí se refleja.
if not df_punt.empty and not df_part.empty:
    _esperado = dict(zip(df_part["folio"], zip(df_part["rfc"], df_part["nombre"])))
    _desfasados = any(
        f in _esperado and (r, n) != _esperado[f]
        for f, r, n in zip(df_punt["folio"], df_punt["rfc"], df_punt["nombre"]))
    if _desfasados:
        try:
            _corregidos = completar_identidad(df_part)
        except Exception as err:
            st.warning("No pude refrescar el R.F.C. y el nombre en PUNTAJES: %s" % err)
        else:
            if _corregidos:
                df_punt = leer("PUNTAJES")
                st.toast("Se actualizó el nombre o el R.F.C. en %d renglón(es) "
                         "de PUNTAJES." % _corregidos)

# Quién ya tiene puntos: evita volver a valorar a alguien por descuido.
puntaje_de = {}
if not df_punt.empty:
    for folio_p, grupo in df_punt.groupby("folio"):
        cant = {}
        for _, p in grupo.iterrows():
            if str(p.get("sin_valor", "")).strip():
                continue
            clave = str(p["concepto"]).strip()
            try:
                n = float(p["cantidad"] or 0)
            except (TypeError, ValueError):
                continue
            if clave and n > 0:
                cant[clave] = cant.get(clave, 0) + n
        if cant:
            puntaje_de[folio_p] = tb.desglose(cant)["TOTAL"]

# Secciones. No se usa st.tabs: pierde la pestaña activa en cada st.rerun() y
# devuelve al operador a la primera a media captura. Un radio con key conserva
# dónde estaba.
REGISTRO, VALORACION, PRELACION = ("1 · Registro", "2 · Valoración",
                                   "3 · Prelación")
seccion = st.radio("Sección", [REGISTRO, VALORACION, PRELACION],
                   horizontal=True, key="seccion", label_visibility="collapsed")
st.divider()

# Se calcula aquí, fuera de las secciones: lo consultan dos bloques distintos.
hay_gente = not df_part.empty


# --- 1. REGISTRO ----------------------------------------------------------
if seccion == REGISTRO:
    st.subheader("Registro de participantes")
    st.caption("La lista se arma a mano, conforme se reciben los expedientes. "
               "No se carga nada de procesos anteriores.")

    with st.form("alta_participante", clear_on_submit=True):
        c1, c2 = st.columns([3, 2])
        nuevo_nombre = c1.text_input("Nombre completo *")
        nuevo_rfc = c2.text_input("R.F.C. con homoclave *")

        # Una persona puede tener horas en más de una categoría: la unidad
        # real es la plaza, no la persona. Se capturan las seis y se deja en
        # cero la que no aplique.
        st.markdown("**Horas que tiene hoy en la DFC, por categoría**")
        horas_por_categoria = {}
        for fila in (CATEGORIAS[:3], CATEGORIAS[3:]):
            cols = st.columns(3)
            for col, categoria in zip(cols, fila):
                horas_por_categoria[categoria] = col.number_input(
                    categoria.title(), min_value=0, max_value=TOPE_TOTAL,
                    step=1, value=0, key="alta_%s" % categoria)

        horas_pide = st.number_input(
            "Horas que solicita *", min_value=1, max_value=TOPE_SOLICITUD,
            step=1, value=1,
            help="Hasta %d horas (Adendum, art. 14, Primera, inciso b)."
                 % TOPE_SOLICITUD)

        dar_alta = st.form_submit_button("Dar de alta", type="primary")

    if dar_alta:
        nombre_alta = " ".join(str(nuevo_nombre).split()).upper()
        rfc_alta = str(nuevo_rfc).strip().upper()
        rfcs = {normalizar(r) for r in df_part["rfc"]} if not df_part.empty else set()
        if not nombre_alta or not rfc_alta:
            st.error("Faltan el nombre y el R.F.C.")
        elif not capturo.strip():
            st.error("Anota en la barra lateral quién está capturando.")
        elif len(rfc_alta) not in (12, 13):
            # 13 para persona física, 12 para moral. Un R.F.C. mal escrito es
            # lo único que después no se corrige sin rehacer el expediente.
            st.error("El R.F.C. tiene %d caracteres; deben ser 13." % len(rfc_alta))
        elif sum(horas_por_categoria.values()) == 0:
            st.error("Captura cuántas horas tiene hoy, al menos en una categoría.")
        elif sum(horas_por_categoria.values()) > TOPE_TOTAL:
            st.error("Las horas suman %d y el tope por persona es %d."
                     % (sum(horas_por_categoria.values()), TOPE_TOTAL))
        elif normalizar(rfc_alta) in rfcs:
            ya = df_part.loc[df_part["rfc"].map(normalizar) == normalizar(rfc_alta),
                             "nombre"].iloc[0]
            st.error("Ese R.F.C. ya está registrado: %s" % ya)
        else:
            try:
                # Categoría y nivel guardan la MAYOR que ostenta: es la que
                # arma el puesto que lee el sistema de asignación. El reparto
                # completo queda en las seis columnas de horas.
                mayor = max((c for c in CATEGORIAS if horas_por_categoria[c] > 0),
                            key=CATEGORIAS.index)
                categoria, nivel = mayor.rsplit(" ", 1)
                folio_nuevo = siguiente_folio(df_part)
                agregar_varios("PARTICIPANTES", [{
                    "folio": folio_nuevo,
                    "fecha_recepcion": ahora(), "nombre": nombre_alta,
                    "rfc": rfc_alta, "estatus": ESTATUS[0], "no_expediente": "",
                    "categoria_actual": categoria, "nivel_actual": nivel,
                    "horas_actuales": sum(horas_por_categoria.values()),
                    "horas_solicitadas": int(horas_pide),
                    "capturo": capturo.strip(), "actualizado": ahora(),
                    **{columna_horas(c): int(h) if h else ""
                       for c, h in horas_por_categoria.items()}}])
            except Exception as err:
                st.error("No se pudo dar de alta: %s" % err)
            else:
                olvidar_lo_leido()
                # Queda seleccionado en Valoración: se registra y se valora de
                # inmediato, uno por uno, sin tener que buscarlo en la lista.
                st.session_state["folio_val"] = folio_nuevo
                avisar("%s dado(a) de alta. En la sección 2 · Valoración ya "
                       "aparece hasta arriba y seleccionado." % nombre_alta)
                st.rerun()

    st.divider()
    if df_part.empty:
        st.info("Todavía no hay nadie registrado.")
    else:
        registro = pd.DataFrame({
            "Folio": df_part["folio"],
            "No. de expediente": df_part["no_expediente"],
            "Nombre del participante": df_part["nombre"],
            "R.F.C.": df_part["rfc"],
            "Categorías que ostenta": [reparto_de(r) for _, r in df_part.iterrows()],
            "Horas actuales": df_part["horas_actuales"],
            "Horas solicitadas": df_part["horas_solicitadas"],
            "Estatus": df_part["estatus"],
            "Capturó": df_part["capturo"],
            "Fecha de registro": df_part["fecha_recepcion"],
        }).sort_values("Nombre del participante")
        st.dataframe(registro, width="stretch", hide_index=True)
        st.caption("%d participante(s) registrados. El nombre se corrige en la "
                   "pestaña PARTICIPANTES del Sheet; la app baja la corrección "
                   "sola a los renglones de puntaje." % len(registro))


# --- 2. VALORACIÓN --------------------------------------------------------
if seccion == VALORACION:
    st.subheader("Valoración conforme al Tabulador de Factores y Puntajes")
    st.caption("Título Octavo, Art. 43 del Reglamento. Sólo se valoran actividades "
               "posteriores a la última promoción (Base PRIMERA.5).")

    if not hay_gente:
        st.info("Todavía no hay participantes registrados. Dalos de alta en la "
                "sección 1 · Registro.")

# La captura se salta cuando no hay a quién valorar. No se usa st.stop():
# detendría el script completo y dejaría la sección de Prelación en blanco.
if hay_gente:
    if seccion == VALORACION:
        # El último registrado va hasta arriba: se registra y se valora de
        # inmediato. El folio es consecutivo, así que ordenarlo al revés es
        # exactamente el orden de registro, del más nuevo al más viejo.
        orden_alfabetico = sorted(df_part["folio"].tolist(), reverse=True)

        # Revisado no es "tiene puntos": alguien puede tener documentos a
        # medio capturar. Revisado es que la Comisión cerró el expediente, y
        # eso se sabe porque le asignó número de expediente.
        expediente_de = {f: str(n).strip() for f, n
                         in zip(df_part["folio"], df_part["no_expediente"])
                         if str(n).strip()}

        def _rotulo_participante(f):
            """Nombre, expediente y estado. El estado va para que nadie vuelva
            a revisar a quien ya quedó listo y se trabaje doble."""
            fila = df_part.loc[df_part.folio == f].iloc[0]
            partes = [fila["nombre"]]
            if f in expediente_de:
                partes.append("exp. %s" % expediente_de[f])
                partes.append("✔ %g pts" % puntaje_de.get(f, 0))
            else:
                partes.append("PENDIENTE")
                if f in puntaje_de:
                    partes.append("%g pts capturados" % puntaje_de[f])
            return "  ·  ".join(partes)

        hechos = sum(1 for f in orden_alfabetico if f in expediente_de)
        st.caption("Revisados **%d de %d**. La lista va del último registrado "
                   "al primero, así que el que acabas de dar de alta queda "
                   "hasta arriba. ✔ es expediente al que la Comisión ya le "
                   "asignó número; PENDIENTE es el que falta."
                   % (hechos, len(orden_alfabetico)))

        folio_v = st.selectbox("Participante a valorar", orden_alfabetico,
                               format_func=_rotulo_participante, key="folio_val")

        if folio_v in expediente_de:
            st.warning("**Este expediente ya fue revisado: expediente %s, %g "
                       "puntos.** Si guardas, se reemplaza por completo lo que "
                       "tenía." % (expediente_de[folio_v], puntaje_de.get(folio_v, 0)))

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
                olvidar_lo_leido()
                # Sin st.rerun(): recargar aquí reconstruiría las tablas de
                # captura desde el Sheet y se perdería lo escrito sin guardar.
                # Se actualiza en memoria lo poco que cambió.
                df_part.loc[df_part.folio == folio_v, "no_expediente"] = \
                    exp_nuevo.strip()
                expediente_de[folio_v] = exp_nuevo.strip()
                st.success("Expediente %s asignado a %s. Lo que lleves "
                           "capturado abajo sigue intacto; recuerda guardar "
                           "la valoración." % (exp_nuevo.strip(),
                                               fila_v.get("nombre", "")))

        # El de arriba y el de abajo son el mismo guardado. Con nueve apartados
        # desplegados, el botón del final queda a varias pantallas de distancia
        # y el operador captura mucho rato sin guardar.
        guardar_arriba = st.button("Guardar valoración", type="primary",
                                   key="bval_top_%s" % folio_v,
                                   width="stretch")
        st.caption("Lo que capturas vive en la pantalla hasta que guardas. "
                   "Conviene guardar al terminar cada apartado.")

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
                "Fecha del título": a_fecha(r.get("fecha_titulo")),
            })

        COLUMNAS = ["No.", "Concepto", "Detalle", "Cantidad", "Sin valor",
                    "Notas y observaciones"]

        capturado = {}
        for factor in tb.FACTORES:
            opciones = [OPCIONES[c.clave] for c in factor.conceptos]
            # La fecha del título sólo tiene sentido en los dos bloques de
            # formación académica; en los demás apartados estorbaría.
            es_formacion = factor.clave in GRADOS_DE_FACTOR
            columnas = COLUMNAS + (["Fecha del título"] if es_formacion else [])
            base = pd.DataFrame(renglones_previos[factor.clave] or [],
                                columns=columnas)
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
                        # Cuando el apartado tiene un solo concepto —c)
                        # Antigüedad, por ejemplo— viene ya elegido: no hay
                        # nada que escoger y sólo falta escribir la cantidad.
                        "Concepto": st.column_config.SelectboxColumn(
                            "Concepto del tabulador", options=opciones,
                            width="large", required=True,
                            default=opciones[0] if len(opciones) == 1 else None),
                        "Detalle": st.column_config.TextColumn(
                            "Detalle (nombre del documento)", width="medium"),
                        "Cantidad": st.column_config.NumberColumn(
                            # En Antigüedad la cantidad son los años de
                            # servicio, no un número de documentos.
                            "Años" if len(opciones) == 1 and
                            "año" in opciones[0].lower() else "Cantidad",
                            min_value=0.0, step=1.0, default=1.0, width="small"),
                        "Sin valor": st.column_config.CheckboxColumn(
                            "Sin valor", default=False,
                            help="Documento recibido que no puntúa. Se conserva en "
                                 "el expediente con cero puntos."),
                        "Notas y observaciones": st.column_config.TextColumn(
                            "Notas y observaciones", width="medium"),
                        "Fecha del título": st.column_config.DateColumn(
                            "Fecha del título", format="DD/MM/YYYY",
                            min_value=date(1950, 1, 1), max_value=CORTE_PERFIL,
                            help="Fecha de expedición. De aquí salen los años "
                                 "de titulación del perfil del artículo 13, "
                                 "aunque el título no sume puntos."),
                    })
                capturado[factor.clave] = editado.to_dict("records")

                # Años de titulación por grado, al margen del puntaje: un
                # título marcado sin valor puede, por su antigüedad, cubrir
                # el perfil de la categoría (Adendum, Título Tercero, art. 13).
                if es_formacion:
                    cuenta = []
                    for fila in editado.to_dict("records"):
                        f = a_fecha(fila.get("Fecha del título"))
                        if f is None:
                            continue
                        n = anios_cumplidos(f)
                        cuenta.append({
                            "Título": str(fila.get("Concepto") or "").split(" — ")[0],
                            "Expedido": f.strftime("%d/%m/%Y"),
                            "Años cumplidos al %s"
                            % CORTE_PERFIL.strftime("%d/%m/%Y"):
                                n if n is not None else "aún no los cumple",
                        })
                    if cuenta:
                        st.dataframe(pd.DataFrame(cuenta), hide_index=True,
                                     width="stretch")
                    else:
                        st.caption("Captura la fecha del título: de ahí salen "
                                   "los años de titulación que piden los "
                                   "perfiles del artículo 13.")

        # Cantidades de todos los factores juntos: es lo que entiende el tabulador.
        cantidades = {}
        for renglones in capturado.values():
            for clave, cantidad in cantidades_de(renglones).items():
                cantidades[clave] = cantidades.get(clave, 0) + cantidad

        desg = tb.desglose(cantidades)
        cuadro = cuadro_general(cantidades)
        st.markdown("#### Puntaje")
        # Los siete apartados del concentrado, en dos renglones de cuatro. Con
        # los nueve factores internos en una sola fila los rótulos no se leían.
        titulos = [t for t, _ in CUADRO] + ["TOTAL DE PUNTOS"]
        for inicio in (0, 4):
            cols = st.columns(4)
            for col, titulo in zip(cols, titulos[inicio:inicio + 4]):
                col.metric(titulo, cuadro.get(titulo, desg["TOTAL"]))

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

        guardar_abajo = st.button("Guardar valoración", type="primary",
                                  key="bval_%s" % folio_v, width="stretch")
        if guardar_arriba or guardar_abajo:
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
                                "fecha_titulo": (
                                    a_fecha(r.get("Fecha del título")).isoformat()
                                    if a_fecha(r.get("Fecha del título")) else ""),
                                "capturo": capturo.strip(), "actualizado": ahora()})
                reemplazar_por_folio("PUNTAJES", folio_v, registros)
                olvidar_lo_leido()
                avisar("Valoración de %s guardada: %d documento(s), %g puntos."
                       % (fila_v.get("nombre", ""), len(registros), desg["TOTAL"]))
                st.rerun()


# --- 3. PRELACIÓN ---------------------------------------------------------
if seccion == PRELACION:
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
            titulos_por_grado, detalle_titulos = titulos_de(pts.to_dict("records"))
            convs = df_conv[df_conv.folio == r["folio"]]["no_convocatoria"].astype(str).tolist()
            participantes.append({
                "reparto": reparto_de(r),
                "titulos_por_grado": titulos_por_grado,
                "titulos": detalle_titulos,
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
                "Categoría mayor": ("%s %s" % (e["categoria_actual"],
                                               e["nivel_actual"])).strip(),
                "Categorías que ostenta": e.get("reparto", ""),
                "Horas actuales": e["horas_actuales"],
                "Horas solicitadas": e["horas_solicitadas"],
                **{"Años de %s" % g: e.get("titulos_por_grado", {}).get(g, "")
                   for g in GRADOS},
                "Títulos y años": e.get("titulos", ""),
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
                   "columna aparte: se suma dentro de a) Formación académica. "
                   "Los años por grado se cuentan al %s y sirven para el perfil "
                   "del Adendum, art. 13, que pide alternativamente doctorado, "
                   "maestría o licenciatura con cierta antigüedad. Cuentan "
                   "aunque el título no sume puntos."
                   % CORTE_PERFIL.strftime("%d/%m/%Y"))

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
        sin_categoria = int((tabla["Categoría mayor"].astype(str).str.strip() == "").sum())
        if sin_categoria:
            st.warning("%d participante(s) no traen categoría ni nivel, así que su "
                       "columna «puesto» sale vacía en ese archivo. Hay que llenarlas "
                       "en la pestaña PARTICIPANTES del Sheet antes de usarlo."
                       % sin_categoria)

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

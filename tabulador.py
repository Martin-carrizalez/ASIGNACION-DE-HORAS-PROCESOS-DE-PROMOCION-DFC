"""
Tabulador de Factores y Puntajes — Título Octavo, capítulo único, Art. 43 del
"Reglamento para la promoción del personal docente del subsistema estatal
adscrito a las instituciones del nivel superior de la SEJ, agremiados a la
Sección 47 del SNTE" (21 de noviembre de 2012).

Ninguno de los dos acuerdos modificatorios vigentes toca el Título Octavo:
- Acuerdo que modifica los artículos 9, 13, 16 y 17.
- Acuerdo que modifica los artículos 8, 14 y 18 y deroga el 15 (16/abr/2018).
Por eso los puntajes de abajo se transcriben del Reglamento original sin cambios.

Módulo sin dependencias: se puede importar y probar fuera de Streamlit.
"""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Concepto:
    """Un renglón capturable del tabulador."""
    clave: str
    etiqueta: str
    puntos: float          # puntos por unidad
    tope_cantidad: int = 0  # 0 = sin tope de unidades


@dataclass(frozen=True)
class Factor:
    """Un bloque del tabulador (los incisos a) a g) del Art. 43)."""
    clave: str
    etiqueta: str
    conceptos: tuple
    excluyente: bool = False   # True = sólo cuenta el concepto de mayor puntaje
    tope_puntos: float = 0.0   # 0 = sin tope de puntos
    nota: str = ""


# --- a) Formación académica ------------------------------------------------
# El Reglamento separa "Estudios en educación" de "Otros estudios" y dice
# expresamente que el segundo se ADICIONA al primero.

F_EDU = Factor(
    clave="FORM_EDU",
    etiqueta="a.1) Formación académica — Estudios en educación",
    excluyente=True,  # "El puntaje no es acumulable entre sí" (nota del Reglamento)
    conceptos=(
        Concepto("LIC_EDU", "Licenciatura (titulado)", 20),
        Concepto("MAE_EDU", "Maestría en Educación (titulado)", 30),
        Concepto("DOC_EDU", "Doctorado en Educación con antecedente de Maestría (titulado)", 40),
    ),
    nota="No acumulable entre sí: sólo se toma el grado de mayor puntaje.",
)

F_OTROS = Factor(
    clave="FORM_OTROS",
    etiqueta="a.2) Formación académica — Otros estudios (no relacionados con educación)",
    excluyente=False,  # ver nota: el Reglamento no lo resuelve expresamente
    conceptos=(
        Concepto("LIC_OTR", "Licenciatura (titulado)", 15),
        Concepto("MAE_OTR", "Maestría con antecedente de Licenciatura (titulado)", 25),
        Concepto("DOC_OTR", "Doctorado con antecedente de Maestría (titulado)", 35),
    ),
    nota=("Este puntaje se adiciona a 'Estudios en educación'. El Reglamento NO dice "
          "si los grados de este bloque son acumulables entre sí; si se captura más "
          "de uno, la Comisión debe dejar constancia del criterio aplicado en el acta."),
)

# --- b) Actualización profesional ------------------------------------------
# El formato histórico de la Comisión separa el inciso b) en dos bloques y
# aplica el tope de 10 puntos SÓLO al segundo (la nota "Sólo se considera
# máximo diez puntos" va al pie de la tabla de cursos como asistente, no de
# los diplomados). El cuadro general suma b1 + b2 sin tope conjunto.

F_ACTUALIZACION = Factor(
    clave="ACTUALIZACION",
    etiqueta="b1) Actualización profesional inherente a la función",
    conceptos=(
        Concepto("DIPLOMADO", "Diplomado acreditado y con reconocimiento oficial", 3),
        Concepto("ESPECIALIDAD", "Especialidad acreditada y con reconocimiento oficial", 5),
    ),
    nota="Sin tope. El tope de 10 puntos corresponde al bloque b2.",
)

F_CURSOS_TOMADOS = Factor(
    clave="CURSOS_TOMADOS",
    etiqueta="b2) Seminario, curso o taller como asistente",
    tope_puntos=10.0,  # "Sólo se considera máximo diez puntos"
    conceptos=(
        Concepto("CURSO_40", "40 o más horas", 1.25),
        Concepto("CURSO_30_39", "30 a 39 horas", 1.0),
        Concepto("CURSO_20_29", "20 a 29 horas", 0.5),
        Concepto("CURSO_MENOR_20", "Menos de 20 horas", 0.25),
    ),
    nota=("Tope de 10 puntos. En el formato de la Comisión el tipo (seminario, "
          "curso o taller) se registra aparte; lo que puntúa es la carga horaria."),
)

# --- c) Antigüedad ---------------------------------------------------------
F_ANTIGUEDAD = Factor(
    clave="ANTIGUEDAD",
    etiqueta="c) Antigüedad en el nivel",
    conceptos=(
        Concepto("ANIO_NS", "Por cada año de servicio en el Nivel Superior", 1),
    ),
    nota="Se acredita con la constancia del Departamento de Registros y Archivo.",
)

# --- d) Experiencia profesional y/o académica ------------------------------
F_EXPERIENCIA = Factor(
    clave="EXPERIENCIA",
    etiqueta="d) Experiencia profesional y/o académica",
    conceptos=(
        Concepto("CURSO_IMP_MENOR_30", "Cursos de actualización impartidos — menos de 30 horas", 2.5),
        Concepto("CURSO_IMP_30_MAS", "Cursos de actualización impartidos — 30 horas o más", 5),
        Concepto("ASIGNATURA_NS", "Asignatura atendida en el nivel superior", 5),
    ),
)

# --- e) Productividad profesional ------------------------------------------
F_PRODUCTIVIDAD = Factor(
    clave="PRODUCTIVIDAD",
    etiqueta="e) Productividad profesional",
    conceptos=(
        Concepto("LIBRO_CIEN_AUTOR", "Libro científico o pedagógico publicado — autor", 5),
        Concepto("LIBRO_CIEN_COAUTOR", "Libro científico o pedagógico publicado — coautor", 2.5),
        Concepto("LIBRO_TEXTO_AUTOR", "Libro de texto publicado — autor", 5),
        Concepto("LIBRO_TEXTO_COAUTOR", "Libro de texto publicado — coautor", 2.5),
        Concepto("PROY_INTER_AUTOR", "Proyecto interinstitucional de la SEJ — autor", 2.0),
        Concepto("PROY_INTER_COAUTOR", "Proyecto interinstitucional de la SEJ — coautor", 1),
        Concepto("PROY_INST_AUTOR", "Proyecto institucional — autor", 1),
        Concepto("PROY_INST_COAUTOR", "Proyecto institucional — coautor", 0.5),
        Concepto("ARTICULO", "Artículo educativo publicado en revista o periódico estatal o nacional", 0.15),
        Concepto("TUTORIA", "Por grupo de tutoría atendido", 0.5),
        Concepto("CA_SIN_PROMEP", "Participar en Cuerpo Académico sin perfil PROMEP", 2),
        Concepto("PERFIL_PROMEP", "Contar con Perfil PROMEP", 3),
        Concepto("CA_LGC_PROMEP", "Participar en CA con Línea de Generación de Conocimiento con perfil PROMEP", 2),
        Concepto("COORD_CA_LGC", "Coordinar un CA con Línea de Generación de Conocimiento", 3),
        Concepto("PDI", "Diseño y ejecución del Plan de Desarrollo Institucional", 2),
        Concepto("TESIS_MAE", "Asesoría de tesis a nivel Maestría", 0.75, tope_cantidad=5),
        Concepto("TESIS_LIC", "Asesoría de tesis a nivel Licenciatura", 0.5, tope_cantidad=10),
        Concepto("SINODAL", "Sinodal a nivel Licenciatura y Maestría", 0.25, tope_cantidad=10),
        Concepto("DIS_CURSO_40", "Diseño de cursos y/o talleres aprobados — 40 horas o más", 2),
        Concepto("DIS_CURSO_30_39", "Diseño de cursos y/o talleres aprobados — 30 a 39 horas", 1.5),
        Concepto("DIS_CURSO_20_29", "Diseño de cursos y/o talleres aprobados — 20 a 29 horas", 1),
        Concepto("DIS_CURSO_10_19", "Diseño de cursos y/o talleres aprobados — 10 a 19 horas", 0.5),
    ),
    nota=("Asesoría de tesis: máximo 5 asesorados en Maestría, 10 en Licenciatura y 10 "
          "participaciones como sinodal. El diseño de cursos debe estar aprobado por la "
          "Coordinación de Formación y Actualización Docente de la SEJ."),
)

# --- f) Congresos, simposios, coloquios y foros ----------------------------
F_CONGRESOS = Factor(
    clave="CONGRESOS",
    etiqueta="f) Congresos, simposios, coloquios y foros",
    conceptos=(
        Concepto("CONG_INTERNACIONAL", "Internacional", 1.5),
        Concepto("CONG_NACIONAL", "Nacional", 1.2),
        Concepto("CONG_REGIONAL", "Regional", 1.0),
        Concepto("CONG_ESTATAL", "Estatal", 0.8),
        Concepto("CONG_MUNICIPAL", "Municipal", 0.5),
        Concepto("CONG_INSTITUCIONAL", "Institucional", 0.2),
    ),
    nota=("Por cada constancia como organizador, ponente, relator, conferenciante o "
          "coordinador, a partir de la última promoción."),
)

# --- g) Diplomados ---------------------------------------------------------
F_DIPLOMADOS = Factor(
    clave="DIPLOMADOS",
    etiqueta="g) Diplomados",
    conceptos=(
        Concepto("AUT_DIPLOMADO", "Autoría en el diseño de diplomados autorizados por la DG de Formación Continua", 5),
        Concepto("AUT_ESPECIALIDAD", "Autoría de especialidad autorizada por la Dirección de Posgrado y Unidades UPN", 7),
    ),
)

FACTORES = (F_EDU, F_OTROS, F_ACTUALIZACION, F_CURSOS_TOMADOS, F_ANTIGUEDAD,
            F_EXPERIENCIA, F_PRODUCTIVIDAD, F_CONGRESOS, F_DIPLOMADOS)

# Índices de consulta rápida.
CONCEPTOS = {c.clave: c for f in FACTORES for c in f.conceptos}
FACTOR_DE = {c.clave: f.clave for f in FACTORES for c in f.conceptos}

# Los incisos a.1) y a.2) son los que deciden el PRIMER criterio de desempate
# ("formación académica", Art. 24 inciso i y Base SÉPTIMA de la convocatoria).
FACTORES_FORMACION = ("FORM_EDU", "FORM_OTROS")


def puntos_factor(factor, cantidades):
    """
    Puntos de un factor dadas las cantidades capturadas {clave_concepto: cantidad}.
    Aplica, en este orden: tope de unidades por concepto, exclusividad del factor
    y tope de puntos del factor.
    """
    parciales = []
    for c in factor.conceptos:
        cant = cantidades.get(c.clave, 0) or 0
        if cant <= 0:
            continue
        if c.tope_cantidad:
            cant = min(cant, c.tope_cantidad)
        parciales.append(round(cant * c.puntos, 4))

    if not parciales:
        return 0.0
    # Factor excluyente: sólo sobrevive el concepto de mayor puntaje.
    total = max(parciales) if factor.excluyente else sum(parciales)
    if factor.tope_puntos:
        total = min(total, factor.tope_puntos)
    return round(total, 2)


def desglose(cantidades):
    """Devuelve {clave_factor: puntos} más la llave 'TOTAL'."""
    res = {f.clave: puntos_factor(f, cantidades) for f in FACTORES}
    res["TOTAL"] = round(sum(res.values()), 2)
    return res


def puntos_formacion(cantidades):
    """Puntaje de formación académica, usado como primer criterio de desempate."""
    d = desglose(cantidades)
    return round(sum(d[k] for k in FACTORES_FORMACION), 2)


def anios_antiguedad(cantidades):
    """Años de servicio en el nivel superior, segundo criterio de desempate."""
    return cantidades.get("ANIO_NS", 0) or 0


def ordenar_prelacion(participantes):
    """
    Ordena la lista de prelación.

    `participantes` es una lista de dicts con al menos:
        folio, nombre, cantidades (dict de conceptos)

    Criterios, en el orden que fija el Art. 24 inciso i del Reglamento y la
    Base SÉPTIMA de la convocatoria:
        1) puntaje total, de mayor a menor
        2) desempate: formación académica
        3) desempate: antigüedad en el sistema de educación superior de la SEJ
    El folio se usa al final sólo para que el orden sea estable y reproducible,
    nunca como criterio normativo.
    """
    enriquecidos = []
    for p in participantes:
        cant = p.get("cantidades", {})
        d = desglose(cant)
        enriquecidos.append({
            **p,
            "desglose": d,
            "total": d["TOTAL"],
            "formacion": puntos_formacion(cant),
            "antiguedad": anios_antiguedad(cant),
        })

    enriquecidos.sort(key=lambda x: (-x["total"], -x["formacion"],
                                     -x["antiguedad"], str(x.get("folio", ""))))

    # Lugar de prelación. Los empates perfectos en los tres criterios comparten
    # lugar: ahí la Comisión debe resolver y dejarlo asentado en el acta.
    lugar = 0
    anterior = None
    for i, e in enumerate(enriquecidos):
        llave = (e["total"], e["formacion"], e["antiguedad"])
        if llave != anterior:
            lugar = i + 1
            anterior = llave
        e["lugar"] = lugar
        e["empate_sin_resolver"] = False
    for e in enriquecidos:
        if sum(1 for o in enriquecidos if o["lugar"] == e["lugar"]) > 1:
            e["empate_sin_resolver"] = True
    return enriquecidos

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_convocatorias.py
Dirección de Formación Continua (DFC) — Área de Recursos Humanos.

Genera UN SOLO archivo .docx con TODAS las convocatorias de incremento de horas,
combinando la plantilla de Word (marcadores <<Marcador>>) con el Excel de vacancia.

La correspondencia NO es una fila = una convocatoria: se agrupa por el campo
CONVOCATORIA, de modo que una convocatoria puede arrastrar 1, 2, 3 o N claves
presupuestales y 1, 2 o 3 categorías distintas.

Uso rápido:
    1) Edita el bloque CONFIGURACIÓN (fechas que acuerde la Comisión Interna).
    2) python generar_convocatorias.py
    3) Se crea Convocatorias_Unificadas.docx junto a los archivos de entrada.

Uso por línea de comandos (opcional, para no tocar el código):
    python generar_convocatorias.py --excel vacancia.xlsx --plantilla plantilla.docx \
        --salida salida.docx --inicio "11 de octubre" --fin "18 de octubre"

Requisito: openpyxl  ->  pip install openpyxl
"""

import argparse
import os
import random
import re
import sys
import zipfile

# =====================================================================
# CONFIGURACIÓN — lo único que normalmente hay que tocar cada proceso
# =====================================================================

ARCHIVO_EXCEL     = "vacancia_unificada_formato_RH_para_incremento_2026.xlsx"
ARCHIVO_PLANTILLA = "plantilla_convocatorias_INCREMENTO_v2.docx"
ARCHIVO_SALIDA    = "Convocatorias_Unificadas.docx"

# Valor de cada marcador <<...>> de la plantilla. Si mañana cambian el formato
# del Word, basta con que el marcador siga escrito igual; el texto de alrededor
# lo pueden reescribir libremente.
FECHAS = {
    # Punto 1 de las BASES y renglón 1 del calendario
    "Fecha inicio":             "11 de octubre",
    "Fecha fin":                "18 de octubre",
    # Renglones 2 a 6 del calendario (SÉPTIMA)
    "Análisis inicio":          "19 de octubre",
    "Análisis fin":             "25 de octubre",
    "Notificación resultados":  "26 de octubre",
    "Reconsideración inicio":   "27 de octubre",
    "Reconsideración fin":      "29 de octubre",
    "Respuesta recursos":       "30 de octubre",
    "Remisión CDE":             "06 de noviembre",
    # Fecha de expedición de la convocatoria (aparece arriba del ATENTAMENTE
    # y se repite en la hoja de firmas)
    "Día convocatoria":         "08",
    "Mes convocatoria":         "octubre",
    # Año que se imprime en todo el documento
    "Año":                      "2026",
}

# True  -> suma las horas de una misma categoría: "15 hrs HORAS DE ASOCIADO B"
# False -> enlista cada registro tal cual: "3 hrs HORAS DE ASOCIADO B, 12 hrs HORAS DE ASOCIADO B"
SUMAR_HORAS_POR_CATEGORIA = True

# Texto que se imprime cuando el Excel no trae Motivo Baja porque no hubo baja
# (p. ej. PLAN DE EXPANSIÓN): la vacante no se genera por una baja de personal.
TEXTO_SIN_MOTIVO = "NO APLICA"

# El Excel trae folios con el año "arrastrado" por el autorrelleno de Excel
# (INC-65-DFC-CEPD/2027, /2028, ... /2036). Con True, el año del folio se
# empareja con FECHAS["Año"] y esas filas vuelven a ser UNA sola convocatoria.
# Ponlo en False si algún año distinto fuera intencional.
NORMALIZAR_ANIO_DEL_FOLIO = True

# True  -> el pie de página cuenta las hojas de CADA convocatoria ("Página 2 de 5")
#          y reinicia la numeración en cada una.
# False -> deja el campo original de Word, que cuenta las hojas de TODO el
#          archivo ("Página 2 de 352").
PAGINAS_POR_CONVOCATORIA = True

# El Excel trae el folio como "INC-01-DFC-CEPD/2026" y el machote nuevo lo imprime
# como "IN-01-Dirección de Formación Continua/2026". {n} es el número que venga en
# el Excel y {anio} el año del proceso. Déjalo en "" para imprimirlo tal cual.
FORMATO_FOLIO = "IN-{n}-Dirección de Formación Continua/{anio}"

# Textos del Excel que significan "sin dato"
VACIOS = ("", "nan", "none", "n/a", "na", "#n/a", "-")

# Nombres de columna esperados en el Excel (cámbialos aquí si el formato cambia)
COL = {
    "cct":         "CCT",
    "escuela":     "Escuela",
    "clave":       "Clave",
    "categoria":   "Categoria",
    "hrs":         "hrs",
    "motivo":      "Motivo Baja",
    "sustituido":  "Sustituido",
    "convocatoria": "CONVOCATORIA",
}

# =====================================================================
# 1. LECTURA DEL EXCEL
# =====================================================================

def limpiar(valor):
    """Quita NaN, espacios dobles y espacios al inicio/fin (el Excel viene con
    nombres rellenados de espacios, que si no se limpian duplican convocatorias)."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in VACIOS:
        return ""
    return re.sub(r"\s+", " ", texto)


def leer_excel(ruta):
    """Devuelve una lista de diccionarios {encabezado: valor} ya limpios."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("Falta la librería openpyxl. Instálala con:  pip install openpyxl")

    hoja = load_workbook(ruta, data_only=True).worksheets[0]
    filas = hoja.iter_rows(values_only=True)
    encabezados = [limpiar(c) for c in next(filas)]
    registros = []
    for fila in filas:
        if all(c is None or limpiar(c) == "" for c in fila):
            continue  # salta renglones vacíos al final de la hoja
        registros.append({enc: limpiar(val) for enc, val in zip(encabezados, fila)})
    return encabezados, registros


def normalizar_folio(folio, corregidos):
    """INC-65-DFC-CEPD/2033 -> INC-65-DFC-CEPD/2026 (arrastre del autorrelleno)."""
    if not NORMALIZAR_ANIO_DEL_FOLIO:
        return con_formato(folio)
    nuevo = re.sub(r"/\d{4}$", "/" + FECHAS["Año"], folio)
    if nuevo != folio:
        corregidos.append((folio, nuevo))
    return con_formato(nuevo)


def con_formato(folio):
    """INC-01-DFC-CEPD/2026 -> IN-01-Dirección de Formación Continua/2026."""
    if not FORMATO_FOLIO:
        return folio
    numero = re.search(r"[A-Za-z]+-(\d+)", folio)
    if not numero:
        return folio
    return FORMATO_FOLIO.format(n=numero.group(1), anio=FECHAS["Año"])


def agrupar_por_convocatoria(registros, corregidos):
    """Agrupa conservando el orden de aparición en el Excel (INC-01, INC-02, ...)."""
    grupos = {}
    orden = []
    for reg in registros:
        clave = normalizar_folio(reg.get(COL["convocatoria"], ""), corregidos)
        if not clave:
            continue
        if clave not in grupos:
            grupos[clave] = []
            orden.append(clave)
        grupos[clave].append(reg)
    return [(c, grupos[c]) for c in orden]


# =====================================================================
# 2. ARMADO DE LOS TEXTOS DE CADA CONVOCATORIA
# =====================================================================

def enlistar(elementos):
    """['a','b','c'] -> 'a, b y c'   (redacción en español)"""
    elementos = [e for e in elementos if e]
    if not elementos:
        return ""
    if len(elementos) == 1:
        return elementos[0]
    return ", ".join(elementos[:-1]) + " y " + elementos[-1]


def texto_plaza(filas):
    """'15 hrs HORAS DE ASOCIADO B y 5 hrs HORAS DE TITULAR A'."""
    if SUMAR_HORAS_POR_CATEGORIA:
        acumulado = {}      # categoría -> horas; dict conserva el orden de inserción
        for f in filas:
            cat = f.get(COL["categoria"], "")
            try:
                hrs = int(float(f.get(COL["hrs"], "0") or 0))
            except ValueError:
                hrs = 0
            acumulado[cat] = acumulado.get(cat, 0) + hrs
        partes = ["%d hrs %s" % (h, c) for c, h in acumulado.items()]
    else:
        partes = ["%s hrs %s" % (f.get(COL["hrs"], ""), f.get(COL["categoria"], ""))
                  for f in filas]
    return enlistar(partes)


def texto_motivo(filas):
    """'la BAJA POR DEFUNCION de AGUILAR FARIAS BENJAMIN'
       y, si el Excel no trae motivo: 'el PLAN DE EXPANSIÓN'."""
    # Una persona puede traer dos motivos distintos: se agrupan para no repetir
    # el nombre ("la BAJA POR X y la BAJA POR Y de FULANO").
    por_persona = {}
    for f in filas:
        persona = f.get(COL["sustituido"], "")
        motivo = f.get(COL["motivo"], "")
        motivos = por_persona.setdefault(persona, [])
        if motivo and motivo not in motivos:
            motivos.append(motivo)
    partes = []
    for persona, motivos in por_persona.items():
        if motivos:
            partes.append("la %s de %s" % (" y la ".join(motivos), persona))
        else:
            partes.append("%s de %s" % (TEXTO_SIN_MOTIVO, persona))
    return enlistar(partes)


def valores_de(convocatoria, filas):
    """Diccionario marcador -> texto, para una convocatoria."""
    def unicos(col):
        vistos = []
        for f in filas:
            v = f.get(COL[col], "")
            if v and v not in vistos:
                vistos.append(v)
        return vistos

    valores = dict(FECHAS)  # las fechas son iguales para todas las convocatorias
    valores.update({
        "CONVOCATORIA": convocatoria,
        "Plaza":        texto_plaza(filas),
        "Motivo":       texto_motivo(filas),
        "Clave":        ", ".join(f.get(COL["clave"], "") for f in filas),
        # Si <<Clave>> está dentro de una fila de tabla, la fila se repite una vez
        # por clave; si está en texto corrido, se usa la cadena de arriba.
        "__listas__":   {"Clave": [f.get(COL["clave"], "") for f in filas]},
        "CCT":          enlistar(unicos("cct")),
        "Escuela":      enlistar(unicos("escuela")),
        # Marcadores de la plantilla original, por si se usa el machote anterior:
        "Categoria":    texto_plaza(filas),
        "hrs":          "",
        "Sustitudo":    enlistar(unicos("sustituido")),   # (así venía escrito, con la errata)
        "Sustituido":   enlistar(unicos("sustituido")),
        "Motivo Baja":  enlistar(unicos("motivo")) + " ",
        "Fecha Inicio": FECHAS["Fecha inicio"],           # variante con mayúscula
    })
    return valores


# =====================================================================
# 3. SUSTITUCIÓN DE MARCADORES DENTRO DEL XML DE WORD
# =====================================================================

ESCAPES = (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"))

def escapar_xml(texto):
    for a, b in ESCAPES:
        texto = texto.replace(a, b)
    return texto


# En el XML, "<<Marcador>>" aparece como "&lt;&lt;Marcador&gt;&gt;", y Word puede
# partirlo en varios <w:r>/<w:t> (por el corrector ortográfico). Este patrón
# tolera etiquetas XML en medio del marcador.
PATRON_MARCADOR = re.compile(r"&lt;&lt;((?:[^<&]|<[^>]*>|&(?!lt;|gt;))*?)&gt;&gt;")

def nombre_limpio(fragmento):
    """Quita las etiquetas XML intercaladas y deja sólo el nombre del marcador."""
    return re.sub(r"\s+", " ", re.sub(r"<[^>]*>", "", fragmento)).strip()


def repetir_filas(xml, listas):
    """Repite la fila de tabla que contiene un marcador, una vez por cada valor."""
    for nombre, elementos in listas.items():
        marca = "&lt;&lt;%s&gt;&gt;" % nombre
        i = xml.find(marca)
        if i == -1:
            continue
        ini = max(xml.rfind("<w:tr ", 0, i), xml.rfind("<w:tr>", 0, i))
        if ini == -1:
            continue                     # no está en una tabla: se deja como texto
        fin = xml.find("</w:tr>", i) + len("</w:tr>")
        fila = xml[ini:fin]
        copias = "".join(fila.replace(marca, escapar_xml(v)) for v in elementos)
        xml = xml[:ini] + (copias or fila.replace(marca, "")) + xml[fin:]
    return xml


def sustituir_marcadores(xml, valores, faltantes):
    """Reemplaza cada <<Marcador>> por su valor conservando la estructura de runs."""
    def reemplazo(m):
        bruto = m.group(1)
        nombre = nombre_limpio(bruto)
        # búsqueda tolerante a mayúsculas/minúsculas y acentos del nombre
        valor = None
        for k, v in valores.items():
            if k.lower() == nombre.lower():
                valor = v
                break
        if valor is None:
            faltantes.add(nombre)
            return m.group(0)        # se deja tal cual para que se note en el Word
        valor = escapar_xml(valor)
        if "<" not in bruto:
            return valor             # caso normal: el marcador está completo en un solo run
        # Marcador partido en varios runs: el valor entra en el primer trozo de
        # texto y los demás trozos se vacían, sin tocar las etiquetas.
        piezas = re.split(r"(<[^>]*>)", bruto)
        salida, puesto = [], False
        for pieza in piezas:
            if pieza.startswith("<"):
                salida.append(pieza)
            elif not puesto:
                salida.append(valor)
                puesto = True
        return "".join(salida)

    return PATRON_MARCADOR.sub(reemplazo, xml)


def ids_unicos(xml, indice):
    """Word exige identificadores únicos en todo el documento; al clonar el cuerpo
    decenas de veces hay que renumerarlos o el archivo se reporta como dañado."""
    # IDs de párrafo (colaboración/comentarios)
    xml = re.sub(r'(w14:paraId|w14:textId)="[0-9A-Fa-f]{8}"',
                 lambda m: '%s="%08X"' % (m.group(1), random.getrandbits(31)), xml)
    # "_GoBack" es sólo la última posición del cursor: se elimina en las copias
    for m in re.finditer(r'<w:bookmarkStart w:id="(\d+)" w:name="_GoBack"\s*/>', xml):
        xml = xml.replace(m.group(0), "")
        xml = xml.replace('<w:bookmarkEnd w:id="%s"/>' % m.group(1), "")
    # Cualquier otro marcador o imagen incrustada se renumera por bloque
    xml = re.sub(r'(<w:bookmark(?:Start|End) w:id=")(\d+)"',
                 lambda m: '%s%d"' % (m.group(1), int(m.group(2)) + indice * 1000), xml)
    xml = re.sub(r'(<w:bookmarkStart [^>]*w:name=")([^"]+)"',
                 lambda m: '%s%s_%d"' % (m.group(1), m.group(2), indice), xml)
    xml = re.sub(r'(<wp:docPr id=")(\d+)"',
                 lambda m: '%s%d"' % (m.group(1), int(m.group(2)) + indice * 1000), xml)
    return xml


def separador_de_seccion(sectpr):
    """Cada convocatoria queda en su propia SECCIÓN de Word. Así el pie de página
    reinicia la numeración y el 'Página X de Y' cuenta las hojas de ESA
    convocatoria, no las del archivo completo."""
    sectpr = re.sub(r'<w:type w:val="[^"]*"/>', "", sectpr)      # nextPage (predeterminado)
    return "<w:p><w:pPr>" + sectpr + "</w:pPr></w:p>"


def listas_independientes(numbering_xml, n_bloques):
    """Las listas numeradas (1., 2., 3. ...) se reiniciarían mal de una convocatoria
    a la siguiente si todas compartieran la misma definición. Aquí se clona la
    DEFINICIÓN COMPLETA de cada lista por convocatoria, con identificadores propios
    (w:nsid y w:tmpl), que es como Word separa dos listas que no deben hablarse.
    Devuelve el numbering.xml nuevo y el mapa de reemplazos de cada bloque."""
    instancias = re.findall(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"/>',
                            numbering_xml)
    definiciones = {re.search(r'w:abstractNumId="(\d+)"', a).group(1): a
                    for a in re.findall(r'<w:abstractNum .*?</w:abstractNum>',
                                        numbering_xml, re.S)}
    if not instancias or not definiciones:
        return numbering_xml, [{} for _ in range(n_bloques)]

    sig_abs = max(int(a) for a in definiciones) + 1
    sig_num = max(int(n) for n, _ in instancias) + 1
    nuevas_def, nuevas_num, mapas = [], [], []

    for _ in range(n_bloques):
        mapa = {}
        for num_id, abs_id in instancias:
            definicion = definiciones.get(abs_id)
            if definicion is None:
                continue
            # identificadores propios: sin esto Word enlaza las copias entre sí
            definicion = re.sub(r'w:abstractNumId="\d+"', 'w:abstractNumId="%d"' % sig_abs,
                                definicion, count=1)
            definicion = re.sub(r'\s*w15:restartNumberingAfterBreak="[^"]*"', "",
                                definicion, count=1)
            definicion = re.sub(r'<w:nsid w:val="[^"]*"/>',
                                '<w:nsid w:val="%08X"/>' % random.getrandbits(31),
                                definicion, count=1)
            definicion = re.sub(r'<w:tmpl w:val="[^"]*"/>',
                                '<w:tmpl w:val="%08X"/>' % random.getrandbits(31),
                                definicion, count=1)
            nuevas_def.append(definicion)
            nuevas_num.append('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
                              % (sig_num, sig_abs))
            mapa[num_id] = str(sig_num)
            sig_abs += 1
            sig_num += 1
        mapas.append(mapa)

    # El esquema de Word exige todos los <w:abstractNum> antes de los <w:num>.
    corte = numbering_xml.index("<w:num ")
    numbering_xml = (numbering_xml[:corte] + "".join(nuevas_def) + numbering_xml[corte:]
                     ).replace("</w:numbering>", "".join(nuevas_num) + "</w:numbering>")
    return numbering_xml, mapas


def aplicar_numeracion(xml, mapa):
    if not mapa:
        return xml
    return re.sub(r'<w:numId w:val="(\d+)"/>',
                  lambda m: '<w:numId w:val="%s"/>' % mapa.get(m.group(1), m.group(1)), xml)


def generar_docx(plantilla, salida, bloques):
    """bloques = lista de diccionarios de valores, uno por convocatoria."""
    with zipfile.ZipFile(plantilla) as zin:
        nombres = zin.namelist()
        if "word/document.xml" not in nombres:
            sys.exit("La plantilla no parece un .docx válido.")
        documento = zin.read("word/document.xml").decode("utf-8")
        otros = {n: zin.read(n) for n in nombres if n != "word/document.xml"}

    # El cuerpo va de <w:body> hasta el <w:sectPr> final (que define márgenes,
    # encabezados con logotipos y pie de página con el número de página).
    ini = documento.index("<w:body>") + len("<w:body>")
    fin = documento.rindex("<w:sectPr")
    cabecera, cuerpo, cola = documento[:ini], documento[ini:fin], documento[fin:]

    sectpr = cola[:cola.index("</w:sectPr>") + len("</w:sectPr>")]
    separador = separador_de_seccion(sectpr)

    numbering = otros.get("word/numbering.xml", b"").decode("utf-8")
    if numbering:
        numbering, mapas = listas_independientes(numbering, len(bloques))
        otros["word/numbering.xml"] = numbering.encode("utf-8")
    else:
        mapas = [{} for _ in bloques]

    faltantes = set()
    partes = []
    for i, valores in enumerate(bloques):
        texto = repetir_filas(cuerpo, valores.get("__listas__", {}))
        texto = sustituir_marcadores(texto, valores, faltantes)
        partes.append(aplicar_numeracion(ids_unicos(texto, i), mapas[i]))
    nuevo_documento = cabecera + separador.join(partes) + cola

    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for nombre, datos in otros.items():       # logotipos, encabezados, pie,
            if (PAGINAS_POR_CONVOCATORIA and nombre.startswith("word/footer")):
                # NUMPAGES cuenta todo el archivo; SECTIONPAGES cuenta la sección,
                # y cada convocatoria es una sección.
                datos = datos.replace(b"NUMPAGES", b"SECTIONPAGES")
            zout.writestr(nombre, datos)          # estilos y fuentes intactos
        zout.writestr("word/document.xml", nuevo_documento.encode("utf-8"))

    return faltantes, set(re.findall(r"&lt;&lt;([^&<]+)&gt;&gt;", cuerpo))


# =====================================================================
# 4. PROGRAMA PRINCIPAL
# =====================================================================

def main():
    ap = argparse.ArgumentParser(description="Genera el Word único de convocatorias.")
    ap.add_argument("--excel", default=ARCHIVO_EXCEL)
    ap.add_argument("--plantilla", default=ARCHIVO_PLANTILLA)
    ap.add_argument("--salida", default=ARCHIVO_SALIDA)
    ap.add_argument("--inicio", help='Fecha de publicación, ej. "11 de octubre"')
    ap.add_argument("--fin", help='Fecha de cierre, ej. "18 de octubre"')
    args = ap.parse_args()

    if args.inicio:
        FECHAS["Fecha inicio"] = args.inicio
    if args.fin:
        FECHAS["Fecha fin"] = args.fin

    for ruta in (args.excel, args.plantilla):
        if not os.path.exists(ruta):
            sys.exit("No encuentro el archivo: %s" % ruta)

    encabezados, registros = leer_excel(args.excel)
    faltan_columnas = [c for c in COL.values() if c not in encabezados]
    if faltan_columnas:
        sys.exit("Al Excel le faltan columnas: %s\nColumnas encontradas: %s"
                 % (", ".join(faltan_columnas), ", ".join(encabezados)))

    corregidos = []
    grupos = agrupar_por_convocatoria(registros, corregidos)
    bloques = [valores_de(conv, filas) for conv, filas in grupos]
    faltantes, marcadores = generar_docx(args.plantilla, args.salida, bloques)

    print("Registros leídos      : %d" % len(registros))
    print("Convocatorias armadas : %d" % len(grupos))
    print("Archivo generado      : %s" % args.salida)
    print("Marcadores de la plantilla: %s" % ", ".join(sorted(marcadores)))

    # --- Avisos de revisión (no se corrige nada en automático) ---
    sin_motivo = [c for c, f in grupos if any(not r.get(COL["motivo"]) for r in f)]
    if sin_motivo:
        print("\nAVISO · sin 'Motivo Baja' en el Excel (revisar redacción): %s"
              % ", ".join(sin_motivo))
    if corregidos:
        print("\nAVISO · se emparejó el año de %d folio(s) con %s:"
              % (len(corregidos), FECHAS["Año"]))
        for viejo, nuevo in corregidos:
            print("   %s -> %s" % (viejo, nuevo))
    else:
        anios = sorted({re.search(r"/(\d{4})$", c).group(1)
                        for c, _ in grupos if re.search(r"/(\d{4})$", c)})
        if len(anios) > 1:
            print("AVISO · hay folios con años distintos: %s" % ", ".join(anios))
    if faltantes:
        print("AVISO · marcadores de la plantilla sin dato (quedaron sin sustituir): %s"
              % ", ".join(sorted(faltantes)))


if __name__ == "__main__":
    main()

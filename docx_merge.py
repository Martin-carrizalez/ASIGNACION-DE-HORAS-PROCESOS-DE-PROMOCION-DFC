# -*- coding: utf-8 -*-
"""
docx_merge.py — motor de combinación para plantillas de Word con marcadores <<Nombre>>.

No contiene ningún dato: recibe la plantilla y los valores, y devuelve el .docx en memoria.

Qué resuelve, que es lo que rompe cuando se hace a mano:
  · marcadores partidos en varios runs por el corrector de Word
  · filas de tabla que deben repetirse una vez por registro (claves, ganadores)
  · identificadores internos duplicados al clonar el cuerpo, que dañan el archivo
  · listas numeradas que continúan entre copias en lugar de reiniciar
  · encabezados, logotipos y pie de página, que se conservan intactos
"""

import io
import random
import re
import zipfile

# ---------------------------------------------------------------------------
# Sustitución de marcadores
# ---------------------------------------------------------------------------

def escapar_xml(texto):
    return (str(texto).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# <<Marcador>> aparece escapado en el XML y Word puede partirlo en varios <w:t>;
# por eso el patrón tolera etiquetas en medio.
PATRON = re.compile(r"&lt;&lt;((?:[^<&]|<[^>]*>|&(?!lt;|gt;))*?)&gt;&gt;")


def sustituir(xml, valores, faltantes=None):
    def reemplazo(m):
        bruto = m.group(1)
        nombre = re.sub(r"\s+", " ", re.sub(r"<[^>]*>", "", bruto)).strip()
        clave = next((k for k in valores if k.lower() == nombre.lower()), None)
        if clave is None:
            if faltantes is not None:
                faltantes.add(nombre)
            return m.group(0)                 # se deja visible para que se note
        valor = escapar_xml(valores[clave])
        if "<" not in bruto:
            return valor
        # Marcador partido: el valor entra en el primer trozo de texto y los
        # demás se vacían, sin tocar las etiquetas para no romper el XML.
        puesto = [False]
        piezas = []
        for pieza in re.split(r"(<[^>]*>)", bruto):
            if pieza.startswith("<"):
                piezas.append(pieza)
            elif not puesto[0]:
                piezas.append(valor)
                puesto[0] = True
        return "".join(piezas)

    return PATRON.sub(reemplazo, xml)


def repetir_filas(xml, anclas):
    """anclas = {marcador_ancla: [ {marcador: valor}, ... ]}

    La fila de tabla que contiene el marcador ancla se repite una vez por cada
    diccionario de la lista. Con la lista vacía la fila se elimina (sirve para
    las convocatorias desiertas, donde la tabla de ganadores no lleva renglones).
    """
    for ancla, registros in anclas.items():
        marca = "&lt;&lt;%s&gt;&gt;" % ancla
        i = xml.find(marca)
        if i == -1:
            continue
        ini = max(xml.rfind("<w:tr ", 0, i), xml.rfind("<w:tr>", 0, i))
        if ini == -1:
            continue                          # el marcador no está en una tabla
        fin = xml.find("</w:tr>", i) + len("</w:tr>")
        fila = xml[ini:fin]
        copias = "".join(sustituir(fila, reg) for reg in registros)
        xml = xml[:ini] + copias + xml[fin:]
    return xml


# ---------------------------------------------------------------------------
# Clonado del cuerpo: una sección por bloque
# ---------------------------------------------------------------------------

def ids_unicos(xml, indice):
    """Word exige identificadores únicos; al clonar el cuerpo hay que renumerarlos
    o el archivo se abre con el aviso de documento dañado."""
    xml = re.sub(r'(w14:paraId|w14:textId)="[0-9A-Fa-f]{8}"',
                 lambda m: '%s="%08X"' % (m.group(1), random.getrandbits(31)), xml)
    for m in re.finditer(r'<w:bookmarkStart w:id="(\d+)" w:name="_GoBack"\s*/>', xml):
        xml = xml.replace(m.group(0), "").replace('<w:bookmarkEnd w:id="%s"/>' % m.group(1), "")
    xml = re.sub(r'(<w:bookmark(?:Start|End) w:id=")(\d+)"',
                 lambda m: '%s%d"' % (m.group(1), int(m.group(2)) + indice * 1000), xml)
    xml = re.sub(r'(<w:bookmarkStart [^>]*w:name=")([^"]+)"',
                 lambda m: '%s%s_%d"' % (m.group(1), m.group(2), indice), xml)
    xml = re.sub(r'(<wp:docPr id=")(\d+)"',
                 lambda m: '%s%d"' % (m.group(1), int(m.group(2)) + indice * 1000), xml)
    return xml


def listas_independientes(numbering, n_bloques):
    """Clona la definición completa de cada lista por bloque, con identificadores
    propios (w:nsid, w:tmpl). Sin esto Word enlaza las copias y la numeración se
    reinicia donde no debe."""
    instancias = re.findall(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"/>',
                            numbering)
    definiciones = {re.search(r'w:abstractNumId="(\d+)"', a).group(1): a
                    for a in re.findall(r"<w:abstractNum .*?</w:abstractNum>", numbering, re.S)}
    if not instancias or not definiciones:
        return numbering, [{} for _ in range(n_bloques)]

    sig_abs = max(int(a) for a in definiciones) + 1
    sig_num = max(int(n) for n, _ in instancias) + 1
    nuevas_def, nuevas_num, mapas = [], [], []
    for _ in range(n_bloques):
        mapa = {}
        for num_id, abs_id in instancias:
            d = definiciones.get(abs_id)
            if d is None:
                continue
            d = re.sub(r'w:abstractNumId="\d+"', 'w:abstractNumId="%d"' % sig_abs, d, count=1)
            d = re.sub(r'\s*w15:restartNumberingAfterBreak="[^"]*"', "", d, count=1)
            d = re.sub(r'<w:nsid w:val="[^"]*"/>',
                       '<w:nsid w:val="%08X"/>' % random.getrandbits(31), d, count=1)
            d = re.sub(r'<w:tmpl w:val="[^"]*"/>',
                       '<w:tmpl w:val="%08X"/>' % random.getrandbits(31), d, count=1)
            nuevas_def.append(d)
            nuevas_num.append('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
                              % (sig_num, sig_abs))
            mapa[num_id] = str(sig_num)
            sig_abs += 1
            sig_num += 1
        mapas.append(mapa)
    corte = numbering.index("<w:num ")           # el esquema exige abstractNum antes de num
    numbering = (numbering[:corte] + "".join(nuevas_def) + numbering[corte:]
                 ).replace("</w:numbering>", "".join(nuevas_num) + "</w:numbering>")
    return numbering, mapas


def aplicar_numeracion(xml, mapa):
    if not mapa:
        return xml
    return re.sub(r'<w:numId w:val="(\d+)"/>',
                  lambda m: '<w:numId w:val="%s"/>' % mapa.get(m.group(1), m.group(1)), xml)


def generar(plantilla_bytes, bloques, paginas_por_bloque=True):
    """bloques = [ (valores, anclas), ... ]  ->  bytes del .docx

    valores: {marcador: texto}
    anclas:  {marcador_ancla: [ {marcador: valor}, ... ]} para filas repetidas
    """
    with zipfile.ZipFile(io.BytesIO(plantilla_bytes)) as zin:
        nombres = zin.namelist()
        documento = zin.read("word/document.xml").decode("utf-8")
        otros = {n: zin.read(n) for n in nombres if n != "word/document.xml"}

    ini = documento.index("<w:body>") + len("<w:body>")
    fin = documento.rindex("<w:sectPr")
    cabecera, cuerpo, cola = documento[:ini], documento[ini:fin], documento[fin:]

    # Cada bloque en su propia sección: el pie reinicia la numeración de página.
    sectpr = cola[:cola.index("</w:sectPr>") + len("</w:sectPr>")]
    separador = ("<w:p><w:pPr>" + re.sub(r'<w:type w:val="[^"]*"/>', "", sectpr) + "</w:pPr></w:p>")

    numbering = otros.get("word/numbering.xml", b"").decode("utf-8")
    if numbering:
        numbering, mapas = listas_independientes(numbering, len(bloques))
        otros["word/numbering.xml"] = numbering.encode("utf-8")
    else:
        mapas = [{} for _ in bloques]

    faltantes = set()
    partes = []
    for i, (valores, anclas) in enumerate(bloques):
        texto = repetir_filas(cuerpo, anclas or {})
        texto = sustituir(texto, valores, faltantes)
        partes.append(aplicar_numeracion(ids_unicos(texto, i), mapas[i]))

    salida = io.BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for nombre, datos in otros.items():
            if paginas_por_bloque and nombre.startswith("word/footer"):
                # NUMPAGES cuenta todo el archivo; SECTIONPAGES cuenta la sección.
                datos = datos.replace(b"NUMPAGES", b"SECTIONPAGES")
            zout.writestr(nombre, datos)
        zout.writestr("word/document.xml",
                      (cabecera + separador.join(partes) + cola).encode("utf-8"))
    return salida.getvalue(), faltantes

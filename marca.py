# -*- coding: utf-8 -*-
"""
marca.py — identidad visual de la aplicación.

Todo el CSS vive aquí para que app.py se quede con la lógica del proceso.
Los dos PNG de recursos/ se embeben en base64 porque Streamlit no sirve rutas
locales dentro de una etiqueta <img>.
"""

import base64
import pathlib

import streamlit as st

RECURSOS = pathlib.Path(__file__).parent / "recursos"

# Paleta tomada del propio logotipo: naranja de las letras, cian del neón.
NARANJA = "#F2720C"
NARANJA_CLARO = "#FF9A3C"
CIAN = "#22C7F0"
TINTA = "#1B2432"
SUAVE = "#6B7688"
LINEA = "#E3E6EC"


@st.cache_data(show_spinner=False)
def _b64(nombre):
    ruta = RECURSOS / nombre
    if not ruta.exists():
        return ""
    return base64.b64encode(ruta.read_bytes()).decode()


def estilos():
    """Hoja de estilos de toda la app. Se llama una sola vez, al arrancar."""
    st.markdown(f"""
<style>
  /* ---- lienzo ---------------------------------------------------------- */
  .stApp {{
      background:
        radial-gradient(1200px 400px at 12% -8%, #FFF1E4 0%, rgba(255,241,228,0) 60%),
        radial-gradient(900px 380px at 92% -4%, #E6F8FE 0%, rgba(230,248,254,0) 55%),
        #FBFAF8;
  }}
  .block-container {{ padding-top: 1.6rem; max-width: 1180px; }}

  /* ---- encabezado ------------------------------------------------------ */
  .cabecera {{
      display: flex; align-items: center; gap: 22px;
      padding: 10px 0 18px; border-bottom: 1px solid {LINEA}; margin-bottom: 18px;
  }}
  .marco-logo {{ position: relative; width: 190px; flex: 0 0 auto; }}
  .capa {{ width: 100%; display: block; }}
  /* La capa de LED se encima en la misma posición que la base y parpadea sola;
     la opacidad baja pero nunca llega a cero, para que se atenúen sin apagarse. */
  .capa-leds {{
      position: absolute; top: 0; left: 0;
      animation: destello 2.6s ease-in-out infinite;
  }}
  @keyframes destello {{
      0%, 100% {{ opacity: 1;    filter: brightness(1.45) drop-shadow(0 0 5px {CIAN}); }}
      45%, 55% {{ opacity: .32;  filter: brightness(.85); }}
  }}
  /* Quien prefiera no ver movimiento, no lo ve. */
  @media (prefers-reduced-motion: reduce) {{
      .capa-leds {{ animation: none; }}
  }}
  .titulo {{
      font-size: 1.65rem; font-weight: 700; color: {TINTA};
      letter-spacing: -.4px; line-height: 1.15; margin: 0;
  }}
  .bajada {{ color: {SUAVE}; font-size: .95rem; margin: 6px 0 0; }}
  .proceso {{
      display: inline-block; margin-top: 10px; padding: 4px 12px;
      border-radius: 999px; font-size: .74rem; font-weight: 700;
      letter-spacing: .6px; text-transform: uppercase;
      color: #8A3B00; background: linear-gradient(90deg, #FFE2C7, #FFF0DF);
      border: 1px solid #FFD2AC;
  }}

  /* ---- pestañas -------------------------------------------------------- */
  .stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {LINEA}; }}
  .stTabs [data-baseweb="tab"] {{
      height: 44px; padding: 0 18px; background: transparent;
      font-weight: 600; color: {SUAVE};
  }}
  .stTabs [aria-selected="true"] {{ color: {NARANJA}; }}
  .stTabs [data-baseweb="tab-highlight"] {{ background: {NARANJA}; height: 3px; }}

  /* ---- indicadores ----------------------------------------------------- */
  div[data-testid="stMetric"] {{
      background: #FFFFFF; border: 1px solid {LINEA}; border-radius: 12px;
      padding: 14px 18px; box-shadow: 0 1px 2px rgba(27,36,50,.05);
  }}
  div[data-testid="stMetric"] [data-testid="stMetricValue"] {{
      color: {NARANJA}; font-weight: 700;
  }}
  div[data-testid="stMetricLabel"] {{
      color: {SUAVE}; font-size: .78rem; text-transform: uppercase; letter-spacing: .5px;
  }}

  /* ---- botones --------------------------------------------------------- */
  .stButton > button, .stDownloadButton > button {{
      border-radius: 9px; border: 1px solid {LINEA}; font-weight: 600;
      padding: .5rem 1.1rem; transition: transform .08s ease, box-shadow .15s ease;
  }}
  .stButton > button:hover, .stDownloadButton > button:hover {{
      border-color: {NARANJA}; color: {NARANJA}; transform: translateY(-1px);
      box-shadow: 0 3px 10px rgba(242,114,12,.14);
  }}
  .stButton > button[kind="primary"] {{
      background: linear-gradient(135deg, {NARANJA}, {NARANJA_CLARO});
      border: 0; color: #fff;
  }}
  .stButton > button[kind="primary"]:hover {{ color: #fff; }}
  .stDownloadButton > button {{ width: 100%; text-align: left; }}

  /* ---- barra lateral --------------------------------------------------- */
  section[data-testid="stSidebar"] {{
      background: #FFFFFF; border-right: 1px solid {LINEA};
  }}
  section[data-testid="stSidebar"] h2 {{
      font-size: .78rem; text-transform: uppercase; letter-spacing: .8px;
      color: {NARANJA}; margin-top: 1.2rem;
  }}

  /* ---- tablas y bloques ------------------------------------------------ */
  div[data-testid="stDataFrame"] {{ border: 1px solid {LINEA}; border-radius: 10px; }}
  .tarjeta {{
      background: #fff; border: 1px solid {LINEA}; border-left: 3px solid {NARANJA};
      border-radius: 10px; padding: 12px 16px; margin-bottom: 10px;
  }}
  .tarjeta b {{ color: {TINTA}; }}
  .tarjeta span {{ color: {SUAVE}; font-size: .88rem; }}
</style>
""", unsafe_allow_html=True)


def cabecera(titulo, bajada, etiqueta=None):
    """Encabezado con el logotipo y sus LED parpadeando."""
    base, leds = _b64("dfc_logo.png"), _b64("dfc_leds.png")
    if base:
        capas = '<img src="data:image/png;base64,%s" class="capa">' % base
        if leds:
            capas += '<img src="data:image/png;base64,%s" class="capa capa-leds">' % leds
        marca = '<div class="marco-logo">%s</div>' % capas
    else:
        marca = ""
    chip = '<div class="proceso">%s</div>' % etiqueta if etiqueta else ""
    st.markdown(
        '<div class="cabecera">%s<div><p class="titulo">%s</p>'
        '<p class="bajada">%s</p>%s</div></div>' % (marca, titulo, bajada, chip),
        unsafe_allow_html=True)

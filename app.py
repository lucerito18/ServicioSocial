import os
import json
import re

import gradio as gr
import pdfplumber
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

# ---------------------------------------------------------------------
# 1. TEMAS (sistema de diseño)
# ---------------------------------------------------------------------
# Cada tema define colores y tipografías. Todo el diseño sale de aquí,
# así que cambiar de tema cambia TODA la presentación de forma coherente.
THEMES = {
    "Corporativo": {
        "bg": "F7F8FA", "surface": "E6ECF4", "text": "1B2A41", "muted": "5C6B80",
        "accent": "1F6FEB", "dark_bg": "0F2742", "dark_text": "FFFFFF",
        "dark_muted": "A9B8CC", "head_font": "Calibri", "body_font": "Calibri",
    },
    "Oscuro": {
        "bg": "12151C", "surface": "1E2430", "text": "E8ECF2", "muted": "98A2B3",
        "accent": "34D1BF", "dark_bg": "0A0C10", "dark_text": "FFFFFF",
        "dark_muted": "98A2B3", "head_font": "Calibri", "body_font": "Calibri",
    },
    "Creativo": {
        "bg": "FFF8F0", "surface": "FFE7D1", "text": "2B2118", "muted": "7A6A5A",
        "accent": "F2542D", "dark_bg": "2B2118", "dark_text": "FFF8F0",
        "dark_muted": "D8C4B0", "head_font": "Georgia", "body_font": "Calibri",
    },
    "Elegante": {
        "bg": "FFFFFF", "surface": "F1EEF9", "text": "1A1A2E", "muted": "6B6B80",
        "accent": "6D28D9", "dark_bg": "1A1033", "dark_text": "FFFFFF",
        "dark_muted": "B9A9E0", "head_font": "Georgia", "body_font": "Calibri",
    },
}

# Lienzo 16:9 (13.333" x 7.5")
SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN_X = 0.9

VALID_LAYOUTS = {"portada", "contenido", "dos_columnas", "dato_grande", "cita", "cierre"}


def rgb(hex_str):
    return RGBColor.from_string(hex_str)


# ---------------------------------------------------------------------
# 2. PROMPT (le pedimos al modelo que ELIJA el layout de cada diapositiva)
# ---------------------------------------------------------------------
SYSTEM_INSTRUCTION = """
Eres un asistente experto en diseño de presentaciones. Tu trabajo es leer
documentos y convertirlos en diapositivas claras, concisas y visualmente
variadas, pensadas para explicarse en voz alta frente a una audiencia.

Reglas que debes seguir SIEMPRE, sin excepción:
- Responde ÚNICAMENTE con JSON válido. Nada de texto antes o después,
  nada de explicaciones, nada de ``` ni marcado de código.
- Escribe en español neutro, con tono claro y profesional.
- Cada "title" debe tener máximo 8 palabras.
- Cada bullet debe tener máximo 12 palabras: una frase corta, nunca un
  párrafo ni una oración copiada tal cual del documento.
- No repitas la misma idea en dos diapositivas distintas.
- Elige para cada diapositiva el "layout" que mejor encaje con su contenido
  y VARÍALOS: no uses el mismo layout en más de dos diapositivas seguidas.

Layouts disponibles y los campos que requiere cada uno:
- "portada": SOLO la primera diapositiva. Campos: title, bullets (1 bullet
  como subtítulo).
- "contenido": lista de ideas. Campos: title, bullets (3 a 5).
- "dos_columnas": comparaciones, pros/contras, antes/después. Campos: title,
  columns = [{"heading": "...", "bullets": ["..."]}, {"heading": "...",
  "bullets": ["..."]}] (2 a 3 bullets por columna).
- "dato_grande": cuando haya una cifra o porcentaje importante. Campos:
  title, stat (la cifra, ej. "85%"), stat_label (qué significa, máx. 10
  palabras), bullets (0 a 3 de contexto).
- "cita": una frase textual o idea clave destacable. Campos: title (etiqueta
  corta), quote (máx. 25 palabras), author (fuente, puede ir vacío).
- "cierre": SOLO la última diapositiva. Campos: title, bullets (3 a 4 puntos
  clave o conclusiones).
"""

model = genai.GenerativeModel(
    "gemini-3.5-flash-lite",
    system_instruction=SYSTEM_INSTRUCTION,
    generation_config={"response_mime_type": "application/json"},
)

# Ejemplo fijo (few-shot): muestra el patrón exacto de salida esperado.
EXAMPLE_INPUT = """
El reciclaje es el proceso de convertir materiales de desecho en nuevos
productos. Ayuda a reducir el consumo de materias primas frescas, reduce
el uso de energía, reduce la contaminación del aire y del agua, y reduce
las emisiones de gases de efecto invernadero. Reciclar una tonelada de
papel ahorra cerca del 60% de la energía necesaria para fabricarlo desde
cero. Los materiales más comunes que se reciclan son el papel, el vidrio,
el plástico y los metales. Para que el reciclaje funcione bien, es
importante separar correctamente los materiales desde el hogar.
"""

EXAMPLE_OUTPUT = {
    "deck_title": "El reciclaje",
    "slides": [
        {"layout": "portada", "title": "El reciclaje",
         "bullets": ["Convertir desechos en nuevos productos"]},
        {"layout": "contenido", "title": "Beneficios ambientales",
         "bullets": ["Reduce el uso de energía",
                     "Disminuye la contaminación del aire y el agua",
                     "Reduce las emisiones de gases de efecto invernadero"]},
        {"layout": "dato_grande", "title": "Ahorro de energía",
         "stat": "60%", "stat_label": "de energía ahorrada al reciclar papel",
         "bullets": ["Menos materia prima nueva", "Menos emisiones en producción"]},
        {"layout": "dos_columnas", "title": "Qué reciclar y cómo",
         "columns": [
             {"heading": "Materiales comunes",
              "bullets": ["Papel y cartón", "Vidrio", "Plástico y metales"]},
             {"heading": "Buenas prácticas",
              "bullets": ["Separar desde el hogar", "Limpiar los envases"]}]},
        {"layout": "cierre", "title": "Conclusión",
         "bullets": ["El reciclaje reduce el impacto ambiental",
                     "La separación correcta desde el hogar es clave"]},
    ],
}


def extract_text(pdf_path):
    """Saca todo el texto del PDF, página por página."""
    text_parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
    return "\n".join(text_parts)


def ask_gemini_for_slides(text):
    """Le pide a Gemini la estructura de diapositivas (JSON con layout por slide)."""
    prompt = f"""
A continuación tienes un ejemplo de cómo debes transformar un texto en
diapositivas, siguiendo tus reglas:

TEXTO DE EJEMPLO:
\"\"\"{EXAMPLE_INPUT}\"\"\"

SALIDA ESPERADA PARA ESE EJEMPLO:
{json.dumps(EXAMPLE_OUTPUT, ensure_ascii=False)}

Ahora haz exactamente lo mismo con el siguiente documento real. Genera
entre 6 y 10 diapositivas (la primera "portada" y la última "cierre") y
devuelve SOLO el JSON con la misma estructura: {{"deck_title": "...",
"slides": [...]}}.

DOCUMENTO REAL:
\"\"\"
{text[:15000]}
\"\"\"
"""
    response = model.generate_content(prompt)
    raw = response.text.strip()

    # Por si el modelo agrega ```json ... ``` a pesar de la instrucción
    raw = re.sub(r"^```json", "", raw)
    raw = re.sub(r"^```", "", raw)
    raw = re.sub(r"```$", "", raw)
    raw = raw.strip()

    data = json.loads(raw)
    if isinstance(data, list):  # por si devuelve solo la lista
        data = {"slides": data}
    return data.get("deck_title", ""), data["slides"]


# ---------------------------------------------------------------------
# 3. NORMALIZACIÓN: el JSON del modelo nunca debe romper el diseño
# ---------------------------------------------------------------------
def clean_list(items, max_items):
    items = [str(i).strip() for i in (items or []) if str(i).strip()]
    return items[:max_items]


def normalize_slide(raw, index, total):
    """Garantiza que cada slide tenga los campos que su layout necesita."""
    layout = raw.get("layout", "contenido")
    if layout not in VALID_LAYOUTS:
        layout = "contenido"

    # La primera es siempre portada y la última siempre cierre
    if index == 0:
        layout = "portada"
    elif index == total - 1 and total > 1:
        layout = "cierre"
    elif layout in ("portada", "cierre"):
        layout = "contenido"

    slide = {
        "layout": layout,
        "title": str(raw.get("title", "Sin título")).strip() or "Sin título",
        "bullets": clean_list(raw.get("bullets"), 5),
    }

    if layout == "dos_columnas":
        cols = raw.get("columns") or []
        cols = [c for c in cols if isinstance(c, dict)][:2]
        if len(cols) < 2:
            layout = "contenido"
        else:
            slide["columns"] = [
                {"heading": str(c.get("heading", "")).strip(),
                 "bullets": clean_list(c.get("bullets"), 4)}
                for c in cols
            ]

    if layout == "dato_grande":
        stat = str(raw.get("stat", "")).strip()
        if not stat:
            layout = "contenido"
        else:
            slide["stat"] = stat
            slide["stat_label"] = str(raw.get("stat_label", "")).strip()
            slide["bullets"] = clean_list(raw.get("bullets"), 3)

    if layout == "cita":
        quote = str(raw.get("quote", "")).strip()
        if not quote:
            layout = "contenido"
        else:
            slide["quote"] = quote
            slide["author"] = str(raw.get("author", "")).strip()

    slide["layout"] = layout
    return slide


# ---------------------------------------------------------------------
# 4. BLOQUES DE DISEÑO REUTILIZABLES
# ---------------------------------------------------------------------
def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(color)
    shp.line.fill.background()  # sin borde
    shp.shadow.inherit = False
    # Quitamos el estilo por defecto del tema de Office (sombras/bordes),
    # para que el diseño se vea plano y limpio en PowerPoint y LibreOffice.
    style = shp._element.find(qn("p:style"))
    if style is not None:
        shp._element.remove(style)
    return shp


def add_text(slide, x, y, w, h, text, size, color, font, bold=False,
             italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = rgb(color)
    return box


def title_size(title, base=36):
    """Reduce el tamaño del título si es largo, para que no se desborde."""
    n = len(title)
    if n > 60:
        return base - 8
    if n > 40:
        return base - 4
    return base


def bullet_size(bullets, base=28):
    """Ajusta el tamaño de letra según cuántos bullets hay y cuán largos son."""
    n = len(bullets)
    size = base if n <= 3 else base - 2 if n == 4 else base - 4
    if any(len(b) > 70 for b in bullets):
        size -= 2
    return max(size, 14)


def bullet_rows(slide, x, y, w, h, bullets, size, text_color, dot_color,
                font, max_row_h=1.0):
    """Cada bullet es una fila: punto de color + texto centrado verticalmente."""
    if not bullets:
        return
    row_h = min(max_row_h, h / len(bullets))
    dot = 0.16
    for i, b in enumerate(bullets):
        top = y + i * row_h
        add_rect(slide, x, top + row_h / 2 - dot / 2, dot, dot, dot_color,
                 MSO_SHAPE.OVAL)
        add_text(slide, x + 0.45, top, w - 0.45, row_h, b, size, text_color,
                 font, anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide, t, deck_title, number, dark=False):
    color = t["dark_muted"] if dark else t["muted"]
    add_text(slide, MARGIN_X, 6.95, 8, 0.3, deck_title, 11, color, t["body_font"])
    add_text(slide, SLIDE_W - MARGIN_X - 1, 6.95, 1, 0.3, str(number), 11,
             color, t["body_font"], align=PP_ALIGN.RIGHT)


def add_title(slide, t, title, color=None):
    """Título + barrita de acento: mismo patrón en todas las diapositivas."""
    add_text(slide, MARGIN_X, 0.65, SLIDE_W - 2 * MARGIN_X, 1.1, title,
             title_size(title), color or t["text"], t["head_font"], bold=True,
             anchor=MSO_ANCHOR.BOTTOM)
    add_rect(slide, MARGIN_X, 1.95, 1.2, 0.08, t["accent"])


# ---------------------------------------------------------------------
# 5. LAYOUTS (uno por tipo de contenido)
# ---------------------------------------------------------------------
def layout_portada(slide, t, s):
    set_bg(slide, t["dark_bg"])
    add_rect(slide, 0, 0, 0.35, SLIDE_H, t["accent"])
    add_text(slide, 1.2, 1.9, 10.5, 2.6, s["title"], title_size(s["title"], 54),
             t["dark_text"], t["head_font"], bold=True, anchor=MSO_ANCHOR.BOTTOM)
    add_rect(slide, 1.2, 4.75, 1.6, 0.1, t["accent"])
    if s["bullets"]:
        add_text(slide, 1.2, 5.1, 10.5, 1.0, s["bullets"][0], 22,
                 t["dark_muted"], t["body_font"])


def layout_contenido(slide, t, s):
    set_bg(slide, t["bg"])
    add_rect(slide, 0, 0, 0.25, SLIDE_H, t["accent"])
    add_title(slide, t, s["title"])
    bullet_rows(slide, MARGIN_X, 2.4, SLIDE_W - 2 * MARGIN_X, 4.3, s["bullets"],
                bullet_size(s["bullets"]), t["text"], t["accent"], t["body_font"])


def layout_dos_columnas(slide, t, s):
    set_bg(slide, t["bg"])
    add_rect(slide, 0, 0, 0.25, SLIDE_H, t["accent"])
    add_title(slide, t, s["title"])
    card_w = (SLIDE_W - 2 * MARGIN_X - 0.4) / 2
    for i, col in enumerate(s["columns"]):
        x = MARGIN_X + i * (card_w + 0.4)
        add_rect(slide, x, 2.4, card_w, 4.25, t["surface"],
                 MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.04
        add_text(slide, x + 0.35, 2.65, card_w - 0.7, 0.6, col["heading"], 22,
                 t["accent"], t["head_font"], bold=True)
        bullet_rows(slide, x + 0.35, 3.4, card_w - 0.7, 3.0, col["bullets"],
                    bullet_size(col["bullets"], 20), t["text"], t["accent"],
                    t["body_font"], max_row_h=0.9)


def layout_dato_grande(slide, t, s):
    set_bg(slide, t["bg"])
    add_rect(slide, 0, 0, 0.25, SLIDE_H, t["accent"])
    add_title(slide, t, s["title"])
    stat_size = 96 if len(s["stat"]) <= 5 else 72 if len(s["stat"]) <= 8 else 54
    add_text(slide, MARGIN_X, 2.5, 5.6, 2.0, s["stat"], stat_size, t["accent"],
             t["head_font"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if s["stat_label"]:
        add_text(slide, MARGIN_X, 4.6, 5.6, 1.4, s["stat_label"], 20,
                 t["muted"], t["body_font"])
    if s["bullets"]:
        add_rect(slide, 6.9, 2.5, 0.04, 3.8, t["surface"])
        bullet_rows(slide, 7.3, 2.5, 5.1, 3.8, s["bullets"],
                    bullet_size(s["bullets"], 22), t["text"], t["accent"],
                    t["body_font"])


def layout_cita(slide, t, s):
    set_bg(slide, t["dark_bg"])
    add_text(slide, 1.2, 0.6, 3, 1.8, "\u201C", 150, t["accent"], t["head_font"],
             bold=True)
    q = s["quote"]
    q_size = 36 if len(q) < 90 else 30 if len(q) < 160 else 26
    add_text(slide, 1.5, 2.3, SLIDE_W - 3.0, 3.0, q, q_size, t["dark_text"],
             t["head_font"], italic=True, anchor=MSO_ANCHOR.MIDDLE)
    if s["author"]:
        add_rect(slide, 1.5, 5.5, 0.8, 0.06, t["accent"])
        add_text(slide, 1.5, 5.7, SLIDE_W - 3.0, 0.5, s["author"], 18,
                 t["dark_muted"], t["body_font"])


def layout_cierre(slide, t, s):
    set_bg(slide, t["dark_bg"])
    add_rect(slide, 0, 0, 0.35, SLIDE_H, t["accent"])
    add_title(slide, t, s["title"], color=t["dark_text"])
    bullet_rows(slide, MARGIN_X + 0.3, 2.4, SLIDE_W - 2 * MARGIN_X - 0.6, 4.3,
                s["bullets"], bullet_size(s["bullets"]), t["dark_text"],
                t["accent"], t["body_font"])


LAYOUT_FUNCS = {
    "portada": layout_portada,
    "contenido": layout_contenido,
    "dos_columnas": layout_dos_columnas,
    "dato_grande": layout_dato_grande,
    "cita": layout_cita,
    "cierre": layout_cierre,
}

DARK_LAYOUTS = {"portada", "cita", "cierre"}


def build_pptx(slides_data, output_path, theme_name="Corporativo", deck_title=""):
    """Arma el .pptx dibujando cada diapositiva con su layout y el tema elegido."""
    t = THEMES.get(theme_name, THEMES["Corporativo"])
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank = prs.slide_layouts[6]  # layout en blanco: dibujamos todo nosotros

    total = len(slides_data)
    slides = [normalize_slide(s, i, total) for i, s in enumerate(slides_data)]
    deck_title = deck_title or slides[0]["title"]

    for i, s in enumerate(slides):
        slide = prs.slides.add_slide(blank)
        LAYOUT_FUNCS[s["layout"]](slide, t, s)
        if s["layout"] != "portada":
            add_footer(slide, t, deck_title, i + 1, dark=s["layout"] in DARK_LAYOUTS)

    prs.save(output_path)


def process_pdf(pdf_file, theme_name):
    if pdf_file is None:
        raise gr.Error("Por favor sube un archivo PDF.")

    text = extract_text(pdf_file)
    if not text.strip():
        raise gr.Error(
            "No se pudo extraer texto de ese PDF (¿está escaneado como imagen?)."
        )

    deck_title, slides_data = ask_gemini_for_slides(text)

    output_path = "/tmp/presentacion_generada.pptx"
    build_pptx(slides_data, output_path, theme_name, deck_title)

    return output_path


demo = gr.Interface(
    fn=process_pdf,
    inputs=[
        gr.File(label="Sube tu PDF", file_types=[".pdf"]),
        gr.Dropdown(choices=list(THEMES), value="Corporativo", label="Tema visual"),
    ],
    outputs=gr.File(label="Descarga tu presentación (.pptx)"),
    title="Agente de IA: PDF a Diapositivas",
    description="Sube un PDF, elige un tema y el agente generará una presentación con diseños variados.",
    flagging_mode="never",
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)

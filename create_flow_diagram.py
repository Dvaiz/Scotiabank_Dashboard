import os
import math
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# -------------------------------------------------------------
# Configuration & Styling Constants
# -------------------------------------------------------------
WIDTH = 1920
HEIGHT = 1080

# Colors (Hex & RGB equivalent)
BG_COLOR = (11, 15, 25)          # Deep dark-blue-grey background
CARD_BG = (17, 24, 39)           # Dark slate-blue for card background
CARD_BORDER = (44, 58, 78)       # Subtle border grey-blue
TEXT_WHITE = (240, 244, 255)     # Primary text off-white
TEXT_MUTED = (136, 146, 176)     # Secondary text muted
TEXT_ACCENT = (200, 210, 230)    # Subtext neutral

# Accent Colors
TEAL_ACCENT = (0, 196, 180)      # Inputs and specific pills
BLUE_ACCENT = (61, 124, 244)     # Python Process
RED_ACCENT = (232, 56, 42)       # Scotiabank Red / Output
ORANGE_ACCENT = (245, 166, 35)   # Projections/Tariff highlights
GREEN_ACCENT = (39, 196, 122)    # Conversions/Quality highlights
GRAY_MUTED = (110, 120, 140)

# Fonts Loader Helper
def get_font(font_name, size):
    try:
        # Standard system font resolution in Pillow
        return ImageFont.truetype(font_name, size)
    except Exception:
        try:
            # Fallbacks for Windows if arial is not resolving as-is
            for alt_name in ["arial.ttf", "segoeui.ttf", "calibri.ttf", "tahoma.ttf"]:
                try:
                    return ImageFont.truetype(alt_name, size)
                except Exception:
                    continue
            return ImageFont.load_default()
        except Exception:
            return ImageFont.load_default()

# Robust font loading
title_font = get_font("arialbd.ttf", 36)      # Arial Bold
subtitle_font = get_font("arial.ttf", 16)      # Arial Regular
section_font = get_font("arialbd.ttf", 20)    # Arial Bold for section headers
card_title_font = get_font("arialbd.ttf", 14) # Arial Bold for card headers
card_sub_font = get_font("ariali.ttf", 10)    # Arial Italic for card subtitles
card_body_font = get_font("arial.ttf", 11)     # Arial Regular for card body
pill_font = get_font("arialbd.ttf", 12)       # Arial Bold for small pills
small_font = get_font("arial.ttf", 10)        # Arial Regular for footnotes

# Dynamic text metrics helpers
def draw_centered_text(draw, text, x_center, y, fill, font):
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]
    except Exception:
        # Fallback for older PIL versions
        w = draw.textlength(text, font=font)
    draw.text((x_center - w / 2, y), text, fill=fill, font=font)

def draw_arrow(draw, start, end, color, width=2, arrow_size=12):
    # Main line
    draw.line([start, end], fill=color, width=width)
    
    # Angle calculation
    x0, y0 = start
    x1, y1 = end
    dx = x1 - x0
    dy = y1 - y0
    angle = math.atan2(dy, dx)
    
    # Arrow head points
    p2_x = x1 - arrow_size * math.cos(angle - math.pi / 6)
    p2_y = y1 - arrow_size * math.sin(angle - math.pi / 6)
    p3_x = x1 - arrow_size * math.cos(angle + math.pi / 6)
    p3_y = y1 - arrow_size * math.sin(angle + math.pi / 6)
    
    draw.polygon([(x1, y1), (p2_x, p2_y), (p3_x, p3_y)], fill=color)

def draw_input_card(draw, x0, y0, x1, y1, filename, desc, accent_color):
    # Card background and border
    draw.rounded_rectangle([x0, y0, x1, y1], radius=8, fill=CARD_BG, outline=CARD_BORDER, width=1)
    # Left vertical accent line
    draw.rectangle([x0 + 2, y0 + 6, x0 + 7, y1 - 6], fill=accent_color)
    # Filename
    draw.text((x0 + 18, y0 + 12), filename, fill=TEXT_WHITE, font=card_title_font)
    # Description
    draw.text((x0 + 18, y0 + 34), desc, fill=TEXT_ACCENT, font=card_body_font)

def draw_pill(draw, x0, y0, x1, y1, label, desc, bg_color, border_color, border_thickness=1):
    # Base rounded rect
    draw.rounded_rectangle([x0, y0, x1, y1], radius=6, fill=bg_color, outline=border_color, width=border_thickness)
    # Bold Label
    draw.text((x0 + 15, y0 + 10), label, fill=TEXT_WHITE, font=pill_font)
    # Separator
    try:
        sep_x = x0 + 15 + draw.textbbox((0, 0), label, font=pill_font)[2] + 8
    except Exception:
        sep_x = x0 + 15 + draw.textlength(label, font=pill_font) + 8
        
    draw.text((sep_x, y0 + 9), "—", fill=TEXT_MUTED, font=card_body_font)
    # Description
    draw.text((sep_x + 15, y0 + 10), desc, fill=TEXT_ACCENT, font=card_body_font)


# -------------------------------------------------------------
# Main Diagram Generation
# -------------------------------------------------------------
def generate_diagram():
    print("Creando lienzo de 1920x1080...")
    img = Image.new("RGB", (WIDTH, HEIGHT), BG_COLOR)
    draw = ImageDraw.Draw(img)

    # 1. HEADER SECTOR
    # Gradient/Underline style
    draw.rectangle([60, 105, WIDTH - 60, 107], fill=RED_ACCENT)
    draw_centered_text(draw, "FASTCO × SCOTIABANK", 175, 45, RED_ACCENT, get_font("arialbd.ttf", 16))
    
    # Main Header
    draw.text((60, 42), "FLUJO DE TRABAJO Y ARQUITECTURA DE DATOS", fill=TEXT_WHITE, font=title_font)
    draw.text((60, 78), "Arquitectura SQL-first con fallback a Excel y salida HTML portable", fill=TEXT_MUTED, font=subtitle_font)

    # 2. COLUMN HEADERS (BACKGROUND CARDS FOR ORGANIZING)
    col_y0 = 135
    col_y1 = 1040
    
    # Setup Columns X boundaries
    X_COL1_START, X_COL1_END = 60, 530
    X_COL2_START, X_COL2_END = 650, 1270
    X_COL3_START, X_COL3_END = 1390, 1860
    
    # 2.1 HEADERS
    draw.rounded_rectangle([X_COL1_START, col_y0, X_COL1_END, col_y0 + 55], radius=8, fill=TEAL_ACCENT)
    draw_centered_text(draw, "1. FUENTES DE DATOS (INPUTS)", (X_COL1_START + X_COL1_END)//2, col_y0 + 16, (15, 23, 42), section_font)
    
    draw.rounded_rectangle([X_COL2_START, col_y0, X_COL2_END, col_y0 + 55], radius=8, fill=BLUE_ACCENT)
    draw_centered_text(draw, "2. MOTOR ETL & PROCESAMIENTO (PYTHON)", (X_COL2_START + X_COL2_END)//2, col_y0 + 16, (255, 255, 255), section_font)
    
    draw.rounded_rectangle([X_COL3_START, col_y0, X_COL3_END, col_y0 + 55], radius=8, fill=RED_ACCENT)
    draw_centered_text(draw, "3. INTERFAZ & VISUALIZACIÓN (DASHBOARD)", (X_COL3_START + X_COL3_END)//2, col_y0 + 16, (255, 255, 255), section_font)

    # 3. COLUMN 1: FUENTES DE DATOS (SQL-first + fallback)
    inputs = [
        ("SQL: v3_Informe_x_ejecutivos_producto_dia", "Gestiones, efectividad, compromisos, monto, tiempos y turnos.", TEAL_ACCENT),
        ("SQL: ALERTAS.dbo.BITACORA", "Operaciones y montos diarios del cliente por producto.", BLUE_ACCENT),
        ("SQL: TBL_CARGAS_POR_PRODUCTO", "Registros de carga para eficiencia y comparativos.", GREEN_ACCENT),
        ("SQL: COMISIONES.dbo.TBL_VENTAS_PERIODO", "Detalle ejecutivo por periodo para ranking y monto.", TEAL_ACCENT),
        ("SQL: ALERTAS.dbo.MAPA", "Facturacion real vs provision por OT y mes.", ORANGE_ACCENT),
        ("SQL: TBL_CIERRE_CALIDAD", "Notas y calidad por agente/ciclo/producto.", GREEN_ACCENT),
        ("Embebido: OT_MAP + Tarifa", "Relacion Producto/OT y tarifas facturacion sin Excel manual.", RED_ACCENT),
        ("Fallback: Excels Data/", "Respaldo automatico cuando SQL no esta disponible.", GRAY_MUTED)
    ]
    
    card_h = 82
    gap_y = 12
    start_y = 210
    
    for i, (fn, desc, color) in enumerate(inputs):
        y0 = start_y + i * (card_h + gap_y)
        y1 = y0 + card_h
        draw_input_card(draw, X_COL1_START, y0, X_COL1_END, y1, fn, desc, color)

    # 4. COLUMN 2: PROCESS MODULES (ETL Scripts and HTML Base fusion)
    # Card 2.1: Python ETL script
    etl_y0 = 210
    etl_y1 = 550
    draw.rounded_rectangle([X_COL2_START, etl_y0, X_COL2_END, etl_y1], radius=8, fill=CARD_BG, outline=BLUE_ACCENT, width=2)
    # Title Ribbon inside Python Card
    draw.rounded_rectangle([X_COL2_START + 5, etl_y0 + 5, X_COL2_END - 5, etl_y0 + 35], radius=4, fill=(30, 41, 59))
    draw_centered_text(draw, "Data/generate_dashboard.py (Pandas ETL Engine)", (X_COL2_START + X_COL2_END)//2, etl_y0 + 10, TEAL_ACCENT, card_title_font)
    
    etl_steps = [
        "1. Extraccion SQL-first: lectura por pyodbc con fallback a Excel local.",
        "2. Normalizacion: tipos, fechas, nulos y homologacion de columnas.",
        "3. Cruce y mapeo OT: correlacion producto cliente vs producto FASTCO.",
        "4. Facturacion/proyeccion: modelo por tarifa, fijos y UF del dia (BCCh).",
        "5. Ejecutivos y tiempos: HC, TMO, productividad y ocupacion hablado/turno.",
        "6. Correlacion mensual: FASTCO vs Cliente vs Facturacion por OT.",
        "7. Serializacion: datos a constantes JS para inyeccion en template."
    ]
    for step_idx, step_txt in enumerate(etl_steps):
        draw.text((X_COL2_START + 25, etl_y0 + 55 + step_idx * 38), step_txt, fill=TEXT_WHITE, font=card_body_font)

    # Card 2.2: static template.html (the HTML source structure)
    tmpl_y0 = 575
    tmpl_y1 = 745
    draw.rounded_rectangle([X_COL2_START + 60, tmpl_y0, X_COL2_END - 60, tmpl_y1], radius=8, fill=CARD_BG, outline=TEAL_ACCENT, width=1)
    draw.rounded_rectangle([X_COL2_START + 65, tmpl_y0 + 5, X_COL2_END - 65, tmpl_y0 + 32], radius=4, fill=(30, 41, 59))
    draw_centered_text(draw, "Data/template.html (Plantilla Web UI)", (X_COL2_START + X_COL2_END)//2, tmpl_y0 + 10, TEXT_WHITE, card_title_font)
    
    tmpl_bullets = [
        "• UI en tabs con filtros de periodo/producto/ejecutivo segun contexto.",
        "• Chart.js (CDN) para visualizacion operativa y financiera.",
        "• Componentes colapsables (+/-) y tablas de detalle con drill-down."
    ]
    for bullet_idx, bullet_txt in enumerate(tmpl_bullets):
        draw.text((X_COL2_START + 82, tmpl_y0 + 48 + bullet_idx * 30), bullet_txt, fill=TEXT_ACCENT, font=card_body_font)

    # Card 2.3: FUSION DE COMPILACIÓN (Combining script outputs with html base)
    fus_y0 = 770
    fus_y1 = 964
    draw.rounded_rectangle([X_COL2_START, fus_y0, X_COL2_END, fus_y1], radius=8, fill=CARD_BG, outline=RED_ACCENT, width=2)
    draw.rounded_rectangle([X_COL2_START + 5, fus_y0 + 5, X_COL2_END - 5, fus_y0 + 35], radius=4, fill=(30, 41, 59))
    draw_centered_text(draw, "PROCESO DE CONSOLIDACIÓN (FUSIÓN)", (X_COL2_START + X_COL2_END)//2, fus_y0 + 10, RED_ACCENT, card_title_font)
    
    fus_steps = [
        "• Localiza el bloque DATA en template.html para inyeccion de constantes.",
        "• Inserta arrays JSON compilados por el ETL para cada modulo del dashboard.",
        "• Empaqueta estilos, JS y datos en un entregable HTML unico y portable.",
        "• Escribe dashboard_scotiabank_fastco.html en la raiz del proyecto."
    ]
    for fus_idx, fus_txt in enumerate(fus_steps):
        draw.text((X_COL2_START + 25, fus_y0 + 55 + fus_idx * 36), fus_txt, fill=TEXT_WHITE, font=card_body_font)


    # 5. COLUMN 3: OUTPUT INTEGRADO Y PESTAÑAS
    # Card 3.1: Output final file
    out_y0 = 210
    out_y1 = 390
    draw.rounded_rectangle([X_COL3_START, out_y0, X_COL3_END, out_y1], radius=8, fill=CARD_BG, outline=RED_ACCENT, width=2)
    # Special header for output
    draw.rounded_rectangle([X_COL3_START + 5, out_y0 + 5, X_COL3_END - 5, out_y0 + 35], radius=4, fill=(45, 15, 23))
    draw_centered_text(draw, "dashboard_scotiabank_fastco.html", (X_COL3_START + X_COL3_END)//2, out_y0 + 10, RED_ACCENT, card_title_font)
    
    out_bullets = [
        "HTML final portable: apertura directa en navegador (file://).",
        "Sin backend web para visualizacion: todo via JavaScript local.",
        "Datos ya consolidados al momento de generar el archivo.",
        "Tema visual y estado de interfaz con persistencia local."
    ]
    for o_idx, o_txt in enumerate(out_bullets):
        draw.text((X_COL3_START + 20, out_y0 + 53 + o_idx * 30), o_txt, fill=TEXT_WHITE, font=card_body_font)

    # Card 3.2: Pestañas container
    tabs_y0 = 410
    tabs_y1 = 964
    draw.rounded_rectangle([X_COL3_START, tabs_y0, X_COL3_END, tabs_y1], radius=8, fill=CARD_BG, outline=CARD_BORDER, width=1)
    draw.rounded_rectangle([X_COL3_START + 5, tabs_y0 + 5, X_COL3_END - 5, tabs_y0 + 35], radius=4, fill=(30, 41, 59))
    draw_centered_text(draw, "ESTRUCTURA DE PESTAÑAS (8 MÓDULOS DE UI)", (X_COL3_START + X_COL3_END)//2, tabs_y0 + 10, TEXT_WHITE, card_title_font)
    
    tab_pills = [
        ("Macro", "Resumen ejecutivo con KPIs centrales y tendencia global.", BLUE_ACCENT),
        ("Campanas", "Comparativos mes/dia con KPI unico y tabla lateral de valores.", BLUE_ACCENT),
        ("Mapa", "Operacion cliente con campanas desde origen SQL (CARTERA | OT).", BLUE_ACCENT),
        ("Facturacion", "Historico y proyeccion de ingresos por OT/producto.", ORANGE_ACCENT),
        ("Oportunidades", "Palancas y escenarios de crecimiento financiero.", ORANGE_ACCENT),
        ("Ejecutivos", "Ranking, HC, tiempos, TMO y ocupacion.", GREEN_ACCENT),
        ("Correlacion", "Cruce FASTCO vs cliente vs facturacion por OT.", TEAL_ACCENT),
        ("Diagnostico", "Lectura ejecutiva de situacion, riesgos y resolucion.", RED_ACCENT)
    ]
    
    tab_h = 48
    tab_gap = 10
    tab_start_y = 460
    
    for t_idx, (lbl, t_desc, accent) in enumerate(tab_pills):
        ty0 = tab_start_y + t_idx * (tab_h + tab_gap)
        ty1 = ty0 + tab_h
        draw_pill(draw, X_COL3_START + 15, ty0, X_COL3_END - 15, ty1, lbl, t_desc, (25, 33, 50), accent, border_thickness=1)

    # 6. ARROWS & CONECTIVIDAD (INTEGRATION FLOW LINES)
    # Connect Column 1 (Inputs) -> Column 2 Python ETL
    # Backbone for Input files
    backbone_x = 590
    backbone_y_start = start_y + card_h//2
    backbone_y_end = start_y + 7 * (card_h + gap_y) + card_h//2
    
    # Draw horizontal branches from each input card to backbone
    for i in range(len(inputs)):
        box_mid_y = start_y + i * (card_h + gap_y) + card_h//2
        draw.line([(X_COL1_END, box_mid_y), (backbone_x, box_mid_y)], fill=TEAL_ACCENT, width=2)
        
    # Draw vertical backbone line
    draw.line([(backbone_x, backbone_y_start), (backbone_x, backbone_y_end)], fill=TEAL_ACCENT, width=2)
    
    # Arrow from backbone to Python ETL Card
    draw_arrow(draw, (backbone_x, 380), (X_COL2_START, 380), TEAL_ACCENT, width=3, arrow_size=14)
    # Text on top of input-to-etl arrow
    draw.text((backbone_x + 10, 350), "Lectura de archivos", fill=TEAL_ACCENT, font=get_font("arialbd.ttf", 11))
    
    # Connect Python ETL & template.html to FUSION Card
    # Down arrow from Python ETL to FUSION passing on the left
    pt_fus_left_x = X_COL2_START + 30
    draw.line([(pt_fus_left_x, etl_y1), (pt_fus_left_x, fus_y0 - 15)], fill=BLUE_ACCENT, width=2)
    draw_arrow(draw, (pt_fus_left_x, fus_y0 - 15), (pt_fus_left_x, fus_y0), BLUE_ACCENT, width=2, arrow_size=10)
    draw.text((pt_fus_left_x + 6, etl_y1 + 10), "Procesamiento", fill=BLUE_ACCENT, font=get_font("arialbd.ttf", 10))
    
    # Down arrow from Python ETL to FUSION passing on the right
    pt_fus_right_x = X_COL2_END - 30
    draw.line([(pt_fus_right_x, etl_y1), (pt_fus_right_x, fus_y0 - 15)], fill=BLUE_ACCENT, width=2)
    draw_arrow(draw, (pt_fus_right_x, fus_y0 - 15), (pt_fus_right_x, fus_y0), BLUE_ACCENT, width=2, arrow_size=10)
    
    # Down arrow from template.html to FUSION Card
    tmpl_mid_x = (X_COL2_START + X_COL2_END) // 2
    draw_arrow(draw, (tmpl_mid_x, tmpl_y1), (tmpl_mid_x, fus_y0), TEAL_ACCENT, width=2, arrow_size=10)
    draw.text((tmpl_mid_x + 10, tmpl_y1 + 8), "Plantilla HTML", fill=TEAL_ACCENT, font=get_font("arialbd.ttf", 10))

    # Connect FUSION Card -> Column 3 (Output Dashboard)
    # Draw elbow arrow from Fusion right center (X_COL2_END, 867) to Output Card left center (X_COL3_START, 300)
    elbow_x0 = X_COL2_END
    elbow_y0 = fus_y0 + (fus_y1 - fus_y0)//2 # 867
    
    elbow_x1 = 1330
    elbow_y1 = out_y0 + (out_y1 - out_y0)//2 # 300
    
    draw.line([(elbow_x0, elbow_y0), (elbow_x1, elbow_y0)], fill=RED_ACCENT, width=3)
    draw.line([(elbow_x1, elbow_y0), (elbow_x1, elbow_y1)], fill=RED_ACCENT, width=3)
    draw_arrow(draw, (elbow_x1, elbow_y1), (X_COL3_START, elbow_y1), RED_ACCENT, width=3, arrow_size=14)
    
    draw.text((elbow_x0 + 15, elbow_y0 - 24), "Generación de Entregable", fill=RED_ACCENT, font=get_font("arialbd.ttf", 11))
    draw.text((elbow_x0 + 15, elbow_y0 - 10), "Unificado y empaquetado", fill=RED_ACCENT, font=get_font("arial.ttf", 10))

    # Connect Column 3 Output Card -> Column 3 Tabs Card
    draw_arrow(draw, ((X_COL3_START + X_COL3_END)//2, out_y1), ((X_COL3_START + X_COL3_END)//2, tabs_y0), RED_ACCENT, width=2, arrow_size=10)
    draw.text(((X_COL3_START + X_COL3_END)//2 + 10, out_y1 + 5), "Despliegue de Módulos", fill=RED_ACCENT, font=get_font("arialbd.ttf", 10))

    # 7. LOAD AND EMBED LOGO (If available)
    try:
        # Check in script directory or parent directory
        logo_path = Path("LOGO1.png")
        if not logo_path.exists():
            logo_path = Path(__file__).parent / "LOGO1.png"
            
        if logo_path.exists():
            print(f"Cargado logo en: {logo_path.resolve()}")
            logo = Image.open(logo_path)
            # Scale logo
            w, h = logo.size
            new_h = 55
            new_w = int(w * (new_h / h))
            logo = logo.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            # Draw subtle container for logo
            logo_x0 = WIDTH - new_w - 90
            logo_y0 = 35
            
            # Mask loading if logo of type RGBA
            if logo.mode in ('RGBA', 'LA') or (logo.mode == 'P' and 'transparency' in logo.info):
                img.paste(logo, (logo_x0, logo_y0), mask=logo.convert('RGBA'))
            else:
                img.paste(logo, (logo_x0, logo_y0))
        else:
            print("No se encontró LOGO1.png en la ejecución del script.")
    except Exception as e:
        print(f"Error al integrar el logo: {e}")

    # Footnote
    draw.text((60, HEIGHT - 35), "© FASTCO × Scotiabank Dashboard ETL Architecture | Generado electrónicamente mediante script Python", fill=TEXT_MUTED, font=small_font)
    
    # Save Image
    output_filename = "flujo_operacional_scotiabank.png"
    print(f"Guardando imagen en {output_filename}...")
    img.save(output_filename, "PNG")
    print("¡Exito!")

if __name__ == "__main__":
    generate_diagram()

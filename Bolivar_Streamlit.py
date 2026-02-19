"""
=============================================================================
SOLICITUDES DE OUTSOURCING - ARL BOLÍVAR
App Streamlit — versión con Supabase activo
=============================================================================
Dependencias (requirements.txt):
    streamlit
    supabase
    pandas
    openpyxl
    holidays
    pillow
    reportlab
    smtplib (built-in)

Variables de entorno requeridas (Streamlit Secrets o .env local):
    SUPABASE_URL
    SUPABASE_KEY
    EMAIL_USER
    EMAIL_PASS
    LOGO_URL   ← URL pública de la imagen del logo (subida a Supabase Storage)
=============================================================================
"""

import streamlit as st
import pandas as pd
import holidays
import datetime
import smtplib
import os
import io
import copy
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import openpyxl
from openpyxl.drawing.image import Image as XLImage

from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

from supabase import create_client, Client
import urllib.request

# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN GENERAL
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Solicitudes Outsourcing — ARL Bolívar",
    page_icon="📋",
    layout="wide"
)

def get_secret(key: str, default: str = "") -> str:
    try:
        return st.secrets[key]
    except Exception:
        return os.getenv(key, default)

SUPABASE_URL  = "https://teams.microsoft.com/l/message/19:meeting_MzE4NTJlOGUtYmI1NC00M2E2LTkyMTgtY2E2NjAzMTk5YzQ3@thread.v2/1771539148546?context=%7B%22contextType%22%3A%22chat%22%7D"
SUPABASE_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVreWZ3dnhta2Fnd2FvbnJiYWZrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjQ2NjE5MTIsImV4cCI6MjA4MDIzNzkxMn0.VD1QFtqxHAkfp1D_TQj4GUCD8YKzmu14oQMpiOkrDX0"
EMAIL_USER    = "notificaciones.bi.adecco@gmail.com"
EMAIL_PASS    = "bgiu ydmq derj ikns"
LOGO_URL      = "https://intranet.uexternado.edu.co/wp-content/uploads/sites/59/2020/11/arl-seguros-bolivar.jpg"

TABLE_NAME    = "solicitudes_bolivar_adecco"
BUCKET_NAME   = "solicitudes-bolivar"

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase: Client = init_supabase()

# ─────────────────────────────────────────────────────────────────────────────
# 2. LISTAS DE CIUDADES Y GRUPOS
# ─────────────────────────────────────────────────────────────────────────────

GRUPO_A_NEISY = [
    "Apia", "Armenia", "Belen De Umbria", "Circasia", "Dosquebradas", "Manizales",
    "Marmato", "Pereira", "Santa Rosa De Cabal", "Viterbo",
    "Barranquilla", "Cartagena", "Chiriguana", "El Paso", "Galapa", "Malambo",
    "Monteria", "Palmar de Varela", "Puerto Colombia", "Santa Marta", "Santo Tomas",
    "Sincelejo", "Soledad", "Turbaco", "Valledupar",
    "Acacias", "Castilla La Nueva", "Guamal", "Puerto Gaitan", "Villavicencio"
]

GRUPO_B_CAMILA = [
    "Amaga", "Apartado", "Bello", "Buritica", "Caldas", "Cisneros", "Copacabana",
    "El Carmen de Atrato", "Envigado", "Girardota", "Guarne", "Itagui", "La Estrella",
    "La Union", "Medellin", "Remedios", "Rionegro", "Sabaneta", "San Pedro De Los Milagros",
    "Santa Barbara", "Santafe De Antioquia", "Segovia",
    "Buenaventura", "Cali", "Cartago", "Dagua", "Jamundi", "Palmira", "Yumbo",
    "Popayan"
]

GRUPO_C_JINETH = [
    "Aguazul", "Anapoima", "Barrancabermeja", "Bogota, D.C.", "Bucaramanga", "Cajica",
    "Chia", "Cota", "Cucuta", "Duitama", "Facatativa", "Floridablanca", "Funza",
    "Fusagasuga", "Gachancipa", "Giron", "Granada", "Ibague", "La Calera", "Los Patios",
    "Los Santos", "Madrid", "Mitu", "Mosquera", "Neiva", "Palermo", "Piedecuesta",
    "Quibdo", "Riohacha", "Santa Maria", "Sibate", "Siberia", "Soacha", "Sogamoso",
    "Sopo", "Tenjo", "Tocancipa", "Tunja", "Usaquen", "Villa Rica", "Villanueva",
    "Villapinzon", "Villeta", "Yopal", "Zipaquira"
]

TODAS_LAS_CIUDADES = sorted(GRUPO_A_NEISY + GRUPO_B_CAMILA + GRUPO_C_JINETH)

CORREOS_ASESORAS = {
    "Neisy Bolanos":  "arelis.bolanos@adecco.com",
    "Camila Londono": "maria.londono@adecco.com",
    "Jineth Cortes":  "jineth.cortes@adecco.com",
}

CC_FIJOS = ["manuel.pimentel@adecco.com", "ingrid.bautista@adecco.com"]

CIUDADES_PRINCIPALES  = ["Bogota, D.C.", "Medellin", "Cali", "Barranquilla",
                          "Cartagena", "Bucaramanga", "Itagui"]
CIUDADES_INTERMEDIAS  = ["Villavicencio", "Neiva", "Ibague", "Pereira",
                          "Manizales", "Armenia", "Cucuta"]

TIEMPOS_RESPUESTA = {"PRINCIPAL": 5, "INTERMEDIA": 7, "ALEJADA": 9}

PROFESIONES = [
    "MEDICO", "ENFERMERO/A", "FISIOTERAPEUTA", "TERAPEUTA OCUPACIONAL",
    "FONOAUDIOLOGO/A", "PSICOLOGO/A", "HIGIENISTA ORAL", "BACTERIOLOGO/A",
    "NUTRICIONISTA", "INGENIERO AMBIENTAL", "INGENIERO AMBIENTAL Y SANITARIO",
    "INGENIERO DE PROCESOS", "INGENIERO DE PRODUCCION", "INGENIERO ELECTRICISTA",
    "INGENIERO ELECTROMECANICO", "INGENIERO INDUSTRIAL", "INGENIERO MECANICO",
    "INGENIERO QUIMICO", "INGENIERO SANITARIO", "OTRO"
]

ESPECIALIDADES = [
    "MEDICO GENERAL", "MEDICO ESPECIALISTA SST", "MEDICO LABORAL",
    "MEDICO OCUPACIONAL", "OTRO"
]

NIVELES_FORMACION = [
    "TECNICO", "TECNOLOGO", "PROFESIONAL", "ESPECIALISTA", "MAGISTER", "OTRO"
]

EXPERIENCIA_ANIOS = [
    "Menos de 2 ANOS", "De 2 a 5 ANOS", "De 5 a 9 ANOS",
    "Mayor a 10 ANOS", "Otra"
]

DIAS_SEMANA = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO", "FESTIVOS"]

SECTORES_ECONOMICOS = [
    "AGRICULTURA", "COMERCIO", "CONSTRUCCION", "EDUCACION", "FINANCIERO",
    "HIDROCARBUROS", "INDUSTRIA MANUFACTURERA", "MINERIA", "SALUD",
    "SERVICIOS", "TRANSPORTE", "OTRO"
]

# ─────────────────────────────────────────────────────────────────────────────
# 3. FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────────────────────

def obtener_asesora_y_clasificacion(ciudad: str):
    """Retorna (nombre_asesora, clasificacion, dias_habiles)."""
    if ciudad in GRUPO_A_NEISY:
        asesora = "Neisy Bolanos"
    elif ciudad in GRUPO_B_CAMILA:
        asesora = "Camila Londono"
    elif ciudad in GRUPO_C_JINETH:
        asesora = "Jineth Cortes"
    else:
        asesora = "Jineth Cortes"

    if ciudad in CIUDADES_PRINCIPALES:
        clasificacion = "PRINCIPAL"
    elif ciudad in CIUDADES_INTERMEDIAS:
        clasificacion = "INTERMEDIA"
    else:
        clasificacion = "ALEJADA"

    return asesora, clasificacion, TIEMPOS_RESPUESTA[clasificacion]


def calcular_fecha_entrega(dias_habiles: int) -> str:
    """Calcula la fecha de entrega en días hábiles colombianos desde hoy."""
    hoy = datetime.date.today()
    anios = [hoy.year, hoy.year + 1]
    festivos = holidays.country_holidays("CO", years=anios)

    if datetime.datetime.now().hour >= 12:
        hoy += datetime.timedelta(days=1)
        while hoy.weekday() >= 5 or hoy in festivos:
            hoy += datetime.timedelta(days=1)

    contador = 0
    fecha = hoy
    while contador < dias_habiles:
        if fecha.weekday() < 5 and fecha not in festivos:
            contador += 1
            if contador == dias_habiles:
                break
        fecha += datetime.timedelta(days=1)

    return fecha.strftime("%d/%m/%Y")


def generar_id_solicitud() -> int:
    """Genera un ID numérico único consultando el último en Supabase."""
    try:
        result = supabase.table(TABLE_NAME).select("id").order("id", desc=True).limit(1).execute()
        if result.data:
            return int(result.data[0]["id"]) + 1
        return 1
    except Exception:
        # Fallback: usar timestamp como número único
        return int(datetime.datetime.now().strftime("%Y%m%d%H%M%S"))


# ─────────────────────────────────────────────────────────────────────────────
# 4. GENERACIÓN DEL EXCEL DILIGENCIADO
# ─────────────────────────────────────────────────────────────────────────────

def column_letter_to_index(col: str) -> int:
    index = 0
    for char in col.upper():
        index = index * 26 + (ord(char) - ord("A")) + 1
    return index


def index_to_column_letter(index: int) -> str:
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_offset_column(col_letter: str, offset: int) -> str:
    return index_to_column_letter(column_letter_to_index(col_letter) + offset)


def diligenciar_formato_excel(datos: dict, plantilla_bytes: bytes, logo_bytes: bytes | None) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes))
    ws = wb["FORMATO"]

    row = datos

    def fill_si_no_block(q_map: dict, displacement: int):
        for q, si_cell in q_map.items():
            response = str(row.get(q, "")).strip().upper()
            if response not in ("SI", "NO"):
                continue
            col_letter = "".join(c for c in si_cell if c.isalpha())
            row_number = "".join(c for c in si_cell if c.isdigit())
            no_col_letter = get_offset_column(col_letter, displacement)
            no_cell = no_col_letter + row_number
            if response == "SI":
                ws[si_cell].value = "X"
            elif response == "NO":
                ws[no_cell].value = "X"

    ws["F8"].value  = row.get("Q2", "")
    if row.get("Q6") == "NUEVO":
        ws["Z8"].value = "X"
    elif row.get("Q6") == "REEMPLAZO":
        ws["AE8"].value = "X"

    ws["P10"].value  = row.get("Q7", "")
    ws["Q12"].value  = row.get("Q8", "")
    ws["AG12"].value = row.get("Q9", "")
    ws["M14"].value  = row.get("Q10", "")
    ws["H16"].value  = row.get("Q11", "")
    ws["AF16"].value = row.get("Q12", "")
    ws["G18"].value  = row.get("Q13", "")
    ws["X18"].value  = row.get("Q14", "")

    formacion = " - ".join(
        str(row[f"Q{i}"]) for i in range(15, 20)
        if row.get(f"Q{i}") and str(row.get(f"Q{i}", "")).strip() not in ("", "nan", "None")
    )
    ws["G23"].value = formacion

    experiencia = " - ".join(
        str(row[f"Q{i}"]) for i in range(20, 22)
        if row.get(f"Q{i}") and str(row.get(f"Q{i}", "")).strip() not in ("", "nan", "None")
    )
    ws["Q29"].value = experiencia

    ws["D35"].value = row.get("Q22", "")

    if row.get("Q23") == "FIJO":
        ws["AG35"].value = "X"
    elif row.get("Q23") == "INTERDISCIPLINARIO":
        ws["Z35"].value = "X"

    if row.get("Q35") == "150 HORAS":
        ws["R37"].value = "X"
    elif row.get("Q35") == "75 HORAS":
        ws["L37"].value = "X"

    ws["Z37"].value  = row.get("Q91", "")
    ws["O39"].value  = row.get("Q36", "")

    dias_raw = str(row.get("Q37", ""))
    dias_seleccionados = [d.strip().upper() for d in dias_raw.split(";") if d.strip()]
    casillas_dias = {
        "LUNES": "H41", "MARTES": "K41", "MIERCOLES": "N41", "JUEVES": "Q41",
        "VIERNES": "T41", "SABADO": "W41", "DOMINGO": "Z41", "FESTIVOS": "AG41"
    }
    for dia, celda in casillas_dias.items():
        ws[celda].value = "X" if dia in dias_seleccionados else ""

    ws["G43"].value = row.get("Q38", "")

    opciones_riesgo = {1: "H45", 2: "K45", 3: "M45", 4: "P45", 5: "S45"}
    try:
        riesgo = int(row.get("Q39", 0))
        if riesgo in opciones_riesgo:
            ws[opciones_riesgo[riesgo]].value = "X"
    except Exception:
        pass

    ws["AA45"].value = row.get("Q40", "")

    ws["H49"].value  = row.get("Q24", "")
    ws["AC49"].value = row.get("Q25", "")
    ws["H51"].value  = row.get("Q27", "")
    ws["AC51"].value = row.get("Q28", "")
    ws["H53"].value  = row.get("Q30", "")
    ws["AC53"].value = row.get("Q31", "")
    ws["H55"].value  = row.get("Q33", "")
    ws["AC55"].value = row.get("Q34", "")

    if row.get("Q41") == "MOTO":
        ws["R59"].value = "X"
    elif row.get("Q41") == "VEHICULO":
        ws["W59"].value = "X"
    ws["AC59"].value = row.get("Q42", "")

    def x(q): return "X" if str(row.get(q, "")).upper() == "SI" else ""
    def no(q): return "X" if str(row.get(q, "")).upper() == "NO" else ""

    ws["I62"].value  = x("Q43");  ws["L62"].value  = no("Q43")
    ws["Q62"].value  = row.get("Q44", "");  ws["AC62"].value = row.get("Q45", "")
    ws["I63"].value  = x("Q46");  ws["L63"].value  = no("Q46")
    ws["Q63"].value  = row.get("Q47", "");  ws["AC63"].value = row.get("Q48", "")
    ws["I65"].value  = x("Q49");  ws["L65"].value  = no("Q49")
    ws["Q65"].value  = row.get("Q50", "");  ws["AC65"].value = row.get("Q51", "")

    ws["H67"].value  = row.get("Q53", "")
    ws["Q67"].value  = row.get("Q54", "")
    ws["AC67"].value = row.get("Q55", "")

    comp_map = {
        "Q56": "O74", "Q57": "O76", "Q58": "O78", "Q59": "O80",
        "Q60": "O82", "Q61": "O84", "Q62": "O86", "Q63": "O88",
        "Q64": "O90", "Q65": "O92",
        "Q66": "AG74", "Q67": "AG76", "Q68": "AG78", "Q69": "AG80"
    }
    fill_si_no_block(comp_map, 3)
    ws["S84"].value = row.get("Q70", "")

    epps_map = {
        "Q71": "O98",  "Q72": "O100", "Q73": "O102", "Q74": "O104",
        "Q75": "O106", "Q76": "O108", "Q77": "O110",
        "Q78": "AG98", "Q79": "AG100","Q80": "AG102","Q81": "AG104",
        "Q82": "AG106","Q83": "AG108","Q84": "AG110"
    }
    fill_si_no_block(epps_map, 3)
    ws["U111"].value = row.get("Q85", "")

    extras_map = {
        "Q86": "O114", "Q87": "O116",
        "Q88": "AG114","Q89": "AG116"
    }
    fill_si_no_block(extras_map, 3)

    ws["A119"].value = row.get("Q90", "")

    if logo_bytes:
        try:
            img = XLImage(io.BytesIO(logo_bytes))
            img.width  = 120
            img.height = 50
            img.anchor = "B2"
            ws.add_image(img)
        except Exception as e:
            st.warning(f"No se pudo insertar el logo en el Excel: {e}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


# ─────────────────────────────────────────────────────────────────────────────
# 5. GENERACIÓN DEL PDF
# ─────────────────────────────────────────────────────────────────────────────

def generar_pdf(datos: dict, logo_bytes: bytes | None) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    COLOR_VERDE   = colors.HexColor("#1a5276")
    COLOR_VERDE_C = colors.HexColor("#d4e6f1")
    COLOR_GRIS    = colors.HexColor("#f2f3f4")

    estilo_titulo = ParagraphStyle(
        "titulo", parent=styles["Normal"],
        fontSize=13, fontName="Helvetica-Bold",
        textColor=COLOR_VERDE, alignment=TA_CENTER, spaceAfter=4
    )
    estilo_seccion = ParagraphStyle(
        "seccion", parent=styles["Normal"],
        fontSize=9, fontName="Helvetica-Bold",
        textColor=colors.white, alignment=TA_LEFT,
        backColor=COLOR_VERDE, spaceBefore=6, spaceAfter=2,
        leftIndent=4
    )
    estilo_campo = ParagraphStyle(
        "campo", parent=styles["Normal"],
        fontSize=7.5, fontName="Helvetica-Bold", textColor=COLOR_VERDE
    )
    estilo_valor = ParagraphStyle(
        "valor", parent=styles["Normal"],
        fontSize=7.5, fontName="Helvetica"
    )
    estilo_nota = ParagraphStyle(
        "nota", parent=styles["Normal"],
        fontSize=7, fontName="Helvetica-Oblique", textColor=colors.grey
    )

    row = datos
    elementos = []

    header_data = [[None, None]]
    if logo_bytes:
        try:
            logo_img = RLImage(io.BytesIO(logo_bytes), width=3*cm, height=1.2*cm)
            header_data = [[logo_img, Paragraph(
                "SOLICITUD DE OUTSOURCING DE SERVICIOS<br/>ESPECIALIZADOS DE GESTION",
                estilo_titulo)]]
        except Exception:
            header_data = [["", Paragraph(
                "SOLICITUD DE OUTSOURCING DE SERVICIOS<br/>ESPECIALIZADOS DE GESTION",
                estilo_titulo)]]
    else:
        header_data = [["", Paragraph(
            "SOLICITUD DE OUTSOURCING DE SERVICIOS<br/>ESPECIALIZADOS DE GESTION",
            estilo_titulo)]]

    tabla_header = Table(header_data, colWidths=[4*cm, 14*cm])
    tabla_header.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",  (1,0), (1,0),  "CENTER"),
        ("BACKGROUND", (0,0), (-1,-1), COLOR_GRIS),
        ("BOX", (0,0), (-1,-1), 0.5, COLOR_VERDE),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    elementos.append(tabla_header)
    elementos.append(Spacer(1, 4))

    asesora, clasificacion, dias_hab = obtener_asesora_y_clasificacion(str(row.get("Q36", "")))
    dias_totales = dias_hab + 2 if row.get("Q17") == "MEDICO" else dias_hab
    fecha_entrega = calcular_fecha_entrega(dias_totales)

    sub_data = [
        [Paragraph("ID SOLICITUD", estilo_campo),
         Paragraph(str(row.get("id_solicitud", "")), estilo_valor),
         Paragraph("FECHA", estilo_campo),
         Paragraph(str(row.get("Q2", datetime.date.today().strftime("%d/%m/%Y"))), estilo_valor)],
        [Paragraph("EMPRESA", estilo_campo),
         Paragraph("ARL SEGUROS BOLIVAR", estilo_valor),
         Paragraph("ASESORA ASIGNADA", estilo_campo),
         Paragraph(asesora, estilo_valor)],
    ]
    tabla_sub = Table(sub_data, colWidths=[3.5*cm, 6*cm, 3.5*cm, 5*cm])
    tabla_sub.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), COLOR_GRIS),
        ("BOX", (0,0), (-1,-1), 0.5, COLOR_VERDE),
        ("INNERGRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
    ]))
    elementos.append(tabla_sub)
    elementos.append(Spacer(1, 6))

    def seccion(titulo):
        p = Paragraph(f"  {titulo}", estilo_seccion)
        t = Table([[p]], colWidths=[18*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), COLOR_VERDE),
            ("TOPPADDING", (0,0), (-1,-1), 3),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))
        return t

    def fila_datos(campo, valor, campo2="", valor2=""):
        return [
            Paragraph(campo, estilo_campo),
            Paragraph(str(valor) if valor else "—", estilo_valor),
            Paragraph(campo2, estilo_campo),
            Paragraph(str(valor2) if valor2 else "", estilo_valor),
        ]

    def tabla_datos(filas, cols=[4.5*cm, 5.5*cm, 4*cm, 4*cm]):
        t = Table(filas, colWidths=cols)
        t.setStyle(TableStyle([
            ("BOX", (0,0), (-1,-1), 0.5, COLOR_VERDE),
            ("INNERGRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
            ("TOPPADDING", (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.white, COLOR_GRIS]),
        ]))
        return t

    elementos.append(seccion("I. INFORMACION GENERAL DE SOLICITUD"))
    filas = [
        fila_datos("Tipo de solicitud", row.get("Q6",""), "A quien reemplaza?", row.get("Q7","")),
        fila_datos("Razon social empresa", row.get("Q8",""), "NIT", row.get("Q9","")),
        fila_datos("AGR Solicitante", row.get("Q10",""), "Correo AGR", row.get("Q11","")),
        fila_datos("Celular AGR", row.get("Q12",""), "Direccion Sectorial", row.get("Q13","")),
        fila_datos("Nombre DS", row.get("Q14",""), "", ""),
    ]
    elementos.append(tabla_datos(filas))
    elementos.append(Spacer(1, 4))

    elementos.append(seccion("II. INFORMACION GENERAL DEL CARGO"))
    formacion = " - ".join(
        str(row.get(f"Q{i}","")) for i in range(15, 20)
        if str(row.get(f"Q{i}","")).strip() not in ("","nan","None")
    )
    experiencia = " - ".join(
        str(row.get(f"Q{i}","")) for i in range(20, 22)
        if str(row.get(f"Q{i}","")).strip() not in ("","nan","None")
    )
    filas = [
        fila_datos("Formacion academica", formacion, "Experiencia requerida", experiencia),
        fila_datos("Salario", row.get("Q22",""), "Tipo asignacion", row.get("Q23","")),
        fila_datos("Tiempo de servicio", row.get("Q35",""), "N vacantes", row.get("Q91","")),
        fila_datos("Ciudad/municipio", row.get("Q36",""), "Horario", row.get("Q38","")),
        fila_datos("Dias de servicio", row.get("Q37",""), "Clase de riesgo", row.get("Q39","")),
        fila_datos("Sector economico", row.get("Q40",""), "Transporte propio", row.get("Q41","")),
        fila_datos("Auxilio transporte", row.get("Q42",""), "", ""),
    ]
    elementos.append(tabla_datos(filas))
    elementos.append(Spacer(1, 4))

    elementos.append(seccion("2.1. DISTRIBUCION RECURSO INTERDISCIPLINARIO"))
    filas = [
        fila_datos("AGR Lider/Responsable", row.get("Q24",""), "Horas mensuales", row.get("Q25","")),
        fila_datos("AGR 1", row.get("Q27",""), "Horas mensuales", row.get("Q28","")),
        fila_datos("AGR 2", row.get("Q30",""), "Horas mensuales", row.get("Q31","")),
        fila_datos("AGR 3", row.get("Q33",""), "Horas mensuales", row.get("Q34","")),
    ]
    elementos.append(tabla_datos(filas))
    elementos.append(Spacer(1, 4))

    elementos.append(seccion("2.2. AUXILIOS AUTORIZADOS"))
    filas = [
        fila_datos("Transporte urbano", row.get("Q43",""), "Frecuencia", row.get("Q44","")),
        fila_datos("Valor transp. urbano", row.get("Q45",""), "Transp. intermunicipal", row.get("Q46","")),
        fila_datos("Frec. intermunicipal", row.get("Q47",""), "Valor intermunicipal", row.get("Q48","")),
        fila_datos("Comunicacion", row.get("Q49",""), "Frecuencia comunicacion", row.get("Q50","")),
        fila_datos("Valor comunicacion", row.get("Q51",""), "Otro auxilio", row.get("Q52","")),
    ]
    elementos.append(tabla_datos(filas))
    elementos.append(Spacer(1, 4))

    elementos.append(seccion("III. COMPETENCIAS REQUERIDAS"))
    comp_labels = {
        "Q56": "Prueba tecnica profesion", "Q57": "Prueba ofimatica",
        "Q58": "Prueba tecnica SIG",       "Q59": "Orientacion al logro",
        "Q60": "Trabajo en equipo",        "Q61": "Atencion al cliente",
        "Q62": "Comunicacion efectiva",    "Q63": "Adaptabilidad",
        "Q64": "Pensamiento analitico",    "Q65": "Innovacion",
        "Q66": "Manejo del conflicto",     "Q67": "Negociacion y persuasion",
        "Q68": "Desarrollo de relaciones", "Q69": "Liderar equipos",
    }
    comp_filas = []
    items = list(comp_labels.items())
    for i in range(0, len(items), 2):
        q1, lbl1 = items[i]
        if i+1 < len(items):
            q2, lbl2 = items[i+1]
            comp_filas.append(fila_datos(lbl1, row.get(q1,""), lbl2, row.get(q2,"")))
        else:
            comp_filas.append(fila_datos(lbl1, row.get(q1,""), "", ""))
    comp_filas.append(fila_datos("Desc. prueba tecnica", row.get("Q70",""), "", ""))
    elementos.append(tabla_datos(comp_filas))
    elementos.append(Spacer(1, 4))

    elementos.append(seccion("IV. DOTACION — EPP's — EQUIPO DE COMPUTO — CURSOS"))
    epp_labels = {
        "Q71": "Casco dielectrico",               "Q72": "Casco con barbuquejo",
        "Q73": "Protector auditivo copa",          "Q74": "Protector auditivo insercion",
        "Q75": "Monogafa antiempanante",           "Q76": "Prot. respiratoria",
        "Q77": "Prot. visual (lente)",             "Q78": "Uniforme anti fluido",
        "Q79": "Chaqueta",                         "Q80": "Camisa",
        "Q81": "Jean",                             "Q82": "Botas dielectr./antidesl.",
        "Q83": "Botas dielectr./antiperforante",   "Q84": "Otro EPP",
    }
    epp_filas = []
    items_epp = list(epp_labels.items())
    for i in range(0, len(items_epp), 2):
        q1, lbl1 = items_epp[i]
        if i+1 < len(items_epp):
            q2, lbl2 = items_epp[i+1]
            epp_filas.append(fila_datos(lbl1, row.get(q1,""), lbl2, row.get(q2,"")))
        else:
            epp_filas.append(fila_datos(lbl1, row.get(q1,""), "", ""))
    epp_filas.append(fila_datos("Otro EPP Cual?", row.get("Q85",""), "", ""))
    epp_filas.append(fila_datos("Equipo computo basico", row.get("Q86",""),
                                "Equipo computo mayor cap.", row.get("Q87","")))
    epp_filas.append(fila_datos("Curso trabajo en alturas", row.get("Q88",""),
                                "Curso espacios confinados", row.get("Q89","")))
    elementos.append(tabla_datos(epp_filas))
    elementos.append(Spacer(1, 4))

    elementos.append(seccion("V. RECOMENDACIONES PARA TENER EN CUENTA DURANTE EL PROCESO"))
    rec_data = [[Paragraph(str(row.get("Q90","—")), estilo_valor)]]
    tabla_rec = Table(rec_data, colWidths=[18*cm])
    tabla_rec.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.5, COLOR_VERDE),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("BACKGROUND", (0,0), (-1,-1), COLOR_GRIS),
    ]))
    elementos.append(tabla_rec)
    elementos.append(Spacer(1, 8))

    pie_data = [[
        Paragraph(f"Clasificacion ciudad: <b>{clasificacion}</b>  |  "
                  f"Tiempo de respuesta: <b>{dias_totales} dias habiles</b>  |  "
                  f"Fecha tentativa de entrega: <b>{fecha_entrega}</b>",
                  estilo_nota)
    ]]
    tabla_pie = Table(pie_data, colWidths=[18*cm])
    tabla_pie.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), COLOR_VERDE_C),
        ("BOX", (0,0), (-1,-1), 0.5, COLOR_VERDE),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
    ]))
    elementos.append(tabla_pie)

    doc.build(elementos)
    buffer.seek(0)
    return buffer.read()


# ─────────────────────────────────────────────────────────────────────────────
# 6. ENVÍO DE CORREO
# ─────────────────────────────────────────────────────────────────────────────

def enviar_correo(datos: dict, pdf_bytes: bytes, xlsx_bytes: bytes) -> bool:
    ciudad       = str(datos.get("Q36", "")).strip()
    asesora, clasificacion, dias_hab = obtener_asesora_y_clasificacion(ciudad)
    dias_totales = dias_hab + 2 if datos.get("Q17") == "MEDICO" else dias_hab
    fecha_entrega = calcular_fecha_entrega(dias_totales)

    correo_agr   = datos.get("Q11", "")
    id_sol       = datos.get("id_solicitud", "N/A")

    formacion = " - ".join(
        str(datos.get(f"Q{i}","")) for i in range(15, 20)
        if str(datos.get(f"Q{i}","")).strip() not in ("","nan","None")
    )

    cuerpo = f"""
    <html><body>
    <p>Estimada Asesora,</p>
    <p>Se adjunta el formulario de Solicitud de Outsourcing para el perfil:</p>
    <p><b>{formacion}</b> en <b>{ciudad}</b>.</p>

    <p><b>NOTA IMPORTANTE:</b> Segun la clasificacion de la ciudad (<b>{clasificacion}</b>),
    el tiempo de respuesta para esta solicitud es de <b>{dias_totales} dias habiles</b>.</p>
    <p>Por lo tanto la fecha tentativa de entrega sera el <b>{fecha_entrega}</b>.</p>
    <p>Numero de vacantes solicitadas: <b>{datos.get("Q91","")}</b>.</p>

    <p>Por favor, revisarlo y dar continuidad al proceso.</p>
    <br>
    <p><b>Informacion del Solicitante:</b></p>
    <ul>
      <li><b>AGR:</b> {datos.get("Q10","")}</li>
      <li><b>Correo AGR:</b> {correo_agr}</li>
      <li><b>Tipo de Afiliacion:</b> {datos.get("Q23","")}</li>
      <li><b>Empresa Afiliada:</b> {datos.get("Q8","")}</li>
    </ul>
    <br>
    <p>Atentamente,<br>Equipo de Notificaciones BI.</p>
    </body></html>
    """

    destinatario = CORREOS_ASESORAS.get(asesora, "jineth.cortes@adecco.com")
    cc_lista     = CC_FIJOS + ([correo_agr] if correo_agr else [])
    nombre_base  = f"{id_sol} - {formacion} en {ciudad}"

    msg = MIMEMultipart("alternative")
    msg["From"]    = EMAIL_USER
    msg["To"]      = destinatario
    msg["Cc"]      = "; ".join(cc_lista)
    msg["Subject"] = f"ID {id_sol} - NUEVA SOLICITUD DE OUTSOURCING: {formacion} en {ciudad}"
    msg.attach(MIMEText(cuerpo, "html"))

    parte_pdf = MIMEBase("application", "octet-stream")
    parte_pdf.set_payload(pdf_bytes)
    encoders.encode_base64(parte_pdf)
    parte_pdf.add_header("Content-Disposition", "attachment",
                         filename=("utf-8", "", f"{nombre_base}.pdf"))
    msg.attach(parte_pdf)

    parte_xlsx = MIMEBase("application", "octet-stream")
    parte_xlsx.set_payload(xlsx_bytes)
    encoders.encode_base64(parte_xlsx)
    parte_xlsx.add_header("Content-Disposition", "attachment",
                          filename=("utf-8", "", f"{nombre_base}.xlsx"))
    msg.attach(parte_xlsx)

    try:
        todos_destinatarios = [destinatario] + cc_lista
        with smtplib.SMTP("smtp.gmail.com", 587) as servidor:
            servidor.starttls()
            servidor.login(EMAIL_USER, EMAIL_PASS)
            servidor.sendmail(EMAIL_USER, todos_destinatarios, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Error al enviar correo: {e}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# 7. SUPABASE: GUARDAR DATOS Y ARCHIVOS
# ─────────────────────────────────────────────────────────────────────────────

def _bool(val: str) -> bool:
    """Convierte 'SI'/'NO' a booleano."""
    return str(val).strip().upper() == "SI"

def _num(val, default=None):
    """Convierte a número o None si está vacío."""
    try:
        v = float(val)
        return v if v != 0 else (0 if default is None else default)
    except Exception:
        return default

def _text(val) -> str | None:
    """Retorna texto limpio o None si está vacío."""
    s = str(val).strip()
    return s if s not in ("", "nan", "None") else None


def construir_registro_supabase(datos: dict, id_solicitud: int) -> dict:
    """
    Mapea todos los campos QXX del formulario a las columnas exactas
    de la tabla solicitudes_bolivar_adecco.
    """
    ahora = datetime.datetime.now()

    registro = {
        # ── Identificación ──────────────────────────────────────────────────
        "id": id_solicitud,
        "hora_de_inicio":       ahora.isoformat(),
        "hora_de_finalizacion": ahora.isoformat(),
        "creado":               ahora.isoformat(),

        # ── I. Información general ──────────────────────────────────────────
        "tipo_de_solicitud":       _text(datos.get("Q6")),
        "remplazo_a_quien":        _text(datos.get("Q7")),
        "razon_social_empresa":    _text(datos.get("Q8")),
        "nit_empresa":             _text(datos.get("Q9")),
        "agr_solicitante":         _text(datos.get("Q10")),
        "correo_electronico_agr":  _text(datos.get("Q11")),
        "numero_celular_agr":      _text(datos.get("Q12")),
        "direccion_sectorial":     _text(datos.get("Q13")),
        "nombre_director_sectorial": _text(datos.get("Q14")),

        # ── II. Información del cargo ───────────────────────────────────────
        # Formación académica: se consolida en un campo de texto
        "formacion_academica": _text(
            " | ".join(
                str(datos.get(f"Q{i}","")) for i in range(15, 20)
                if str(datos.get(f"Q{i}","")).strip() not in ("", "nan", "None")
            )
        ),
        # Q15: profesion principal → profesional_sst, profesional_especialista, ingeniero_especialista
        # Se mapea según la profesion seleccionada
        "profesional_sst":          True if "PSICOLOGO" in str(datos.get("Q15","")).upper()
                                         or "ENFERMERO" in str(datos.get("Q15","")).upper()
                                         or "FISIOTERAPIA" in str(datos.get("Q15","")).upper()
                                         else False,
        "profesional_especialista": True if "MEDICO" in str(datos.get("Q15","")).upper() else False,
        "ingeniero_especialista":   True if "INGENIERO" in str(datos.get("Q15","")).upper() else False,
        "otra_profesion":           _text(datos.get("Q19")),  # Otra formación/certificación

        # Experiencia
        "experiencia_requerida": _text(datos.get("Q20")),
        "experiencia_adicional": _text(datos.get("Q21")),

        # Salario y asignación
        "salario_fuera_tabla": _num(datos.get("Q22"), 0),
        "asignacion":          _text(datos.get("Q23")),

        # ── 2.1 Distribución AGRs ───────────────────────────────────────────
        "agr1_lider":          _text(datos.get("Q24")),
        "asignacion_horas_agr1": _num(datos.get("Q25"), 0),

        "adicionar_otro_agr":  bool(_text(datos.get("Q27"))),
        "agr2":                _text(datos.get("Q27")),
        "asignacion_horas_agr2": _num(datos.get("Q28"), 0),

        "adicionar_otro_agr1": bool(_text(datos.get("Q30"))),
        "agr3":                _text(datos.get("Q30")),
        "asignacion_horas_agr3": _num(datos.get("Q31"), 0),

        "adicionar_otro_agr2": bool(_text(datos.get("Q33"))),
        "agr4":                _text(datos.get("Q33")),
        "asignacion_horas_agr4": _num(datos.get("Q34"), 0),

        # Servicio
        "tiempo_prestacion_servicio": _text(datos.get("Q35")),
        "ciudad_municipio":           _text(datos.get("Q36")),
        "dias_de_servicio":           _text(datos.get("Q37")),
        "horario_de_servicio":        _text(datos.get("Q38")),
        "clase_de_riesgo":            _text(str(datos.get("Q39",""))),
        "sector_economico":           _text(datos.get("Q40")),
        "requiere_transporte_propio": str(datos.get("Q41","")).upper() != "NINGUNO",
        "auxilio":                    _text(str(_num(datos.get("Q42"), 0))),

        # ── 2.2 Auxilios ────────────────────────────────────────────────────
        "transporte_urbano":               _bool(datos.get("Q43")),
        "frecuencia_transporte_urbano":    _text(datos.get("Q44")),
        "valor_transporte_urbano":         _num(datos.get("Q45"), 0),

        "transporte_intermunicipal":            _bool(datos.get("Q46")),
        "frecuencia_transporte_intermunicipal": _text(datos.get("Q47")),
        "valor_transporte_intermunicipal":      _num(datos.get("Q48"), 0),

        "comunicacion":           _bool(datos.get("Q49")),
        "frecuencia_comunicacion": _text(datos.get("Q50")),
        "valor_comunicacion":     _num(datos.get("Q51"), 0),

        "otro":          _bool(datos.get("Q52")),
        "cual_otro":     _text(datos.get("Q53")),
        "frecuencia_otro": _text(datos.get("Q54")),
        "valor_otro":    _num(datos.get("Q55"), 0),

        # ── III. Competencias técnicas ──────────────────────────────────────
        "prueba_tecnica_profesion": _bool(datos.get("Q56")),
        "prueba_ofimatica":         _bool(datos.get("Q57")),
        "prueba_tecnica_sig":       _bool(datos.get("Q58")),
        "descripcion_prueba":       _text(datos.get("Q70")),

        # Competencias comportamentales → columnas de texto (SI/NO)
        "competencia_orientacion_resultados":   _text(datos.get("Q59")),
        "competencia_orientacion_cliente":      _text(datos.get("Q61")),
        "competencia_analisis_problemas":       _text(datos.get("Q64")),
        "competencia_adaptacion_cambio":        _text(datos.get("Q63")),
        "competencia_automanejo":               _text(datos.get("Q65")),
        "competencia_comunicacion":             _text(datos.get("Q62")),
        "competencia_trabajo_equipo":           _text(datos.get("Q60")),
        "competencia_desarrollo_relaciones":    _text(datos.get("Q68")),
        "competencia_liderar_equipos":          _text(datos.get("Q69")),
        "competencia_planificacion_estrategica":_text(datos.get("Q70_plan", datos.get("Q70"))),

        # ── IV. EPPs ────────────────────────────────────────────────────────
        "epp_casco_dielectrico":          _bool(datos.get("Q71")),
        "epp_casco_dielectrico_barbuquejo": _bool(datos.get("Q72")),
        "epp_protector_auditivo_copa":    _bool(datos.get("Q73")),
        "epp_protector_auditivo_insercion": _bool(datos.get("Q74")),
        "epp_monogafa_antiempanante":     _bool(datos.get("Q75")),
        "epp_proteccion_respiratoria":    _bool(datos.get("Q76")),
        "epp_proteccion_visual":          _bool(datos.get("Q77")),

        # Dotación
        "dotacion_uniforme_antifluido":  _bool(datos.get("Q78")),
        "dotacion_chaqueta":             _bool(datos.get("Q79")),
        "dotacion_camisa":               _bool(datos.get("Q80")),
        "dotacion_jean":                 _bool(datos.get("Q81")),
        "dotacion_botas_seguridad":      _bool(datos.get("Q82")),
        "dotacion_botas_antiperforante": _bool(datos.get("Q83")),
        "dotacion_otro":                 _bool(datos.get("Q84")),
        "dotacion_cual":                 _text(datos.get("Q85")),

        # Equipos
        "equipo_computo_basico":          _bool(datos.get("Q86")),
        "equipo_computo_mayor_capacidad": _bool(datos.get("Q87")),

        # Cursos
        "curso_trabajo_alturas":    _bool(datos.get("Q88")),
        "curso_espacios_confinados": _bool(datos.get("Q89")),

        # ── V. Recomendaciones ──────────────────────────────────────────────
        "recomendaciones":   _text(datos.get("Q90")),
        "numero_de_vacantes": int(datos.get("Q91", 1) or 1),
    }

    return registro


def guardar_solicitud_supabase(datos: dict, id_solicitud: int) -> bool:
    """Guarda los datos del formulario en la tabla solicitudes_bolivar_adecco."""
    try:
        registro = construir_registro_supabase(datos, id_solicitud)
        supabase.table(TABLE_NAME).insert(registro).execute()
        return True
    except Exception as e:
        st.error(f"Error guardando en Supabase: {e}")
        return False


def subir_archivo_supabase(file_bytes: bytes, path: str, content_type: str) -> bool:
    """Sube un archivo al bucket en Supabase Storage."""
    try:
        supabase.storage.from_(BUCKET_NAME).upload(
            path=path,
            file=file_bytes,
            file_options={"content-type": content_type}
        )
        return True
    except Exception as e:
        st.warning(f"No se pudo subir archivo a Storage: {e}")
        return False


def obtener_logo() -> bytes | None:
    if not LOGO_URL:
        return None
    try:
        with urllib.request.urlopen(LOGO_URL) as resp:
            return resp.read()
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# 8. INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.markdown("""
    <style>
    .stApp { background-color: #f8f9fa; }
    div[data-testid="stForm"] {
        background: white;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #dee2e6;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .section-header {
        background: #1a5276;
        color: white;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        margin: 20px 0 10px 0;
    }
    .info-box {
        background: #d4e6f1;
        border-left: 4px solid #1a5276;
        padding: 10px 14px;
        border-radius: 4px;
        font-size: 0.88rem;
        margin: 8px 0;
    }
    </style>
    """, unsafe_allow_html=True)

    col_logo, col_titulo = st.columns([1, 4])
    with col_titulo:
        st.markdown("## Solicitud de Outsourcing de Servicios Especializados de Gestion")
        st.markdown("**ARL Seguros Bolivar** — Complete el formulario y presione **Enviar Solicitud**.")

    st.divider()

    with st.form("solicitud_form", clear_on_submit=False):

        # ── I. Información general ─────────────────────────────────────────
        st.markdown('<div class="section-header">I. INFORMACION GENERAL DE SOLICITUD</div>',
                    unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            Q6  = st.selectbox("Tipo de solicitud *", ["NUEVO", "REEMPLAZO"])
        with c2:
            Q7  = st.text_input("En caso de reemplazo, a quien reemplaza?")

        c1, c2 = st.columns(2)
        with c1:
            Q8  = st.text_input("Razon social empresa afiliada *")
        with c2:
            Q9  = st.text_input("NIT de la empresa *")

        c1, c2, c3 = st.columns(3)
        with c1:
            Q10 = st.text_input("AGR Solicitante (Lider o Responsable) *")
        with c2:
            Q11 = st.text_input("Correo electronico AGR *")
        with c3:
            Q12 = st.text_input("Numero de celular AGR *")

        c1, c2 = st.columns(2)
        with c1:
            Q13 = st.text_input("Direccion Sectorial *")
        with c2:
            Q14 = st.text_input("Nombre DS *")

        # ── II. Información del cargo ──────────────────────────────────────
        st.markdown('<div class="section-header">II. INFORMACION GENERAL DEL CARGO</div>',
                    unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            Q15 = st.selectbox("Profesion / Perfil principal *", PROFESIONES)
        with c2:
            Q16 = st.text_input("Especialidad medica (si aplica)")

        c1, c2 = st.columns(2)
        with c1:
            Q17 = st.selectbox("Nivel de formacion *", NIVELES_FORMACION)
        with c2:
            Q18 = st.text_input("Titulo especifico requerido")

        Q19 = st.text_input("Otra formacion / certificacion adicional")

        c1, c2 = st.columns(2)
        with c1:
            Q20 = st.selectbox("Experiencia requerida (anos minimos) *", EXPERIENCIA_ANIOS)
        with c2:
            Q21 = st.text_input("Mencione experiencia diferente a lo anterior")

        c1, c2, c3 = st.columns(3)
        with c1:
            Q22 = st.number_input("Salario *", min_value=0, step=50000, value=0)
        with c2:
            Q23 = st.selectbox("Asignacion *", ["INTERDISCIPLINARIO", "FIJO"])
        with c3:
            Q91 = st.number_input("Numero de vacantes *", min_value=1, step=1, value=1)

        c1, c2 = st.columns(2)
        with c1:
            Q35 = st.selectbox("Tiempo de prestacion del servicio *", ["75 HORAS", "150 HORAS"])
        with c2:
            Q36 = st.selectbox("Ciudad/Municipio donde se prestara el servicio *",
                               TODAS_LAS_CIUDADES)

        c1, c2 = st.columns(2)
        with c1:
            Q37_dias = st.multiselect("Dias de servicio *", DIAS_SEMANA)
        with c2:
            Q38 = st.text_input("Horario de servicio *")

        c1, c2 = st.columns(2)
        with c1:
            Q39 = st.selectbox("Clase de riesgo *", [1, 2, 3, 4, 5])
        with c2:
            Q40 = st.selectbox("Sector economico *", SECTORES_ECONOMICOS)

        c1, c2 = st.columns(2)
        with c1:
            Q41 = st.selectbox("Requiere transporte propio?", ["NINGUNO", "MOTO", "VEHICULO"])
        with c2:
            Q42 = st.number_input("Auxilio transporte propio ($)", min_value=0, step=10000, value=0)

        # ── 2.1 Distribución interdisciplinaria ───────────────────────────
        st.markdown('<div class="section-header">2.1. DISTRIBUCION RECURSO INTERDISCIPLINARIO</div>',
                    unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            Q24 = st.text_input("AGR Lider o Responsable")
        with c2:
            Q25 = st.number_input("Horas mensuales (Lider)", min_value=0, step=1, value=0)

        agregar_agr2 = st.checkbox("Adicionar AGR 1")
        Q27, Q28 = "", 0
        if agregar_agr2:
            c1, c2 = st.columns(2)
            with c1: Q27 = st.text_input("AGR 1")
            with c2: Q28 = st.number_input("Horas mensuales (AGR 1)", min_value=0, step=1, value=0)

        agregar_agr3 = st.checkbox("Adicionar AGR 2")
        Q30, Q31 = "", 0
        if agregar_agr3:
            c1, c2 = st.columns(2)
            with c1: Q30 = st.text_input("AGR 2")
            with c2: Q31 = st.number_input("Horas mensuales (AGR 2)", min_value=0, step=1, value=0)

        agregar_agr4 = st.checkbox("Adicionar AGR 3")
        Q33, Q34 = "", 0
        if agregar_agr4:
            c1, c2 = st.columns(2)
            with c1: Q33 = st.text_input("AGR 3")
            with c2: Q34 = st.number_input("Horas mensuales (AGR 3)", min_value=0, step=1, value=0)

        # ── 2.2 Auxilios autorizados ──────────────────────────────────────
        st.markdown('<div class="section-header">2.2. AUXILIOS AUTORIZADOS</div>',
                    unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1: Q43 = st.selectbox("Transporte Urbano", ["NO", "SI"])
        with c2: Q44 = st.selectbox("Frecuencia T. Urbano", ["", "QUINCENAL","MENSUAL","ANUAL"])
        with c3: Q45 = st.number_input("Valor T. Urbano ($)", min_value=0, step=10000, value=0)

        c1, c2, c3 = st.columns(3)
        with c1: Q46 = st.selectbox("Transporte Intermunicipal", ["NO", "SI"])
        with c2: Q47 = st.selectbox("Frecuencia T. Intermunicipal", ["", "QUINCENAL","MENSUAL","ANUAL"])
        with c3: Q48 = st.number_input("Valor T. Intermunicipal ($)", min_value=0, step=10000, value=0)

        c1, c2, c3 = st.columns(3)
        with c1: Q49 = st.selectbox("Comunicacion", ["NO", "SI"])
        with c2: Q50 = st.selectbox("Frecuencia Comunicacion", ["", "QUINCENAL","MENSUAL","ANUAL"])
        with c3: Q51 = st.number_input("Valor Comunicacion ($)", min_value=0, step=10000, value=0)

        c1, c2, c3 = st.columns(3)
        with c1: Q52 = st.selectbox("Otro auxilio", ["NO", "SI"])
        with c2: Q53_texto = st.text_input("Cual otro auxilio?")
        with c3: Q54_frec = st.selectbox("Frecuencia otro auxilio", ["", "QUINCENAL","MENSUAL","ANUAL"])
        Q55_valor = st.number_input("Valor otro auxilio ($)", min_value=0, step=10000, value=0)

        # ── III. Competencias ─────────────────────────────────────────────
        st.markdown('<div class="section-header">III. COMPETENCIAS TECNICAS REQUERIDAS</div>',
                    unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            Q56 = st.selectbox("Prueba tecnica especifica segun profesion", ["NO","SI"])
            Q57 = st.selectbox("Prueba ofimatica (minimo Excel intermedio)", ["NO","SI"])
            Q58 = st.selectbox("Prueba tecnica SIG", ["NO","SI"])
        with c2:
            Q53 = st.text_area("Competencias tecnicas generales (describe brevemente)")
            Q54 = st.selectbox("Otra prueba tecnica?", ["NO","SI"])
            Q55 = st.text_input("Describe la prueba solicitada")

        st.markdown("**Competencias comportamentales requeridas:**")
        comp_labels_form = [
            ("Q59","Orientacion al logro"),     ("Q60","Trabajo en equipo"),
            ("Q61","Atencion al cliente"),      ("Q62","Comunicacion efectiva"),
            ("Q63","Adaptabilidad"),            ("Q64","Pensamiento analitico"),
            ("Q65","Innovacion"),               ("Q66","Manejo del conflicto"),
            ("Q67","Negociacion y persuasion"), ("Q68","Desarrollo de relaciones"),
            ("Q69","Liderar equipos"),          ("Q70","Planificacion estrategica"),
        ]
        comp_vals = {}
        for i in range(0, len(comp_labels_form), 3):
            cols = st.columns(3)
            for j, (qkey, qlabel) in enumerate(comp_labels_form[i:i+3]):
                with cols[j]:
                    comp_vals[qkey] = st.selectbox(qlabel, ["NO","SI"], key=f"comp_{qkey}")

        Q70_extra = st.text_input("Descripcion adicional competencias tecnicas especificas")

        # ── IV. EPPs / Dotación / Equipos / Cursos ────────────────────────
        st.markdown('<div class="section-header">IV. DOTACION — EPP\'s — EQUIPO DE COMPUTO — CURSOS ADICIONALES</div>',
                    unsafe_allow_html=True)

        st.markdown("**4.1. EPP's requeridos:**")
        epp_items = [
            ("Q71","Casco dielectrico"), ("Q72","Casco con barbuquejo"),
            ("Q73","Protector auditivo de copa"), ("Q74","Protector auditivo de insercion"),
            ("Q75","Monogafa antiempanante"), ("Q76","Proteccion respiratoria"),
            ("Q77","Proteccion visual (lente claro/oscuro)"),
        ]
        epp_vals = {}
        for i in range(0, len(epp_items), 3):
            cols = st.columns(3)
            for j, (qk, ql) in enumerate(epp_items[i:i+3]):
                with cols[j]:
                    epp_vals[qk] = st.selectbox(ql, ["NO","SI"], key=f"epp_{qk}")

        st.markdown("**4.2. Dotacion:**")
        dot_items = [
            ("Q78","Uniforme anti fluido"), ("Q79","Chaqueta"),
            ("Q80","Camisa"), ("Q81","Jean"),
            ("Q82","Botas dielectr./antideslizante"), ("Q83","Botas dielectr./antiperforante"),
            ("Q84","Otro elemento dotacion"),
        ]
        dot_vals = {}
        for i in range(0, len(dot_items), 3):
            cols = st.columns(3)
            for j, (qk, ql) in enumerate(dot_items[i:i+3]):
                with cols[j]:
                    dot_vals[qk] = st.selectbox(ql, ["NO","SI"], key=f"dot_{qk}")

        Q85 = st.text_input("Cual otro elemento de dotacion/EPP?")

        st.markdown("**4.3. Equipo de computo:**")
        c1, c2 = st.columns(2)
        with c1: Q86 = st.selectbox("Equipo de computo basico", ["NO","SI"])
        with c2: Q87 = st.selectbox("Equipo de computo mayor capacidad", ["NO","SI"])

        st.markdown("**4.4. Cursos especiales:**")
        c1, c2 = st.columns(2)
        with c1: Q88 = st.selectbox("Curso trabajo seguro en alturas", ["NO","SI"])
        with c2: Q89 = st.selectbox("Curso espacios confinados", ["NO","SI"])

        # ── V. Recomendaciones ────────────────────────────────────────────
        st.markdown('<div class="section-header">V. RECOMENDACIONES PARA TENER EN CUENTA DURANTE EL PROCESO</div>',
                    unsafe_allow_html=True)
        Q90 = st.text_area("Recomendaciones generales")

        st.divider()
        submitted = st.form_submit_button(
            "Enviar Solicitud",
            use_container_width=True,
            type="primary"
        )

    # ─────────────────────────────────────────────────────────────────────────
    # 9. PROCESAMIENTO AL ENVIAR
    # ─────────────────────────────────────────────────────────────────────────
    if submitted:
        errores = []
        if not Q8.strip():  errores.append("Razon social empresa es obligatorio.")
        if not Q10.strip(): errores.append("Nombre del AGR es obligatorio.")
        if not Q11.strip(): errores.append("Correo del AGR es obligatorio.")
        if not Q36:         errores.append("Debe seleccionar una ciudad.")
        if not Q37_dias:    errores.append("Debe seleccionar al menos un dia de servicio.")

        if errores:
            for e in errores:
                st.error(e)
            st.stop()

        with st.spinner("Procesando solicitud..."):

            # Generar ID numérico
            id_sol = generar_id_solicitud()
            fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")

            # Consolidar datos con clave Q→valor
            datos = {
                "id_solicitud": id_sol,
                "Q2":  fecha_hoy,
                "Q6":  Q6,  "Q7":  Q7,  "Q8":  Q8,  "Q9":  Q9,
                "Q10": Q10, "Q11": Q11, "Q12": Q12, "Q13": Q13, "Q14": Q14,
                "Q15": Q15, "Q16": Q16, "Q17": Q17, "Q18": Q18, "Q19": Q19,
                "Q20": Q20, "Q21": Q21, "Q22": Q22, "Q23": Q23,
                "Q24": Q24, "Q25": Q25,
                "Q27": Q27, "Q28": Q28,
                "Q30": Q30, "Q31": Q31,
                "Q33": Q33, "Q34": Q34,
                "Q35": Q35, "Q36": Q36,
                "Q37": "; ".join(Q37_dias),
                "Q38": Q38, "Q39": Q39, "Q40": Q40, "Q41": Q41, "Q42": Q42,
                "Q43": Q43, "Q44": Q44, "Q45": Q45,
                "Q46": Q46, "Q47": Q47, "Q48": Q48,
                "Q49": Q49, "Q50": Q50, "Q51": Q51,
                "Q52": Q52, "Q53": Q53_texto, "Q54": Q54_frec, "Q55": str(Q55_valor),
                "Q56": Q56, "Q57": Q57, "Q58": Q58,
                **comp_vals,
                "Q70": Q70_extra,
                **epp_vals,
                **dot_vals,
                "Q85": Q85,
                "Q86": Q86, "Q87": Q87,
                "Q88": Q88, "Q89": Q89,
                "Q90": Q90,
                "Q91": Q91,
            }

            logo_bytes = obtener_logo()

            plantilla_path = os.path.join(os.path.dirname(__file__), "FORMATO.xlsx")
            try:
                with open(plantilla_path, "rb") as f:
                    plantilla_bytes = f.read()
            except FileNotFoundError:
                st.error("No se encontro FORMATO.xlsx junto al script.")
                st.stop()

            xlsx_bytes = diligenciar_formato_excel(datos, plantilla_bytes, logo_bytes)
            pdf_bytes  = generar_pdf(datos, logo_bytes)

            # ── Guardar en Supabase ──────────────────────────────────────
            guardado_ok = guardar_solicitud_supabase(datos, id_sol)
            if not guardado_ok:
                st.warning("No se pudo guardar en la base de datos. El correo se enviara de todas formas.")

            asesora, _, _ = obtener_asesora_y_clasificacion(Q36)
            nombre_base   = f"{id_sol} - {Q15} en {Q36}"
            carpeta       = asesora

            subir_archivo_supabase(
                xlsx_bytes,
                f"{carpeta}/{nombre_base}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            subir_archivo_supabase(
                pdf_bytes,
                f"{carpeta}/{nombre_base}.pdf",
                "application/pdf"
            )

            exito_correo = enviar_correo(datos, pdf_bytes, xlsx_bytes)

        if exito_correo:
            st.success(f"Solicitud **{id_sol}** enviada exitosamente a **{asesora}**.")
        else:
            st.warning(f"Solicitud registrada como **{id_sol}** pero hubo un problema al enviar el correo. Contacta al administrador.")

        asesora_info, clasificacion_info, dias_info = obtener_asesora_y_clasificacion(Q36)
        dias_totales = dias_info + 2 if Q17 == "MEDICO" else dias_info
        fecha_entrega = calcular_fecha_entrega(dias_totales)

        st.markdown(f"""
        <div class="info-box">
        Ciudad: <b>{Q36}</b> ({clasificacion_info}) &nbsp;|&nbsp;
        Tiempo de respuesta: <b>{dias_totales} dias habiles</b> &nbsp;|&nbsp;
        Fecha tentativa de entrega: <b>{fecha_entrega}</b>
        </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "Descargar PDF",
                data=pdf_bytes,
                file_name=f"{nombre_base}.pdf",
                mime="application/pdf"
            )
        with col2:
            st.download_button(
                "Descargar Excel",
                data=xlsx_bytes,
                file_name=f"{nombre_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


if __name__ == "__main__":
    main()
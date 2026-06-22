import base64
import os
import re
import socket
import zipfile
from datetime import date, datetime
from html import escape
from io import BytesIO

import streamlit as st
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


OUTPUT_DIR = "INFORMES_GENERADOS"
UPLOAD_DIR = "IMAGENES_SUBIDAS"
LOGO_PATH = "LOGOO.png"
COUNTER_FILE = "ultimo_informe_autogood.txt"
HISTORY_FILE = "historial_autogood.xlsx"

PAGE_W, PAGE_H = A4
MARGIN_X = 28
CONTENT_START_Y = 675
TABLE_BOTTOM_Y = 70

BLUE = colors.HexColor("#082440")
DARK_BLUE = colors.HexColor("#06172B")
LIGHT_BLUE = colors.HexColor("#D9EEF8")
GREY = colors.HexColor("#E8ECEF")
LIGHT_GREY = colors.HexColor("#F7F9FA")
BORDER = colors.HexColor("#8D9AA3")
TEXT = colors.HexColor("#20252A")
GREEN = colors.HexColor("#21A366")
RED = colors.HexColor("#C0392B")
AMBER = colors.HexColor("#F39C12")
MUTED = colors.HexColor("#64717D")

AUTHORIZED_USERS = {
    "tecnico1": "autogood2026",
    "tecnico2": "autogood2026",
    "asesor": "autogood2026",
    "admin": "autogoodadmin2026",
}

STATE_OPTIONS = [
    "Buen estado",
    "Mal estado",
    "Seguimiento",
    "Se cambió",
    "No aplica",
]

HISTORY_HEADERS = [
    "Código de informe",
    "Fecha",
    "Cliente",
    "DNI/RUC",
    "Teléfono",
    "Placa",
    "Marca",
    "Modelo",
    "Año",
    "Kilometraje",
    "Técnico",
    "Informe realizado por",
    "Observaciones del cliente",
    "Observaciones del técnico",
    "Resumen ejecutivo",
    "Usuario",
    "PDF generado",
]

OBSERVATION_LIBRARY = {
    "Bujes de trapecio": ["Buen estado", "Resecos", "Rajados", "Requiere cambio"],
    "Amortiguadores": ["Buen estado", "Fuga hidráulica", "Rendido", "Requiere cambio"],
    "Revisión de amortiguadores": ["Buen estado", "Fuga hidráulica", "Rendido", "Requiere cambio"],
    "Pastillas delanteras": ["Buen estado", "Desgaste 50%", "Desgaste 80%", "Requiere cambio inmediato"],
    "Pastillas posteriores": ["Buen estado", "Desgaste 50%", "Desgaste 80%", "Requiere cambio inmediato"],
    "Bieletas": ["Buen estado", "Juego leve", "Juego excesivo", "Requiere cambio"],
    "Terminales": ["Buen estado", "Juego axial", "Guardapolvo roto", "Requiere cambio"],
}

MOTOR_COMPONENTS = [
    "Bujías de cobre",
    "Bujías de platino",
    "Bujías de iridio",
    "Kit de distribución",
    "Kit de accesorios",
    "Bobinas de encendido",
    "Batería",
    "Tapón de cárter",
    "Cambio de filtro de aire de motor",
    "Cambio de filtro de aceite",
    "Caña de bobinas",
    "Conector de bobinas",
    "Refrigerante",
    "Shampoo de parabrisas",
    "Arandela de tapón de cárter",
]

SUSPENSION_FRONT = [
    "Rótula superior",
    "Rótula inferior",
    "Bujes de trapecio",
    "Resortes",
    "Amortiguadores",
    "Jebes de barra estabilizadora",
    "Guardapolvos de amortiguador",
]

SUSPENSION_REAR = [
    "Revisión de bujes de puente posterior",
    "Revisión de bujes secundarios",
    "Revisión de resortes",
    "Revisión de amortiguadores",
    "Revisión de jebes de barra estabilizadora",
    "Revisión de guardapolvos de amortiguador",
]

BRAKES = [
    "Pastillas delanteras",
    "Pastillas posteriores",
    "Zapatas",
    "Rectificado de discos delanteros",
    "Rectificado de discos posteriores",
    "Rectificado de tambores",
    "Calipers delanteros",
    "Calipers posteriores",
    "Bombines de freno",
    "Regulación freno de mano",
    "Líquido de frenos",
    "Neumáticos delanteros",
    "Neumáticos posteriores",
]

STEERING = [
    "Bieletas",
    "Terminales",
    "Guardapolvo de cremallera",
    "Fuga de hidrolina",
]

TRANSMISSION = [
    "Aceite de caja / grado",
    "Aceite de corona / grado",
    "Crucetas",
    "Palieres",
    "Guardapolvos de palier",
    "Cambio de retén / requiere",
    "Fuga de aceite de corona",
    "Fuga de aceite de caja",
]

ELECTRONIC_TUNE = [
    "Inyectores",
    "O'rings de inyectores",
    "Sensores MAP, MAF, IAC, EGR, TPS",
    "Cuerpo de aceleración",
    "Prueba de inyectores",
]


def ensure_dirs():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)


def safe_filename(value):
    value = (value or "AUTOGOOD").strip().upper()
    value = re.sub(r"[^A-Z0-9_-]+", "_", value)
    return value.strip("_") or "AUTOGOOD"


def component_key(component):
    key = component.lower()
    replacements = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
        "'": "",
        "/": "",
    }
    for old, new in replacements.items():
        key = key.replace(old, new)
    return re.sub(r"[^a-z0-9]+", "_", key).strip("_")


def get_next_report_code():
    current_year = datetime.now().year
    last_year = current_year
    last_number = 0
    if os.path.exists(COUNTER_FILE):
        raw = open(COUNTER_FILE, "r", encoding="utf-8").read().strip()
        match = re.match(r"^(\d{4}),(\d+)$", raw)
        if match:
            last_year = int(match.group(1))
            last_number = int(match.group(2))
    next_number = last_number + 1 if last_year == current_year else 1
    with open(COUNTER_FILE, "w", encoding="utf-8") as counter:
        counter.write(f"{current_year},{next_number}")
    return f"AG-{current_year}-{next_number:04d}"


def automatic_recommendation(state, observation=""):
    if state == "Buen estado":
        return "No requiere intervención por el momento."
    if state == "Seguimiento":
        return "Se recomienda realizar seguimiento en el próximo mantenimiento."
    if state == "Mal estado":
        return "Se recomienda reemplazo para evitar daños mayores."
    if state == "Se cambió":
        return "Componente reemplazado durante el servicio."

    observation = observation or ""
    if observation in ["Resecos", "Desgaste 50%", "Juego leve", "Rendido"]:
        return "Se recomienda realizar seguimiento en el próximo mantenimiento."
    if observation in [
        "Rajados",
        "Fuga hidráulica",
        "Desgaste 80%",
        "Juego excesivo",
        "Juego axial",
        "Guardapolvo roto",
        "Requiere cambio",
        "Requiere cambio inmediato",
    ]:
        return "Se recomienda reemplazo para evitar daños mayores."
    if observation == "Buen estado":
        return "No requiere intervención por el momento."
    return ""


def prepare_uploaded_photos(files, label):
    photos = []
    for index, uploaded in enumerate(files or [], start=1):
        try:
            image = Image.open(BytesIO(uploaded.getvalue())).convert("RGB")
            buf = BytesIO()
            image.save(buf, format="JPEG", quality=88)
            photos.append(
                {
                    "label": f"{label} {index}",
                    "name": uploaded.name,
                    "bytes": buf.getvalue(),
                    "width": image.width,
                    "height": image.height,
                }
            )
        except Exception:
            continue
    return photos


def xlsx_col_name(index):
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def sheet_xml(rows):
    xml_rows = []
    for row_idx, row in enumerate(rows, start=1):
        cells = []
        for col_idx, value in enumerate(row, start=1):
            ref = f"{xlsx_col_name(col_idx)}{row_idx}"
            value = "" if value is None else str(value)
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
        xml_rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        + "".join(xml_rows)
        + "</sheetData></worksheet>"
    )


def write_history_xlsx(rows):
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Historial" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    with zipfile.ZipFile(HISTORY_FILE, "w", zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types_xml)
        xlsx.writestr("_rels/.rels", rels_xml)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml(rows))


def read_history_rows():
    if not os.path.exists(HISTORY_FILE):
        return [HISTORY_HEADERS]
    try:
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(HISTORY_FILE, "r") as xlsx:
            content = xlsx.read("xl/worksheets/sheet1.xml")
        root = ET.fromstring(content)
        ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rows = []
        for row in root.findall(".//x:row", ns):
            values = []
            for cell in row.findall("x:c", ns):
                text_node = cell.find(".//x:t", ns)
                values.append(text_node.text if text_node is not None and text_node.text is not None else "")
            rows.append(values)
        return rows or [HISTORY_HEADERS]
    except Exception:
        return [HISTORY_HEADERS]


def normalize_history_rows(rows):
    if not rows:
        return [HISTORY_HEADERS]
    if rows[0] == HISTORY_HEADERS:
        return rows

    old_headers = [header for header in HISTORY_HEADERS if header != "Usuario"]
    if rows[0] == old_headers:
        normalized = [HISTORY_HEADERS]
        for row in rows[1:]:
            padded = row + [""] * (len(old_headers) - len(row))
            old_record = dict(zip(old_headers, padded))
            normalized.append(
                [
                    old_record.get("Código de informe", ""),
                    old_record.get("Fecha", ""),
                    old_record.get("Cliente", ""),
                    old_record.get("DNI/RUC", ""),
                    old_record.get("Teléfono", ""),
                    old_record.get("Placa", ""),
                    old_record.get("Marca", ""),
                    old_record.get("Modelo", ""),
                    old_record.get("Año", ""),
                    old_record.get("Kilometraje", ""),
                    old_record.get("Técnico", ""),
                    old_record.get("Informe realizado por", ""),
                    old_record.get("Observaciones del cliente", ""),
                    old_record.get("Observaciones del técnico", ""),
                    old_record.get("Resumen ejecutivo", ""),
                    "",
                    old_record.get("PDF generado", ""),
                ]
            )
        return normalized
    return [HISTORY_HEADERS] + rows


def append_history(data, pdf_path):
    required = [data.get("cliente", ""), data.get("placa", ""), data.get("telefono", "")]
    if any(not str(value).strip() for value in required):
        return False

    rows = normalize_history_rows(read_history_rows())
    rows.append(
        [
            data.get("codigo_informe", ""),
            data.get("fecha", ""),
            data.get("cliente", ""),
            data.get("dni_ruc", ""),
            data.get("telefono", ""),
            data.get("placa", ""),
            data.get("marca", ""),
            data.get("modelo", ""),
            data.get("anio", ""),
            data.get("km_actual", ""),
            data.get("tecnico_responsable", ""),
            data.get("informe_realizado_por", ""),
            data.get("obs_cliente", ""),
            data.get("obs_tecnico", ""),
            data.get("resumen", ""),
            data.get("usuario_generador", ""),
            pdf_path,
        ]
    )
    write_history_xlsx(rows)
    return True


def search_history(query):
    query = (query or "").strip().lower()
    rows = normalize_history_rows(read_history_rows())
    if len(rows) <= 1:
        return []
    records = []
    for row in rows[1:]:
        padded = row + [""] * (len(HISTORY_HEADERS) - len(row))
        record = dict(zip(HISTORY_HEADERS, padded))
        haystack = " ".join(
            [
                record.get("Placa", ""),
                record.get("Cliente", ""),
                record.get("Teléfono", ""),
                record.get("Código de informe", ""),
            ]
        ).lower()
        if query and query in haystack:
            records.append(record)
    return list(reversed(records))


def pdf_link(path):
    if not path or not os.path.exists(path):
        return ""
    with open(path, "rb") as pdf_file:
        encoded = base64.b64encode(pdf_file.read()).decode("ascii")
    return f"data:application/pdf;base64,{encoded}"


def get_local_ip():
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "IP_DEL_EQUIPO"


def apply_mobile_styles():
    st.markdown(
        """
        <style>
        div.stButton > button,
        div.stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] button {
            min-height: 3.25rem;
            font-size: 1.05rem;
            font-weight: 700;
            border-radius: 8px;
        }

        div[data-baseweb="input"] input,
        textarea,
        div[data-baseweb="select"] {
            min-height: 2.75rem;
            font-size: 1rem;
        }

        div[data-testid="stExpander"] summary {
            min-height: 3rem;
            font-size: 1.05rem;
            font-weight: 700;
        }

        @media (max-width: 760px) {
            .block-container {
                padding-left: 0.85rem;
                padding-right: 0.85rem;
                padding-top: 1rem;
            }

            h1 {
                font-size: 1.55rem;
                line-height: 1.2;
            }

            div[data-testid="column"] {
                min-width: 100% !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def require_login():
    if st.session_state.get("authenticated"):
        user = st.session_state.get("active_user", "")
        c1, c2 = st.columns([0.72, 0.28])
        with c1:
            st.success(f"Sesión iniciada: {user}")
        with c2:
            if st.button("Cerrar sesión", use_container_width=True):
                st.session_state["authenticated"] = False
                st.session_state["active_user"] = ""
                st.rerun()
        return True

    st.subheader("Acceso privado AUTOGOOD")
    st.caption("Ingrese sus credenciales para continuar.")
    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        login = st.form_submit_button("Iniciar sesión", use_container_width=True)

    if login:
        if AUTHORIZED_USERS.get(username.strip()) == password:
            st.session_state["authenticated"] = True
            st.session_state["active_user"] = username.strip()
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")
    return False


def fit_text(c, text, max_width, font="Helvetica", size=8):
    text = str(text or "").strip()
    if not text:
        return ""
    if stringWidth(text, font, size) <= max_width:
        return text
    ellipsis = "..."
    while text and stringWidth(text + ellipsis, font, size) > max_width:
        text = text[:-1]
    return text + ellipsis if text else ellipsis


def wrap_text(c, text, max_width, font="Helvetica", size=8, max_lines=None):
    words = str(text or "").replace("\r", "").split()
    lines = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if stringWidth(candidate, font, size) <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
        if max_lines and len(lines) >= max_lines:
            break
    if current and (not max_lines or len(lines) < max_lines):
        lines.append(current)
    return lines[:max_lines] if max_lines else lines


def state_color(state):
    return {
        "Buen estado": GREEN,
        "Se cambió": GREEN,
        "Seguimiento": AMBER,
        "Mal estado": RED,
        "No aplica": colors.HexColor("#A7ADB3"),
    }.get(state, MUTED)


def header_logo_asset():
    image = Image.open(LOGO_PATH).convert("RGBA")
    bg = image.getpixel((0, 0))[:3]
    pixels = image.load()
    width, height = image.size
    min_x, min_y = width, height
    max_x, max_y = 0, 0

    for y in range(height):
        for x in range(width):
            r, g, b, a = pixels[x, y]
            is_logo_color = r > 140 and g > 90 and b < 80
            if a and is_logo_color:
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

    if max_x <= min_x or max_y <= min_y:
        cropped = image
    else:
        pad = 12
        cropped = image.crop(
            (
                max(0, min_x - pad),
                max(0, min_y - pad),
                min(width, max_x + pad),
                min(height, max_y + pad),
            )
        )

    buffer = BytesIO()
    cropped.save(buffer, format="PNG")
    buffer.seek(0)
    return ImageReader(buffer), colors.Color(bg[0] / 255, bg[1] / 255, bg[2] / 255)


def draw_header(c, page_num, total_pages=4, title="INFORME TÉCNICO VEHICULAR", report_code="", order_number=""):
    c.setFillColor(colors.white)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    header_down = 16
    logo_x = MARGIN_X + 10
    logo_y = PAGE_H - 92 - header_down
    logo_w = 296
    logo_h = 80
    logo_fill = DARK_BLUE
    logo_reader = None
    if os.path.exists(LOGO_PATH):
        logo_reader, logo_fill = header_logo_asset()

    c.setFillColor(logo_fill)
    c.rect(logo_x, logo_y, logo_w, logo_h, fill=1, stroke=0)
    if logo_reader:
        c.drawImage(
            logo_reader,
            logo_x + 12,
            logo_y + 13,
            width=logo_w - 24,
            height=logo_h - 26,
            preserveAspectRatio=True,
            mask="auto",
        )
    else:
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(logo_x + logo_w / 2, logo_y + 32, "AUTOGOOD")

    c.setFillColor(TEXT)
    c.setFont("Helvetica", 9.5)
    contact_x = PAGE_W - 220
    contact_y = logo_y + 51
    c.drawString(contact_x, contact_y, "Av. Ferrocarril 2825")
    c.drawString(contact_x, contact_y - 15, "El Tambo")
    c.drawString(contact_x, contact_y - 34, "915 176 566")

    line_y = PAGE_H - 101 - header_down
    title_y = PAGE_H - 132 - header_down
    c.setFillColor(BLUE)
    c.rect(MARGIN_X, line_y, PAGE_W - MARGIN_X * 2, 7, fill=1, stroke=0)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(PAGE_W / 2, title_y, title.replace("TÉCNICO VEHICULAR", "DE SERVICIO"))
    c.setLineWidth(1.2)
    c.line(PAGE_W / 2 - 118, title_y - 5, PAGE_W / 2 + 118, title_y - 5)
    c.setFont("Helvetica", 7)
    c.setFillColor(MUTED)
    c.drawRightString(PAGE_W - MARGIN_X, title_y, f"Página {page_num} de {total_pages}")
    c.setFillColor(TEXT)
    c.setFont("Helvetica-Bold", 8)
    if report_code:
        c.drawString(MARGIN_X + 10, title_y, f"Código: {report_code}")
    if order_number:
        c.drawString(MARGIN_X + 10, title_y - 14, f"N° orden: {order_number}")


def draw_footer(c):
    c.setFillColor(BLUE)
    c.rect(MARGIN_X + 40, 38, PAGE_W - MARGIN_X * 2 - 80, 3, fill=1, stroke=0)
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 8)
    c.drawCentredString(PAGE_W / 2, 23, "Centro de diagnóstico Automotriz")


def section_title(c, x, y, w, title, fill=LIGHT_BLUE):
    c.setFillColor(fill)
    c.setStrokeColor(colors.black)
    c.rect(x, y - 18, w, 18, fill=1, stroke=1)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 8.2)
    c.drawString(x + 7, y - 12, title.upper())
    return y - 18


def draw_info_grid(c, x, y, w, rows, row_h=19):
    label_w = w * 0.34
    total_h = row_h * len(rows)
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.55)
    for i, (label, value) in enumerate(rows):
        yy = y - row_h * (i + 1)
        c.setFillColor(GREY)
        c.rect(x, yy, label_w, row_h, fill=1, stroke=1)
        c.setFillColor(colors.white)
        c.rect(x + label_w, yy, w - label_w, row_h, fill=1, stroke=1)
        c.setFillColor(TEXT)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(x + 5, yy + 6, fit_text(c, label, label_w - 9, "Helvetica-Bold", 7))
        c.setFont("Helvetica", 7)
        c.drawString(x + label_w + 5, yy + 6, fit_text(c, value, w - label_w - 10, "Helvetica", 7))
    return y - total_h


def draw_text_box(c, x, y, w, h, title, text, fill=LIGHT_BLUE, max_lines=5):
    section_title(c, x, y, w, title, fill)
    box_y = y - 18 - h
    c.setFillColor(colors.white)
    c.setStrokeColor(BORDER)
    c.rect(x, box_y, w, h, fill=1, stroke=1)
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 7.5)
    yy = box_y + h - 12
    for line in wrap_text(c, text, w - 14, "Helvetica", 7.5, max_lines=max_lines):
        c.drawString(x + 7, yy, line)
        yy -= 10
    return box_y


def draw_component_table(c, x, y, w, title, components, data, row_h=18, title_fill=LIGHT_BLUE):
    section_title(c, x, y, w, title, title_fill)
    y -= 18
    half = w / 2
    comp_w = half * 0.50
    good_w = half * 0.18
    bad_w = half * 0.18
    obs_w = half - comp_w - good_w - bad_w

    def draw_half_header(xx):
        c.setFillColor(colors.HexColor("#07A9D8"))
        c.setStrokeColor(colors.black)
        c.rect(xx, y - row_h, half, row_h, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 6.2)
        c.drawCentredString(xx + comp_w / 2, y - 12, "MOTOR" if "MOTOR" in title.upper() else "COMPONENTE")
        c.drawCentredString(xx + comp_w + good_w / 2, y - 8, "BUEN")
        c.drawCentredString(xx + comp_w + good_w / 2, y - 15, "ESTAD.")
        c.drawCentredString(xx + comp_w + good_w + bad_w / 2, y - 8, "MAL")
        c.drawCentredString(xx + comp_w + good_w + bad_w / 2, y - 15, "ESTADO")
        c.drawCentredString(xx + comp_w + good_w + bad_w + obs_w / 2, y - 12, "OBS")

    draw_half_header(x)
    draw_half_header(x + half)
    y -= row_h
    rows = (len(components) + 1) // 2

    def draw_component_cell(component, xx, yy):
        if not component:
            c.setFillColor(colors.white)
            c.rect(xx, yy - row_h, half, row_h, fill=1, stroke=1)
            return
        key = component_key(component)
        state = data.get(f"{key}_estado", "No aplica")
        obs = data.get(f"{key}_obs", "")
        c.setStrokeColor(colors.black)
        c.setFillColor(colors.white)
        c.rect(xx, yy - row_h, half, row_h, fill=1, stroke=1)
        c.line(xx + comp_w, yy, xx + comp_w, yy - row_h)
        c.line(xx + comp_w + good_w, yy, xx + comp_w + good_w, yy - row_h)
        c.line(xx + comp_w + good_w + bad_w, yy, xx + comp_w + good_w + bad_w, yy - row_h)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 6.1)
        lines = wrap_text(c, component.upper(), comp_w - 8, "Helvetica-Bold", 6.1, max_lines=2)
        text_y = yy - 8 if len(lines) == 1 else yy - 6
        for line in lines:
            c.drawString(xx + 4, text_y, line)
            text_y -= 7

        if state == "Buen estado":
            c.setFillColor(GREEN)
            c.rect(xx + comp_w, yy - row_h, good_w, row_h, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 6.5)
            c.drawCentredString(xx + comp_w + good_w / 2, yy - 12, "OK")
        elif state == "Mal estado":
            c.setFillColor(RED)
            c.rect(xx + comp_w + good_w, yy - row_h, bad_w, row_h, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 6.5)
            c.drawCentredString(xx + comp_w + good_w + bad_w / 2, yy - 12, "X")
        elif state == "Se cambió":
            c.setFillColor(GREEN)
            c.rect(xx + comp_w + good_w + bad_w, yy - row_h, obs_w, row_h, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 5.7)
            c.drawCentredString(xx + comp_w + good_w + bad_w + obs_w / 2, yy - 12, "SE CAMBIÓ")
        elif state == "Seguimiento":
            c.setFillColor(AMBER)
            c.rect(xx + comp_w + good_w + bad_w, yy - row_h, obs_w, row_h, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 5.7)
            c.drawCentredString(xx + comp_w + good_w + bad_w + obs_w / 2, yy - 12, "SEGUIM.")
        elif state == "No aplica":
            c.setFillColor(GREY)
            c.rect(xx + comp_w + good_w + bad_w, yy - row_h, obs_w, row_h, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 5.7)
            c.drawCentredString(xx + comp_w + good_w + bad_w + obs_w / 2, yy - 12, "N/A")

        if obs and state not in ["Se cambió", "Seguimiento", "No aplica"]:
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 5.4)
            c.drawCentredString(
                xx + comp_w + good_w + bad_w + obs_w / 2,
                yy - 12,
                fit_text(c, obs.upper(), obs_w - 4, "Helvetica", 5.4),
            )

    for row in range(rows):
        left = components[row]
        right_idx = row + rows
        right = components[right_idx] if right_idx < len(components) else None
        draw_component_cell(left, x, y)
        draw_component_cell(right, x + half, y)
        y -= row_h
    return y


def component_table_height(components, row_h):
    rows = (len(components) + 1) // 2
    return 18 + row_h + rows * row_h


def table_fits_on_page(y, components, row_h):
    return y - component_table_height(components, row_h) >= TABLE_BOTTOM_Y


def estimate_base_pages():
    page_num = 1
    if not table_fits_on_page(280 - 76, MOTOR_COMPONENTS, 14):
        page_num += 1
    page_num += 1
    for y, components, row_h in [
        (CONTENT_START_Y, SUSPENSION_FRONT, 20),
        (535, SUSPENSION_REAR, 20),
        (410, BRAKES, 16),
    ]:
        if not table_fits_on_page(y, components, row_h):
            page_num += 1
    page_num += 1
    for y, components, row_h in [
        (CONTENT_START_Y, STEERING, 22),
        (565, TRANSMISSION, 20),
        (435, ELECTRONIC_TUNE, 22),
    ]:
        if not table_fits_on_page(y, components, row_h):
            page_num += 1
    page_num += 1
    return page_num


def draw_component_table_guarded(c, page_num, total_pages, data, x, y, w, title, components, row_h=18, title_fill=LIGHT_BLUE):
    if not table_fits_on_page(y, components, row_h):
        draw_footer(c)
        c.showPage()
        page_num += 1
        draw_header(
            c,
            page_num,
            total_pages=total_pages,
            report_code=data.get("codigo_informe", ""),
            order_number=data.get("numero_orden", ""),
        )
        y = CONTENT_START_Y
    bottom = draw_component_table(c, x, y, w, title, components, data, row_h=row_h, title_fill=title_fill)
    return bottom, page_num


def draw_legend(c, x, y, w):
    section_title(c, x, y, w, "LEYENDA DE ESTADOS", GREY)
    y -= 35
    item_w = w / 5
    for i, state in enumerate(STATE_OPTIONS):
        xx = x + i * item_w
        swatch_x = xx + 8
        label_x = swatch_x + 20
        c.setFillColor(state_color(state))
        c.roundRect(swatch_x, y, 14, 10, 3, fill=1, stroke=0)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 6.6)
        c.drawString(label_x, y + 2, fit_text(c, state, item_w - 32, "Helvetica", 6.6))
    return y - 12


def draw_signature_box(c, x, y, w, h, title, value):
    c.setFillColor(colors.white)
    c.setStrokeColor(BORDER)
    c.rect(x, y - h, w, h, fill=1, stroke=1)
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 7)
    c.drawCentredString(x + w / 2, y - h + 10, title)
    c.line(x + 18, y - h + 28, x + w - 18, y - h + 28)
    c.setFont("Helvetica-Bold", 7)
    c.drawCentredString(x + w / 2, y - h + 34, fit_text(c, value, w - 24, "Helvetica-Bold", 7))


def draw_image_preserved(c, image_bytes, x, y, w, h):
    image = Image.open(BytesIO(image_bytes))
    iw, ih = image.size
    scale = min(w / iw, h / ih)
    draw_w = iw * scale
    draw_h = ih * scale
    draw_x = x + (w - draw_w) / 2
    draw_y = y + (h - draw_h) / 2
    c.drawImage(ImageReader(BytesIO(image_bytes)), draw_x, draw_y, width=draw_w, height=draw_h, mask="auto")


def draw_photo_pages(c, photos, start_page, total_pages, data):
    if not photos:
        return
    for page_offset, chunk_start in enumerate(range(0, len(photos), 4)):
        page_num = start_page + page_offset
        draw_header(
            c,
            page_num,
            total_pages=total_pages,
            report_code=data.get("codigo_informe", ""),
            order_number=data.get("numero_orden", ""),
        )
        y = CONTENT_START_Y
        section_title(c, MARGIN_X, y, PAGE_W - MARGIN_X * 2, "EVIDENCIA FOTOGRÁFICA")
        grid_top = y - 42
        gap = 18
        cell_w = (PAGE_W - MARGIN_X * 2 - gap) / 2
        cell_h = 210
        chunk = photos[chunk_start : chunk_start + 4]
        for index, photo in enumerate(chunk):
            col = index % 2
            row = index // 2
            x = MARGIN_X + col * (cell_w + gap)
            cell_top = grid_top - row * (cell_h + 34)
            c.setFillColor(colors.white)
            c.setStrokeColor(BORDER)
            c.rect(x, cell_top - cell_h, cell_w, cell_h, fill=1, stroke=1)
            draw_image_preserved(c, photo["bytes"], x + 7, cell_top - cell_h + 20, cell_w - 14, cell_h - 32)
            c.setFillColor(TEXT)
            c.setFont("Helvetica-Bold", 7)
            c.drawString(x + 7, cell_top - cell_h + 8, fit_text(c, photo["label"], cell_w - 14, "Helvetica-Bold", 7))
            c.setFont("Helvetica", 6)
            c.setFillColor(MUTED)
            c.drawRightString(x + cell_w - 7, cell_top - cell_h + 8, fit_text(c, photo["name"], cell_w / 2, "Helvetica", 6))
        draw_footer(c)
        if chunk_start + 4 < len(photos):
            c.showPage()


def build_recommendation_summary(data, limit=5):
    components = MOTOR_COMPONENTS + SUSPENSION_FRONT + SUSPENSION_REAR + BRAKES + STEERING + TRANSMISSION + ELECTRONIC_TUNE
    lines = []
    for component in components:
        key = component_key(component)
        state = data.get(f"{key}_estado", "No aplica")
        obs = data.get(f"{key}_obs", "")
        rec = data.get(f"{key}_rec") or automatic_recommendation(state, obs)
        if state in ["Mal estado", "Seguimiento", "Se cambió"] or (obs and obs != "Buen estado"):
            lines.append(f"{component}: {obs or state}. {rec}")
        if len(lines) >= limit:
            break
    return " ".join(lines) if lines else "No se registran recomendaciones críticas adicionales."


def generate_pdf(data):
    ensure_dirs()
    placa = safe_filename(data.get("placa"))
    report_code = safe_filename(data.get("codigo_informe"))
    photos = data.get("photos", [])
    photo_pages = (len(photos) + 3) // 4 if photos else 0
    base_pages = estimate_base_pages()
    total_pages = base_pages + photo_pages
    output_path = os.path.join(OUTPUT_DIR, f"Informe_{report_code}_{placa}.pdf")
    c = canvas.Canvas(output_path, pagesize=A4)
    page_num = 1

    draw_header(c, page_num, total_pages=total_pages, report_code=data.get("codigo_informe", ""), order_number=data.get("numero_orden", ""))
    y = CONTENT_START_Y
    client_rows = [
        ("Nombre y apellido", data.get("cliente")),
        ("Distrito", data.get("distrito")),
        ("Teléfono", data.get("telefono")),
        ("Fecha", data.get("fecha")),
        ("DNI/RUC", data.get("dni_ruc")),
        ("Email", data.get("email")),
    ]
    vehicle_rows = [
        ("Placa", data.get("placa")),
        ("Color", data.get("color")),
        ("Marca", data.get("marca")),
        ("Km actual", data.get("km_actual")),
        ("Modelo", data.get("modelo")),
        ("Km próximo mant.", data.get("km_proximo")),
        ("Año", data.get("anio")),
        ("Combustible", data.get("combustible")),
        ("Mantenimiento", data.get("tipo_mantenimiento")),
        ("GDI", data.get("gdi")),
        ("Tipo vehículo", data.get("tipo_vehiculo")),
    ]
    left_w = 260
    right_w = PAGE_W - MARGIN_X * 2 - left_w - 12
    section_title(c, MARGIN_X, y, left_w, "1. INFORMACIÓN DEL CLIENTE")
    draw_info_grid(c, MARGIN_X, y - 18, left_w, client_rows, row_h=17)
    section_title(c, MARGIN_X + left_w + 12, y, right_w, "2. DATOS DEL VEHÍCULO")
    draw_info_grid(c, MARGIN_X + left_w + 12, y - 18, right_w, vehicle_rows, row_h=17)

    y = 455
    draw_text_box(c, MARGIN_X, y, 260, 58, "3. OBSERVACIONES DEL CLIENTE", data.get("obs_cliente"), max_lines=4)
    draw_text_box(c, MARGIN_X + 272, y, PAGE_W - MARGIN_X * 2 - 272, 58, "4. OBSERVACIONES DEL TÉCNICO", data.get("obs_tecnico"), max_lines=4)
    y = 355
    draw_text_box(c, MARGIN_X, y, PAGE_W - MARGIN_X * 2, 44, "5. RESUMEN EJECUTIVO", data.get("resumen"), fill=GREY, max_lines=3)
    y = 280
    oil_rows = [
        ("Grado de aceite de motor", data.get("grado_aceite")),
        ("Marca de aceite de motor", data.get("marca_aceite")),
        ("Litros de aceite", data.get("litros_aceite")),
    ]
    section_title(c, MARGIN_X, y, PAGE_W - MARGIN_X * 2, "6. REPORTE DEL SERVICIO - MOTOR")
    draw_info_grid(c, MARGIN_X, y - 18, PAGE_W - MARGIN_X * 2, oil_rows, row_h=16)
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, y - 76, PAGE_W - MARGIN_X * 2, "COMPONENTES DE MOTOR", MOTOR_COMPONENTS, row_h=14, title_fill=LIGHT_BLUE)
    draw_footer(c)
    c.showPage()
    page_num += 1

    draw_header(c, page_num, total_pages=total_pages, report_code=data.get("codigo_informe", ""), order_number=data.get("numero_orden", ""))
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, CONTENT_START_Y, PAGE_W - MARGIN_X * 2, "7. SISTEMA DE SUSPENSIÓN - COMPONENTES DELANTEROS", SUSPENSION_FRONT, row_h=20)
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, 535, PAGE_W - MARGIN_X * 2, "SISTEMA DE SUSPENSIÓN - COMPONENTES POSTERIORES", SUSPENSION_REAR, row_h=20)
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, 410, PAGE_W - MARGIN_X * 2, "8. SISTEMA DE FRENOS", BRAKES, row_h=16)
    draw_footer(c)
    c.showPage()
    page_num += 1

    draw_header(c, page_num, total_pages=total_pages, report_code=data.get("codigo_informe", ""), order_number=data.get("numero_orden", ""))
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, CONTENT_START_Y, PAGE_W - MARGIN_X * 2, "9. SISTEMA DE DIRECCIÓN", STEERING, row_h=22)
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, 565, PAGE_W - MARGIN_X * 2, "10. TRANSMISIÓN", TRANSMISSION, row_h=20)
    _, page_num = draw_component_table_guarded(c, page_num, total_pages, data, MARGIN_X, 435, PAGE_W - MARGIN_X * 2, "11. AFINAMIENTO ELECTRÓNICO", ELECTRONIC_TUNE, row_h=22)
    draw_footer(c)
    c.showPage()
    page_num += 1

    draw_header(c, page_num, total_pages=total_pages, report_code=data.get("codigo_informe", ""), order_number=data.get("numero_orden", ""))
    y = CONTENT_START_Y
    validation_rows = [
        ("Técnico responsable", data.get("tecnico_responsable")),
        ("Informe realizado por", data.get("informe_realizado_por")),
        ("Fecha de emisión", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    section_title(c, MARGIN_X, y, PAGE_W - MARGIN_X * 2, "12. VALIDACIÓN")
    validation_bottom = draw_info_grid(c, MARGIN_X, y - 18, PAGE_W - MARGIN_X * 2, validation_rows, row_h=22)
    y = validation_bottom - 25
    legend_bottom = draw_legend(c, MARGIN_X, y, PAGE_W - MARGIN_X * 2)
    y = legend_bottom - 20
    recommendations_bottom = draw_text_box(c, MARGIN_X, y, PAGE_W - MARGIN_X * 2, 46, "RECOMENDACIONES TÉCNICAS AUTOMÁTICAS", build_recommendation_summary(data), fill=LIGHT_BLUE, max_lines=3)
    y = recommendations_bottom - 14
    notes_bottom = draw_text_box(
        c,
        MARGIN_X,
        y,
        PAGE_W - MARGIN_X * 2,
        42,
        "NOTAS DEL INFORME",
        "Este documento resume la evaluación registrada por el técnico en el formulario AUTOGOOD. Los estados y observaciones deben revisarse con el cliente antes de autorizar trabajos adicionales.",
        fill=GREY,
        max_lines=3,
    )
    next_service = data.get("proximo_mantenimiento") or data.get("km_proximo") or "No especificado"
    next_bottom = draw_text_box(c, MARGIN_X, notes_bottom - 12, PAGE_W - MARGIN_X * 2, 28, "PRÓXIMO MANTENIMIENTO", next_service, fill=LIGHT_BLUE, max_lines=1)
    y = next_bottom - 24
    box_w = (PAGE_W - MARGIN_X * 2 - 18) / 2
    draw_signature_box(c, MARGIN_X, y, box_w, 105, "FIRMA DEL TÉCNICO", data.get("firma_tecnico"))
    draw_signature_box(c, MARGIN_X + box_w + 18, y, box_w, 105, "FIRMA DEL CLIENTE", data.get("firma_cliente"))
    c.setFillColor(LIGHT_GREY)
    c.setStrokeColor(BORDER)
    c.rect(MARGIN_X, 48, PAGE_W - MARGIN_X * 2, 48, fill=1, stroke=1)
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(PAGE_W / 2, 82, "AUTOGOOD")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 8)
    c.drawCentredString(PAGE_W / 2, 68, "Informe técnico vehicular generado desde formulario digital.")
    c.drawCentredString(PAGE_W / 2, 56, "Documento válido para revisión de servicio, seguimiento y entrega al cliente.")
    draw_footer(c)
    if photos:
        c.showPage()
        draw_photo_pages(c, photos, page_num + 1, total_pages, data)
    c.save()
    return output_path


def component_inputs(title, components):
    st.markdown(f"**{title}**")
    for component in components:
        key = component_key(component)
        col1, col2, col3 = st.columns([0.30, 0.34, 0.36])
        with col1:
            st.selectbox(component, STATE_OPTIONS, key=f"{key}_estado")
        with col2:
            options = OBSERVATION_LIBRARY.get(component)
            if options:
                st.selectbox("Observación", [""] + options, key=f"{key}_obs", label_visibility="collapsed")
            else:
                st.text_input("Observación corta", key=f"{key}_obs", placeholder="Detalle breve", label_visibility="collapsed")
        with col3:
            state = st.session_state.get(f"{key}_estado", "No aplica")
            obs = st.session_state.get(f"{key}_obs", "")
            st.text_input(
                "Recomendación automática",
                value=automatic_recommendation(state, obs),
                key=f"{key}_rec_preview",
                disabled=True,
                label_visibility="collapsed",
            )


def collect_form_data():
    keys = [
        "cliente",
        "numero_orden",
        "distrito",
        "telefono",
        "fecha",
        "dni_ruc",
        "email",
        "placa",
        "color",
        "marca",
        "km_actual",
        "modelo",
        "km_proximo",
        "proximo_mantenimiento",
        "anio",
        "combustible",
        "tipo_mantenimiento",
        "gdi",
        "tipo_vehiculo",
        "obs_cliente",
        "obs_tecnico",
        "resumen",
        "grado_aceite",
        "marca_aceite",
        "litros_aceite",
        "tecnico_responsable",
        "firma_tecnico",
        "firma_cliente",
        "informe_realizado_por",
        "cargo",
    ]
    data = {key: st.session_state.get(key, "") for key in keys}
    if isinstance(data["fecha"], date):
        data["fecha"] = data["fecha"].strftime("%d/%m/%Y")
    all_components = MOTOR_COMPONENTS + SUSPENSION_FRONT + SUSPENSION_REAR + BRAKES + STEERING + TRANSMISSION + ELECTRONIC_TUNE
    for component in all_components:
        key = component_key(component)
        data[f"{key}_estado"] = st.session_state.get(f"{key}_estado", "No aplica")
        data[f"{key}_obs"] = st.session_state.get(f"{key}_obs", "")
        data[f"{key}_rec"] = automatic_recommendation(data[f"{key}_estado"], data[f"{key}_obs"])
    return data


def main():
    st.set_page_config(page_title="AUTOGOOD - Informe de Servicio", page_icon="🚗", layout="wide")
    st.write("INICIO MAIN")
    st.write("DEBUG: app_reconstruido.py cargó main() correctamente.")
    try:
        apply_mobile_styles()
        ensure_dirs()
        st.write("DEBUG: estilos y carpetas cargados.")

        st.title("AUTOGOOD - Informe técnico vehicular")
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=260)
        else:
            st.warning("LOGOO.png no se encontró. La app seguirá funcionando sin logo.")
        st.caption("Formulario digital para técnicos. No usa OCR ni lectura de fotos; todos los datos se completan manualmente.")
        local_ip = get_local_ip()
        st.info("Si la app está publicada en internet, abra el enlace público desde cualquier celular o tablet.")
        st.caption("Si la usa dentro del taller en red local, también puede abrir esta dirección:")
        st.code(f"http://{local_ip}:8501", language="text")

        st.write("DEBUG: antes de login.")
        if not require_login():
            st.write("DEBUG: login renderizado; esperando credenciales.")
            return
        st.write("DESPUÉS DEL LOGIN")
        st.write("DEBUG: login correcto; formulario renderizado.")
    except Exception as exc:
        st.error("Error al cargar la pantalla inicial de AUTOGOOD.")
        st.exception(exc)
        return

    with st.form("service_report_form"):
        st.write("SECCIÓN FORMULARIO PRINCIPAL")
        with st.expander("1. Información del cliente", expanded=True):
            st.write("SECCIÓN 1: INFORMACIÓN DEL CLIENTE")
            c1, c2 = st.columns(2)
            with c1:
                st.text_input("Nombre y apellido", key="cliente")
                st.text_input("Número de orden", key="numero_orden")
                st.text_input("Distrito", key="distrito")
                st.text_input("Teléfono", key="telefono")
            with c2:
                st.date_input("Fecha", value=date.today(), key="fecha", format="DD/MM/YYYY")
                st.text_input("DNI/RUC", key="dni_ruc")
                st.text_input("Email", key="email")

        with st.expander("2. Datos del vehículo", expanded=True):
            st.write("SECCIÓN 2: DATOS DEL VEHÍCULO")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Placa", key="placa")
                st.text_input("Color", key="color")
                st.text_input("Marca", key="marca")
                st.text_input("Km actual", key="km_actual")
            with c2:
                st.text_input("Modelo", key="modelo")
                st.text_input("Km próximo mantenimiento", key="km_proximo")
                st.text_input("Próximo mantenimiento", key="proximo_mantenimiento", placeholder="Ej.: En 5,000 km o 6 meses")
                st.text_input("Año", key="anio")
                st.selectbox("Tipo de combustible", ["Gasolina", "Diésel", "GLP", "GNV", "Híbrido", "Eléctrico", "Otro"], key="combustible")
            with c3:
                st.text_input("Tipo de mantenimiento", key="tipo_mantenimiento")
                st.selectbox("GDI", ["No aplica", "Sí", "No"], key="gdi")
                st.text_input("Tipo de vehículo", key="tipo_vehiculo")

        with st.expander("3. Observaciones", expanded=True):
            st.write("SECCIÓN 3: OBSERVACIONES")
            st.text_area("Observaciones del cliente", key="obs_cliente", height=90)
            st.text_area("Observaciones del técnico", key="obs_tecnico", height=90)
            st.text_area("Resumen ejecutivo", key="resumen", height=90)

        with st.expander("4. Reporte del servicio - Motor", expanded=True):
            st.write("SECCIÓN 4: REPORTE DEL SERVICIO - MOTOR")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Grado de aceite de motor", key="grado_aceite")
            with c2:
                st.text_input("Marca de aceite de motor", key="marca_aceite")
            with c3:
                st.text_input("Litros de aceite", key="litros_aceite")
            component_inputs("Componentes de motor", MOTOR_COMPONENTS)

        with st.expander("5. Sistema de suspensión"):
            st.write("SECCIÓN 5: SISTEMA DE SUSPENSIÓN")
            component_inputs("Componentes delanteros", SUSPENSION_FRONT)
            st.divider()
            component_inputs("Componentes posteriores", SUSPENSION_REAR)

        with st.expander("6. Sistema de frenos"):
            st.write("SECCIÓN 6: SISTEMA DE FRENOS")
            component_inputs("Componentes de frenos", BRAKES)

        with st.expander("7. Sistema de dirección"):
            st.write("SECCIÓN 7: SISTEMA DE DIRECCIÓN")
            component_inputs("Componentes de dirección", STEERING)

        with st.expander("8. Transmisión"):
            st.write("SECCIÓN 8: TRANSMISIÓN")
            component_inputs("Componentes de transmisión", TRANSMISSION)

        with st.expander("9. Afinamiento electrónico"):
            st.write("SECCIÓN 9: AFINAMIENTO ELECTRÓNICO")
            component_inputs("Componentes de afinamiento electrónico", ELECTRONIC_TUNE)

        with st.expander("10. Evidencia fotográfica"):
            st.write("SECCIÓN 10: EVIDENCIA FOTOGRÁFICA")
            st.file_uploader("Fotos antes", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="fotos_antes")
            st.file_uploader("Fotos después", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="fotos_despues")

        with st.expander("11. Validación", expanded=True):
            st.write("SECCIÓN 11: VALIDACIÓN")
            c1, c2 = st.columns(2)
            with c1:
                st.text_input("Técnico responsable", key="tecnico_responsable")
                st.text_input("Firma del técnico", key="firma_tecnico")
                st.text_input("Informe realizado por", key="informe_realizado_por")
            with c2:
                st.text_input("Firma del cliente", key="firma_cliente")
                st.text_input("Cargo", key="cargo")

        submitted = st.form_submit_button("Generar PDF de servicio", use_container_width=True)

    if submitted:
        st.write("SECCIÓN GENERACIÓN PDF")
        form_data = collect_form_data()
        form_data["codigo_informe"] = get_next_report_code()
        form_data["usuario_generador"] = st.session_state.get("active_user", "")
        form_data["photos"] = prepare_uploaded_photos(st.session_state.get("fotos_antes", []), "Antes")
        form_data["photos"].extend(prepare_uploaded_photos(st.session_state.get("fotos_despues", []), "Después"))
        pdf_path = generate_pdf(form_data)
        history_saved = append_history(form_data, pdf_path)
        st.success(f"PDF generado correctamente: {pdf_path}")
        if not history_saved:
            st.warning("El informe no se guardó en historial porque falta cliente, placa o teléfono.")
        with open(pdf_path, "rb") as pdf_file:
            st.download_button("Descargar PDF de servicio", data=pdf_file, file_name=os.path.basename(pdf_path), mime="application/pdf", use_container_width=True)

    st.divider()
    st.write("SECCIÓN HISTORIAL")
    st.subheader("Buscar historial")
    search_query = st.text_input("Buscar por placa, cliente, teléfono o código de informe", key="historial_busqueda")
    results = search_history(search_query)
    if results:
        for index, record in enumerate(results):
            pdf_path = record.get("PDF generado", "")
            pdf_exists = os.path.exists(pdf_path)
            with st.container(border=True):
                st.markdown(
                    f"**{record.get('Código de informe', '')}** | "
                    f"{record.get('Fecha', '')} | "
                    f"{record.get('Cliente', '')} | "
                    f"{record.get('Placa', '')} | "
                    f"{record.get('Teléfono', '')}"
                )
                st.caption(
                    f"{record.get('Marca', '')} {record.get('Modelo', '')} | "
                    f"Técnico: {record.get('Técnico', '')} | "
                    f"Usuario: {record.get('Usuario', '')}"
                )
                c1, c2 = st.columns(2)
                with c1:
                    if pdf_exists:
                        st.markdown(f"[Abrir PDF]({pdf_link(pdf_path)})")
                    else:
                        st.caption("PDF no encontrado en la ruta guardada.")
                with c2:
                    if pdf_exists:
                        with open(pdf_path, "rb") as pdf_file:
                            st.download_button(
                                "Descargar PDF",
                                data=pdf_file,
                                file_name=os.path.basename(pdf_path),
                                mime="application/pdf",
                                key=f"download_historial_{index}_{record.get('Código de informe', '')}",
                                use_container_width=True,
                            )
    elif search_query:
        st.info("No se encontraron informes con ese criterio.")
    else:
        st.caption("Ingresa una placa, cliente, teléfono o código de informe para consultar el historial.")


if __name__ == "__main__":
    main()

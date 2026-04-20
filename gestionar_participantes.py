#!/usr/bin/env python3
"""
Gestión de participantes del curso OVA Preprácticas ITM en Google Sheets.

Uso:
  python gestionar_participantes.py setup
  python gestionar_participantes.py agregar "<nombre>" <cedula> <correo>
  python gestionar_participantes.py listar

Configuración previa:
  1. Crea un proyecto en https://console.cloud.google.com
  2. Activa la API de Google Sheets y Google Drive
  3. Crea una Cuenta de Servicio y descarga el JSON como 'credentials.json'
  4. Comparte el sheet con el email de la cuenta de servicio (con permisos de editor)
"""

import sys
import json
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# ── Configuración ──────────────────────────────────────────────────────────────
SPREADSHEET_ID  = "1etkSENFncJgRmQdnoSkJMihk4B_i1OnF80L7-FKDZU0"
CREDENTIALS_FILE = "credentials.json"
SHEET_NAME      = "Participantes"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MODULES = [
    {"id": 1, "title": "Compromiso y Reglamentos",     "xp": 100, "quiz": 3},
    {"id": 2, "title": "Seguridad de la Información",  "xp": 100, "quiz": 3},
    {"id": 3, "title": "Habilidades Blandas",           "xp": 150, "quiz": 3},
    {"id": 4, "title": "El Mundo Organizacional",       "xp": 100, "quiz": 3},
    {"id": 5, "title": "Metodologías Ágiles",           "xp": 150, "quiz": 3},
    {"id": 6, "title": "Manejo del Tiempo",             "xp": 100, "quiz": 3},
    {"id": 7, "title": "IA y Uso Responsable",          "xp": 150, "quiz": 3},
    {"id": 8, "title": "Herramientas y Actualización",  "xp": 150, "quiz": 3},
]

# Paleta de colores ITM
_AZUL_ITM    = {"red": 0.051, "green": 0.129, "blue": 0.404}   # #0D2167
_VERDE_ITM   = {"red": 0.102, "green": 0.420, "blue": 0.227}   # #1a6b3a
_MORADO      = {"red": 0.361, "green": 0.090, "blue": 0.588}   # #5C1796
_AZUL_CLARO  = {"red": 0.933, "green": 0.945, "blue": 1.000}   # #EEF2FF
_BLANCO      = {"red": 1.000, "green": 1.000, "blue": 1.000}
_GRIS_BORDE  = {"red": 0.800, "green": 0.800, "blue": 0.800}

# Colores alternados para cada módulo
_COLORES_MOD = [
    {"red": 0.051, "green": 0.357, "blue": 0.678},   # azul medio
    {"red": 0.102, "green": 0.420, "blue": 0.227},   # verde ITM
    {"red": 0.463, "green": 0.043, "blue": 0.510},   # morado oscuro
    {"red": 0.710, "green": 0.396, "blue": 0.114},   # naranja oscuro
    {"red": 0.051, "green": 0.129, "blue": 0.404},   # azul ITM
    {"red": 0.020, "green": 0.514, "blue": 0.494},   # verde azulado
    {"red": 0.569, "green": 0.118, "blue": 0.102},   # rojo oscuro
    {"red": 0.220, "green": 0.290, "blue": 0.380},   # gris azulado
]

FIXED_COLS = 7   # Nombre, Cédula, Correo, Mód. Completados, Progreso%, XP Total, Fecha
MOD_COLS   = 4   # Contenido, Quiz, Puntaje, XP  (por módulo)


# ── Conexión ───────────────────────────────────────────────────────────────────

def _connect():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def _get_or_create_sheet(spreadsheet):
    try:
        return spreadsheet.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(SHEET_NAME, rows=1000, cols=50)


# ── Helpers de formato (Sheets API v4) ────────────────────────────────────────

def _fmt(sheet_id, r0, r1, c0, c1, bg, fg=_BLANCO, bold=False, size=10, wrap="WRAP", valign="MIDDLE"):
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                      "startColumnIndex": c0, "endColumnIndex": c1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": bg,
                "textFormat": {"foregroundColor": fg, "bold": bold, "fontSize": size,
                               "fontFamily": "Arial"},
                "verticalAlignment": valign,
                "horizontalAlignment": "CENTER",
                "wrapStrategy": wrap,
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,"
                      "horizontalAlignment,wrapStrategy)",
        }
    }


def _merge(sheet_id, r0, r1, c0, c1):
    return {"mergeCells": {
        "range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                  "startColumnIndex": c0, "endColumnIndex": c1},
        "mergeType": "MERGE_ALL",
    }}


def _col_width(sheet_id, c0, c1, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                  "startIndex": c0, "endIndex": c1},
        "properties": {"pixelSize": px},
        "fields": "pixelSize",
    }}


def _row_height(sheet_id, r0, r1, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                  "startIndex": r0, "endIndex": r1},
        "properties": {"pixelSize": px},
        "fields": "pixelSize",
    }}


def _borders(sheet_id, r0, r1, c0, c1):
    s = {"style": "SOLID", "color": _GRIS_BORDE}
    return {"updateBorders": {
        "range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                  "startColumnIndex": c0, "endColumnIndex": c1},
        "top": s, "bottom": s, "left": s, "right": s,
        "innerHorizontal": s, "innerVertical": s,
    }}


def _freeze(sheet_id):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                       "gridProperties": {"frozenRowCount": 2, "frozenColumnCount": 0}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
    }}


def _align_left(sheet_id, r0, r1, c0, c1):
    return {"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r1,
                  "startColumnIndex": c0, "endColumnIndex": c1},
        "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT"}},
        "fields": "userEnteredFormat.horizontalAlignment",
    }}


# ── Comandos principales ───────────────────────────────────────────────────────

def setup_template():
    """Crea los encabezados con formato en el Google Sheet."""
    spreadsheet = _connect()
    sheet = _get_or_create_sheet(spreadsheet)
    sid = sheet.id
    total_cols = FIXED_COLS + len(MODULES) * MOD_COLS

    # ── Fila 1: grupos de secciones ──
    row1 = (
        ["DATOS DEL ESTUDIANTE", "", ""]
        + ["PROGRESO GENERAL", "", "", ""]
    )
    for i, m in enumerate(MODULES):
        row1 += [f"M{m['id']}: {m['title']}", "", "", ""]

    # ── Fila 2: nombres de columnas ──
    row2 = [
        "Nombre", "Cédula", "Correo",
        "Módulos\nCompletados", "Progreso\n(%)", "XP\nTotal", "Fecha\nRegistro",
    ]
    for m in MODULES:
        row2 += [
            "Contenido\nLeído",
            "Quiz\nAprobado",
            f"Puntaje\n(/{m['quiz']})",
            f"XP\n(/{m['xp']})",
        ]

    sheet.clear()
    sheet.update([row1, row2], "A1")

    # ── Requests de formato ──
    req = []

    # Congelar filas y columna
    req.append(_freeze(sid))

    # Tamaños de fila
    req.append(_row_height(sid, 0, 1, 40))
    req.append(_row_height(sid, 1, 2, 56))

    # Ancho de columnas fijas
    req.append(_col_width(sid, 0, 1, 220))   # Nombre
    req.append(_col_width(sid, 1, 2, 110))   # Cédula
    req.append(_col_width(sid, 2, 3, 210))   # Correo
    req.append(_col_width(sid, 3, 7, 95))    # Progreso (4 cols)
    # Ancho columnas por módulo
    req.append(_col_width(sid, FIXED_COLS, total_cols, 78))

    # ── Formato fila 1 ──
    req.append(_fmt(sid, 0, 1, 0, 3,         _AZUL_ITM,  bold=True, size=11))  # Datos estudiante
    req.append(_fmt(sid, 0, 1, 3, FIXED_COLS, _MORADO,   bold=True, size=11))  # Progreso general
    for i, m in enumerate(MODULES):
        c0 = FIXED_COLS + i * MOD_COLS
        req.append(_fmt(sid, 0, 1, c0, c0 + MOD_COLS,
                        _COLORES_MOD[i], bold=True, size=10))

    # ── Formato fila 2 ──
    req.append(_fmt(sid, 1, 2, 0, total_cols,
                    {"red": 0.20, "green": 0.22, "blue": 0.25},
                    bold=True, size=10))

    # ── Fusión de celdas fila 1 ──
    req.append(_merge(sid, 0, 1, 0, 3))            # Datos del estudiante
    req.append(_merge(sid, 0, 1, 3, FIXED_COLS))   # Progreso general
    for i in range(len(MODULES)):
        c0 = FIXED_COLS + i * MOD_COLS
        req.append(_merge(sid, 0, 1, c0, c0 + MOD_COLS))

    # ── Bordes encabezados ──
    req.append(_borders(sid, 0, 2, 0, total_cols))

    # ── Alinear columna Nombre a la izquierda en datos ──
    req.append(_align_left(sid, 2, 1000, 0, 3))

    # ── Fondo alternado para filas de datos (bandas) ──
    req.append({
        "addBanding": {
            "bandedRange": {
                "bandedRangeId": 1,
                "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 1000,
                          "startColumnIndex": 0, "endColumnIndex": total_cols},
                "rowProperties": {
                    "firstBandColor":  _BLANCO,
                    "secondBandColor": _AZUL_CLARO,
                },
            }
        }
    })

    spreadsheet.batch_update({"requests": req})

    print(f"Plantilla configurada correctamente.")
    print(f"  Sheet: '{SHEET_NAME}'")
    print(f"  Columnas totales: {total_cols}  ({FIXED_COLS} fijas + {len(MODULES)} módulos × {MOD_COLS})")
    print(f"  URL: {spreadsheet.url}")


def agregar_participante(nombre, cedula, correo, progreso=None):
    """
    Agrega o actualiza un participante en el sheet.

    progreso (opcional): dict con el estado exportado por el OVA:
    {
        "modulos_completados": 5,
        "progreso_pct": 62,
        "xp_total": 600,
        "modulos": [
            {"contenido": True,  "quiz_aprobado": True,  "puntaje": 3, "xp": 100},
            {"contenido": True,  "quiz_aprobado": False, "puntaje": 1, "xp": 0},
            ...  (8 entradas, una por módulo en orden)
        ]
    }
    """
    spreadsheet = _connect()
    sheet = _get_or_create_sheet(spreadsheet)

    # Buscar si ya existe por cédula (columna B = índice 2)
    cedulas = sheet.col_values(2)
    existing_row = None
    for i, val in enumerate(cedulas[2:], start=3):
        if str(val).strip() == str(cedula).strip():
            existing_row = i
            break

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

    row = [nombre, str(cedula), correo]

    if progreso:
        mods_prog = progreso.get("modulos", [{}] * len(MODULES))
        row += [
            progreso.get("modulos_completados", 0),
            progreso.get("progreso_pct", 0),
            progreso.get("xp_total", 0),
            fecha,
        ]
        for idx, mp in enumerate(mods_prog[:len(MODULES)]):
            m = MODULES[idx]
            puntaje = mp.get("puntaje", 0)
            row += [
                "✓" if mp.get("contenido") else "—",
                "✓" if mp.get("quiz_aprobado") else "—",
                f"{puntaje}/{m['quiz']}",
                mp.get("xp", 0),
            ]
        # Rellenar si faltan módulos
        for _ in range(len(MODULES) - len(mods_prog)):
            row += ["—", "—", f"0/{MODULES[0]['quiz']}", 0]
    else:
        row += [0, "0%", 0, fecha]
        for m in MODULES:
            row += ["—", "—", f"0/{m['quiz']}", 0]

    if existing_row:
        sheet.update(f"A{existing_row}", [row], value_input_option="USER_ENTERED")
        print(f"Participante actualizado: {nombre}  (fila {existing_row})")
    else:
        sheet.append_row(row, value_input_option="USER_ENTERED")
        print(f"Participante agregado: {nombre}")


def listar_participantes():
    """Muestra todos los participantes registrados."""
    spreadsheet = _connect()
    sheet = _get_or_create_sheet(spreadsheet)

    all_rows = sheet.get_all_values()
    data_rows = all_rows[2:]  # Saltar 2 filas de encabezado

    if not data_rows or not any(r[0] for r in data_rows):
        print("No hay participantes registrados aún.")
        return

    print(f"\n{'#':<4} {'Nombre':<30} {'Cédula':<12} {'Correo':<30} {'Mód.':<6} {'XP'}")
    print("─" * 90)
    count = 0
    for i, row in enumerate(data_rows, start=1):
        if not row[0]:
            continue
        nombre   = row[0] if len(row) > 0 else ""
        cedula   = row[1] if len(row) > 1 else ""
        correo   = row[2] if len(row) > 2 else ""
        mod_comp = row[3] if len(row) > 3 else "0"
        xp_total = row[5] if len(row) > 5 else "0"
        print(f"{i:<4} {nombre:<30} {cedula:<12} {correo:<30} {mod_comp:<6} {xp_total}")
        count += 1
    print(f"\nTotal: {count} participante(s)")


# ── CLI ────────────────────────────────────────────────────────────────────────

def _usage():
    print(__doc__)
    print("Comandos disponibles:")
    print("  setup              Configura encabezados y formato en el sheet")
    print("  agregar <nombre> <cedula> <correo>")
    print("                     Registra un nuevo participante (sin progreso)")
    print("  listar             Muestra todos los participantes")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        _usage()
        sys.exit(0)

    cmd = sys.argv[1].lower()

    if cmd == "setup":
        setup_template()

    elif cmd == "agregar":
        if len(sys.argv) < 5:
            print("Error: faltan argumentos.")
            print('Uso: python gestionar_participantes.py agregar "<nombre completo>" <cedula> <correo>')
            sys.exit(1)
        agregar_participante(sys.argv[2], sys.argv[3], sys.argv[4])

    elif cmd == "listar":
        listar_participantes()

    else:
        print(f"Comando desconocido: '{cmd}'")
        _usage()
        sys.exit(1)

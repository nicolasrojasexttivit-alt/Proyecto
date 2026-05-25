#!/usr/bin/env python3
"""
buscar_usuario.py
=================
Busca un usuario por Hostname en un .xlsx y genera su archivo .md
con preguntas interactivas por sección.
"""

import argparse
import subprocess
import sys
from pathlib import Path

import pandas as pd


# ══════════════════════════════════════════════════════════════
#  ★  CONFIGURACIÓN — EDITA ESTA SECCIÓN  ★
# ══════════════════════════════════════════════════════════════
print("\n" + "="*60)
fecha = input("Ingrese la fecha de toma de evidencias: ")
print("\n" + "="*60)
print("\n" + "="*60)
opc = input("Ingrese el tipo de Template a llenar: 1. OFICINA 2. Comex&Logistica: ")
tipo_template = "OFICINA" if opc == "1"  else "Comex&Logistica"
print("\n" + "="*60)
PROPIEDADES_FIJAS = {
    "OFI_COMMEX":   tipo_template,
    "GESTOR":        "GESTOR: Jhoan Nicolas Cruz Sierra",
    "FECHA":           "FECHA: " + fecha,
    "LUGAR":        "LUGAR: Bogotá",
    "CARGO":          "CARGO: Analista Soporte en Sitio",
    "CIUDAD":           "Bogotá",
    "tipo_actividad": "Inventario de equipos",
}

PLANTILLA_IMAGENES = {
    "IMAGEN_1": "{Hostname}/1 (1).png",
    "IMAGEN_2": "{Hostname}/1 (2).png",
    "IMAGEN_3": "{Hostname}/1 (3).png",
    "IMAGEN_4": "{Hostname}/1 (4).png",
    "IMAGEN_5": "{Hostname}/1 (5).png",
    "IMAGEN_6": "{Hostname}/1 (6).png",
    "IMAGEN_7": "{Hostname}/1 (7).png",
    "IMAGEN_8": "{Hostname}/1 (8).png",
    "IMAGEN_9": "{Hostname}/1 (9).png",
    "IMAGEN_10": "{Hostname}/1 (10).png",
    "IMAGEN_11": "{Hostname}/1 (11).png",
    "IMAGEN_12": "{Hostname}/1 (12).png",
    "IMAGEN_13": "{Hostname}/1 (13).png",
    "IMAGEN_14": "{Hostname}/1 (14).png",
    "IMAGEN_15": "{Hostname}/1 (15).png",
    "IMAGEN_16": "{Hostname}/1 (16).png",
}

COLUMNAS = {
    "Hostname": "Hostname",
    "USUARIO":   "Nombre",
    "correo":   "Correo",
    "ip":       "IP",
}


# ══════════════════════════════════════════════════════════════
#  ★  NUEVA FUNCIONALIDAD: PREGUNTAS POR SECCIÓN  ★
# ══════════════════════════════════════════════════════════════

def preguntar_si_no(mensaje: str, default: str = "N") -> bool:
    """Pregunta S/N con valor por defecto."""
    while True:
        resp = input(f"{mensaje} (S/N) [{default}]: ").strip().upper()
        if resp == "":
            resp = "N"
        if resp in ("S", "N"):
            return resp == "S"
        print("   Por favor responde S o N.")


def obtener_observaciones(seccion: str) -> str:
    """Pide observaciones. Si solo Enter → cadena vacía."""
    print(f"\n📝 Observaciones para {seccion} (Enter = ninguna):")
    obs = input("> ").strip()
    mensaje = "Observaciones: "
    return "Observaciones: " + obs if obs else "Observaciones: Sin observaciones"


def seccion_1_pna():
    """Sección 1: Páginas no autorizadas - Default = NO"""
    print("\n" + "="*60)
    print("1. PÁGINAS NO AUTORIZADAS (Default: TODO en NO)")
    print("="*60)

    if not preguntar_si_no("¿Hay algún cambio respecto al default (NO)?"):
        return {
            "FACEBOOK_SI": "' '", "FACEBOOK_NO": "X",
            "INSTA_SI": "' '", "INSTA_NO": "X",
            "X_SI": "' '", "X_NO": "X",
            "TIKTOK_SI": "' '", "TIKTOK_NO": "X",
            "OBSERVACIONES_PNAUTORIZADAS": "Observaciones: Sin observaciones"
        }

    datos = {}
    for app in ["FACEBOOK", "INSTAGRAM", "X / TWITTER", "TIKTOK"]:
        key_si = app.replace(" / ", "_").replace(" ", "_").upper() + "_SI"
        key_no = app.replace(" / ", "_").replace(" ", "_").upper() + "_NO"
        
        tiene_acceso = preguntar_si_no(f"¿Permite acceso a {app}?")
        datos[key_si] = "X" if tiene_acceso else "' '"
        datos[key_no] = "' '" if tiene_acceso else "X"

    datos["OBSERVACIONES_PNAUTORIZADAS"] = obtener_observaciones("Páginas no autorizadas")
    return datos


def seccion_2_chat():
    """Sección 2: Chat en línea - Default = NO"""
    print("\n" + "="*60)
    print("2. PLATAFORMAS DE CHAT EN LÍNEA (Default: NO)")
    print("="*60)

    if not preguntar_si_no("¿Hay algún cambio?"):
        return {
            "WHAT_WEB_S": "' '", "WHAT_WEB_N": "X",
            "WHAT_APP_S": "' '", "WHAT_APP_N": "X",
            "OBSERVACIONES_WHATSAPP": "Observaciones: Sin observaciones"
        }

    datos = {}
    for app in ["WHATSAPP WEB", "WHATSAPP APP"]:
        prefix = "WHAT_WEB" if "WEB" in app else "WHAT_APP"
        permite = preguntar_si_no(f"¿Permite acceso a {app}?")
        datos[f"{prefix}_S"] = "X" if permite else "' '"
        datos[f"{prefix}_N"] = "' '" if permite else "X"

    datos["OBSERVACIONES_WHATSAPP"] = obtener_observaciones("Chat en línea")
    return datos


def seccion_3_antivirus():
    """Sección 3: Antivirus/EDR - Default = SÍ"""
    print("\n" + "="*60)
    print("3. ANTIVIRUS Y EDR (Default: McAfee y Symantec SÍ)")
    print("="*60)

    if not preguntar_si_no("¿Hay algún cambio en CrowdStrike u otros?"):
        return {
            "CROWD_SI": "X", "CROWD_NO": "' '",
            "OBSERVACIONES_ANTI": "Observaciones: Sin observaciones"
        }

    crowd_si = preguntar_si_no("¿Tiene CrowdStrike EDR instalado?")
    datos = {
        "CROWD_SI": "X" if crowd_si else "' '",
        "CROWD_NO": "' '" if crowd_si else "X",
        "OBSERVACIONES_ANTI": obtener_observaciones("Antivirus/EDR")
    }
    return datos


def seccion_4_plugin():
    """Sección 4: Plugin navegador - Default = SÍ"""
    print("\n" + "="*60)
    print("4. PLUGIN NAVEGADOR (Default: SÍ)")
    print("="*60)

    if not preguntar_si_no("¿Hay algún cambio?"):
        return {
            "TRELL_SI": "X", "TRELL_NO": "' '",
            "OBSERVACIONES_TRELL": "Observaciones: Sin observaciones"
        }

    trell_si = preguntar_si_no("¿Tiene Trellix Control Web instalado y activo?")
    datos = {
        "TRELL_SI": "X" if trell_si else "' '",
        "TRELL_NO": "' '" if trell_si else "X",
        "OBSERVACIONES_TRELL": obtener_observaciones("Plugin Navegador")
    }
    return datos


def seccion_5_control_remoto():
    """Sección 5: Instalación Control Remoto - Default = NO"""
    print("\n" + "="*60)
    print("5. INSTALACIÓN CONTROL REMOTO LOCAL (Default: NO)")
    print("="*60)

    if not preguntar_si_no("¿Hay algún cambio?"):
        return {
            "ANY_SI": "' '", "ANY_NO": "X",
            "TEAM_SI": "' '", "TEAM_NO": "X",
            "OBSERVACIONES_CREMOTO": "Observaciones: Sin observaciones"
        }

    datos = {}
    for tool in ["ANYDESK", "TEAM VIEWER"]:
        prefix = "ANY" if "ANY" in tool else "TEAM"
        instalado = preguntar_si_no(f"¿Tiene instalado {tool}?")
        datos[f"{prefix}_SI"] = "X" if instalado else "' '"
        datos[f"{prefix}_NO"] = "' '" if instalado else "X"

    datos["OBSERVACIONES_CREMOTO"] = obtener_observaciones("Control Remoto Instalado")
    return datos


def seccion_6_acceso_web():
    """Sección 6: Acceso web a control remoto - Default = NO"""
    print("\n" + "="*60)
    print("6. ACCESO WEB CONTROL REMOTO (Default: NO)")
    print("="*60)

    if not preguntar_si_no("¿Hay algún cambio?"):
        return {
            "ANYW_SI": "' '", "ANYW_NO": "X",
            "TEAMW_SI": "' '", "TEAMW_NO": "X",
            "OBSERVACIONS_CRWEB": "Observaciones: Sin observaciones"
        }

    datos = {}
    for tool in ["ANYDESK - URL WEB", "TEAM VIEWER - URL WEB"]:
        prefix = "ANYW" if "ANY" in tool else "TEAMW"
        permite = preguntar_si_no(f"¿Permite acceso a {tool}?")
        datos[f"{prefix}_SI"] = "X" if permite else "' '"
        datos[f"{prefix}_NO"] = "' '" if permite else "X"

    datos["OBSERVACIONS_CRWEB"] = obtener_observaciones("Acceso Web Control Remoto Web")
    return datos


def seccion_7_panel():
    """Sección 7: Programas instalados"""
    print("\n" + "="*60)
    print("7. PROGRAMAS INSTALADOS - PANEL DE CONTROL")
    print("="*60)
    tiene_no_autorizados = preguntar_si_no("¿El equipo contiene programas NO autorizados?")
    obs = obtener_observaciones("Programas instalados")
    print(obs)
    return {
        "OBSERVACIONES_PANEL": obs,
        # Puedes agregar más campos si quieres marcar SI/NO explícitamente
    }


def seccion_8_sap():
    """Sección 8: SAP"""
    print("\n" + "="*60)
    print("8. AUTENTICACIÓN EN SAP")
    print("="*60)
    obs = obtener_observaciones("SAP")
    return {"OBSERVACIONES_SAP": obs}


# ══════════════════════════════════════════════════════════════

def construir_md(registro: dict, Hostname: str, datos_secciones: dict) -> str:
    """Genera el contenido del .md con todos los campos."""
    lineas = ["---"]

    # Datos del Excel
    for clave_md, col_xlsx in COLUMNAS.items():
        valor = registro.get(col_xlsx, "").strip()
        if valor:
            lineas.append(f"{clave_md}: {_escapar_yaml(valor)}")

    # Propiedades fijas
    for clave, valor in PROPIEDADES_FIJAS.items():
        lineas.append(f"{clave}: {_escapar_yaml(str(valor))}")

    # Datos de las secciones interactivas
    for k, v in datos_secciones.items():
        lineas.append(f"{k}: {_escapar_yaml(str(v))}")

    lineas.append("---")
    lineas.append("")

    # Imágenes
    lineas.append("imagenes:")
    for placeholder, plantilla in PLANTILLA_IMAGENES.items():
        ruta = plantilla.replace("{Hostname}", Hostname)
        lineas.append(f"  {placeholder}: {ruta}")

    lineas.append("")
    return "\n".join(lineas)


def _escapar_yaml(valor: str) -> str:
    especiales = (':', '#', '{', '}', '[', ']', '&', '*', '!', '@')
    if any(c in valor for c in especiales):
        return f'"{valor}"'
    return valor


def main():
    parser = argparse.ArgumentParser(description="Busca usuario y genera .md con preguntas interactivas.")
    parser.add_argument("-x", "--xlsx", required=True, help="Ruta al archivo Excel")
    parser.add_argument("--Hostname", help="Hostname a buscar")
    parser.add_argument("-o", "--salida", help="Carpeta de salida")
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.exists():
        sys.exit(f"ERROR: No se encontró {xlsx_path}")

    df = pd.read_excel(xlsx_path, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    df = df.fillna("").apply(lambda col: col.str.strip())

    # Buscar usuario
    Hostname = args.Hostname or input("🔍 Ingresa el Hostname: ").strip()
    col_h = COLUMNAS["Hostname"]
    coincidencias = df[df[col_h].str.upper() == Hostname.upper()]

    if coincidencias.empty:
        sys.exit(f"❌ No encontrado: {Hostname}")

    registro = coincidencias.iloc[0].to_dict()
    Hostname_real = registro[col_h]

    print(f"\n✅ Usuario encontrado: {Hostname_real} - {registro.get('Nombre', '')}\n")

    # === EJECUTAR SECCIONES INTERACTIVAS ===
    datos = {}
    datos.update(seccion_1_pna())
    datos.update(seccion_2_chat())
    datos.update(seccion_3_antivirus())
    datos.update(seccion_4_plugin())
    datos.update(seccion_5_control_remoto())
    datos.update(seccion_6_acceso_web())
    datos.update(seccion_7_panel())
    datos.update(seccion_8_sap())

    # Generar MD
    contenido = construir_md(registro, Hostname_real, datos)

    salida_dir = Path(args.salida) if args.salida else xlsx_path.parent
    salida_dir.mkdir(parents=True, exist_ok=True)
    md_path = salida_dir / f"{Hostname_real}.md"

    md_path.write_text(contenido, encoding="utf-8")
    print(f"\n📄 Markdown generado: {md_path}")

    # ── Paso 2: fill_template.py → genera el .docx ──────────────────
    docx_path = salida_dir / f"{Hostname_real}.docx"
    script_dir = Path(__file__).parent   # misma carpeta que buscar_usuario.py
    fill_script = script_dir / "fill_template.py"

    print(f"\n⚙️  Generando DOCX: {docx_path.name} ...")
    resultado = subprocess.run(
        [sys.executable, str(fill_script),
         "-t", "seguridad_template.docx",
         "-d", str(md_path),
         "-o", str(docx_path)],
        capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if resultado.returncode != 0:
        print("❌ Error en fill_template.py:")
        print(resultado.stderr)
        sys.exit(1)
    print(f"✅ DOCX generado: {docx_path.name}")

    # ── Paso 3: LibreOffice → convierte el .docx a PDF ──────────────
    LIBRE_OFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"

    print(f"\n⚙️  Convirtiendo a PDF ...")
    resultado_pdf = subprocess.run(
        [LIBRE_OFFICE, "--headless", "--convert-to", "pdf",
         str(docx_path), "--outdir", str(salida_dir)],
        capture_output=True, text=True, encoding='utf-8', errors='replace'
    )
    if resultado_pdf.returncode != 0:
        print("❌ Error al convertir a PDF:")
        print(resultado_pdf.stderr)
        sys.exit(1)

    pdf_path = salida_dir / f"{Hostname_real}.pdf"
    print(f"✅ PDF generado: {pdf_path.name}")
    print(f"\n{'─'*50}")
    print(f"  .md   → {md_path}")
    print(f"  .docx → {docx_path}")
    print(f"  .pdf  → {pdf_path}")
    print(f"{'─'*50}\n")


if __name__ == "__main__":
    main()
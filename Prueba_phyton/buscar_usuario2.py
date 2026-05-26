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
class GeneradorReporteSeguridad:
    def __init__(self, xlsx_path: str, hostname: str = None, salida: str = None):
        self.xlsx_path = Path(xlsx_path)
        self.hostname = hostname
        self.salida_dir = Path(salida) if salida else None
        self.tipo_template = None
        self.fecha = None
        self.PROPIEDADES_FIJAS = {}
        self.PLANTILLA_IMAGENES = {
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
        self.columnas = {
            "Hostname": "Hostname",
            "USUARIO":   "Nombre",
            "correo":   "Correo",
            "ip":       "IP",
        }
        self.df = None
        self.registro = None
        self.hostname_real = None

    def cargar_configuracion(self):
        print("\n" + "="*60)
        self.fecha = input("Ingrese la fecha de toma de evidencias: ")
        print("\n" + "="*60)
        opc = input("Ingrese el tipo de Template a llenar: 1. OFICINA 2. Comex&Logistica: ")
        self.tipo_template = "OFICINA" if opc == "1" else "Comex&Logistica"
        print("\n" + "="*60)
        self.PROPIEDADES_FIJAS = {
            "OFI_COMMEX":   self.tipo_template,
            "GESTOR":        "GESTOR: Jhoan Nicolas Cruz Sierra",
            "FECHA":           f"FECHA: {self.fecha}",
            "LUGAR":        "LUGAR: Bogotá",
            "CARGO":          "CARGO: Analista Soporte en Sitio",
            "CIUDAD":           "Bogotá",
            "tipo_actividad": "Inventario de equipos",
        }
    
    def cargar_excel(self):
        """Carga y prepara el DataFrame"""
        if not self.xlsx_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo: {self.xlsx_path}")
        
        self.df = pd.read_excel(self.xlsx_path, dtype=str, keep_default_na=False)
        self.df.columns = [c.strip() for c in self.df.columns]
        self.df = self.df.fillna("").apply(lambda col: col.str.strip())

    def buscar_usuario(self):
        """Busca el hostname en el Excel"""
        if not self.hostname:
            self.hostname = input("🔍 Ingresa el Hostname: ").strip()

        col_h = self.columnas["Hostname"]
        coincidencias = self.df[self.df[col_h].str.upper() == self.hostname.upper()]

        if coincidencias.empty:
            raise ValueError(f"❌ No encontrado: {self.hostname}")

        self.registro = coincidencias.iloc[0].to_dict()
        self.hostname_real = self.registro[col_h]
        print(f"\n✅ Usuario encontrado: {self.hostname_real} - {self.registro.get('Nombre', '')}\n")

    def seccion_1_pna(self) -> dict:
        """Sección 1: Páginas no autorizadas - Default = NO"""
        print("\n" + "="*60)
        print("1. PÁGINAS NO AUTORIZADAS (Default: TODO en NO)")
        print("="*60)

        if not self.preguntar_si_no("¿Hay algún cambio respecto al default (NO)?"):
            return {
                "FACEBOOK_SI": "' '", "FACEBOOK_NO": "X",
                "INSTA_SI": "' '", "INSTA_NO": "X",
                "X_SI": "' '", "X_NO": "X",
                "TIKTOK_SI": "' '", "TIKTOK_NO": "X",
                "OBSERVACIONES_PNAUTORIZADAS": "Observaciones: Sin observaciones"
            }

        datos = {}
        for app in ["FACEBOOK", "INSTA", "X", "TIKTOK"]:
            key_si = app.replace(" / ", "_").replace(" ", "_").upper() + "_SI"
            key_no = app.replace(" / ", "_").replace(" ", "_").upper() + "_NO"
            
            tiene_acceso = self.preguntar_si_no(f"¿Permite acceso a {app}?")
            datos[key_si] = "X" if tiene_acceso else "' '"
            datos[key_no] = "' '" if tiene_acceso else "X"

        datos["OBSERVACIONES_PNAUTORIZADAS"] = self.obtener_observaciones("Páginas no autorizadas")
        return datos

    def seccion_2_chat(self) -> dict:
        """Sección 2: Chat en línea - Default = NO"""
        print("\n" + "="*60)
        print("2. PLATAFORMAS DE CHAT EN LÍNEA (Default: NO)")
        print("="*60)

        if not self.preguntar_si_no("¿Hay algún cambio?"):
            return {
                "WHAT_WEB_S": "' '", "WHAT_WEB_N": "X",
                "WHAT_APP_S": "' '", "WHAT_APP_N": "X",
                "OBSERVACIONES_WHATSAPP": "Observaciones: Sin observaciones"
            }

        datos = {}
        for app in ["WHATSAPP WEB", "WHATSAPP APP"]:
            prefix = "WHAT_WEB" if "WEB" in app else "WHAT_APP"
            permite = self.preguntar_si_no(f"¿Permite acceso a {app}?")
            datos[f"{prefix}_S"] = "X" if permite else "' '"
            datos[f"{prefix}_N"] = "' '" if permite else "X"

        datos["OBSERVACIONES_WHATSAPP"] = self.obtener_observaciones("Chat en línea")
        return datos

    def seccion_3_antivirus(self) -> dict:
        """Sección 3: Antivirus/EDR - Default = SÍ"""
        print("\n" + "="*60)
        print("3. ANTIVIRUS Y EDR (Default: McAfee y Symantec SÍ)")
        print("="*60)

        if not self.preguntar_si_no("¿Hay algún cambio en CrowdStrike u otros?"):
            return {
                "CROWD_SI": "X", "CROWD_NO": "' '",
                "OBSERVACIONES_ANTI": "Observaciones: Sin observaciones"
            }

        crowd_si = self.preguntar_si_no("¿Tiene CrowdStrike EDR instalado?")
        datos = {
            "CROWD_SI": "X" if crowd_si else "' '",
            "CROWD_NO": "' '" if crowd_si else "X",
            "OBSERVACIONES_ANTI": self.obtener_observaciones("Antivirus/EDR")
        }
        return datos

    def seccion_4_plugin(self) -> dict:
        """Sección 4: Plugin navegador - Default = SÍ"""
        print("\n" + "="*60)
        print("4. PLUGIN NAVEGADOR (Default: SÍ)")
        print("="*60)

        if not self.preguntar_si_no("¿Hay algún cambio?"):
            return {
                "TRELL_SI": "X", "TRELL_NO": "' '",
                "OBSERVACIONES_TRELL": "Observaciones: Sin observaciones"
            }

        trell_si = self.preguntar_si_no("¿Tiene Trellix Control Web instalado y activo?")
        datos = {
            "TRELL_SI": "X" if trell_si else "' '",
            "TRELL_NO": "' '" if trell_si else "X",
            "OBSERVACIONES_TRELL": self.obtener_observaciones("Plugin Navegador")
        }
        return datos

    def seccion_5_control_remoto(self) -> dict:
        """Sección 5: Instalación Control Remoto - Default = NO"""
        print("\n" + "="*60)
        print("5. INSTALACIÓN CONTROL REMOTO LOCAL (Default: NO)")
        print("="*60)

        if not self.preguntar_si_no("¿Hay algún cambio?"):
            return {
                "ANY_SI": "' '", "ANY_NO": "X",
                "TEAM_SI": "' '", "TEAM_NO": "X",
                "OBSERVACIONES_CREMOTO": "Observaciones: Sin observaciones"
            }

        datos = {}
        for tool in ["ANYDESK", "TEAM VIEWER"]:
            prefix = "ANY" if "ANY" in tool else "TEAM"
            instalado = self.preguntar_si_no(f"¿Tiene instalado {tool}?")
            datos[f"{prefix}_SI"] = "X" if instalado else "' '"
            datos[f"{prefix}_NO"] = "' '" if instalado else "X"

        datos["OBSERVACIONES_CREMOTO"] = self.obtener_observaciones("Control Remoto Instalado")
        return datos

    def seccion_6_acceso_web(self) -> dict:
        """Sección 6: Acceso web a control remoto - Default = NO"""
        print("\n" + "="*60)
        print("6. ACCESO WEB CONTROL REMOTO (Default: NO)")
        print("="*60)

        if not self.preguntar_si_no("¿Hay algún cambio?"):
            return {
                "ANYW_SI": "' '", "ANYW_NO": "X",
                "TEAMW_SI": "' '", "TEAMW_NO": "X",
                "OBSERVACIONS_CRWEB": "Observaciones: Sin observaciones"
            }

        datos = {}
        for tool in ["ANYDESK - URL WEB", "TEAM VIEWER - URL WEB"]:
            prefix = "ANYW" if "ANY" in tool else "TEAMW"
            permite = self.preguntar_si_no(f"¿Permite acceso a {tool}?")
            datos[f"{prefix}_SI"] = "X" if permite else "' '"
            datos[f"{prefix}_NO"] = "' '" if permite else "X"

        datos["OBSERVACIONS_CRWEB"] = self.obtener_observaciones("Acceso Web Control Remoto Web")
        return datos

    def seccion_7_panel(self) -> dict:
        """Sección 7: Programas instalados"""
        print("\n" + "="*60)
        print("7. PROGRAMAS INSTALADOS - PANEL DE CONTROL")
        print("="*60)
        tiene_no_autorizados = self.preguntar_si_no("¿El equipo contiene programas NO autorizados?")
        obs = self.obtener_observaciones("Programas instalados")
        print(obs)
        return {
            "OBSERVACIONES_PANEL": obs,
            # Puedes agregar más campos si quieres marcar SI/NO explícitamente
        }

    def seccion_8_sap(self) -> dict:
        """Sección 8: SAP"""
        print("\n" + "="*60)
        print("8. AUTENTICACIÓN EN SAP")
        print("="*60)
        obs = self.obtener_observaciones("SAP")
        return {"OBSERVACIONES_SAP": obs}

    def preguntar_si_no(self,mensaje: str, default: str = "N") -> bool:
        """Pregunta S/N con valor por defecto."""
        while True:
            resp = input(f"{mensaje} (S/N) [{default}]: ").strip().upper()
            if resp == "":
                resp = "N"
            if resp in ("S", "N"):
                return resp == "S"
            print("   Por favor responde S o N.")

    def obtener_observaciones(self, seccion: str) -> str:
        """Pide observaciones. Si solo Enter → cadena vacía."""
        print(f"\n📝 Observaciones para {seccion} (Enter = ninguna):")
        obs = input("> ").strip()
        return f"Observaciones: {obs}" if obs else "Observaciones: Sin observaciones"

    def construir_md(self, registro: dict, hostname: str, datos_secciones: dict) -> str:
        """Genera el contenido del .md con todos los campos."""
        lineas = ["---"]

        # Datos del Excel
        for clave_md, col_xlsx in self.columnas.items():
            valor = registro.get(col_xlsx, "").strip()
            if valor:
                lineas.append(f"{clave_md}: {self._escapar_yaml(valor)}")

        # Propiedades fijas
        for clave, valor in self.PROPIEDADES_FIJAS.items():
            lineas.append(f"{clave}: {self._escapar_yaml(str(valor))}")

        # Datos de las secciones interactivas
        for k, v in datos_secciones.items():
            lineas.append(f"{k}: {self._escapar_yaml(str(v))}")

        lineas.append("---")
        lineas.append("")

        # Imágenes
        lineas.append("imagenes:")
        for placeholder, plantilla in self.PLANTILLA_IMAGENES.items():
            ruta = plantilla.replace("{Hostname}", hostname)
            lineas.append(f"  {placeholder}: {ruta}")

        lineas.append("")
        return "\n".join(lineas)

    def _escapar_yaml(self, valor: str) -> str:
        """Escapa caracteres especiales para YAML."""
        especiales = (':', '#', '{', '}', '[', ']', '&', '*', '!', '@')
        if any(c in valor for c in especiales):
            return f'"{valor}"'
        return valor

    def generar_archivos(self, datos_secciones: dict):
        """Orquesta la generación completa: MD → DOCX → PDF"""
        if not self.registro or not self.hostname_real:
            raise ValueError("No hay registro de usuario cargado. Ejecuta buscar_usuario() primero.")

        # ====================== GENERAR MARKDOWN ======================
        print("\n📝 Generando archivo Markdown...")
        contenido_md = self.construir_md(self.registro, self.hostname_real, datos_secciones)

        # Definir carpeta de salida
        salida_dir = self.salida_dir if self.salida_dir else self.xlsx_path.parent
        salida_dir.mkdir(parents=True, exist_ok=True)

        md_path = salida_dir / f"{self.hostname_real}.md"
        md_path.write_text(contenido_md, encoding="utf-8")

        print(f"✅ Markdown generado: {md_path.name}")

        # ====================== GENERAR DOCX ======================
        print("\n⚙️  Generando documento Word (.docx)...")

        script_dir = Path(__file__).parent
        fill_script = script_dir / "fill_template.py"

        if not fill_script.exists():
            raise FileNotFoundError(f"No se encontró fill_template.py en: {fill_script.parent}")

        docx_path = salida_dir / f"{self.hostname_real}.docx"

        resultado = subprocess.run(
            [sys.executable, str(fill_script),
             "-t", "seguridad_template.docx",
             "-d", str(md_path),
             "-o", str(docx_path)],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        if resultado.returncode != 0:
            print("❌ Error al ejecutar fill_template.py:")
            print(resultado.stderr)
            raise RuntimeError("Falló la generación del DOCX")

        print(f"✅ DOCX generado: {docx_path.name}")

        # ====================== CONVERTIR A PDF ======================
        print("\n⚙️  Convirtiendo a PDF...")

        LIBRE_OFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"

        if not Path(LIBRE_OFFICE).exists():
            print("⚠️ LibreOffice no encontrado en la ruta predeterminada.")
            raise FileNotFoundError("LibreOffice no encontrado")

        resultado_pdf = subprocess.run(
            [LIBRE_OFFICE, "--headless", "--convert-to", "pdf",
             str(docx_path), "--outdir", str(salida_dir)],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        if resultado_pdf.returncode != 0:
            print("❌ Error al convertir a PDF:")
            print(resultado_pdf.stderr)
            raise RuntimeError("Falló la conversión a PDF")

        pdf_path = salida_dir / f"{self.hostname_real}.pdf"

        # ====================== RESUMEN FINAL ======================
        print(f"\n{'─'*60}")
        print("🎉 ¡Reporte generado exitosamente!")
        print(f"{'─'*60}")
        print(f"📄 Markdown : {md_path}")
        print(f"📝 Word     : {docx_path}")
        print(f"📕 PDF      : {pdf_path}")
        print(f"{'─'*60}\n")

def main():
    parser = argparse.ArgumentParser(description="Generador de reportes de seguridad")
    parser.add_argument("-x", "--xlsx", required=True, help="Ruta al archivo Excel")
    parser.add_argument("--Hostname", help="Hostname a buscar")
    parser.add_argument("-o", "--salida", help="Carpeta de salida")
    args = parser.parse_args()

    try:
        # Creamos la instancia de la clase
        generador = GeneradorReporteSeguridad(
            xlsx_path=args.xlsx,
            hostname=args.Hostname,
            salida=args.salida
        )

        # Flujo principal
        generador.cargar_configuracion() 
        generador.cargar_excel()
        generador.buscar_usuario()

        # Ejecutar todas las secciones interactivas
        datos_secciones = {}
        datos_secciones.update(generador.seccion_1_pna())
        datos_secciones.update(generador.seccion_2_chat())
        datos_secciones.update(generador.seccion_3_antivirus())
        datos_secciones.update(generador.seccion_4_plugin())
        datos_secciones.update(generador.seccion_5_control_remoto())
        datos_secciones.update(generador.seccion_6_acceso_web())
        datos_secciones.update(generador.seccion_7_panel())
        datos_secciones.update(generador.seccion_8_sap())

        # Generar archivos (esta parte la haremos juntos en el próximo paso)
        generador.generar_archivos(datos_secciones)   # ← crearemos este método

    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
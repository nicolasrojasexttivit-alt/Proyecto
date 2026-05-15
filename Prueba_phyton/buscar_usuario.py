#!/usr/bin/env python3
"""
buscar_usuario.py
=================
Busca un usuario por Hostname en un .xlsx y genera su archivo .md
listo para usar con fill_template.py.

USO
───
  # Modo interactivo (te pregunta el Hostname)
  python buscar_usuario.py -x inventario.xlsx

  # Directo por argumento
  python buscar_usuario.py -x inventario.xlsx --Hostname PC-LAB-03

  # Con carpeta de salida específica
  python buscar_usuario.py -x inventario.xlsx --Hostname PC-LAB-03 -o salida/

ESTRUCTURA DEL XLSX
────────────────────
  Columna A: Hostname
  Columna B: Nombre
  Columna C: Correo
  Columna D: IP
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


# ══════════════════════════════════════════════════════════════
#  ★  CONFIGURACIÓN — EDITA ESTA SECCIÓN  ★
# ══════════════════════════════════════════════════════════════

# ── Propiedades fijas ──────────────────────────────────────────
# Estas aparecen igual en TODOS los archivos .md, sin importar el usuario.
# Agrega o quita las que necesites.
PROPIEDADES_FIJAS = {
    "GESTOR":        "GESTOR: Jhoan Nicolas Cruz Sierra",
    "FECHA":           "FECHA 00/00/00",
    "LUGAR":        "LUGAR XD",
    "CARGO":          "CARGO Analista Soporte en Sitio",
    "CIUDAD":           "Bogotá",
    "tipo_actividad": "Inventario de equipos",
    # Agrega más aquí:
    # "otro_campo": "otro valor",
}

# ── Plantilla de imágenes ──────────────────────────────────────
# Usa {Hostname} donde debe ir el Hostname del usuario.
# Agrega o quita líneas según cuántas imágenes necesites.
# Formato:  "NOMBRE_PLACEHOLDER": "{Hostname}/nombre_archivo.png"
# Con ancho: "IMAGEN_3": "{Hostname}/captura.png @ 10"   (10 cm de ancho)
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

    # Agrega más aquí:
    # "IMAGEN_6": "{Hostname}/1 (6).png",
}

# ── Mapeo de columnas del xlsx ─────────────────────────────────
# Clave   = nombre que tendrá en el .md (placeholder en el docx)
# Valor   = nombre EXACTO de la columna en el xlsx (cabecera)
COLUMNAS = {
    "Hostname": "Hostname",   # columna A — también se usa para buscar
    "USUARIO":   "Nombre",     # columna B
    "correo":   "Correo",     # columna C
    "ip":       "IP",         # columna D
    # Si tienes más columnas en el xlsx, agrégalas aquí:
    # "cargo": "Cargo",
}

# ══════════════════════════════════════════════════════════════


def leer_xlsx(xlsx_path: Path) -> pd.DataFrame:
    """Carga el xlsx y valida que existan las columnas esperadas."""
    try:
        df = pd.read_excel(xlsx_path, dtype=str, keep_default_na=False)
    except Exception as e:
        sys.exit(f"ERROR al abrir el archivo: {e}")

    df.columns = [c.strip() for c in df.columns]
    df = df.fillna("").apply(lambda col: col.str.strip())

    # Verificar columnas requeridas
    cols_requeridas = list(COLUMNAS.values())
    faltantes = [c for c in cols_requeridas if c not in df.columns]
    if faltantes:
        sys.exit(
            f"ERROR: Columnas no encontradas en el xlsx: {faltantes}\n"
            f"Columnas disponibles: {list(df.columns)}\n"
            f"Revisa la sección COLUMNAS en el script."
        )
    return df


def buscar_Hostname(df: pd.DataFrame, Hostname: str) -> dict | None:
    """
    Busca el Hostname (insensible a mayúsculas/espacios).
    Devuelve el primer registro como dict, o None si no existe.
    """
    col_Hostname = COLUMNAS["Hostname"]
    mascara = df[col_Hostname].str.upper() == Hostname.strip().upper()
    coincidencias = df[mascara]

    if coincidencias.empty:
        return None
    return coincidencias.iloc[0].to_dict()


def construir_md(registro: dict, Hostname: str) -> str:
    """Genera el contenido completo del archivo .md."""
    lineas = ["---"]

    # 1. Datos extraídos del xlsx
    for clave_md, col_xlsx in COLUMNAS.items():
        valor = registro.get(col_xlsx, "").strip()
        if valor:
            lineas.append(f"{clave_md}: {_escapar_yaml(valor)}")

    # 2. Propiedades fijas
    for clave, valor in PROPIEDADES_FIJAS.items():
        lineas.append(f"{clave}: {_escapar_yaml(str(valor))}")

    lineas.append("---")
    lineas.append("")

    # 3. Sección de imágenes con Hostname reemplazado
    lineas.append("imagenes:")
    for placeholder, plantilla in PLANTILLA_IMAGENES.items():
        ruta = plantilla.replace("{Hostname}", Hostname)
        lineas.append(f"  {placeholder}: {ruta}")

    lineas.append("")
    return "\n".join(lineas)


def _escapar_yaml(valor: str) -> str:
    """Envuelve en comillas si el valor contiene caracteres especiales YAML."""
    especiales = (':', '#', '{', '}', '[', ']', '&', '*', '!', '@')
    if any(c in valor for c in especiales):
        return f'"{valor}"'
    return valor


def main():
    parser = argparse.ArgumentParser(
        description="Genera un .md desde un xlsx buscando por Hostname.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("-x", "--xlsx", required=True, metavar="ARCHIVO.xlsx",
                        help="Ruta al archivo Excel")
    parser.add_argument("--Hostname", metavar="Hostname",
                        help="Hostname a buscar (si se omite, el script lo pregunta)")
    parser.add_argument("-o", "--salida", metavar="CARPETA/",
                        help="Carpeta de salida (por defecto: misma carpeta del xlsx)")
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.exists():
        sys.exit(f"ERROR: No se encontró el archivo: {xlsx_path}")

    print(f"\n📂 Cargando: {xlsx_path}")
    df = leer_xlsx(xlsx_path)
    print(f"   {len(df)} usuario(s) encontrados en el archivo.\n")

    # Obtener Hostname (argumento o input interactivo)
    if args.Hostname:
        Hostname = args.Hostname.strip()
    else:
        Hostname = input("🔍 Ingresa el Hostname a buscar: ").strip()

    if not Hostname:
        sys.exit("ERROR: El Hostname no puede estar vacío.")

    # Buscar usuario
    registro = buscar_Hostname(df, Hostname)
    if registro is None:
        print(f"\n❌ No se encontró ningún usuario con Hostname: '{Hostname}'")
        # Mostrar Hostnames similares como sugerencia
        col_h = COLUMNAS["Hostname"]
        todos = df[col_h].str.upper().tolist()
        similares = [h for h in todos if Hostname.upper()[:4] in h]
        if similares:
            print(f"   ¿Quisiste decir? → {', '.join(similares[:5])}")
        sys.exit(1)

    # Usar el Hostname tal como está en el xlsx (respetando mayúsculas originales)
    Hostname_real = registro[COLUMNAS["Hostname"]]

    print(f"\n✅ Usuario encontrado:")
    for clave_md, col_xlsx in COLUMNAS.items():
        print(f"   {clave_md:<15} {registro.get(col_xlsx, '')}")

    # Construir contenido del .md
    contenido_md = construir_md(registro, Hostname_real)

    # Ruta de salida
    salida_dir = Path(args.salida) if args.salida else xlsx_path.parent
    salida_dir.mkdir(parents=True, exist_ok=True)
    md_path = salida_dir / f"{Hostname_real}.md"

    md_path.write_text(contenido_md, encoding="utf-8")

    print(f"\n📄 Archivo generado: {md_path}")
    print("\n── Contenido ──────────────────────────────────")
    print(contenido_md)
    print("───────────────────────────────────────────────")
    print(f"\nSiguiente paso:")
    print(f"  python fill_template.py -t formato.docx -d {md_path}\n")


if __name__ == "__main__":
    main()

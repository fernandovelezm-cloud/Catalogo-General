#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
comparar_y_validar.py

Ejecuta el comando "comparar" + validación de la galería:

1. Recorre recursivamente C:\\Users\\ferna\\Dropbox\\nueva colección\\Catálogo
   e identifica códigos de producto (nombre de archivo sin extensión).
2. Compara por código (sin distinguir mayúsculas/minúsculas, sin importar
   la extensión) contra los archivos en C:\\Users\\ferna\\Dropbox\\nueva colección\\Costos Ficha técnica:
   - Si el código ya existe en destino -> REEMPLAZA esa foto (convirtiendo
     el formato si es necesario para conservar la extensión del destino).
   - Si no existe -> la AGREGA como "<código>.jpg".
   - Archivos cuyo nombre no empieza con al menos 2 dígitos (sin código
     reconocible, ej. "Gemini_Generated_Image_xxxx", "Plana", "altiva azul")
     se EXCLUYEN (no se copian, solo se informan).
   - Sufijos "(1)", "(2)", etc. al final del nombre se ignoran para la
     comparación (ej. "57040 (1)" se compara como "57040").
3. Sincroniza el array `items` de index.html con el contenido real de la
   carpeta "Costos Ficha técnica":
   - Agrega entradas para archivos de imagen nuevos que no estén en la galería.
   - Elimina entradas cuyo archivo ya no exista en la carpeta (huérfanos).
4. Escribe un log con el detalle de lo hecho (comparar_log.txt).

Requiere: Python 3 + Pillow (el .bat lo instala si falta).
"""

import os
import re
import sys
from datetime import datetime

try:
    from PIL import Image
except ImportError:
    print("ERROR: falta la librería Pillow. Ejecuta: pip install Pillow")
    sys.exit(1)

CATALOGO = r"C:\Users\ferna\Dropbox\nueva colección\Catálogo"
DEST = r"C:\Users\ferna\Dropbox\nueva colección\Costos Ficha técnica"
INDEX_HTML = os.path.join(DEST, "index.html")
LOG_PATH = os.path.join(DEST, "comparar_log.txt")

IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff"}

RE_CODE = re.compile(r"^\d{2,}")
RE_SUFFIX = re.compile(r"\s*\(\d+\)\s*$")


def normalize_code(stem):
    s = RE_SUFFIX.sub("", stem)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_recognizable(code):
    return bool(RE_CODE.match(code))


def save_as(img, path):
    """Guarda img en path, eligiendo formato según la extensión de path."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg"):
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        img.save(path, "JPEG", quality=90)
    elif ext == ".png":
        img.save(path, "PNG")
    else:
        if img.mode in ("RGBA", "P", "LA") and ext in (".bmp",):
            img = img.convert("RGB")
        img.save(path)


def main():
    log_lines = []

    def log(msg=""):
        print(msg)
        log_lines.append(msg)

    log("=" * 60)
    log("comparar_y_validar.py - %s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log("=" * 60)

    # ------------------------------------------------------------------
    # 1) Indexar Catálogo (recursivo)
    # ------------------------------------------------------------------
    catalog_codes = {}  # norm(lower) -> (ruta_origen, codigo_visible)
    sin_codigo = []

    for root, dirs, files in os.walk(CATALOGO):
        for fn in files:
            stem, ext = os.path.splitext(fn)
            if ext.lower() not in IMG_EXTS:
                continue
            code = normalize_code(stem)
            if not is_recognizable(code):
                sin_codigo.append(os.path.relpath(os.path.join(root, fn), CATALOGO))
                continue
            norm = code.lower()
            if norm not in catalog_codes:
                catalog_codes[norm] = (os.path.join(root, fn), code)

    # ------------------------------------------------------------------
    # 2) Indexar destino (carpeta "Costos Ficha técnica")
    # ------------------------------------------------------------------
    destino_index = {}  # norm(lower) -> filename
    for fn in os.listdir(DEST):
        stem, ext = os.path.splitext(fn)
        if ext.lower() not in IMG_EXTS:
            continue
        destino_index[stem.lower()] = fn

    # ------------------------------------------------------------------
    # 3) Reemplazar / agregar fotos
    # ------------------------------------------------------------------
    reemplazados = []
    agregados = []
    errores = []

    for norm, (src_path, code) in catalog_codes.items():
        try:
            img = Image.open(src_path)
            img.load()
        except Exception as e:
            errores.append((src_path, str(e)))
            continue

        if norm in destino_index:
            dest_fn = destino_index[norm]
            dest_path = os.path.join(DEST, dest_fn)
            try:
                save_as(img, dest_path)
                reemplazados.append((code, dest_fn))
            except Exception as e:
                errores.append((src_path, str(e)))
        else:
            dest_fn = code + ".jpg"
            dest_path = os.path.join(DEST, dest_fn)
            try:
                if img.mode in ("RGBA", "P", "LA"):
                    img = img.convert("RGB")
                img.save(dest_path, "JPEG", quality=90)
                agregados.append((code, dest_fn))
                destino_index[norm] = dest_fn
            except Exception as e:
                errores.append((src_path, str(e)))

    # ------------------------------------------------------------------
    # 4) Sincronizar galería (items en index.html) con la carpeta destino
    # ------------------------------------------------------------------
    html_agregados = []
    html_eliminados = []

    if os.path.exists(INDEX_HTML):
        with open(INDEX_HTML, "r", encoding="utf-8") as f:
            html = f.read()

        m = re.search(r"const items = \[(.*?)\];", html, re.DOTALL)
        if not m:
            log("ADVERTENCIA: no se encontró el array `items` en index.html, no se actualizó la galería.")
        else:
            pairs_str = m.group(1)
            pairs = re.findall(r'\["([^"]*)",\s*"([^"]*)"\]', pairs_str)

            destino_files = [f for f in os.listdir(DEST) if os.path.splitext(f)[1].lower() in IMG_EXTS]
            destino_files_lower = {f.lower(): f for f in destino_files}

            new_pairs = []
            seen_lower = set()

            for code, filename in pairs:
                if filename.lower() in destino_files_lower:
                    new_pairs.append((code, filename))
                    seen_lower.add(filename.lower())
                else:
                    html_eliminados.append((code, filename))

            for f in destino_files:
                if f.lower() not in seen_lower:
                    stem = os.path.splitext(f)[0]
                    new_pairs.append((stem, f))
                    html_agregados.append((stem, f))
                    seen_lower.add(f.lower())

            new_pairs_str = ", ".join('["%s", "%s"]' % (c, fn) for c, fn in new_pairs)
            new_html = html[:m.start(1)] + new_pairs_str + html[m.end(1):]

            if new_html != html:
                with open(INDEX_HTML, "w", encoding="utf-8") as f:
                    f.write(new_html)
    else:
        log("ADVERTENCIA: no se encontró index.html en %s" % DEST)

    # ------------------------------------------------------------------
    # 5) Resumen
    # ------------------------------------------------------------------
    log("")
    log("Reemplazados (%d):" % len(reemplazados))
    for code, fn in reemplazados:
        log("  - %s -> %s" % (code, fn))

    log("")
    log("Agregados como fotos nuevas (%d):" % len(agregados))
    for code, fn in agregados:
        log("  - %s -> %s" % (code, fn))

    log("")
    log("Excluidos por no tener código reconocible (%d):" % len(sin_codigo))
    for fn in sin_codigo:
        log("  - %s" % fn)

    log("")
    log("Errores al procesar (%d):" % len(errores))
    for src, err in errores:
        log("  - %s: %s" % (src, err))

    log("")
    log("Galería (index.html) - entradas agregadas (%d):" % len(html_agregados))
    for code, fn in html_agregados:
        log("  - %s -> %s" % (code, fn))

    log("")
    log("Galería (index.html) - entradas eliminadas por huérfanas (%d):" % len(html_eliminados))
    for code, fn in html_eliminados:
        log("  - %s -> %s" % (code, fn))

    log("")
    log("Listo.")

    with open(LOG_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines) + "\n")


if __name__ == "__main__":
    main()

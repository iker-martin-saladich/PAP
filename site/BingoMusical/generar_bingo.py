"""
Generador de cartrons de Bingo Musical
=======================================

Substitueix Graelles.py + pdffinal.py per un únic script que:

  - NO necessita cap Excel: les cançons es defineixen directament en una
    llista (o es poden llegir d'un fitxer de text, una cançó per línia).
  - Genera automàticament l'assignació de cançons a cada cartró, evitant
    cartrons duplicats i, dins del possible, línies (files) repetides entre
    cartrons diferents.
  - Dibuixa el contingut directament amb reportlab (vectorial, no amb
    imatges PNG intermèdies), de manera que es pot canviar el nombre de
    files/columnes de cada cartró sense perdre qualitat: la mida de lletra
    s'adapta automàticament a la mida de cada cel·la.
  - Solapa els cartrons sobre la plantilla de fons (per defecte 3 per full
    DIN A4) i genera directament el PDF final.

Com adaptar-lo cada any
------------------------
1. Canvia CANCONS per la llista d'aquest any (o activa l'opció de fitxer .txt).
2. Canvia PLANTILLA pel nom de la imatge de fons nova.
3. Si vols cartrons de 3x3, 5x4, etc. només cal canviar N_FILES / N_COLUMNES.

Per aconseguir la llista de cançons, pots exportar-la des de Spotify amb Spotlistr.
"""

import random
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

# ============================================================
# CONFIGURACIÓ — el que caldrà tocar cada any
# ============================================================

# --- Cançons ---
# Es llegeixen directament del fitxer de text exportat (p.ex. amb Spotlistr),
# amb una cançó per línia i el format "Artista - Títol".
# Si una cançó té diversos artistes, al fitxer ja apareixen separats per "; "
# dins del mateix camp d'artista (p.ex. "Figa Flawas; Lluís Gavaldà"), cosa
# que aquest lector gestiona sense cap canvi.

FITXER_CANCONS = "cançons.txt"


def carregar_cancons_de_fitxer(path):
    cancons = []
    with open(path, encoding="utf-8") as f:
        for linia in f:
            linia = linia.strip()
            if not linia:
                continue
            artista, titol = linia.split(" - ", 1)
            cancons.append((artista.strip(), titol.strip()))
    return cancons


CANCONS = carregar_cancons_de_fitxer(FITXER_CANCONS)

# --- Mida de la graella de cada cartró ---
N_FILES = 3       # files per cartró
N_COLUMNES = 4    # columnes per cartró

# --- Nombre de cartrons a generar ---
N_CARTRONS = 150

# --- Plantilla de fons (canviarà cada any) ---
PLANTILLA = "BingoMusical.png"

# --- Sortida ---
OUTPUT_PDF = "BingoMusical_final2.pdf"

# --- Mida de pàgina ---
# Es llegeix automàticament de la mida real del fitxer PLANTILLA, per evitar
# que es deformi/desquadri si la plantilla d'un any té una mida diferent de
# la de l'any anterior. NO cal tocar això mai.
# Mida del full A4 en píxels a 300dpi
PAGE_W, PAGE_H = 2480, 3508

# Coordenades dels rectangles (ajusta segons la teva plantilla)
RECTANGLES = [
    (313, 253, 1860, 901),   # (x, y, ample, alt) primer rectangle
    (313, 1398, 1860, 901),  # segon rectangle
    (315, 2544, 1860, 901),  # tercer rectangle
]


FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
MIDA_FONT_MAX = 30
MIDA_FONT_MIN = 15
MARGE_CELLA = 10


# ============================================================
# GENERACIÓ DE L'ASSIGNACIÓ DE CANÇONS ALS CARTRONS
# ============================================================

def generar_cartrons(cancons, n_files, n_columnes, n_cartrons, max_intents=300):
    """
    Genera `n_cartrons` cartrons de mida n_files x n_columnes a partir de la
    llista de cançons, intentant evitar:
      - cartrons idèntics entre ells (sempre es garanteix)
      - línies (files) idèntiques entre cartrons diferents (es garanteix
        "dins del possible": si hi ha poques cançons i molts cartrons, en
        algun moment és matemàticament inevitable repetir alguna línia)

    Retorna una llista de cartrons; cada cartró és una llista de files, i
    cada fila una llista de tuples (artista, títol).
    """
    caselles_per_cartro = n_files * n_columnes
    if len(cancons) < caselles_per_cartro:
        raise ValueError(
            f"Necessites almenys {caselles_per_cartro} cançons diferents per "
            f"omplir un cartró de {n_files}x{n_columnes}; ara en tens "
            f"{len(cancons)}."
        )

    cartrons = []
    bingos_vists = set()
    files_vistes = set()

    for _ in range(n_cartrons):
        cartro_final = None
        for intent in range(max_intents):
            seleccio = random.sample(cancons, caselles_per_cartro)
            random.shuffle(seleccio)
            files = [
                tuple(seleccio[f * n_columnes:(f + 1) * n_columnes])
                for f in range(n_files)
            ]
            bingo_tuple = tuple(files)

            if bingo_tuple in bingos_vists:
                continue  # cartró idèntic a un ja generat, ho tornem a provar

            hi_ha_linia_repetida = any(fila in files_vistes for fila in files)
            if hi_ha_linia_repetida and intent < max_intents - 1:
                continue  # provem una altra combinació

            cartro_final = files
            break

        if cartro_final is None:
            raise RuntimeError(
                "No s'ha pogut generar un cartró únic. Prova d'afegir més "
                "cançons a la llista o de reduir el nombre de cartrons."
            )

        bingos_vists.add(tuple(cartro_final))
        for fila in cartro_final:
            files_vistes.add(fila)
        cartrons.append(cartro_final)

    return cartrons


def validar_cartrons(cartrons):
    """Mostra per pantalla el mateix tipus d'avisos que l'script original."""
    bingos = [tuple(c) for c in cartrons]
    if len(set(bingos)) == len(bingos):
        print("OK: tots els cartrons són diferents.")
    else:
        print("ATENCIÓ: hi ha cartrons duplicats!")

    totes_files = []
    for c in cartrons:
        totes_files.extend(c)

    duplicades = len(totes_files) - len(set(totes_files))
    if duplicades == 0:
        print("OK: no hi ha cap línia idèntica entre cartrons.")
    else:
        print(
            f"AVÍS: hi ha {duplicades} línia(es) repetida(es) entre cartrons "
            f"diferents. Amb poques cançons i molts cartrons, no sempre és "
            f"possible evitar-ho del tot."
        )


# ============================================================
# DIBUIX VECTORIAL DEL CARTRÓ (reportlab -> sense pèrdua de qualitat)
# ============================================================

def _wrap_text(text, font_name, font_size, max_width):
    """Parteix `text` en línies que caben dins `max_width` amb font_name/font_size."""
    paraules = text.split()
    linies = []
    for paraula in paraules:
        if not linies:
            linies.append(paraula)
        else:
            prova = linies[-1] + " " + paraula
            if stringWidth(prova, font_name, font_size) <= max_width:
                linies[-1] = prova
            else:
                linies.append(paraula)
    return linies


def _cap_cella(text, font_name, font_size, max_width, max_linies=3):
    linies = _wrap_text(text, font_name, font_size, max_width)
    if len(linies) > max_linies:
        linies = linies[:max_linies]
        linies[-1] += "..."
    return linies


def _mida_font_optima(files, cell_w, cell_h):
    """
    Busca la mida de font més gran (entre MIDA_FONT_MIN i MIDA_FONT_MAX) amb
    la qual el text de TOTES les cel·les hi cap sense sortir-se. Això és el
    que permet canviar N_FILES/N_COLUMNES sense haver de retocar cap mida a
    mà: com més gran la graella, més petita (automàticament) la lletra.
    """
    max_amplada = cell_w - 2 * MARGE_CELLA
    max_alcada = cell_h - 2 * MARGE_CELLA

    for mida in range(MIDA_FONT_MAX, MIDA_FONT_MIN - 1, -1):
        alcada_linia = mida * 1.25
        espai_entre = mida * 1.4
        cap_problema = False
        for fila in files:
            for artista, titol in fila:
                l_titol = _cap_cella(titol, FONT_BOLD, mida, max_amplada)
                l_artista = _cap_cella(artista, FONT_REGULAR, mida, max_amplada)
                alt_total = (len(l_titol) + len(l_artista)) * alcada_linia + espai_entre
                if alt_total > max_alcada:
                    cap_problema = True
                    break
            if cap_problema:
                break
        if not cap_problema:
            return mida
    return MIDA_FONT_MIN


def dibuixar_cartro(c, x0, y0, w, h, files, n_files, n_columnes):
    """
    Dibuixa un cartró directament al canvas de reportlab.
    (x0, y0) = cantonada INFERIOR esquerra del cartró (sistema reportlab).
    """
    cell_w = w / n_columnes
    cell_h = h / n_files

    mida_font = _mida_font_optima(files, cell_w, cell_h)
    alcada_linia = mida_font * 1.25
    espai_entre = mida_font * 1.4
    max_amplada = cell_w - 2 * MARGE_CELLA

    c.setLineWidth(2)
    for fila_idx, fila in enumerate(files):
        for col_idx, (artista, titol) in enumerate(fila):
            cx0 = x0 + col_idx * cell_w
            cy1 = y0 + h - fila_idx * cell_h   # vora superior de la cel·la
            cy0 = cy1 - cell_h                  # vora inferior de la cel·la

            c.rect(cx0, cy0, cell_w, cell_h, stroke=1, fill=0)

            l_titol = _cap_cella(titol, FONT_BOLD, mida_font, max_amplada)
            l_artista = _cap_cella(artista, FONT_REGULAR, mida_font, max_amplada)
            alt_total = (len(l_titol) + len(l_artista)) * alcada_linia + espai_entre

            y_text = cy0 + cell_h / 2 + alt_total / 2 - mida_font * 0.35
            cx_centre = cx0 + cell_w / 2

            c.setFont(FONT_BOLD, mida_font)
            for linia in l_titol:
                c.drawCentredString(cx_centre, y_text, linia)
                y_text -= alcada_linia

            y_text -= (espai_entre - alcada_linia)

            c.setFont(FONT_REGULAR, mida_font)
            for linia in l_artista:
                c.drawCentredString(cx_centre, y_text, linia)
                y_text -= alcada_linia


# ============================================================
# GENERACIÓ DEL PDF FINAL (graelles + plantilla en un sol pas)
# ============================================================

def generar_pdf(cartrons, n_files, n_columnes, plantilla, rectangles,
                 output_pdf, page_w, page_h):
    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    for i in range(0, len(cartrons), len(rectangles)):
        c.drawImage(plantilla, 0, 0, width=page_w, height=page_h)

        for j, (x, y_des_de_dalt, w, h) in enumerate(rectangles):
            if i + j >= len(cartrons):
                break
            y0 = page_h - y_des_de_dalt - h  # "des de dalt" -> reportlab
            dibuixar_cartro(c, x, y0, w, h, cartrons[i + j], n_files, n_columnes)

        c.showPage()

    c.save()
    print(f"PDF generat: {output_pdf}")


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

if __name__ == "__main__":
    cartrons = generar_cartrons(CANCONS, N_FILES, N_COLUMNES, N_CARTRONS)
    validar_cartrons(cartrons)
    generar_pdf(
        cartrons, N_FILES, N_COLUMNES, PLANTILLA, RECTANGLES, OUTPUT_PDF,
        PAGE_W, PAGE_H,
    )
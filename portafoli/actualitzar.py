#!/usr/bin/env python3
"""
actualitzar.py
Llegeix els Excel PLM i incrusta les dades directament als HTML.
Funciona tant localment com a GitHub Actions.

Us:  python actualitzar.py
"""

from pathlib import Path
from openpyxl import load_workbook
import re

BASE = Path(__file__).parent

def cel(ws, row, col):
    v = ws.cell(row=row, column=col).value
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v

def es_num(v):
    try:
        int(str(v).strip())
        return True
    except:
        return False

def files_dades(ws, fila_ini=5):
    return [r for r in range(fila_ini, ws.max_row + 1)
            if es_num(ws.cell(r, 1).value)]

def pct_num(v):
    try:
        return float(v)
    except:
        return 0.0

def progress_global(ws_fases, c_pct):
    vals = [pct_num(cel(ws_fases, r, c_pct)) for r in files_dades(ws_fases)]
    return round(sum(vals) / len(vals), 1) if vals else 0

ESTAT_COLOR = {
    "Completat":  ("#4A7C59", "#E8F2EB"),
    "En progres": ("#8B5E3C", "#F0E6D6"),
    "Pendent":    ("#8C7B68", "#F0EEEB"),
}
PRIOR_COLOR = {
    "Alta":  ("#8B3A3A", "#FDEAEA"),
    "Mitja": ("#A07820", "#FEF6E0"),
    "Baixa": ("#4A7C59", "#E8F2EB"),
}
RISC_COLOR = {
    "Alta":  ("#8B3A3A", "#FDEAEA"),
    "Mitja": ("#A07820", "#FEF6E0"),
    "Baixa": ("#4A7C59", "#E8F2EB"),
}
VOCAB_COLOR = {
    "Apres":    ("#4A7C59", "#E8F2EB"),
    "En repas": ("#A07820", "#FEF6E0"),
    "Pendent":  ("#8C7B68", "#F0EEEB"),
}

def get_estat_color(estat):
    # Cerca amb i sense accents
    for k, v in ESTAT_COLOR.items():
        if k.lower().replace('\xe0','a').replace('\xe8','e').replace('\xe9','e') == \
           estat.lower().replace('\xe0','a').replace('\xe8','e').replace('\xe9','e'):
            return v
    return ("#8C7B68", "#F0EEEB")

def badge(text, fg, bg):
    return (
        '<span style="font-size:0.7rem;font-weight:500;padding:0.2rem 0.65rem;'
        'border-radius:2rem;background:' + bg + ';color:' + fg + ';white-space:nowrap">'
        + text + '</span>'
    )

def barra(pct, color="#8B5E3C"):
    p = min(100, max(0, pct_num(pct)))
    return (
        '<div style="height:5px;background:#E2D9CE;border-radius:3px;overflow:hidden;margin-top:0.4rem">'
        '<div style="width:' + str(int(p)) + '%;height:100%;background:' + color + ';border-radius:3px"></div></div>'
        '<div style="font-size:0.72rem;color:#8C7B68;text-align:right;margin-top:2px">' + str(int(p)) + '%</div>'
    )

def icon_estat(estat):
    if "Complet" in estat:
        return "OK"
    if "progr" in estat.lower():
        return "WIP"
    return "..."

def html_fases(ws, c_titol, c_desc, c_pct, c_estat, color_accent="#8B5E3C"):
    parts = []
    for r in files_dades(ws):
        titol = str(cel(ws, r, c_titol))
        desc  = str(cel(ws, r, c_desc))
        pct   = pct_num(cel(ws, r, c_pct))
        estat = str(cel(ws, r, c_estat))
        fg, bg = get_estat_color(estat)
        label = icon_estat(estat)
        parts.append(
            '<div style="display:flex;align-items:center;gap:1rem;padding:0.9rem 1.1rem;'
            'background:#FEFCF9;border:1px solid #E2D9CE;border-radius:0.75rem;margin-bottom:0.6rem">'
            '<div style="flex:1">'
            '<strong style="font-size:0.88rem;display:block">' + titol + '</strong>'
            '<span style="font-size:0.76rem;color:#8C7B68">' + desc + '</span>'
            '</div>'
            + badge(estat, fg, bg)
            + '<div style="width:80px">' + barra(pct, fg) + '</div>'
            '</div>'
        )
    return "\n".join(parts) if parts else "<p style='color:#8C7B68'>Sense dades.</p>"

def html_tasques(ws, c_titol, c_fase, c_prior, c_estat, c_pct, c_notes):
    parts = []
    for r in files_dades(ws):
        titol = str(cel(ws, r, c_titol))
        fase  = str(cel(ws, r, c_fase))
        prior = str(cel(ws, r, c_prior))
        estat = str(cel(ws, r, c_estat))
        notes = str(cel(ws, r, c_notes))
        done  = "Complet" in estat
        fg_p, bg_p = PRIOR_COLOR.get(prior, ("#8C7B68", "#F0EEEB"))
        check = "[OK]" if done else "[ ]"
        td = 'text-decoration:line-through;color:#8C7B68' if done else 'color:#1A1208'
        notes_html = ('<span style="font-size:0.72rem;color:#8C7B68">' + notes + '</span>') if notes else ""
        parts.append(
            '<div style="display:flex;gap:0.75rem;padding:0.6rem 0.4rem;border-bottom:1px solid #E2D9CE">'
            '<span style="font-size:0.85rem;margin-top:1px">' + check + '</span>'
            '<div style="flex:1">'
            '<div style="font-size:0.86rem;' + td + '">' + titol + '</div>'

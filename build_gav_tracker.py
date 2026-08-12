#!/usr/bin/env python3
"""
Pipeline: F01CGPPTO -> baseline JSON -> seguimiento_gavs.html

Lee el archivo de presupuesto de GAVs/Indirectos de Exportaciones y genera el
baseline que consume el tracker HTML:

La hoja `Formato` es la única fuente: trae los montos por línea y por mes y,
en las columnas O y P, la Denominación y el Texto de cabecera de cada línea —
que es el desglose fino que se asigna en los recursos de FOCO.

El control se concentra en los recursos de FOCO; el resto del presupuesto queda
como referencia de solo lectura.

Uso:
    python build_gav_tracker.py /ruta/al/F01CGPPTO_..._Consolidado.xlsx
"""

import hashlib
import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as ci
from openpyxl.utils import get_column_letter as gcl

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "gav_template.html"
OUT_HTML = HERE / "seguimiento_gavs.html"
OUT_JSON = HERE / "gav_baseline.json"

# --- recursos bajo control activo -------------------------------------------
FOCO = ["Gastos De Viaje", "Gastos de Feria y Eventos", "Refrigerios"]

# --- temporada ---------------------------------------------------------------
MONTHS = [f"2026_{m:02d}" for m in range(7, 13)] + [f"2027_{m:02d}" for m in range(1, 7)]
QTY_FIRST = "AD"   # (Q)2026_07 en la hoja Formato
USD_FIRST = "BN"   # (M)2026_07 en USD -> la fila 13 trae el TC presupuestado
SEASON = {"id": "T2627", "label": "Temporada 26/27"}

HEADER_ROW, TC_ROW, FIRST_DATA_ROW = 15, 13, 16
TC_FALLBACK = 3.36
NA = re.compile(r"#N/?A", re.I)


def val(row, letter):
    v = row[ci(letter) - 1]
    if isinstance(v, str):
        return v.strip() or None
    return v


def col_at(first, offset):
    return gcl(ci(first) + offset)


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def txt(v):
    s = "" if v is None else str(v).strip()
    return "" if s.lower() in ("none", "nan", "#n/a") else s


# ---------------------------------------------------------------- hoja Formato
def read_formato(wb):
    ws = wb["Formato"]
    head = list(ws.iter_rows(min_row=1, max_row=5, max_col=2, values_only=True))
    header = {
        "entidad": txt(head[0][1]),
        "presupuesto": txt(head[1][1]),
        "ambito": txt(head[2][1]),
        "gerencia": txt(head[3][1]),
    }

    rows = list(
        ws.iter_rows(min_row=TC_ROW, max_row=ws.max_row, max_col=ci("BZ"), values_only=True)
    )
    tc = [num(val(rows[0], col_at(USD_FIRST, i))) or TC_FALLBACK for i in range(len(MONTHS))]

    lines = []
    for r in rows[FIRST_DATA_ROW - TC_ROW :]:
        if not val(r, "E"):
            continue
        qty = [num(val(r, col_at(QTY_FIRST, i))) for i in range(len(MONTHS))]
        if not any(qty):
            continue

        recurso = val(r, "H")
        if not recurso or NA.search(str(recurso)):
            # el VLOOKUP del formato falla para algunas cuentas; AgrCtas lo resuelve
            recurso = val(r, "G") or "Sin clasificar"

        pu, moneda = num(val(r, "U")), (val(r, "X") or "USD")
        usd = [(q * pu / tc[i]) if moneda == "PEN" else q * pu for i, q in enumerate(qty)]

        lines.append(
            {
                "id": f"L{len(lines)+1:03d}",
                "ceco": val(r, "E"),
                "recurso": recurso,
                "claseCoste": txt(val(r, "J")),
                "denomClaseCoste": txt(val(r, "K")),
                "descripClaseCoste": txt(val(r, "L")),
                "material": txt(val(r, "M")),
                "textoMaterial": txt(val(r, "N")),
                "denominacion": txt(val(r, "O")),
                "texto": txt(val(r, "P")),
                "qty": [round(q, 4) for q in qty],
                "tipo": txt(val(r, "Q")),
                "umb": txt(val(r, "R")),
                "pu": round(pu, 6),
                "moneda": moneda,
                "usd": [round(x, 2) for x in usd],
            }
        )
    return header, [round(x, 4) for x in tc], lines


# ------------------------------------------------------------------- partidas
def make_parts(lines):
    """Explota cada línea del Formato en una partida por mes con cantidad.

    La Denominación (col. O) y el Texto de cabecera (col. P) viven en el
    Formato, no en BD_Opex: esa hoja es un extracto parcial y deja fuera casi
    todo el detalle de Gastos de Feria y Eventos.
    """
    parts = []
    for l in lines:
        if l["recurso"] not in FOCO:
            continue
        for i, monto in enumerate(l["usd"]):
            if abs(monto) < 0.005:
                continue
            parts.append(
                {
                    "id": f"P{len(parts)+1:03d}",
                    "recurso": l["recurso"],
                    "mi": i,
                    "claseCoste": l["claseCoste"],
                    "denomClaseCoste": l["denomClaseCoste"],
                    "descripClaseCoste": l["descripClaseCoste"] or l["denomClaseCoste"],
                    "denominacion": l["denominacion"],
                    "texto": l["texto"],
                    "monto": round(monto, 2),
                }
            )
    return parts


def extract(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    header, tc, lines = read_formato(wb)

    parts = make_parts(lines)

    # tope por recurso = el Formato manda
    topes = {r: 0.0 for r in FOCO}
    for l in lines:
        if l["recurso"] in topes:
            topes[l["recurso"]] += sum(l["usd"])
    topes = {k: round(v, 2) for k, v in topes.items()}

    # red de seguridad: si por algo las partidas no cuadran con el tope,
    # la diferencia entra como partida libre en vez de perderse
    for rec in FOCO:
        cubierto = sum(p["monto"] for p in parts if p["recurso"] == rec)
        resto = round(topes[rec] - cubierto, 2)
        if abs(resto) < 0.5:
            continue
        cand = [l for l in lines if l["recurso"] == rec]
        parts.append(
            {
                "id": f"P{len(parts)+1:03d}",
                "recurso": rec,
                "mi": 0,
                "claseCoste": cand[0]["claseCoste"] if cand else "",
                "denomClaseCoste": cand[0]["denomClaseCoste"] if cand else "",
                "descripClaseCoste": cand[0]["descripClaseCoste"] if cand else "",
                "denominacion": "",
                "texto": "Por asignar",
                "monto": resto,
            }
        )

    # catálogo de clases de coste por recurso, para los selectores de la app
    clases = {}
    for rec in FOCO:
        seen, out = set(), []
        for l in lines:
            if l["recurso"] != rec:
                continue
            k = (l["claseCoste"], l["denomClaseCoste"], l["descripClaseCoste"])
            if k in seen:
                continue
            seen.add(k)
            out.append(
                {
                    "cc": k[0],
                    "denomClaseCoste": k[1],
                    "descripClaseCoste": k[2] or k[1],  # la descripción es la etiqueta
                }
            )
        clases[rec] = out

    # huella del plan base: si cambia, la app avisa que hay que resembrar
    stamp = hashlib.sha1(
        json.dumps([topes, parts], ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()[:12]

    return {
        "header": header,
        "stamp": stamp,
        "season": {**SEASON, "months": MONTHS, "tc": tc},
        "foco": FOCO,
        "topes": topes,
        "clases": clases,
        "parts": parts,
        "lines": lines,
        "totalFoco": round(sum(topes.values()), 2),
        "totalPpto": round(sum(sum(l["usd"]) for l in lines), 2),
    }


def build(xlsx_path):
    p = extract(xlsx_path)
    OUT_JSON.write_text(json.dumps(p, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"{p['season']['label']}: {len(p['lines'])} líneas · USD {p['totalPpto']:,.2f}")
    print(f"En foco ({len(FOCO)} recursos): USD {p['totalFoco']:,.2f} · "
          f"{len(p['parts'])} partidas · huella {p['stamp']}")
    for r in FOCO:
        ps = [x for x in p["parts"] if x["recurso"] == r]
        den = len({x["denominacion"] for x in ps if x["denominacion"]})
        print(
            f"   {r:28} tope {p['topes'][r]:>9,.2f}  ·  {len(ps):3} partidas"
            f"  ·  {den:2} denominaciones  ·  suma {sum(x['monto'] for x in ps):>9,.2f}"
        )

    if TEMPLATE.exists():
        html = TEMPLATE.read_text(encoding="utf-8")
        html = html.replace("/*__BASELINE__*/", json.dumps(p, ensure_ascii=False))
        OUT_HTML.write_text(html, encoding="utf-8")
        print(f"\n{OUT_HTML.name} generado ({OUT_HTML.stat().st_size/1024:,.0f} KB)")
    else:
        print(f"\n[aviso] falta {TEMPLATE.name}; solo se generó {OUT_JSON.name}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    build(sys.argv[1])

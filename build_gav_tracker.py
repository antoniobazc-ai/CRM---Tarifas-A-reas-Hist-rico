#!/usr/bin/env python3
"""
Pipeline: F01CGPPTO (Formato) -> baseline JSON -> seguimiento_gavs.html

Lee la hoja `Formato` del archivo de presupuesto de GAVs/Indirectos y genera el
baseline congelado que consume el tracker HTML.

Uso:
    python build_gav_tracker.py /ruta/al/F01CGPPTO_..._Consolidado.xlsx
"""

import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as ci

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "gav_template.html"
OUT_HTML = HERE / "seguimiento_gavs.html"
OUT_JSON = HERE / "gav_baseline.json"

HEADER_ROW = 15          # fila de encabezados de la tabla `Data`
TC_ROW = 13              # fila con el TC presupuestado (columnas USD)
FIRST_DATA_ROW = 16

# Temporadas: (id, etiqueta, primera columna de cantidades, meses)
SEASONS = [
    {
        "id": "T2627",
        "label": "Temporada 26/27",
        "qty_first": "AD",              # (Q)2026_07
        "usd_first": "BN",              # (M)2026_07 en USD -> fila TC
        "months": [f"2026_{m:02d}" for m in range(7, 13)]
        + [f"2027_{m:02d}" for m in range(1, 7)],
    },
    {
        "id": "T2526",
        "label": "Cola 25/26 (feb–jun 26)",
        "qty_first": "Y",               # (Q)2026_02
        "usd_first": None,              # el formato no trae TC para estos meses
        "months": [f"2026_{m:02d}" for m in range(2, 7)],
    },
]

TC_FALLBACK = 3.36  # se usa donde el formato no trae TC presupuestado

# El VLOOKUP de `Recurso` falla para algunas cuentas que sí existen en el
# maestro de Control de Gestion. Se resuelve con AgrCtas (C&G).
NA = re.compile(r"#N/?A", re.I)


def val(row, letter):
    v = row[ci(letter) - 1]
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def col_at(first, offset):
    from openpyxl.utils import get_column_letter as gcl

    return gcl(ci(first) + offset)


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def extract(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb["Formato"]

    rows = list(
        ws.iter_rows(
            min_row=TC_ROW, max_row=ws.max_row, max_col=ci("BZ"), values_only=True
        )
    )
    tc_row = rows[0]
    meta_rows = rows[: FIRST_DATA_ROW - TC_ROW]
    data_rows = rows[FIRST_DATA_ROW - TC_ROW :]

    header = {
        "entidad": val(meta_rows[0], "B") if len(meta_rows) else None,
    }
    # cabecera del formato (B1:B5)
    head = list(ws.iter_rows(min_row=1, max_row=5, max_col=2, values_only=True))
    header = {
        "entidad": (head[0][1] or "").strip() if head[0][1] else "",
        "presupuesto": (head[1][1] or "").strip() if head[1][1] else "",
        "ambito": (head[2][1] or "").strip() if head[2][1] else "",
        "gerencia": (head[3][1] or "").strip() if head[3][1] else "",
    }

    seasons_out = []
    for s in SEASONS:
        n = len(s["months"])
        if s["usd_first"]:
            tc = [num(val(tc_row, col_at(s["usd_first"], i))) or TC_FALLBACK
                  for i in range(n)]
        else:
            tc = [TC_FALLBACK] * n

        lines = []
        for r in data_rows:
            ceco = val(r, "E")
            if not ceco:
                continue
            qty = [num(val(r, col_at(s["qty_first"], i))) for i in range(n)]
            if not any(qty):
                continue

            recurso = val(r, "H")
            if not recurso or NA.search(str(recurso)):
                recurso = val(r, "G") or "Sin clasificar"

            pu = num(val(r, "U"))
            moneda = val(r, "X") or "USD"
            usd = [
                (q * pu / tc[i]) if moneda == "PEN" else (q * pu)
                for i, q in enumerate(qty)
            ]

            lines.append(
                {
                    "id": f"{s['id']}-{len(lines)+1:03d}",
                    "ceco": ceco,
                    "oe": val(r, "F") or "",
                    "area": val(r, "C") or "",
                    "subarea": val(r, "D") or "",
                    "claseCostoGasto": val(r, "B") or "",
                    "recurso": recurso,
                    "agrCtas": val(r, "G") or "",
                    "claseCoste": str(val(r, "J") or ""),
                    "denomClaseCoste": val(r, "K") or "",
                    "descripClaseCoste": val(r, "L") or "",
                    "material": val(r, "M") or "",
                    "textoMaterial": val(r, "N") or "",
                    "tipo": val(r, "Q") or "",
                    "umb": val(r, "R") or "",
                    "pu": round(pu, 6),
                    "moneda": moneda,
                    "qty": qty,
                    "usd": [round(x, 2) for x in usd],
                }
            )

        seasons_out.append(
            {
                "id": s["id"],
                "label": s["label"],
                "months": s["months"],
                "tc": [round(x, 4) for x in tc],
                "lines": lines,
                "total": round(sum(sum(l["usd"]) for l in lines), 2),
            }
        )

    return {"header": header, "seasons": seasons_out}


def build(xlsx_path):
    payload = extract(xlsx_path)
    OUT_JSON.write_text(
        json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    for s in payload["seasons"]:
        recursos = sorted({l["recurso"] for l in s["lines"]})
        print(
            f"{s['label']:28} {len(s['lines']):3} líneas · "
            f"{len(recursos):2} recursos · USD {s['total']:,.2f}"
        )

    if TEMPLATE.exists():
        html = TEMPLATE.read_text(encoding="utf-8")
        html = html.replace(
            "/*__BASELINE__*/", json.dumps(payload, ensure_ascii=False)
        )
        OUT_HTML.write_text(html, encoding="utf-8")
        print(f"\n{OUT_HTML.name} generado ({OUT_HTML.stat().st_size/1024:,.0f} KB)")
    else:
        print(f"\n[aviso] falta {TEMPLATE.name}; solo se generó {OUT_JSON.name}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    build(sys.argv[1])

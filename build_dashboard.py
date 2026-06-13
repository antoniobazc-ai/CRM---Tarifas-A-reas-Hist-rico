"""
Pipeline: HISTORICO_TARIFAS.xlsx -> index.html (dashboard autónomo)

Uso:
    pip install pandas openpyxl
    npm install chart.js@4.4.4
    python build_dashboard.py HISTORICO_TARIFAS.xlsx

Reglas de limpieza (acordadas):
- Tarifa: se captura solo la tarifa base USD/kg (recargos fijos por AWB excluidos)
- Quarter asignado por ETD
- Se incluyen cancelados (filtrables en el dashboard)
- Excluidos: registros sin tarifa, outliers > 20 USD/kg, registros previos a 2023
- Skids: se toma solo el número principal (ej. "5(10 tripas)" -> 5)
"""
import sys, re, json
import pandas as pd

TEMPLATE = 'dashboard_template.html'
CHARTJS = 'node_modules/chart.js/dist/chart.umd.js'
OUT = '../index.html'

# Tarifario proyectado Sem 24-27 2026 (tarifa base, idéntico en las 4 semanas)
PROJ = [
 ("AMS","DIRECTO",2.41),  # opción única: tarifa del vuelo directo
 ("BKK","AFKL",3.70),("BKK","KOREAN",4.11),("BKK","CATHAY",3.60),
 ("BOG","LATAM",0.83),("BOG","AIR CARIBE",0.83),
 ("BOM","AFKL",3.70),("BOM","DHL",3.05),("BOM","MERCURY",3.55),("BOM","QATAR",3.50),
 ("DEL","AFKL",3.70),("DEL","DHL",2.90),("DEL","MERCURY",3.45),
 ("DXB","QATAR",4.60),("DXB","IBERIA",3.10),
 ("GRU","LATAM",2.60),("GRU","DELTA",2.50),
 ("HKG","AFKL",3.70),("HKG","CATHAY",2.75),("HKG","KOREAN",3.81),
 ("KUL","KOREAN",4.21),("KUL","AFKL",3.70),
 ("BLR","AFKL",4.05),
 ("LHR","AIR FRANCE",2.31),("LHR","AIR CANADA",2.05),("LHR","DELTA",1.95),
 ("MAD","IBERIA",1.05),("MAD","LATAM",1.05),("MAD","AIR EUROPA",1.10),("MAD","PLUS ULTRA",1.00),
 ("SIN","AFKL",3.70),("SIN","CATHAY",3.85),("SIN","KOREAN",4.21),
 ("TLV","MERCURY",5.75),
 ("JFK","LATAM",2.60),
]

def norm_air(a):
    s = str(a).strip().upper()
    return {'AF':'AIR FRANCE','LAN':'LATAM','CORGOJET':'CARGOJET'}.get(s, s)

def parse_tarifa(t):
    if pd.isna(t): return None
    m = re.search(r'(\d+\.?\d*)', str(t).replace(',',''))
    return float(m.group(1)) if m else None

def parse_skids(s):
    if pd.isna(s): return 0
    if isinstance(s,(int,float)): return int(s)
    m = re.match(r'\s*(\d+)', str(s))
    return int(m.group(1)) if m else 0

def main(xlsx):
    df = pd.read_excel(xlsx)
    df['Aerolinea_N'] = df['Aerolinea'].apply(norm_air)
    df['POD_N'] = df['POD'].astype(str).str.strip().str.upper().replace({'HK':'HKG'})
    df['Status_N'] = df['Status'].replace({'Confirmada':'Confirmado'})
    df['TarifaBase'] = df['Tarifa'].apply(parse_tarifa)
    df['SkidsN'] = df['Skids'].apply(parse_skids)
    df['ETD'] = pd.to_datetime(df['ETD'])
    df['Anio'] = df['ETD'].dt.year
    df['AnioQ'] = df['Anio'].astype(str) + '-Q' + df['ETD'].dt.quarter.astype(str)

    clean = df[(df['TarifaBase'].notna()) & (df['TarifaBase'] < 20) & (df['Anio'] >= 2023)].copy()
    recs = clean[['POD_N','ETD','Aerolinea_N','Agente','SkidsN','TarifaBase','Status_N','Anio','AnioQ']].copy()
    recs.columns = ['pod','etd','air','agente','skids','tarifa','status','anio','anioq']
    recs['etd'] = recs['etd'].dt.strftime('%Y-%m-%d')

    html = open(TEMPLATE, encoding='utf-8').read()
    html = html.replace('<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.4/chart.umd.min.js"></script>',
                        '<script>\n' + open(CHARTJS, encoding='utf-8').read() + '\n</script>')
    html = html.replace('const DATA = /*__DATA__*/[];', 'const DATA = ' + json.dumps(recs.to_dict(orient='records'), ensure_ascii=False, separators=(',',':')) + ';')
    html = html.replace('const PROJ = /*__PROJ__*/[];', 'const PROJ = ' + json.dumps([{"pod":p,"air":a,"tarifa":t} for p,a,t in PROJ], separators=(',',':')) + ';')
    open(OUT, 'w', encoding='utf-8').write(html)
    print(f'OK -> {OUT} | {len(recs)} registros')

if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'HISTORICO_TARIFAS.xlsx')

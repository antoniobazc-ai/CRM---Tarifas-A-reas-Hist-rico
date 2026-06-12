# Dashboard — Tarifas Aéreas Histórico

Dashboard interactivo y autónomo (un solo archivo HTML, sin dependencias externas) para el análisis de tarifas aéreas de exportación por destino, aerolínea y quarter, con comparativa del tarifario proyectado vs. histórico.

## Contenido

| Archivo | Descripción |
|---|---|
| `index.html` | Dashboard final autónomo (datos y librería de gráficos embebidos). Se abre en cualquier navegador, sin internet. |
| `src/dashboard_template.html` | Plantilla fuente con placeholders `/*__DATA__*/` y `/*__PROJ__*/` |
| `src/build_dashboard.py` | Pipeline completo: Excel crudo → limpieza → dashboard |
| `data/data.json` | Histórico limpio (1,028 registros, 2023–2026) |
| `data/proj.json` | Tarifario proyectado Sem 24–26 · 2026 (tarifa base) |

## Vistas del dashboard

- **Evolución por Quarter** — promedio, banda mín–máx y volumen; pestañas General / Por destino
- **Matriz Destino × Quarter** — heatmap, color relativo al promedio de cada ruta
- **Evolución Ruta-Aerolínea × Quarter** — heatmap con filas ruta·aerolínea (nunca mezcla destinos)
- **GAP proyectado vs histórico** — barras divergentes con benchmark conmutable (histórico total / solo Q2)
- **Close-up Sem 24–26** — tarifa proyectada por opción vs. benchmark histórico de la ruta-aerolínea
- Filtros en cascada: el selector de aerolíneas se limita a los carriers reales del destino elegido

## Reglas de datos

- Tarifa = **tarifa base USD/kg** (recargos fijos por AWB excluidos)
- Quarter asignado por **ETD**
- Cancelados incluidos (filtrables)
- Excluidos: registros sin tarifa, outliers >20 USD/kg, registros pre-2023
- Promedios simples por embarque (no ponderados por peso)

## Regenerar el dashboard con datos nuevos

```bash
pip install pandas openpyxl
npm install chart.js@4.4.4
cd src
python build_dashboard.py /ruta/a/HISTORICO_TARIFAS.xlsx
```

## ⚠️ Confidencialidad

Los datos de tarifas embebidos en `index.html` y `data/` son información comercial. **Verificar que el repositorio sea privado antes de subir.**

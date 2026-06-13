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

- **Evolución por Quarter** — promedio, banda mín–máx y volumen; pestañas General / Por aerolínea / Por destino, con las semanas proyectadas S24–26 extendiendo cada serie en línea punteada
- **Matriz Destino × Quarter** — heatmap, color relativo al promedio de cada ruta
- **Matriz Aerolínea × Quarter** — heatmap de carriers (filtrar por destino para comparar en una misma ruta)
- **Close-up histórico → forecast** — vista secuencial por destino: histórico por quarter (trazo sólido) y forecast S24·S25·S26 (trazo punteado sobre zona sombreada), con tabla de detalle colapsable
- Filtros por destino, aerolínea, año y status

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

## Acceso

El dashboard pide ID y contraseña al abrir. **Nota técnica honesta**: al ser un HTML estático, el login es disuasivo, no seguridad real — credenciales y datos están en el código fuente del archivo. La protección efectiva es mantener el repositorio privado.

## ⚠️ Confidencialidad

Los datos de tarifas embebidos en `index.html` y `data/` son información comercial. **Verificar que el repositorio sea privado antes de subir.**

# Seguimiento de GAVs — Exportaciones

Herramienta de control de ejecución del presupuesto de Indirectos/GAV
(`F01CGPPTO 26/27 · Exportaciones · CECO PEA2620102`).

Es **un solo archivo HTML** (`seguimiento_gavs.html`): se abre con doble clic en
cualquier navegador, sin internet, sin Excel abierto y sin instalar nada.

| Archivo | Qué es |
|---|---|
| `seguimiento_gavs.html` | La herramienta final, lista para usar |
| `gav_template.html` | Plantilla fuente (placeholder `/*__BASELINE__*/`) |
| `build_gav_tracker.py` | Excel del presupuesto → baseline → HTML |
| `gav_baseline.json` | Presupuesto congelado extraído del formato |

## La regla de control

El **tope duro es por Recurso** (`MD.Tpo Recurso`, las mismas 12 categorías que
usa Control de Gestión: Asesorías y Consultorias, Combustibles, Gastos De Viaje…).

- Los meses **se pueden mover libremente** dentro de un recurso.
- La suma de los 12 meses de un recurso **no puede pasar su presupuesto**: si se
  intenta, la celda se recorta al máximo disponible y avisa cuánto quedaba.
- **No hay traspasos entre recursos.** Cada categoría se controla contra su propio tope.
- Bajo cada recurso se abre el detalle en **Denom. clase de coste → Descripción
  clase de coste**, igual que la tabla dinámica del área.

## Las cuatro vistas

**Resumen y control** — KPIs (presupuesto, plan reprogramado, ejecutado,
comprometido, disponible, desvío vs. plan a la fecha, recursos excedidos),
plan vs. ejecución por mes, consumo acumulado contra el tope, y la tabla
jerárquica de tres niveles con semáforo por recurso.

Semáforo: 🔴 consumido > presupuesto · 🟡 del 90 % al 100 % · 🟢 bajo el 90 %.

**Plan mensual (reprogramar)** — grilla Recurso × 12 meses, editable celda a
celda. En rojo los meses donde lo gastado ya superó lo planificado. El botón
**«Traer saldo vencido al mes destino»** recoge, de los meses ya cerrados, el
plan que no llegó a ejecutarse ni comprometerse y lo reubica en el mes que
elijas — sin tocar el tope anual.

**Ejecución** — dos formas de cargar:

1. *Carga rápida*: fecha, recurso, clase de coste, concepto, monto, moneda,
   estado y documento. Enter guarda. Al cargar dice cuánto queda en ese recurso
   y avisa si lo dejó excedido. En PEN convierte a USD con el TC presupuestado
   del mes.
2. *Importar de SAP*: pegas la extracción de OPEX con su fila de encabezados
   (layout `BD_Opex`) y se mapea sola contra el presupuesto usando la clase de
   coste. Lo que caiga fuera de la temporada se ignora, y las clases de coste
   que no existan en el presupuesto quedan marcadas como «Sin clasificar» al
   final del control en vez de desaparecer.

Estados: **Comprometido** (pedido/OC, aún no contabilizado) y **Ejecutado**
(contabilizado). El disponible descuenta los dos, que es lo que evita pasarse
sin darse cuenta. Un clic alterna un movimiento de comprometido a ejecutado.

**Detalle del presupuesto** — las 60 líneas aprobadas con PU, UMB, moneda,
material y su reparto mensual.

## Datos y supuestos

- Control en **USD**. Las líneas en PEN se convierten con el **TC presupuestado
  de cada mes** que trae el propio formato (3,36 → 3,38 de jul-26 a jun-27).
- **T26/27 = jul-2026 a jun-2027**, 60 líneas, **USD 21.946**.
- El selector de temporada incluye además la **cola 25/26** (feb–jun 2026, 17
  líneas, USD 2.278). Ese tramo no trae TC en el formato: se convierte a 3,36 y
  queda anotado como supuesto.
- El `Recurso` de las líneas de teléfono/internet viene como `#N/A` en el Excel
  (la cuenta 4120000217 no está en el maestro para ese CECO); se resuelve con
  `AgrCtas (C&G)` → **Telefono y Comunicaciones**.
- Diferencias de menos de USD 0,5 se tratan como redondeo, no como desvío.

## Dónde se guarda

Todo lo que cargas (plan reprogramado + movimientos) queda en el
almacenamiento local del navegador de esa PC, y **no viaja dentro del HTML**.

- **Respaldo** descarga un `.json` con todo tu avance.
- **Restaurar** lo vuelve a cargar — así se pasa el seguimiento a otra PC o se
  comparte con el área.
- **CSV** exporta el control por recurso y el detalle de movimientos para pegar
  en Excel.

> Si se usa desde varias PCs, cada una lleva su propia copia: conviene definir
> un responsable de la carga y compartir el respaldo, o dejar el HTML y el JSON
> en una carpeta compartida.

## Regenerar con un presupuesto nuevo

```bash
pip install openpyxl
python build_gav_tracker.py /ruta/al/F01CGPPTO_..._Consolidado.xlsx
```

Lee la hoja `Formato` (encabezados en la fila 15, TC en la 13) y reescribe
`gav_baseline.json` y `seguimiento_gavs.html`. Los movimientos ya cargados
siguen en el navegador; si cambian los nombres de los recursos, hay que
reasignar los movimientos afectados.

## ⚠️ Confidencialidad

El presupuesto embebido es información interna. **Verificar que el repositorio
sea privado.**

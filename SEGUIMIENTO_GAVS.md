# Seguimiento de GAVs — Viajes, Ferias y Refrigerios

Herramienta de asignación y control de tres recursos del presupuesto de
Indirectos/GAV de Exportaciones (`F01CGPPTO 26/27 · CECO PEA2620102`):

| Recurso | Tope temporada 26/27 |
|---|---:|
| Gastos De Viaje | USD 5.000,00 |
| Gastos de Feria y Eventos | USD 1.710,19 |
| Refrigerios | USD 2.570,46 |
| **Total gestionado** | **USD 9.280,65** |

Los otros USD 12.665 del presupuesto quedan como referencia de solo lectura en
la pestaña *Resto del presupuesto*: acá no se asignan ni se controlan.

Es **un solo archivo HTML** (`seguimiento_gavs.html`): doble clic, sin internet,
sin Excel abierto, sin instalar nada.

| Archivo | Qué es |
|---|---|
| `seguimiento_gavs.html` | La herramienta final, lista para usar |
| `gav_template.html` | Plantilla fuente (placeholder `/*__BASELINE__*/`) |
| `build_gav_tracker.py` | Hoja `Formato` del presupuesto → baseline → HTML |
| `gav_baseline.json` | Presupuesto y partidas congelados |

## La regla de control

El **tope duro es por Recurso**. Dentro de un recurso hay libertad total; entre
recursos, ninguna.

- Dentro de un recurso puedes crear, dividir, duplicar, borrar y mover partidas;
  cambiarles el mes, la **Denominación**, el **Texto de cabecera**, la clase de
  coste y el monto.
- La suma de las partidas de un recurso **no puede pasar su tope**: si se
  intenta, el monto se recorta al disponible y avisa cuánto quedaba.
- **No hay traspasos entre recursos.** Cada uno se controla contra el suyo.

## Las cuatro vistas

### Asignación

Arriba, una **barra de meses** con los doce del ejercicio más «Todo»: cada uno
muestra lo asignado y una barrita de lo consumido, el mes en curso va marcado y
los cerrados atenuados. Un clic filtra toda la pantalla a ese mes — así ves
«para este mes tengo esto y esto» sin desplegar las 54 partidas, y sin perder de
vista cómo viene el resto del año. Al lado, un **buscador** por denominación,
texto de cabecera o clase de coste, que se combina con el mes.

Con un mes activo, cada recurso muestra un recuadro con *asignado del mes*,
*ejec.+comp. del mes* y *saldo del mes*, además de sus totales anuales. Las
partidas nuevas y el «mover en bloque» toman por defecto el mes que estás
mirando.

El botón **Panorama Mes × Recurso** abre la tabla de los doce meses contra los
tres recursos, con asignado, ejecutado y saldo de cada mes; se hace clic en un
mes para saltar a él.

Una tarjeta por recurso con tope, asignado, sin asignar, comprometido,
ejecutado, disponible, barra de consumo y semáforo; debajo, la tabla de partidas
editable celda a celda, y al pie el reparto por mes (también clicable).

Por fila: **÷** divide la partida en dos (preguntando cuánto separar), **⧉** la
duplica y **✕** la borra. Con los checkbox se seleccionan varias y la barra azul
permite **moverlas de mes**, duplicarlas o borrarlas en bloque. *Restaurar
original* devuelve ese recurso a las partidas del presupuesto.

Dividir es la herramienta para abrir una partida gruesa en las que hagan falta
sin pelear con el tope: separa un monto hacia una partida nueva y descuenta el
resto de la original.

### Control

KPIs, asignado vs. ejecutado por mes, consumo acumulado contra el tope, y la
tabla por recurso con **agrupación configurable**: los dos selectores «Abrir
por» arman la jerarquía con Denominación, Texto de cabecera, Clase de coste o
Mes, en el orden que quieras.

Semáforo por recurso: 🔴 consumido sobre el tope · 🟡 del 90 % al 100 % ·
🟢 bajo el 90 %.

### Ejecución

Dos formas de cargar:

1. *Carga rápida*: fecha, recurso, clase de coste, Denominación, Texto de
   cabecera, monto, moneda, estado y documento. Enter guarda. Al cargar dice
   cuánto queda en ese recurso y avisa si lo dejó excedido. En PEN convierte a
   USD con el TC presupuestado del mes. Denominación y Texto autocompletan con
   los valores que ya usaste.
2. *Importar de SAP*: pegas la extracción de OPEX con su fila de encabezados
   (layout `BD_Opex`) y se mapea sola por clase de coste, trayendo Denominación
   y Texto de cabecera. Lo que cae fuera de la temporada o fuera de los tres
   recursos se descarta y se reporta cuánto.

Estados: **Comprometido** (pedido/OC, aún no contabilizado) y **Ejecutado**
(contabilizado). El disponible descuenta los dos, que es lo que evita pasarse
sin darse cuenta. Un clic alterna un movimiento entre ambos.

### Resto del presupuesto

Los nueve recursos fuera de foco, por mes, solo como referencia.

## Datos y supuestos

- Control en **USD**. Las líneas en PEN se convierten con el **TC presupuestado
  de cada mes** del propio formato (3,36 → 3,38 de jul-26 a jun-27).
- **Temporada 26/27 = jul-2026 a jun-2027.** La cola feb–jun 2026 quedó fuera.
- **Todo sale de la hoja `Formato`**, que es la fuente autoritativa: los topes,
  el reparto mensual y —en las columnas **O (Denominación)** y **P (Texto de
  cabecera de documento)**— el detalle de cada línea. Cada línea se explota en
  una partida por mes con cantidad, así que las 63 partidas suman exactamente
  el tope de cada recurso. `BD_Opex` no se usa para el plan: es un extracto
  parcial que deja fuera casi todo el detalle de Gastos de Feria y Eventos.
- La **clase de coste** se muestra siempre por `Descrip.clases coste`, nunca por
  `Denom.clase de coste`: SAP trunca esa columna a 20 caracteres y se presta a
  confusión — «Gastos de Feria» en vez de *Gastos de Feria y Eventos*, «Consumos
  del Persona» en vez de *Consumos del Personal*, «Gs.: Vje./Vuelo naci» en vez
  de *Gastos de viajes nacionales*. Al importar de SAP la clase se normaliza
  contra el catálogo del presupuesto usando el número de cuenta.
- Diferencias de menos de USD 0,5 se tratan como redondeo, no como desvío.

## Dónde se guarda

Todo lo que cargas (partidas + movimientos) queda en el almacenamiento local del
navegador de esa PC, y **no viaja dentro del HTML**.

- **Respaldo** descarga un `.json` con todo tu avance.
- **Restaurar** lo vuelve a cargar — así se pasa el seguimiento a otra PC.
- **CSV** exporta partidas, movimientos y resumen por recurso para Excel.

> Si se usa desde varias PCs, cada una lleva su propia copia: conviene definir un
> responsable de la carga y compartir el respaldo, o dejar el HTML y el JSON en
> una carpeta compartida.

## Cuando se regenera el HTML

El plan vive en el navegador, así que abrir un `seguimiento_gavs.html` nuevo no
pisa por sí solo lo que ya tenías guardado. Cada build lleva una **huella** del
plan base y la app la compara al abrir:

- Si no habías tocado nada, las partidas se actualizan solas y avisa con un
  aviso al pie.
- Si ya habías asignado, aparece un banner: **Actualizar partidas** reemplaza el
  plan por el del presupuesto vigente (los movimientos cargados no se tocan), o
  **Seguir con las mías** lo deja como está y no vuelve a preguntar.
- *Restaurar original* en un recurso y el botón ⟲ también dejan el plan alineado
  con el presupuesto vigente.

## Regenerar con un presupuesto nuevo

```bash
pip install openpyxl
python build_gav_tracker.py /ruta/al/F01CGPPTO_..._Consolidado.xlsx
```

Lee la hoja `Formato` (encabezados en la fila 15, TC en la 13) y reescribe
`gav_baseline.json` y `seguimiento_gavs.html`. Para cambiar los recursos bajo
control, edita la lista `FOCO` al inicio de `build_gav_tracker.py`.

Lo ya cargado sigue en el navegador; si cambian los nombres de los recursos hay
que reasignar los movimientos afectados.

## ⚠️ Confidencialidad

El presupuesto embebido es información interna. **Verificar que el repositorio
sea privado.**

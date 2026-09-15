# Dashboard Sell Out — resumen del proyecto

Documento de traspaso. Estado al 15 de septiembre de 2026, semanas 27 a 37 cargadas.

---

## 1. Qué es

Un archivo HTML único, `Dashboard sell out.html` (79 MB), que reemplaza el reporte semanal de sell out de Grupo Depor. Se abre con doble clic, no necesita internet ni Excel, y vive en `C:\Users\Jonathan.Muñoz\Documents\Claude\Projects\sell out semanal`.

**Regla de oro del proyecto:** los números son los del Excel, tal cual. No se calcula, no se pondera, no se estima nada por fuera. Cuando un dato no viene en el reporte, la celda muestra un guión. Los totales, cuando no hay filtros puestos, se copian literal de la fila "Total general" de cada hoja.

---

## 2. Archivos fuente, uno por semana

Los cinco se dejan en la carpeta del proyecto con el número de semana en el nombre.

| Archivo | Alimenta | Estructura |
|---|---|---|
| `Reporte de zapatilla SellOut wNN.xlsx` | Pestaña CALZADO | 6 hojas, una fila por modelo, clientes como bloques de columnas |
| `Reporte de Ropa SellOut wNN.xlsx` | Pestaña ROPA | igual |
| `Reporte de Accesorios SellOut WNN.xlsx` | Pestaña ACC | igual, sin Semanas de Stock (usa Stk Final UN) |
| `Reporte por tienda wNN.xlsx` | Pestaña TIENDAS | hoja `tiendas`, pivot jerárquico con Nro Semana como campo de fila (trae todas las semanas) |
| `reporte por SKU por tienda wNN.xlsx` | Top-5 sucursales | una hoja, Cliente → Sucursal → Semana → Modelo, solo tiendas físicas |

Aparte, sin periodicidad fija:

- `Reservas WNN.xlsx` — tres hojas dinámicas (Bo02, KAMS, Regionales), modelo contra Suma de Available. Alimenta las columnas BO02, KAM y Regionales.
- `Excel_Macro.xlsx` — hoja Imagenes, mapea modelo a ID de Google Drive para las fotos.

**Las seis hojas de los reportes de sell out**, en este orden fijo:
1. `WNN(tienda + .com)` — la semana completa
2. `Tienda WNN` — solo locales
3. `WNN.com` — solo internet
4. `Acumulado (tienda + .com)` — desde la semana 27
5. `Acumulado tienda`
6. `Acumulado.com`

---

## 3. Cómo se construye

El dashboard no se edita a mano. Se compila desde una plantilla con marcadores.

```
template_v2.html  +  cat_data_*.json  →  build.py  →  Dashboard sell out.html
```

Todo vive en la carpeta de trabajo de Claude (`outputs`). El comando:

```
python3 build.py template_v2.html "cat_" "Dashboard sell out.html"
```

`build.py` reemplaza cada marcador `/*__DATA_XXX__*/` por el contenido de su JSON, incrusta los dos logos y escribe el archivo final.

### Los datos van como JSON, no como código

Este punto es crítico y ya rompió el archivo una vez. Los datos viajan en bloques `<script type="application/json" id="...">` que el dashboard lee con `JSON.parse`. **No** como literales dentro del JavaScript.

Cuando estaban como código, a los 84 MB el navegador no alcanzaba a compilar el script, no corría nada y se veía solo el cascarón: los filtros dibujados, sin pestañas, sin tabla, con la semana en "S29" (el texto que trae el HTML estático).

### Catálogo de modelos

Los diez atributos de cada modelo —descripción, status, duración, temporada, género, marca, full price, BO02, KAM, Regionales— se guardan **una sola vez** en `data_catalogo.json` (4.877 modelos) y se reparten a las filas al abrir el archivo.

Cuando un modelo cambia un atributo entre semanas, ese valor distinto queda escrito en la fila como excepción. Nada se aplana.

Ahorra 23 MB. Verificado: las 186.000+ filas reconstruyen idénticas a las originales, campo por campo, 0 diferencias.

---

## 4. Estructura de los datos

### Líneas de producto — `data_v2_NN.json`, `data_ropa_NN.json`, `data_acc_NN.json`

```json
{
  "W37(tienda + .com)": {
    "retailers": ["BELSPORT", "DEPOR", "FALABELLA", "..."],
    "rows": [{
      "m": "M9160C-001", "d": "CT AS CORE HI SLT", "st": "IN LINE",
      "du": "CORE", "te": "CORE", "g": "MUJER", "b": "CONVERSE",
      "dep": "CALZADO MUJER", "fp": 57990,
      "bo": 2137, "kam": 7427, "reg": 885,
      "r": { "DEPOR": [vtaUN, rotacion, margen, stockActual, semanasStock, vtaNeta] }
    }],
    "gt": { "DEPOR": [...], "_TOTAL": [...] }
  }
}
```

`gt` se copia verbatim de la fila "Total general". `_TOTAL` se deriva con la misma fórmula de la dinámica.

### TIENDAS — `data_tiendas.json` (26 MB)

Comprimido por diccionario. 280.221 filas.

```json
{ "dim": ["t","l","cli","m","te","g","du","st","s","w","d"],
  "vals": { "cli": ["DEPOR","FALABELLA","..."], "...": [] },
  "mets": ["Vta UN","Vta UN LY","Vta Neta","...16 métricas"],
  "rows": [[idx,idx,...,val,val,...]],
  "gt": [...] }
```

### Top-5 sucursales — `data_topsuc.json` (4,3 MB)

```json
{ "W": { "37": { "M9160C-001": { "D": [["107 - CONV.COSTANERA", 13, 633503], ...] } } },
  "ACC": { "M9160C-001": { "D": [...] } },
  "accWeeks": ["27","...","37"] }
```

Claves de cliente: `D` Depor, `F` Falabella, `R` Ripley, `P` Paris, `L` La Polar, `H` Hites, `G` GE2.

### Otros

- `data_catalogo.json` — atributos por modelo
- `data_reservas.json` — modelo a {bo, kam, reg}
- `images.json` — 4.267 fotos; la mayoría como ID de Drive, 10 embebidas como data URI
- `data_sucdet.json` (23 MB) — detalle por sucursal de calzado, semanas 32 a 34. **Está obsoleto**, solo alimenta el filtro Tienda de Calzado

---

## 5. Reglas de negocio

### Colores

| Indicador | Verde | Rojo |
|---|---|---|
| Semanas de stock | 16 o menos | sobre 16 |
| Rotación | 7% o más | bajo 7% |
| Margen | 25% o más | bajo 25% |

Los mismos colores bajan al Excel que se descarga.

### Fórmulas

Nunca se promedian porcentajes. Se suman los valores absolutos y recién ahí se divide.

- **Rotación** = Vta UN total ÷ Stk Inicial total
- **Margen** = Σ(margen × venta neta) ÷ venta neta total — ponderado, no promedio simple
- **Semanas de stock** = Stk Inicial total ÷ Vta UN total
- **Crecimiento %** = (actual − año pasado) ÷ año pasado
- **Stock Actual** nunca se suma entre semanas; se muestra el más reciente

En TIENDAS el Stk Inicial viene como columna y se suma directo. En las líneas de producto esa columna no existe, así que se reconstruye por fila: Vta UN ÷ rotación, o semanas × venta, y Stock Actual para las filas sin venta.

### Colores de cliente

```
DEPOR #ea580c · FALABELLA #0e9f6e · RIPLEY #a855f7 · PARIS #1a56db
HITES #14224a · LA POLAR #dc2626 · GE2 #0891b2
```

---

## 6. Qué tiene el dashboard

**Cinco pestañas:** CALZADO, ROPA, ACC, TIENDAS y GUÍA.

**Filtros** (todos de selección múltiple): Semana, Buscar (varios SKU separados por coma), Marca, Género, Status, Duración, Temporada, Cliente, Departamento Interno, Tienda (solo Calzado), y rangos de Rotación %, Margen % y Semanas de stock.

**Indicadores** arriba, una tarjeta por cliente cuando se filtra por varios, nunca sumados entre ellos. Con varias semanas seleccionadas, una tarjeta por semana.

**Panel por duración**: NEW, CARRYOVER, CORE, OLD con venta, venta neta, margen y semanas de stock.

**Cinco alertas** clicables: riesgo de quiebre, sobre-stock sobre 26 semanas, stock sin venta, margen negativo y top rotación.

**Top-5 de sucursales** al hacer clic en una fila. Siete columnas por cliente. En las hojas de semana muestra esa semana, en los acumulados el acumulado. Para productos IN LINE agrega dos etiquetas: primera venta y primer stock, **del cliente de la fila que se abrió**, no del global.

**TIENDAS**: fila de TOTAL fija arriba de los títulos, columna de ranking que se renumera con cada filtro y orden, y las tres primeras columnas fijas al desplazarse.

**Descarga a Excel** con el formato del reporte original: una fila por modelo, clientes como bloques de columnas, pesos chilenos, porcentajes reales, los mismos colores y las fotos incrustadas.

---

## 7. Lecciones aprendidas

**Relleno jerárquico de las dinámicas.** Excel escribe el valor solo en la primera fila del grupo y deja las demás en blanco. Hay que arrastrar el nivel superior y resetear los niveles más profundos ausentes. Sin esto se perdieron 4.867 filas de calzado en la semana 34 — un modelo con dos departamentos internos aparecía con el código vacío y su venta quedaba huérfana.

**Celdas vacías entre hojas.** A veces una hoja del mismo archivo trae la fila sin atributos y otra sí. Se rellena con el valor que el propio modelo tiene en las otras hojas. Fueron 2.938 celdas.

**Filtros que recalculaban de más.** Los filtros de atributos en Calzado recalculaban todo desde el detalle por sucursal en vez de filtrar las filas del reporte. Al marcar CONVERSE la tabla quedaba vacía. Ahora solo el filtro Tienda usa el detalle.

**Stock Actual es la foto de cierre.** Un producto que llegó y se vendió completo muestra cero. Por eso el "primer stock" usa stock de cierre más lo vendido, que es el stock al inicio de la semana.

**Dos dinámicas distintas dan números distintos.** Calzado sale del reporte de zapatilla y TIENDAS del reporte por tienda. Si un cliente no cuadra entre las dos pestañas, casi siempre es un filtro de página distinto en el Excel de origen: Grupo Sell(In-Out), Estado Semana, Periodo Semana o Estación. Pasó con 259 unidades de CLOSE OUT.

**Archivos que no llegan.** Subir dos veces un archivo con el mismo nombre y contenido idéntico lo deja fuera. Hay que renombrarlo, y tenerlo cerrado en Excel.

---

## 8. Estado actual

- Semanas 27 a 37, tres líneas, TIENDAS y top-5 al día
- 79 MB. Techo cómodo 90 MB, riesgoso sobre 130, muy probable que falle sobre 150
- Crece entre 4 y 5 MB por semana
- El límite real no es el disco sino la memoria: los datos se expanden entre 3 y 5 veces al abrirse

### Para bajar peso, en orden de rendimiento

1. **Sacar `data_sucdet.json`** (23 MB). Está obsoleto y su información está en `reporte por SKU por tienda`, que pesa 13 MB comprimido y cubre todas las semanas y las tres líneas. Se probó y funciona.
2. **Comprimir los datos dentro del archivo** con gzip y descomprimir al abrir. Reduce unas cinco veces el tamaño en disco, pero no la memoria.
3. **Ventana móvil de semanas**: dejar las últimas seis u ocho. Las hojas de Acumulado conservan el histórico completo.

### Pendiente conversado

Cargar el primer semestre, semanas 1 a 26. A la tasa actual son 125 MB más y no entra. La salida propuesta es un segundo archivo histórico, sin top-5 ni detalle por tienda, y exportar con Nro Semana como campo de fila para que sean 5 archivos en vez de 130.

---

## 9. Rutina semanal

1. Dejar los cinco Excel en la carpeta con el número de semana en el nombre
2. Pedir "agrega la semana NN"
3. Se agregan los marcadores de la semana a `template_v2.html` y a `build.py`, se extraen los JSON, se regenera el catálogo, se compila y se verifica que cada hoja cuadre con su "Total general"

Las reservas se actualizan cuando llega un archivo nuevo, y se aplican a todas las semanas a la vez porque son una foto del stock actual, no un dato semanal.

---

## 10. Nota sobre el entorno

Entre el 8 y el 15 de septiembre el entorno Linux donde corren los scripts estuvo caído por una actualización de Windows, y no se podían procesar Excel ni recompilar el archivo. Durante ese lapso sí se pudo parchear el HTML ya construido con las herramientas de archivo, que es como se aplicaron las columnas fijas de TIENDAS.

---

## 11. Pipeline semanal en GitHub (desde la semana 38)

Estado al 15 de septiembre de 2026. `build.py` y `template_v2.html` ya no existen (vivían en el entorno Linux de Claude). Los reemplaza un conversor en JavaScript que corre en el navegador.

### Cómo queda

```
cargar.html  →  js/convertir.js (en un Web Worker)  →  data/wNN/*.json  →  GitHub  →  el dashboard los descarga
```

- El HTML sigue trayendo embebidas las semanas 27–37. **No se recompila.**
- Cada semana nueva son 5 JSON en `data/wNN/` (`calzado`, `ropa`, `acc`, `tiendas`, `topsuc`, ~8 MB) más su registro en `data/semanas.json`.
- Al abrir, el dashboard lee `data/semanas.json`, descarga la última semana publicada y la deja como semana por defecto. Las demás semanas publicadas aparecen en el selector con ☁ y se descargan al elegirlas (también desde el filtro Semana de TIENDAS).
- Abierto con doble clic (`file://`) no hay `fetch`: se ve solo lo embebido.
- Los Excel no se suben (60 y 52 MB semanales; GitHub no lo aguanta). Se guardan en `Archivos sell out/WNN/`, carpeta ignorada por git.

### Qué hace el conversor (verificado campo por campo contra la W37 embebida)

- Líneas de producto: 6 hojas, bloques por cliente, `[vtaUN, rotación, margen, stock, semanas, vtaNeta]`, redondeo a 4 decimales con la regla del par de Python. Dinámica anidada Modelo > Departamento Interno (calzado): las filas sin modelo heredan el modelo y sus atributos. Atributos vacíos en una hoja se completan con los de otra hoja del mismo archivo. Precio 0 = sin precio (toma el del catálogo).
- `gt`: solo clientes con algún valor distinto de cero. `_TOTAL`: suma de UN, $ y stock; rotación y margen ponderados por venta **sobre el total completo** (un cliente sin indicador pesa 0); semanas de stock ponderadas por stock **solo de los clientes que traen semanas**, o stock/venta si ninguno trae.
- TIENDAS: relleno jerárquico con reset (el nivel más a la izquierda con valor es el que cambió; los de su derecha se toman de la fila aunque estén vacíos). Solo las filas de la semana. Diccionario propio por archivo; el dashboard lo traduce al suyo al cargar.
- Top‑5: solo filas con `Vta UN ≠ 0` (las devoluciones cuentan), orden venta UN desc y nombre de sucursal asc, venta neta redondeada por fila antes de sumar. El acumulado se recalcula completo desde el archivo (trae todas las semanas) y reemplaza al anterior.
- Las filas de las semanas nuevas llevan los atributos inline (no dependen del catálogo `CAT`, que sigue rellenando `bo/kam/reg`).

### Lo que no se actualiza con las semanas nuevas

- `DATA_SUCDET` / `DATA_SUCFILTRO` (filtro Tienda de Calzado): estaban obsoletos; para semanas 38+ ese filtro queda vacío.
- Reservas (`bo`, `kam`, `reg`) e imágenes (`IMG`): siguen siendo las embebidas. Un modelo nuevo sin foto muestra el ícono.
- Dashboard Steve Madden: sigue siendo un HTML aparte (`Reporte Steve madden WNN.xlsx` no entra en este pipeline).

### Diferencias conocidas con la W37 embebida

Al reconvertir la W37 desde los Excel de hoy: métricas idénticas en las 3 líneas y en TIENDAS (24.116 filas). En Top‑5, Falabella/Ripley/Paris difieren porque esos retailers restatan datos y el Excel actual ya no es el que se usó para construir el HTML. 25 filas de ropa `.com` traían atributos vacíos que el original resolvía con el catálogo (NEW) y el conversor con las otras hojas del mismo archivo (OLD).

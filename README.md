
Reportes semanales de sell out de Grupo Depor empaquetados como archivos HTML únicos. Se abren con doble clic en cualquier navegador; no necesitan internet, servidor ni Excel.

**Sitio publicado:** https://jmunozjm93-blip.github.io/pagina-web/

## Contenido

| Archivo | Descripción | Tamaño |
|---|---|---|
| `index.html` | Portada con enlaces a los dashboards | — |
| `Dashboard sell out.html` | Sell out semanal: Calzado, Ropa, Accesorios, Tiendas y Guía. Semanas 27–36 embebidas; 37+ desde `data/` | 74 MB |
| `Dashboard Steve Madden.html` | Sell out de la marca Steve Madden | 1,4 MB |
| `curvas.html` | Guía de curvas de tallas (Converse, Fila, Umbro); se muestra en la pestaña CURVA del dashboard sell out | 60 KB |
| `RESUMEN PROYECTO.md` | Documento de traspaso: fuentes, estructura de datos, reglas de negocio y rutina semanal | — |

## Qué tiene el dashboard principal

- **Cinco pestañas:** CALZADO, ROPA, ACC, TIENDAS y GUÍA.
- **Filtros múltiples:** semana, SKU, marca, género, status, duración, temporada, cliente, departamento, tienda y rangos de rotación, margen y semanas de stock.
- **Indicadores por cliente**, nunca sumados entre retailers.
- **Alertas clicables:** riesgo de quiebre, sobre-stock, stock sin venta, margen negativo y top rotación.
- **Top‑5 de sucursales** por modelo y cliente.
- **Descarga a Excel** con el formato del reporte original, colores y fotos.

## Reglas de negocio

| Indicador | Verde | Rojo |
|---|---|---|
| Semanas de stock | ≤ 16 | > 16 |
| Rotación | ≥ 7 % | < 7 % |
| Margen | ≥ 25 % | < 25 % |

Los porcentajes nunca se promedian: se suman los valores absolutos y luego se divide. Los totales sin filtros se copian literal de la fila "Total general" del Excel de origen.

## Cómo se construye

El dashboard no se edita a mano. Se compila desde una plantilla con marcadores y los datos en JSON:

```
template_v2.html + cat_data_*.json → build.py → Dashboard sell out.html
```

Los datos van incrustados como `<script type="application/json">` y se leen con `JSON.parse`, no como literales de JavaScript (a partir de ~84 MB el navegador no alcanza a compilar el script).

Detalle completo en [`RESUMEN PROYECTO.md`](RESUMEN%20PROYECTO.md).

## Rutina semanal (desde la semana 37)

1. Abrir **https://jmunozjm93-blip.github.io/pagina-web/cargar.html** (también hay un enlace "cargar semana" al pie de la portada).
2. Arrastrar los 6 Excel de la semana (5 de sell out + `Reporte Steve madden WNN.xlsx`), uno por uno. La app valida el nombre (`Reporte de zapatilla SellOut w38.xlsx`, etc.), la estructura de las hojas y que el contenido sea de esa semana; luego los convierte a JSON en el navegador.
3. Pulsar **Subir a GitHub** en cada uno. El JSON queda en `data/w38/` y la semana se registra en `data/semanas.json`.
4. En 1–2 minutos el dashboard muestra la semana nueva. Los Excel no se suben: se quedan en el PC (carpeta `Archivos sell out/`, ignorada por git).

Las semanas 27–36 siguen embebidas en el HTML; la 37 en adelante se descargan solo cuando el usuario las selecciona (la última publicada se carga al abrir).

| Archivo | Descripción |
|---|---|
| `cargar.html` | App de carga: valida, convierte y sube |
| `js/convertir.js` | Conversión Excel → JSON (misma lógica y mismos redondeos que el HTML original; verificada contra la semana 37) |
| `js/worker-convertir.js` | Web Worker que lee el Excel sin congelar la pantalla |
| `data/semanas.json` | Índice de semanas publicadas y qué archivos tiene cada una |
| `data/wNN/*.json` | `calzado`, `ropa`, `acc`, `tiendas`, `topsuc` de la semana NN (~8 MB en total) y `steve` (Steve Madden, ~110 KB) |

La app necesita un token de GitHub (fine-grained, permiso *Contents: Read and write* solo sobre este repositorio). Se pega una vez en la app y queda guardado en el navegador.

## Nota sobre tamaño

GitHub advierte con archivos sobre 50 MB y rechaza los que superan 100 MB. El dashboard principal crece 4–5 MB por semana; cuando se acerque al límite, las opciones son quitar `data_sucdet.json` (obsoleto, 23 MB), comprimir los datos con gzip o dejar una ventana móvil de semanas.

# Dashboards Sell Out

Reportes semanales de sell out de Grupo Depor empaquetados como archivos HTML únicos. Se abren con doble clic en cualquier navegador; no necesitan internet, servidor ni Excel.

**Sitio publicado:** https://jmunozjm93-blip.github.io/pagina-web/

## Contenido

| Archivo | Descripción | Tamaño |
|---|---|---|
| `index.html` | Portada con enlaces a los dashboards | — |
| `Dashboard sell out.html` | Sell out semanal: Calzado, Ropa, Accesorios, Tiendas y Guía. Semanas 27–37 | 79 MB |
| `Dashboard Steve Madden.html` | Sell out de la marca Steve Madden | 1,4 MB |
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

## Rutina semanal

1. Dejar los cinco Excel de la semana en la carpeta de trabajo con el número de semana en el nombre.
2. Extraer los JSON, regenerar el catálogo y compilar.
3. Verificar que cada hoja cuadre con su "Total general".
4. Hacer commit y push; GitHub Pages se actualiza solo.

## Nota sobre tamaño

GitHub advierte con archivos sobre 50 MB y rechaza los que superan 100 MB. El dashboard principal crece 4–5 MB por semana; cuando se acerque al límite, las opciones son quitar `data_sucdet.json` (obsoleto, 23 MB), comprimir los datos con gzip o dejar una ventana móvil de semanas.

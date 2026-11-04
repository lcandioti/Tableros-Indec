# Tablero INDEC

Página con gráficos y explicaciones de la economía argentina, armada con el [Catálogo API-INDEC](https://api.indec.gob.ar/) y la [API del BCRA](https://www.bcra.gob.ar/).

| Sección | Servicio de la API |
|---|---|
| EMAE | `estimador-mensual-de-actividad-economica-nivel-general`, `...-sectores` |
| Exportaciones de bienes | `comercio_exterior_exportacion_bienes` |
| Importaciones de bienes | `comercio_exterior_importacion_bienes` |
| IPC | `ipc_nacional` |
| Índice de salarios | `indice_salarios` |
| Costo de la construcción | `icc_precios_construccion` |
| Principales tasas (mercado de trabajo) | `mercado_trabajo_eph` |
| Precios mayoristas | `sipm_precios_mayoristas` |
| Dólar, tasas y reservas (BCRA) | `estadisticas/v4.0/monetarias`, variables 1, 4, 5, 7, 12, 14 y 44 |
| Tipo de cambio real (BCRA) | Excel [`ITCRMSerie.xlsx`](https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/ITCRMSerie.xlsx), hoja mensual |

Fuentes: INDEC y BCRA. Las variaciones, sumas de 12 meses y series reales son cálculos propios a partir de los datos publicados.

## Cómo funciona

- `index.html` es la página, sin dependencias: se abre con doble clic o se publica en GitHub Pages.
- `scripts/actualizar.ps1` baja los datos y genera `data/datos.js` y `data/datos.json`.
- `.github/workflows/actualizar.yml` corre ese script de lunes a viernes a las 17:30 de Argentina y guarda los cambios.

El token nunca llega a la página: solo lo usa el script para el INDEC, y la página lee el archivo de datos ya generado. La API del BCRA no pide token.

## Publicar en GitHub

1. Creá un repositorio público y subí todos estos archivos, incluida la carpeta `.github`.
2. En **Settings → Secrets and variables → Actions**, creá el secret `INDEC_TOKEN` con tu token.
3. En **Settings → Pages**, elegí **Deploy from a branch**, rama `main`, carpeta `/ (root)`.
4. En **Actions**, abrí **Actualizar datos INDEC** y tocá **Run workflow** para probarlo.

La página queda en `https://TU-USUARIO.github.io/NOMBRE-DEL-REPO/`.

## Actualizar a mano

```powershell
$env:INDEC_TOKEN = Read-Host 'Token INDEC'
.\scripts\actualizar.ps1
```

- `-Completo` vuelve a bajar comercio exterior desde 2016 (unos 6 minutos).
- `-SoloBCRA` actualiza solo dólar, tasas, reservas e ITCRM, sin token.
- `-Descubrir` lista los valores de cada columna en `data/_catalogo.txt`, útil si el INDEC cambia alguna categoría.

Si falta una serie imprescindible, el script se detiene sin tocar los datos publicados, y GitHub avisa del fallo por correo.

## Token

Genera uno nuevo en "Gestionar la cuenta" y actualizá el secret antes de que venza. La API está en beta, así que sus categorías pueden cambiar.

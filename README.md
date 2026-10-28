# Tablero de Pauta Publicitaria · GCBA

Seguimiento anual de la pauta publicitaria del Gobierno de la Ciudad (2017 en adelante): por rubro, medio y grupo de medios, en pesos corrientes y constantes, y contra la partida presupuestaria *Publicidad y propaganda*.

- `index.html` — el tablero (página única; lee `pauta_data.json` con fetch).
- `pauta_data.json` — agregados por año, mes y medio (lo genera `data/build.py`).
- `data/descargar.py` — baja todos los recursos del dataset [Pauta Publicitaria](https://data.buenosaires.gob.ar/dataset/pauta-publicitaria) de BA Data a `data/raw/` (saltea los que ya están).
- `data/build.py` — normaliza los archivos (formatos distintos por año), unifica nombres de medios, asigna grupos, deflacta con el IPCBA mensual y escribe `pauta_data.json` y `data/pauta.csv.gz` (una fila por orden).
- `data/build_presupuesto.py` — partida 3.6 *Publicidad y propaganda* de la ejecución presupuestaria, usando los lectores del repo del tablero de presupuesto (`../_repo_tablero_presupuesto`, o la ruta en `REPO_PRESUPUESTO`). Escribe `data/presupuesto_publicidad.json`.
- `data/IPCBA_empalmado.xlsx` — IPCBA mensual (copiado del repo del tablero de presupuesto).

## Actualizar (cuando BA Data publica un año o un rubro nuevo)

```
cd data
python descargar.py
python build_presupuesto.py
python build.py
```

`build.py` imprime la tabla año × rubro y los controles CSV vs XLSX; revisar que no aparezcan avisos `!`.
Si BA Data publica el año con un nombre de recurso nuevo, `rubro_de()` lo clasifica por palabras clave (tv abierta, cable, vía pública, web, radio, gráfica, vecinal, cine).
Para tener el IPC del último año completo, reemplazar `data/IPCBA_empalmado.xlsx` por la versión nueva.

## Trampas de la fuente

- Fechas: los XLSX de 2021 traen día y mes invertidos; los CSV mezclan `d/m/aaaa` y `m/d/aaaa`. El orden se deduce archivo por archivo.
- Importes: `1,234.56`, `1.234,56` y enteros redondeados (CSV 2021-2022). Formato deducido por archivo.
- XLSX vacíos en 2024 (gráfica, vecinales, cine): se usa el CSV.
- 2017 viene en un solo archivo con columna `tipo` (y nombres con caracteres basura, p. ej. `CLARÍ\x8dN`). TV 2018-2019 trae abierta y cable juntas.
- Vía pública 2019 no está publicada.
- En vía pública los "medios" son formatos (LED, grandes formatos), no empresas.

# CLAUDE.md

## Descripción

`ft_buffer_remaining_cap.py` hace scraping del sitio de First Trust para los ETFs de buffer de las series F y G (FJAN–FDEC, GJAN–GDEC), extrae métricas clave de cada uno y genera un Excel con tres hojas: datos completos, resumen serie F y resumen serie G.

## Cómo ejecutar

```bash
python ft_buffer_remaining_cap.py
```

No requiere archivos de entrada — los datos se obtienen directamente de `ftportfolios.com`. El Excel de salida se genera en la misma carpeta que el script.

## Dependencias

```bash
pip install requests beautifulsoup4 pandas openpyxl
```

## Arquitectura

El script tiene cuatro secciones claramente delimitadas:

### 1. CONFIGURACIÓN GENERAL
- `RUTA_EXCEL`: ruta de salida (relativa al script)
- `TICKERS_F` / `TICKERS_G`: listas de tickers de cada serie (JAN→DEC)
- `TITULOS_OBJETIVO`: métricas que aparecen en las hojas de resumen

### 2. SCRAPING (`scrapear_ticker`)
- HTTP GET a `ftportfolios.com/Retail/Etf/EtfSummary.aspx?Ticker=<ticker>`
- Parsea el bloque `div.ftmas.CEFFieldLabel`
- Extrae fecha "As of" y la convierte de MM/DD/YYYY → DD/MM/YYYY
- Devuelve `(fecha, list[dict])` con `Ticker`, `Titulo`, `Valor`
- Si el sitio cambia su HTML, solo esta sección necesita actualización

### 3. TRANSFORMACIÓN (`construir_dataframe`, `construir_resumen`)
- `construir_dataframe`: pivota la lista plana → DataFrame ancho (Titulo × Ticker)
- `valor_prioritario`: extrae el valor significativo de strings como `"X (Y)"` → `"Y"` o `"X/Y"` → `"Y"`
- `construir_resumen`: filtra solo `TITULOS_OBJETIVO`, pivota, renombra columnas y ordena JAN→DEC con `pd.Categorical`

### 4. EXCEL / FORMATO (`escribir_excel`, `aplicar_formato_excel`)
- Escribe tres hojas con `pd.ExcelWriter`
- Re-abre con `openpyxl` para inyectar fecha "As of", negrita en encabezados, tablas `TableStyleLight9`, fuente roja para negativos y botones de filtro ocultos

## Hojas del Excel de salida

| Hoja | Contenido |
|---|---|
| `Datos ETFs` | Todos los datos crudos de los 24 tickers (todas las métricas) |
| `Resumen F` | Métricas clave de FJAN–FDEC, con fecha "As of" |
| `Resumen G` | Métricas clave de GJAN–GDEC, con fecha "As of" |

## Parámetros modificables

| Variable | Línea | Descripción |
|---|---|---|
| `RUTA_EXCEL` | 16 | Ruta del archivo de salida |
| `TICKERS_F` | 20 | Tickers de la serie F |
| `TICKERS_G` | 21 | Tickers de la serie G |
| `TITULOS_OBJETIVO` | 25 | Métricas a incluir en el resumen |

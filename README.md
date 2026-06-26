# FT Buffer Remaining Cap

Herramienta de scraping automático para los ETFs buffer de **First Trust** (series F y G). Extrae métricas clave de cada fondo directamente desde el sitio oficial y genera un reporte Excel listo para análisis.

---

## ¿Qué hace este script?

Para cada uno de los 24 ETFs buffer de First Trust (FJAN–FDEC y GJAN–GDEC):

1. Consulta automáticamente la página del fondo en `ftportfolios.com`.
2. Extrae la fecha de actualización ("As of") y las métricas de buffer/cap.
3. Construye un DataFrame consolidado con todos los tickers.
4. Genera un Excel con tres hojas: datos completos + resumen por serie.

No requiere ningún archivo de entrada — todo viene del scraping en tiempo real.

---

## Requisitos previos

### Python
Versión mínima recomendada: **Python 3.8+**

```bash
python --version
```

### Librerías
```bash
pip install requests beautifulsoup4 pandas openpyxl
```

| Librería | Uso |
|---|---|
| `requests` | Peticiones HTTP al sitio de First Trust |
| `beautifulsoup4` | Parseo del HTML de cada página |
| `pandas` | Transformación y pivoteo de datos |
| `openpyxl` | Escritura y formato del Excel de salida |

---

## Cómo ejecutar

```bash
python ft_buffer_remaining_cap.py
```

Durante la ejecución verás el progreso en consola:

```
Procesando FJAN...
Procesando FFEB...
Procesando FMAR...
...
Procesando GDEC...
Proceso terminado | Fecha: As of 25/06/2026
```

Al terminar se genera `FT buffer remaining cap.xlsx` en la misma carpeta que el script.

---

## ETFs incluidos

### Serie F (Defined Outcome — Buffer)
| Ticker | Mes |
|--------|-----|
| FJAN | Enero |
| FFEB | Febrero |
| FMAR | Marzo |
| FAPR | Abril |
| FMAY | Mayo |
| FJUN | Junio |
| FJUL | Julio |
| FAUG | Agosto |
| FSEP | Septiembre |
| FOCT | Octubre |
| FNOV | Noviembre |
| FDEC | Diciembre |

### Serie G (Defined Outcome — Buffer)
Misma estructura mensual con prefijo G: GJAN, GFEB, GMAR... GDEC.

---

## Métricas extraídas (hojas de resumen)

| Métrica | Descripción |
|---|---|
| **Remaining Buffer (Net)** | Cuánto buffer queda antes de absorber pérdidas |
| **Remaining Cap (Net)** | Cuánto upside queda disponible antes de llegar al cap |
| **Fund Return** | Retorno actual del fondo |
| **Reference Asset Return** | Retorno del activo de referencia (S&P 500) |
| **Reference Asset Return to Realize Cap** | Retorno que necesita el benchmark para alcanzar el cap |
| **Downside Before Buffer (Net)** | Caída que absorbe el inversionista antes de activarse el buffer |

---

## Archivo de salida: `FT buffer remaining cap.xlsx`

Se genera automáticamente en la misma carpeta. Contiene **tres hojas**:

### Hoja `Datos ETFs`
Todos los datos crudos extraídos del sitio para los 24 tickers. Incluye todas las métricas disponibles, no solo las del resumen.

### Hoja `Resumen F`
- Fecha "As of" en la celda A1 (negrita).
- Una fila por cada ETF de la serie F (ordenado FJAN → FDEC).
- Columnas: Ticker + 6 métricas clave.

### Hoja `Resumen G`
- Misma estructura que Resumen F para la serie G (GJAN → GDEC).

### Formato visual
- Tablas con estilo `TableStyleLight9` con filas alternadas.
- Valores negativos en **rojo**.
- Botones de filtro ocultos (tabla limpia visualmente).
- Fecha "As of" en encabezado de hojas de resumen.

---

## Personalización

### Agregar un nuevo ticker

```python
TICKERS_F = ["FJAN", ..., "FDEC", "FNUEVO"]
```

### Cambiar las métricas del resumen

```python
TITULOS_OBJETIVO = [
    "Remaining Buffer (Net)",
    "Remaining Cap (Net)",
    # agregar o quitar métricas aquí
    "Buffer Annualized"
]
```

### Cambiar el nombre del archivo de salida

```python
RUTA_EXCEL = Path(__file__).parent / "mi_reporte_custom.xlsx"
```

---

## Errores comunes

| Error | Causa probable | Solución |
|---|---|---|
| `ConnectionError` / `Timeout` | Sin conexión a internet o el sitio está caído | Verifica conexión e intenta de nuevo |
| Hoja con datos vacíos | El HTML del sitio cambió de estructura | Revisar y actualizar `scrapear_ticker()` |
| `PermissionError` al guardar Excel | El archivo está abierto en Excel | Cierra el archivo y vuelve a ejecutar |
| Tabla Excel con nombre duplicado | El archivo de salida ya existe con tablas previas | Elimina el archivo de salida y vuelve a ejecutar |
| Fecha `None` en el Excel | El sitio no devolvió la etiqueta "As of" | Revisar `extraer_fecha_asof()` |

---

## Fuente de datos

Los datos se obtienen de:
```
https://www.ftportfolios.com/Retail/Etf/EtfSummary.aspx?Ticker=<TICKER>
```

Si First Trust modifica la estructura HTML de su sitio, solo la función `scrapear_ticker()` necesita actualización.

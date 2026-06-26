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

## Automatización con GitHub Actions

El script puede correr automáticamente en la nube usando GitHub Actions. Al terminar, el Excel actualizado se guarda directamente en el repositorio — sin necesidad de correrlo manualmente.

### Cómo configurarlo

#### Paso 1 — Crear el archivo del workflow

Dentro del repositorio crea la carpeta `.github/workflows/` y dentro el archivo `update_report.yml` con este contenido:

```yaml
name: Actualizar FT Buffer Report

on:
  schedule:
    - cron: "0 14 * * 1-5"   # Lunes a viernes a las 14:00 UTC (8:00 AM CST / 9:00 AM CDT)
  workflow_dispatch:           # También se puede lanzar manualmente desde GitHub

jobs:
  run-script:
    runs-on: ubuntu-latest

    permissions:
      contents: write          # Necesario para hacer commit del Excel generado

    steps:
      - name: Checkout del repositorio
        uses: actions/checkout@v4

      - name: Configurar Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Instalar dependencias
        run: pip install requests beautifulsoup4 pandas openpyxl

      - name: Ejecutar script
        run: python ft_buffer_remaining_cap.py

      - name: Guardar Excel en el repositorio
        run: |
          git config user.name  "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add "FT buffer remaining cap.xlsx"
          git diff --cached --quiet || git commit -m "chore: actualizar reporte FT buffer [$(date +'%Y-%m-%d')]"
          git push
```

#### Paso 2 — Subir el workflow al repositorio

```bash
git add .github/workflows/update_report.yml
git commit -m "Add GitHub Actions workflow for automated report"
git push
```

#### Paso 3 — Verificar que funciona

1. Ve a tu repositorio en GitHub.
2. Haz clic en la pestaña **Actions**.
3. Verás el workflow `Actualizar FT Buffer Report`.
4. Para probarlo manualmente: haz clic en el workflow → **Run workflow** → **Run workflow**.

---

### Ajustar el horario (`cron`)

El campo `cron` usa la sintaxis estándar de Unix en **UTC**. México está en **UTC-6** (CST) o **UTC-5** (CDT en verano).

| Quieres que corra a... | Zona | Cron (UTC) |
|---|---|---|
| 8:00 AM CST (invierno) | UTC-6 | `0 14 * * 1-5` |
| 8:00 AM CDT (verano) | UTC-5 | `0 13 * * 1-5` |
| 6:00 AM CST todos los días | UTC-6 | `0 12 * * *` |
| Cada hora en días hábiles | — | `0 * * * 1-5` |

Formato del cron: `minuto hora día-del-mes mes día-de-la-semana`

---

### Descargar el Excel generado

Una vez que el workflow corre exitosamente, el archivo `FT buffer remaining cap.xlsx` aparece actualizado directamente en el repositorio. Puedes descargarlo desde GitHub o hacer `git pull` en tu máquina local.

---

### Nota sobre costos

GitHub Actions es **gratuito** para repositorios públicos sin límite de minutos. Para repositorios privados, incluye **2,000 minutos gratis al mes** — este script tarda aproximadamente 1–2 minutos por ejecución, por lo que con una ejecución diaria (≈ 22 días hábiles) usa alrededor de 44 minutos al mes.

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

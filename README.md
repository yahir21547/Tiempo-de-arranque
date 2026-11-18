# Motor Start-up Studio (Streamlit)

Nueva interfaz web para la calculadora original de tiempo de arranque de motores.

## Despliegue en Streamlit Community Cloud

1. Sube `streamlit_app.py` y `requirements.txt` a la rama `main` del repositorio de GitHub.
2. Mantén `base_motores.xlsx` en la raíz para conservar las curvas existentes. Si tienes `PMD NEMA FS 56 to FS 500 - V47-R1.xlsx`, súbelo a la misma carpeta para convertirlo en el catálogo predeterminado.
3. En Streamlit Community Cloud: **Create app** → selecciona el repositorio y la rama `main` → entrypoint `streamlit_app.py`.
4. También puedes cargar un Excel nuevo desde la barra lateral, sin necesidad de desplegar de nuevo. El archivo subido solo existe en la sesión en curso.

Para uso local: `pip install -r requirements.txt` seguido de `streamlit run streamlit_app.py`.

## Tratamiento de curvas

La aplicación usa el factor en la columna BT del catálogo activo dividido entre 100, conforme al código original. Cuando el factor no existe, intenta localizarlo en `base_motores.xlsx` por coincidencia exacta del MLFB. No calcula promedios nuevos, no infiere factores ausentes y deja el campo editable si no hay curva.

**La estructura exacta del V47-R1 no se ha verificado:** selecciona en la barra lateral el nombre de hoja, fila de encabezados y letras de columnas apropiadas si la versión nueva no conserva el esquema original.

## Importante: unidades

La interfaz muestra por separado las entradas y sus valores normalizados, también en el PDF. Se conservaron los factores de conversión del programa original: `lb·ft² / 23.731 → kg·m²` y `lb·ft / 7.233 → kgf·m`. Se mantuvo asimismo el algoritmo de `GD² = 4 × J`, los redondeos intermedios y la ecuación de tiempo de arranque. Es necesario confirmar con una referencia de ingeniería qué representa exactamente el campo `WR²` (momento de inercia J o convención WR²/GD²).

## Estado

El código se ha comprobado sintácticamente y las principales fórmulas pueden verificarse sin interfaz. La aplicación visual completa y los datos contenidos en los Excel no han sido validados aquí.

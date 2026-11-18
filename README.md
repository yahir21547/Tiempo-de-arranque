# Calculadora de Tiempo de Arranque · Streamlit V6

## Arranque sin cargar archivos

Coloca `streamlit_app.py`, `requirements.txt` y estos Excel **en la raíz del repositorio nuevo** (no dentro de carpetas):

- `PMD NEMA FS 56 to FS 500 - V47-R1.xlsx` (catálogo predeterminado)
- `base_motores.xlsx` (respaldo de curvas)

Los puedes descargar desde el repositorio anterior:

- https://github.com/yahir21547/CALCULO-DE-TIEMPO-DE-ARRANQUE-STREAMLIT/blob/main/PMD%20NEMA%20FS%2056%20to%20FS%20500%20-%20V47-R1.xlsx
- https://github.com/yahir21547/CALCULO-DE-TIEMPO-DE-ARRANQUE-STREAMLIT/blob/main/base_motores.xlsx

La aplicación selecciona automáticamente el V47-R1 y solo ofrece la carga de archivos nuevos como función opcional. Si los Excel todavía no se han copiado, intentará leerlos del repositorio anterior, pero no conviene depender de internet para operar la app. **Este ZIP no incluye los dos archivos binarios Excel**.

En Streamlit Community Cloud usa `streamlit_app.py` como punto de entrada.

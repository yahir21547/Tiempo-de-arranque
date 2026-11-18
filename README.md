# Calculadora de arranque de motores · Streamlit V4

Archivos para la raíz de un repositorio GitHub nuevo:
- `streamlit_app.py`: aplicación principal.
- `requirements.txt`: dependencias.
- Excel recomendados en la misma raíz: `PMD NEMA FS 56 to FS 500 - V47-R1.xlsx` y `base_motores.xlsx`.

## Cambios V4
- El catálogo ya NO se convierte completo en un DataFrame durante el arranque.
- Búsqueda MLFB por lectura de filas con `openpyxl` en modo `read_only`, solo cuando se pulsa BUSCAR.
- Botón para detectar hojas cuando el catálogo no usa la hoja predeterminada.
- Se evita sobrescribir `session_state.factor_mmv` luego de crear su widget.
- Una excepción Python se muestra dentro de la aplicación para diagnosticarla.

## Despliegue
Sube estos archivos a la raíz y selecciona `streamlit_app.py` en Streamlit Cloud. Agrega ambos Excel a este mismo repositorio para evitar depender de descargas de red. En caso de nuevo error, copia el detalle de excepción y los logs de Streamlit Cloud.

# Cálculo de Tiempo de Arranque de Motores · Streamlit

## Publicar en un repositorio nuevo

1. Descomprime el paquete, coloca **los archivos** en la raíz de tu nuevo repositorio de GitHub (no subas el ZIP sin descomprimir).
2. Copia también desde el repositorio original estos archivos (recomendado, para que nunca dependa de Internet):
   - `PMD NEMA FS 56 to FS 500 - V47-R1.xlsx`
   - `base_motores.xlsx` (respaldo de curvas)
   - `logo.png`, `motor_header.png`, `fondo_motor.png` (opcionales)
3. Confirma que `streamlit_app.py` y `requirements.txt` quedaron en la raíz.
4. En Streamlit Community Cloud, despliega `streamlit_app.py` y selecciona **Reboot app** después de cada actualización.
5. Si todavía falla, abre **Manage app > Logs** y revisa el primer `Traceback` o error de dependencias.

## Importante

Los Excel pueden precargarse si están en el repositorio. Si no están, la app intenta recuperarlos del repositorio público original con un tiempo de espera acotado. También puedes subir y actualizar el Excel desde la interfaz. Si no hay conexión, aparece una advertencia en pantalla en lugar de impedir el inicio.

La evaluación usa el modelo de cálculo heredado de Tkinter, con comprobaciones de torque acelerante y conversiones explícitas. Verifica los resultados contra documentación técnica antes de usarlos para seleccionar equipos.

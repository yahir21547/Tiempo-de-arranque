# Calculadora de tiempo de arranque de motores | Streamlit

## Nuevo repositorio independiente (despliegue sencillo)

Sube **todos los archivos de esta carpeta** a la raíz de un nuevo repositorio GitHub. En Streamlit Community Cloud elige `streamlit_app.py` como archivo principal y `main` como rama.

La aplicación intenta descargar al iniciar, si faltan, los catálogos `PMD NEMA FS 56 to FS 500 - V47-R1.xlsx`, `base_motores.xlsx`, `promedios.xlsx` y los gráficos originales del repositorio público de origen. **Esto requiere que dicho repositorio siga siendo público y accesible**. Una vez descargados, se usan como archivos locales del despliegue. Para independizarlo por completo, descarga o copia esos archivos del repositorio anterior y súbelos también al nuevo repositorio.

**Importante:** el catálogo V47-R1 es el predeterminado cuando está disponible; si falla su descarga, se usa `base_motores.xlsx`. El respaldo del factor de curvas funciona por coincidencia MLFB. La carga de archivos Excel desde la interfaz es por sesión, no actualiza GitHub permanentemente.

## Archivos para independencia completa

- `PMD NEMA FS 56 to FS 500 - V47-R1.xlsx` (catálogo reciente)
- `base_motores.xlsx` (respaldo de curvas)
- `promedios.xlsx` (datos auxiliares originales)
- `logo.png`, `motor_header.png`, `fondo_motor.png`, `logo.ico` (imágenes originales)

Se conservaron el formulario por etapas, los campos, la exportación PDF, el comparador MLFB y las fórmulas del código original. Validar unidades y resultados contra documentación técnica antes de su uso operativo.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

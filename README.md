# AUTOGOOD - Informe técnico vehicular

Aplicación Streamlit para que técnicos llenen informes desde celular, tablet o computadora y generen PDFs de servicio AUTOGOOD.

## Requisitos

- Python 3.10 o superior
- `LOGOO.png` en la raíz del proyecto
- Dependencias en `requirements.txt`

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Para red local del taller:

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

En Windows también puedes usar:

```bat
iniciar_autogood.bat
```

## Carpetas y archivos generados

La app crea automáticamente si no existen:

- `INFORMES_GENERADOS`
- `IMAGENES_SUBIDAS`

También puede generar:

- `historial_autogood.xlsx`
- `ultimo_informe_autogood.txt`

## Despliegue en Streamlit Community Cloud

1. Sube este proyecto a un repositorio de GitHub.
2. Verifica que estén incluidos:
   - `app.py`
   - `requirements.txt`
   - `LOGOO.png`
   - `.streamlit/config.toml`
3. En Streamlit Community Cloud, crea una nueva app desde el repositorio.
4. Selecciona `app.py` como archivo principal.
5. Publica la app y comparte el enlace público con los técnicos.

Nota: en Streamlit Community Cloud los archivos generados pueden no ser persistentes entre reinicios. Para historial permanente en producción, usar una base de datos o almacenamiento externo.

## Despliegue en Render

1. Sube este proyecto a GitHub.
2. Crea un Web Service en Render.
3. Usa como build command:

```bash
pip install -r requirements.txt
```

4. Usa como start command:

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port $PORT
```

5. Publica el servicio y comparte la URL pública.

Nota: si Render reinicia el servicio, los archivos locales pueden perderse según el tipo de disco configurado. Para conservar historial/PDFs, configurar disco persistente o almacenamiento externo.

## Uso desde celular o tablet

Cuando la app esté desplegada, los técnicos solo deben abrir el enlace público en el navegador del celular o tablet. No necesitan estar conectados al mismo WiFi.


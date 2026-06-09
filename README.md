# Dashboard CXC / BEESCARE - Streamlit

App ejecutiva para capacitacion JDV y SPV sobre gestion CXC / BEESCARE del Manual GALAXIA Nivel 1.

## Contenido

- KPIs reales de tickets CXC.
- ON TIME mensual y acumulado.
- Adopcion CXC mensual.
- Tickets cerrados dentro/fuera SLA.
- Pendientes dentro SLA, vencidos y riesgo +10 dias.
- Reclamos que corresponden / no corresponden.
- Clientes criticos y planes de accion.
- Checklist de auditoria 100%.
- Descarga de evidencia en CSV, Excel y PPT.

## Archivos principales

- `app.py`: aplicacion Streamlit.
- `requirements.txt`: dependencias para Streamlit Cloud.
- `data/dashboard-data.json`: base limpia usada por la app.
- `data/CXC_BEESCARE_DelValle_analisis.xlsx`: Excel de evidencia.
- `data/Capacitacion_JDV_SPV_CXC_BEESCARE_GALAXIA_DelValle.pptx`: presentacion de capacitacion.

## Ejecutar localmente

```powershell
cd streamlit-cxc-app
python -m pip install -r requirements.txt
streamlit run app.py
```

## Subir a GitHub

1. Crear un repositorio nuevo en GitHub.
2. Subir el contenido de la carpeta `streamlit-cxc-app`.
3. Confirmar que en la raiz del repo queden `app.py`, `requirements.txt` y la carpeta `data`.

## Publicar en Streamlit Cloud

1. Entrar a https://share.streamlit.io/.
2. Seleccionar el repositorio de GitHub.
3. En `Main file path`, escribir:

```text
app.py
```

4. Deploy.

## Actualizar datos

Cuando se regenere la informacion CXC:

1. Reemplazar `data/dashboard-data.json`.
2. Opcionalmente reemplazar el Excel y la PPT dentro de `data`.
3. Subir cambios a GitHub.
4. Streamlit Cloud actualiza la app automaticamente.

## Nota sobre planes de accion

La tabla de planes de accion es editable en pantalla y permite descargar CSV como evidencia.
Streamlit Cloud no guarda cambios multiusuario de forma permanente sin una base externa.
Para persistencia compartida real se recomienda conectar Google Sheets, Supabase o una base similar.

## Firma

by QπU

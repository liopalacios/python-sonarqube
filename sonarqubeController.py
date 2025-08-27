from fastapi import FastAPI, Query, Path
from fastapi.responses import FileResponse, JSONResponse
from reporte_sonarqube import obtener_datos_sonarqube, generar_reporte_excel, obtener_condiciones_fallidas, obtener_code_smells, obtener_reglas_afectadas
import tempfile
from typing import List
from datetime import datetime
import os

app = FastAPI()

COMPONENTES = [
        "mr-api-neg-gestion-notificacion", "mr-api-neg-gestion-pago", "mr-api-neg-gestion-solicitud",
        "mr-api-neg-gestion-usuarios", "mr-api-neg-portal", "mr-api-neg-proceso-carga", "mr-api-sp-buzon",
        "mr-api-sp-control-pago", "mr-api-sp-convivencia-legado", "mr-api-sp-filenet",
        "mr-api-sp-gestion-notificacion", "mr-api-sp-gestion-pago", "mr-api-sp-gestion-solicitud",
        "mr-api-sp-gestion-solicitud-leg", "mr-api-sp-gestor-procedimientos", "mr-api-sp-interoperabilidad",
        "mr-api-sp-maestros", "mr-api-sp-pasarela-pago", "mr-api-sp-portal", "mr-api-sp-proceso-carga",
        "mr-api-sp-sunat", "mr-api-sp-usuarios", "mr-api-neg-evaluacion-tramite", "mr-api-neg-resolucion-tramite",
        "mr-api-sp-evaluacion-tramite", "mr-api-sp-resolucion-tramite", "mr-api-sp-solicitud-norel",
        "mr-api-sp-transmisiones", "mr-ui"
    ]

# Carpeta temporal donde se guardarán los archivos Excel generados
TEMP_FOLDER = "reportes_temp"
os.makedirs(TEMP_FOLDER, exist_ok=True)


@app.get("/")
def root():
    return {"message": "API de Reportes SonarQube activa"}

@app.get("/reporte")
def generar_reporte():
    datos = [obtener_datos_sonarqube(comp) for comp in COMPONENTES]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"reporte_sonarqube_{timestamp}.xlsx"
    ruta_archivo = os.path.join(TEMP_FOLDER, nombre_archivo)

    generar_reporte_excel(datos, COMPONENTES, ruta_archivo)
    #os.rename(nombre_archivo, ruta_excel)  # Si tu función guarda por defecto con este nombre

    return FileResponse(
        path=ruta_archivo,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="reporte_sonarqube.xlsx"
    )

@app.get("/metricas")
def metricas_componente(componente: str = Query(..., description="Nombre del componente en SonarQube")):
    try:
        datos = obtener_datos_sonarqube(componente)
        return JSONResponse(content=datos)
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": "Error al obtener métricas", "detalle": str(e)}
        )
    
@app.get("/sonarqube/condiciones/{componente}")
def condiciones_fallidas(componente: str = Path(..., description="Nombre del componente")):
    return obtener_condiciones_fallidas(componente)


@app.get("/sonarqube/code-smells/{componente}")
def condiciones_fallidas(componente: str = Path(..., description="Nombre del componente")):
    return obtener_code_smells(componente)

@app.get("/sonarqube/reglas-afectadas/{componente}")
def reglas_afectadas_por_componente(componente):
    return obtener_reglas_afectadas(componente)

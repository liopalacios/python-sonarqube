from flask import Flask, jsonify, request
import requests
from fastapi.responses import JSONResponse
from fastapi import HTTPException
import math 
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import math
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
SONARQUBE_URL_ISSUES = "http://172.23.12.34:9000/api/issues/search"
SONARQUBE_URL_METRICS = "http://172.23.12.34:9000/api/measures/component"
# Configuración de SonarQube
SONARQUBE_URL = "http://172.23.12.34:9000/api/issues/search"
SONARQUBE_URL_BASE = "http://172.23.12.34:9000"
SONARQUBE_PARAMS_ISSUES = {
    "componentKeys": "-",
    "types": "CODE_SMELL",
    "statuses": "OPEN"
}
SONARQUBE_PARAMS = {
    "component": "-",
    "metricKeys": "code_smells,ncloc,coverage,comment_lines,alert_status,rules,duplicated_lines_density"
}
SONARQUBE_PARAMS_METRICS = {
    "component": "-",
    "metricKeys": "ncloc,duplicated_lines_density,comment_lines_density,coverage,security_hotspots,sqale_index"
}
SONARQUBE_PARAMS_RULES = {
    "componentKeys": "-", 
    "ps": 1, 
    "facets": "rules",
    "statuses": "OPEN"
}
SONARQUBE_HEADERS = {"Authorization": "Bearer TU_TOKEN"}

# Configuración de Google Sheets
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "ruta/credenciales.json"
SPREADSHEET_ID = "ID_DE_TU_HOJA"

def obtener_datos_sonarqube(comp):
    SONARQUBE_PARAMS_ISSUES["componentKeys"] = comp
    SONARQUBE_PARAMS_METRICS["component"] = comp
    SONARQUBE_PARAMS["component"] = comp
    SONARQUBE_PARAMS_RULES["componentKeys"] = comp
    # Obtener estado del Quality Gate
    response_qg = requests.get(f"{SONARQUBE_URL_BASE}/api/qualitygates/project_status", params={"projectKey": comp})
    if response_qg.status_code == 200:
        data_qg = response_qg.json()
        estado_quality_gate = data_qg.get("projectStatus", {}).get("status", "DESCONOCIDO")
    else:
        print("Error al consultar SonarQube (Quality Gate)", response_qg.text)
        estado_quality_gate = "ERROR"
    response = requests.get(SONARQUBE_URL_ISSUES, params=SONARQUBE_PARAMS_ISSUES)
    if response.status_code == 200:
        data = response.json()
        issues = data.get("issues", [])
        cantidad_code_smells = data.get("total",0)
    else:
        print("Error al consultar SonarQube", response.text)
        cantidad_code_smells = 0
    
    response_metrics = requests.get(SONARQUBE_URL_METRICS, params=SONARQUBE_PARAMS_METRICS)
    if response_metrics.status_code == 200:
        data_metrics = response_metrics.json()
        measures = data_metrics.get("component", {}).get("measures", [])
        lineas_codigo = next((m["value"] for m in measures if m["metric"] == "ncloc"), "0")
        porcentaje_duplicidad = next((m["value"] for m in measures if m["metric"] == "duplicated_lines_density"), "0")
        porcentaje_comentarios = next((m["value"] for m in measures if m["metric"] == "comment_lines_density"), "0")
        coverage = next((m["value"] for m in measures if m["metric"] == "coverage"), "0")
        security_hotspots = next((m["value"] for m in measures if m["metric"] == "security_hotspots"), "0")
        deuda_tecnica_minutos_str = next((m["value"] for m in measures if m["metric"] == "sqale_index"), "0")
        try:
            deuda_tecnica_minutos = int(deuda_tecnica_minutos_str)
        except ValueError:
            print(f"Error al convertir deuda técnica: valor recibido '{deuda_tecnica_minutos_str}' no es un número válido.")
            deuda_tecnica_minutos = 0
        if deuda_tecnica_minutos < 480:
            # Menos de 8 horas: mostramos el total en horas (redondeando hacia arriba)
            deuda_tecnica_horas = math.ceil(deuda_tecnica_minutos / 60)
            deuda_tecnica = f"{deuda_tecnica_horas}h"
        else:
            dias = deuda_tecnica_minutos // 480
            minutos_restantes = deuda_tecnica_minutos % 480
            if minutos_restantes > 0:
                horas_restantes = math.ceil(minutos_restantes / 60)
                deuda_tecnica = f"{dias}d {horas_restantes}h"
            else:
                deuda_tecnica = f"{dias}d"
    
    else:
        print("Error al consultar SonarQube (metrics)", response_metrics.text)
        lineas_codigo = "0"

     # Obtener cantidad de reglas afectadas
    response_issues = requests.get(f"{SONARQUBE_URL}", params=SONARQUBE_PARAMS_RULES)
    if response_issues.status_code == 200:
        data_issues = response_issues.json()
        facets = data_issues.get("facets", [])
        reglas_afectadas = next((facet["values"] for facet in facets if facet["property"] == "rules"), [])
        cantidad_reglas_afectadas = len(reglas_afectadas)
    else:
        print("Error al consultar SonarQube (issues)", response_issues.text)
        cantidad_reglas_afectadas = 0

    return {
        "Estado": estado_quality_gate,
        "Cantidad Code Smells": cantidad_code_smells,
        "Líneas de Código": lineas_codigo,
        "Duplicidad": porcentaje_duplicidad,
        "Documentación": porcentaje_comentarios,
        "Coverage": coverage,
        "Vulnerabilidades": security_hotspots,
        "Deuda Tecnica": deuda_tecnica,
        "Reglas Afectadas": cantidad_reglas_afectadas
    }

def generar_reporte_excel(datos, nombres_componentes, ruta_salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte SonarQube"

    # Encabezados
    encabezados = [
        "Componente","Code Smells", "Estado", "Líneas de Código", "Duplicidad", "Documentación",
        "Coverage", "Vulnerabilidades", "Deuda Técnica", "Reglas Afectadas"
    ]
    ws.append(encabezados)

    # Colores
    verde_claro = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    rosa_claro = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    celeste = PatternFill(start_color="87CEFA", end_color="87CEFA", fill_type="solid")
    negrita_mayor = Font(bold=True, size=12)
    normal = Font(size=11)

    borde_medio = Border(
        left=Side(border_style="medium", color="000000"),
        right=Side(border_style="medium", color="000000"),
        top=Side(border_style="medium", color="000000"),
        bottom=Side(border_style="medium", color="000000")
    )

    for col_idx, encabezado in enumerate(encabezados, start=1):
        celda = ws.cell(row=1, column=col_idx)
        celda.fill = celeste
        celda.font = negrita_mayor
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_medio

    for componente, d in zip(nombres_componentes, datos):
        fila = [
            componente,
            d["Cantidad Code Smells"],
            "PASSED" if d["Estado"] == "OK" else "FAILED" if d["Estado"] == "ERROR" else d["Estado"],
            int(d["Líneas de Código"]),
            f"{d['Duplicidad']}%",
            f"{d['Documentación']}%",
            f"{d['Coverage']}%",
            str(d["Vulnerabilidades"]),
            d["Deuda Tecnica"],
            d["Reglas Afectadas"]
        ]
        ws.append(fila)

    # Formato
    for fila in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fila[0].alignment = Alignment(horizontal="left")     # Componente
        fila[1].alignment = Alignment(horizontal="center")   # Estado
        fila[2].alignment = Alignment(horizontal="center")   # Code Smells
        fila[3].alignment = Alignment(horizontal="right")    # Líneas de Código
         # Duplicidad, Documentación, Coverage y Vulnerabilidades → formato texto y centrado
        for i in [4, 5, 6, 7]:
            fila[i].alignment = Alignment(horizontal="center")
            fila[i].number_format = '@'  # texto

        # Deuda Técnica y Reglas Afectadas
        fila[8].alignment = Alignment(horizontal="center")
        fila[9].alignment = Alignment(horizontal="center")

        # Color según estado
        estado_celda = fila[2]
        if estado_celda.value == "PASSED":
            estado_celda.fill = verde_claro
        elif estado_celda.value == "FAILED":
            estado_celda.fill = rosa_claro

        # Color rosa claro si duplicidad >= 1
        duplicidad_str = fila[4].value
        if duplicidad_str.endswith('%'):
            try:
                duplicidad_valor = float(duplicidad_str.strip('%'))
                if duplicidad_valor >= 1:
                    fila[4].fill = rosa_claro
            except ValueError:
                pass  # Si no se puede convertir a número, lo ignoramos

        # Colorear documentación < 20%
        documentacion_str = fila[5].value
        if documentacion_str.endswith('%'):
            try:
                documentacion_valor = float(documentacion_str.strip('%'))
                if documentacion_valor < 20:
                    fila[5].fill = rosa_claro
            except ValueError:
                pass

        # Colorear coverage < 80%
        coverage_str = fila[6].value
        if coverage_str.endswith('%'):
            try:
                coverage_valor = float(coverage_str.strip('%'))
                if coverage_valor < 80:
                    fila[6].fill = rosa_claro
            except ValueError:
                pass
            
        for cell in fila:
            cell.border = borde_medio

    # Ajuste automático del ancho de columna (manual)
    for col_idx, _ in enumerate(encabezados, 1):
        col_letter = get_column_letter(col_idx)
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in ws[col_letter])
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(ruta_salida)

def obtener_condiciones_fallidas(componente: str):
    URL = SONARQUBE_URL_BASE + "/api/qualitygates/project_status"
    PARAMS = {"projectKey": componente}
    
    response = requests.get(URL, params=PARAMS)
    if response.status_code != 200:
        return {"estado": "ERROR", "mensaje": f"No se pudo obtener condiciones para {componente}"}

    condiciones = response.json().get("projectStatus", {}).get("conditions", [])
    condiciones_fallidas = [c for c in condiciones if c["status"] == "ERROR"]

    return {
        "componente": componente,
        "estado": "ERROR",
        "condiciones_fallidas": condiciones_fallidas,
        "total_fallidas": len(condiciones_fallidas)
    }

def obtener_code_smells(componente: str):
    URL = SONARQUBE_URL_BASE + "/api/issues/search"
    PARAMS = {
        "componentKeys": componente,
        "types": "CODE_SMELL",
        "resolved": "false",
        "sinceLeakPeriod": "true",
        "p": 1,
        "ps": 100
    }
    
    try:
        response = requests.get(URL, params=PARAMS)
        if response.status_code != 200:
            raise HTTPException(status_code=response.status_code, detail="Error consultando SonarQube")

        data = response.json()
        issues = data.get("issues", [])
        result = []

        for issue in issues:
            result.append({
                "message": issue.get("message"),
                "component": issue.get("component"),
                "author": issue.get("author"),
                "line": issue.get("line")
            })

        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
def obtener_reglas_afectadas(componente: str):
    try:
        resultados = []
        page = 1
        page_size = 500  # máximo permitido por SonarQube
        reglas_map = {}

        while True:
            response = requests.get(
                f"{SONARQUBE_URL}",
                params={
                    "componentKeys": componente,
                    "p": page,
                    "ps": page_size
                }
            )

            if response.status_code != 200:
                return JSONResponse(
                    content={"error": f"No se pudo obtener información para el componente {componente}"},
                    status_code=400
                )

            data = response.json()
            issues = data.get("issues", [])

            for issue in issues:
                regla = issue.get("rule")
                severity = issue.get("severity")

                if regla in reglas_map:
                    reglas_map[regla]["issues"] += 1
                else:
                    reglas_map[regla] = {
                        "regla": regla,
                        "descripcion": None,  # lo llenaremos luego
                        "severity": severity,
                        "issues": 1
                    }

            if len(issues) < page_size:
                break
            page += 1

        # Obtener descripciones de las reglas
        for regla in reglas_map:
            resp = requests.get(
                f"{SONARQUBE_URL_BASE}/api/rules/show",
                params={"key": regla}
            )
            if resp.status_code == 200:
                detalle = resp.json()
                reglas_map[regla]["descripcion"] = detalle.get("rule", {}).get("name")

        resultados = list(reglas_map.values())
        return JSONResponse(content={"reglas_afectadas": resultados}, status_code=200)

    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)
import os
import time
import io
import requests
import pandas as pd
import urllib.parse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
import plotly.express as px

# ReportLab para la exportación de tablas a PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# ====================================================================
# --- 1. CONFIGURACIÓN DE LA JERARQUÍA MAESTRA Y METADATOS ---
# ====================================================================
NOMBRE_PROYECTO_GENERAL = "Control Documentario - Shougang Hierro Perú"

ESTRUCTURA_MAESTRA = {
    "Proyectos Mayores": {
        "994440 - Tercera Línea": {
            "config": {
                "campos_personalizados": {
                    "6011837": "REV",
                    "6011840": "PENDIENTE POR",
                    "6011839": "FECHA",
                    "6011838": "ESTADO"
                },
                "id_campo_fecha": "6011839",
                "id_campo_pendiente": "6011840"
            },
            "Obras Civiles": {
                "Paquete 01 - CGI-SHP-019-2025": {
                    "nombre_proyecto": "994440 Paquete CIVIL 01 - C519 - SGMC",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Luis Rojas",
                    "coordinador_shp": "Wilbert Salas",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.4UH-sGf0Rj-R934_tj1S8w&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.oruowhl0STa9w2M9Vo3gZQ&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.8XNGFh6XTcy8HwGVqgx0tA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.qGCrERygRvWHVF__AF2hzQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1mZq3p5tTFmIqwiWR8mk5Q&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.dq0X5F6lT-e8xtT57iEn5Q&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1mZq3p5tTFmIqwiWR8mk5Q&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.IgD5_slUQmCzIDI3-N5_ag&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                },
                "Paquete 02 CIVIL - CGI-SHP-027-2025": {
                    "nombre_proyecto": "994440 Paquete CIVIL 02 - C527 - CHEC",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Luis Rojas",
                    "coordinador_shp": "W. Salas V. Injante",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.3y89r3vUSNK38q9RBNMcTw&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.5B-xkz92QcKaWuRIW-4jOQ&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.nfZJis6zTwmuhwSkHy3q6Q&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Z3sU9UeUS6-hA3FmXbsBzA&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.O_EGlk_BR3mKQ7iXlFpJHg&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.HU3HnztHR369GIhLTsDCSA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.nBwFxIOaSRGRiL9Ka5URxQ&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Cm_uBP0wRHyvXG3hRMJnGg&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                },
                "Paquete 03 - CGI-SHP-030-2025": {
                    "nombre_proyecto": "994440 Paquete CIVIL 03 - C530 - CHEC",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Luis Rojas",
                    "coordinador_shp": "W. Salas V. Injante",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.2BAc3d6BQd6QPx9NYLsR8g&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.s06ONvkCTMOwnyHS5GqFdQ&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.YFrF_lHvS468ojtMthPv-Q&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.-6wYCzgjSm-wSEPFjpbkYA&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.cDG2JObXQYyLlYKomvxEVw&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.n_1rlmgQS0OG626yut5eeA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.YYQ1fUa6SW-P7sXpddjTig&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.4lKp7659QoG7BWZnbXveGg&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                },
                "Paquete 04 - CGI-SHP-039-2025": {
                    "nombre_proyecto": "994440 Paquete CIVIL 43 - C539 - SGMC",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Luis Rojas",
                    "coordinador_shp": "W. Salas V. Injante",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.MlBMDlJUQaeIxce1dhS8bA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.T7aNEAb4QvWxGBWXWgyiwg&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.JNNb8aQaTEqc5T0UDKGBBA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.j1ehadOcT6KCCIt72R_D0w&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.RV_SxUg6Tf2akD0TUsK7cA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.qn6sW6ogTq63uLBirE16lA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.nlH_q9GRRuacmubJOuKzbA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.yLLA_nYMR0WCmTbqIQksVQ&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                },
                "Paquete 05 - CGI-SHP-034-2025": {
                    "nombre_proyecto": "994440 Paquete CIVIL 05 - C534 - COVEC",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Luis Rojas",
                    "coordinador_shp": "Norman Chura",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.IxQiNfJ5TQia_4CkkcWlow&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.z8sjrzhJSY633BycxBdtOw&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.COHiEf-5Rv-1_GdCyjRIWg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.sm5zScUMS_eGMUzBTZH-UA&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.QtkTHz6ZRhCMo_q8EICxJg&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1kdLDO2tTFuxq5Fgu0EG7g&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.IvwY65OaQMuJloGhZnSRdg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.WxYPFm9uQ-2zV3vSNEh8yg&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                }
            },
            "Obras Electromecanicas": {
                "Paquete 01 - CGI-SHP-001-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 01 - C601 CHEC BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "Joseph Beltran",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Bji4Uop7SSGwBSj2kKdmLg&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ZUirAImDTluEAH7D1z2xCA&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.6ayy3aotR1iZErdo8ZQZpA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.cWeBjsVeQJOe9T2dgiDLRQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.iLK0CFTcRE-f5vPgXsuWGQ&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Sgn_HpAZSfK84UcH7Od1Pw&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.J8zAgoEdRFOlCo4FwXmVHg&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 02 - CGI-SHP-003-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 02 - C602 SGMC BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "Ronald Valdivia",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.hVkYLUDGTBWxaaPw1rDMwA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.XxU1AWhiR2GfKKRlKDC5xw&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.KZ0ThTUiSGaev5XQ8aBmlA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ubR1IJdXQJOxKBgw_MUgRg&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.9315eaIrR-SN4m22KNQ6WA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.P3K04-MYRZ27fNfyl4grnQ&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ODILcCHhS0qbGcZL7dP8TA&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 03 - CGI-SHP-005-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 03 - C602 SGMC BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "Ronald Valdivia",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.5iNN6rXrRG-ZwrgECCcwDQ&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.BvbTrUAfR0i5ZlYplxKUTA&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.6ZQfNN_NSlilQsUjxSX-kQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.TojUUNCqQuihYkLFClAAeA&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.8JZKPjJfQBWZhZC-45fAKg&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.DOHVT-hYS9OW-w-oIz7Quw&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.4Pex6mj3Q2aUH8kNnMVicA&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 04 - CGI-SHP-006-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 04 - C605 CHEC BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "Ronald Valdivia",
                    "administrador_contratos": "Victor Calvo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.78FSqSiiT2OingJX1Ja9aQ&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ZvZsTnO8QlSWDcnjw_H30w&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.NWcSlYaQTxeFj0sATS_U-A&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco._kdxfyRsRUGba669vzphTA&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.xlVNeXVnSXe-4O8j0x-PGA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.E0DpZUPrQ9GjyRfP6BjmTA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.zr5NBPPPREegjBM3tgREkw&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 05A - CGI-SHP-004-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 05A - C604 CCCC BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "A. Nina / J. Salazar",
                    "administrador_contratos": "Gonzalo Vargas",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.UJKXYjhISuytzIB_rqyjng&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.H3kuOcFhTk6pJOW9IJ0DhA&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.g756st7KT3yVaRZOztyTSg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.TgPRi37tQnqLGHFE6GLelQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.YVxMOXejTX-HS_LbD6dHkA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.XqxxFA-MQWaF83BGQfVKMg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.bKUwofRQS2uZ-NKyB3YhpA&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 05B - CGI-SHP-010-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 05B - C610 COVEC BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "A. Nina / J. Salazar",
                    "administrador_contratos": "Gonzalo Vargas",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.bQj2WFhhQs62aGUvEcTxRA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.cAHQmaVtTbyYtwANgqCjYQ&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.5zGmicrWQv6DSKhv-QGlQA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.cx_ccOZRRCGHb3J8EfSl2Q&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.RPXN3sO1QDmvJS-5hn_kBg&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.MtSekBToRhGs0xIKfvL9rg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Y31XHPY2Sy265fvyKhPt_g&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 06 - CGI-SHP-011-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 06 - C611 AID BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "Mauro Yaranga",
                    "administrador_contratos": "Gonzalo Vargas",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.4i25QgAQQ9WgM7aAPcqK0A&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.HuWUE7r6QI2wE9H1T94VQQ&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.HH6wISmLSaWAxr-FUw8lBg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.3OYObaXeSvCljs2HQFYaEw&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.AUdUCKq0QFeHxlE-PPTOQQ&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.CKikIYrkSQmMF3wyt62jOw&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Zba0ob9tTk6PCFB-k9akUA&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Paquete 07 - CGI-SHP-012-2026": {
                    "nombre_proyecto": "994440 Paquete OEM 06 - C611 AID BISA",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Ing. BISA / Supervisor Asignado",
                    "coordinador_shp": "Misael Castillo",
                    "administrador_contratos": "Ivan Lima",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.OyYj6aw6SASOsJXmEBmKAA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.VP0NGSxmRBmB6PqVttQK7g&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.vKhVxRjjQv2LkLLTFfbaEA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1dXsqtcOTPKzK3RKlZv0rg&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.NPx2OF7SQUemcJj4g-MJ6w&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.NJUKhA54SiCmLn2X-xXaqw&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.wsxEGRDMSJ-KFVuU5c463w&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                }
            }
        },
        "998316 Nuevo Muelle San Nicolas": {
            "config": {
                "campos_personalizados": {
                    "6011837": "REV",
                    "6011840": "PENDIENTE POR",
                    "6011839": "FECHA",
                    "6011838": "ESTADO"
                },
                "id_campo_fecha": "6011839",
                "id_campo_pendiente": "6011840"
            }
        },
        "I25G50 Subestacion Mina 02": {
            "config": {
                "campos_personalizados": {
                    "6011837": "REV",
                    "6011840": "PENDIENTE POR",
                    "6011839": "FECHA",
                    "6011838": "ESTADO"
                },
                "id_campo_fecha": "6011839",
                "id_campo_pendiente": "6011840"
            }
        }
    },
    "Proyectos Menores": {
        "I26R20 Reparacion de Tanque C402-026": {
            "config": {
                "campos_personalizados": {
                    "11981550": "REV",
                    "11981548": "PENDIENTE POR",
                    "11981549": "FECHA",
                    "11981551": "ESTADO"
                },
                "id_campo_fecha": "11981549",
                "id_campo_pendiente": "11981548"
            },
            "General": {
                "C619 CGI-SHP-019-2026": {
                    "nombre_proyecto": "I26R20 Reparacion de Tanque C402-026 - C619 - CHEC",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "N/A",
                    "coordinador_shp": "Juan Purilla / Erick Mayta",
                    "administrador_contratos": "Hugo Cardenas",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.yCuIPfheQcOEb1lL9eobqw&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.8L99-KY-SEeFu7qGx677YA&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.waNHhkfZRV2EMpm5TvMaqw&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Fj8vsQg8R4KtDg1262Al3Q&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.z4k_ZGrnTMeQlOJyTN9p0A&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.f5BtCYtRS6yeRdPNgA5-qA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/6030afc6-005c-4bb8-937e-c9a97850d14b?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.CGg_avSiQXKunwiqy7mSKg&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                }
            }
        }
    },
    "Plan de Conservacion": {
        "Paquete 03 Mina": {
            "config": {
                "campos_personalizados": {
                    "5737743": "REV",
                    "5737740": "PENDIENTE POR",
                    "5737741": "FECHA",
                    "5737742": "ESTADO"
                },
                "id_campo_fecha": "5737741",
                "id_campo_pendiente": "5737740"
            },
            "Mina": {
                "C624 CGI-SHP-024-2026": {
                    "nombre_proyecto": "PC2026 Paquete 03 Mina - C624 / CVC",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "",
                    "coordinador_shp": "Victor Injante",
                    "administrador_contratos": "Hugo Cardenas",
                    "project_id": "3fe40740-4483-4ed9-895b-9cb579cd7b1d",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.L17F4gQ5QeeRgtrsF6oQTg&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.DH6CRKemTJqfDz2Je3zE9g&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.6f4Tpw-jTu-2_z8ysdZ3rw&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ut8q9MX_STmGYaXM5TLLng&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.dp9F3hlAQMmHsYn4xmjqtA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.drLhaXctTaeQSBS5bU7hjg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.TkNgo2NmRWiAjICC9hftHA&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                }
            }
        }
    }
}

COL_CODIGO = "CODIGO"
COL_DESCRIPCION = "DESCRIPCION"

# ====================================================================
# --- 2. CREDENCIALES (SEGURAS NUBE / LOCAL) ---
# ====================================================================
try:
    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
except:
    CLIENT_ID = "v1qczmMgnll6AUsLAPKuPVC31GaKLNnqmFsVvc6OA2Seqx4H"
    CLIENT_SECRET = "MAKRfMaqiXWBTJiFY5sJP3BLqwcRDWClc0LjFN79HYKBoyOcBfC7ObjCdRkQzebW"

# ====================================================================
# --- 3. FUNCIONES DE EXPORTACIÓN A PDF (REPORTLAB) ---
# ====================================================================
def exportar_df_a_pdf(df: pd.DataFrame, titulo: str) -> io.BytesIO:
    """Convierte cualquier DataFrame de pandas a un documento PDF descargable."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=25,
        rightMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    elements = []
    styles = getSampleStyleSheet()

    titulo_style = styles["Heading2"]
    titulo_style.alignment = 0

    cell_style = ParagraphStyle(
        name="CellText",
        parent=styles["Normal"],
        fontSize=7,
        leading=9
    )
    header_style = ParagraphStyle(
        name="HeaderCellText",
        parent=styles["Normal"],
        fontSize=7,
        leading=9,
        textColor=colors.whitesmoke,
        fontName="Helvetica-Bold"
    )

    elements.append(Paragraph(titulo, titulo_style))
    elements.append(Spacer(1, 10))

    encabezados = [Paragraph(str(col), header_style) for col in df.columns]
    data_matrix = [encabezados]

    for _, row in df.iterrows():
        fila = []
        for val in row:
            texto = "" if pd.isna(val) else str(val)
            fila.append(Paragraph(texto, cell_style))
        data_matrix.append(fila)

    tabla = Table(data_matrix, repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))

    elements.append(tabla)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ====================================================================
# --- 4. LÓGICA DE API AUTODESK ---
# ====================================================================
def obtener_ids_excluidos(urls):
    ids = set()
    for url in urls:
        if "folderUrn=" in url:
            ids.add(url.split("folderUrn=")[1].split("&")[0].replace("%3A", ":"))
        elif "folders/" in url:
            ids.add(url.split("folders/")[1].split("/")[0].replace("%3A", ":"))
    return ids

def obtener_token():
    url = "https://developer.api.autodesk.com/authentication/v2/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
    data = {"grant_type": "client_credentials", "scope": "data:read"}
    response = requests.post(url, headers=headers, data=data, auth=(CLIENT_ID, CLIENT_SECRET))
    return response.json().get("access_token") if response.status_code == 200 else None

def deconstruir_diccionario(datos, prefijo, resultado):
    if isinstance(datos, dict):
        for llave, valor in datos.items():
            nueva_llave = f"{prefijo}_{llave.strip().upper()}" if prefijo else llave.strip().upper()
            if any(x in nueva_llave for x in ["LINKS", "SCHEMA", "HREF", "DERIVATIVES"]):
                continue
            deconstruir_diccionario(valor, nueva_llave, resultado)
    elif isinstance(datos, list):
        for i, elemento in enumerate(datos):
            if isinstance(elemento, dict):
                name = elemento.get("name") or elemento.get("displayName") or elemento.get("title")
                val = elemento.get("value") or elemento.get("displayValue")
                if name and val is not None:
                    resultado[f"CAMPO_CUSTOM_{str(name).strip().upper()}"] = val
                else:
                    deconstruir_diccionario(elemento, f"{prefijo}_{i}", resultado)
            else:
                resultado[prefijo] = datos
    else:
        resultado[prefijo] = datos

def obtener_tip_version_metadata(token, project_id, item_id):
    clean_item_id = item_id.split("?version=")[0]
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{clean_item_id}/tip"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json", "x-ads-region": "US"}
    
    while True:
        headers["x-ads-region"] = "US" 
        response = requests.get(url, headers=headers)
        if response.status_code == 429:
            time.sleep(2)
            continue
        elif response.status_code == 401:
            nuevo_token = obtener_token()
            if nuevo_token:
                headers["Authorization"] = f"Bearer {nuevo_token}"
                continue
        elif response.status_code != 200:
            headers["x-ads-region"] = "EMEA"
            response = requests.get(url, headers=headers)
            if response.status_code == 429:
                time.sleep(2)
                continue
        break
    return response.json().get("data", {}) if response.status_code == 200 else {}

def procesar_archivo_pdf(item, project_id, headers, ruta_actual, name, campos_personalizados):
    item_id = item.get("id", "")
    token_val = headers["Authorization"].split(" ")[1]
    
    registro_plano = {
        "NOMBRE_DEL_ARCHIVO": name,
        "RUTA_EN_AUTODESK": ruta_actual,
        "ITEM_ID": item_id
    }
    deconstruir_diccionario(item, "ITEM", registro_plano)
    
    tip_version = obtener_tip_version_metadata(token=token_val, project_id=project_id, item_id=item_id)
    if tip_version:
        deconstruir_diccionario(tip_version, "VERSION", registro_plano)

    vid = item.get("relationships", {}).get("tip", {}).get("data", {}).get("id", "")
    if vid:
        vid_encoded = urllib.parse.quote(vid, safe='')
        url_attr = f"https://developer.api.autodesk.com/bim360/docs/v1/projects/{project_id.replace('b.', '')}/versions/{vid_encoded}/custom-attributes"
        
        while True:
            res_attr = requests.get(url_attr, headers={"Authorization": headers["Authorization"], "Content-Type": "application/json"})
            if res_attr.status_code == 429:
                time.sleep(2)
                continue
            elif res_attr.status_code == 401:
                nuevo_token_attr = obtener_token()
                if nuevo_token_attr:
                    headers["Authorization"] = f"Bearer {nuevo_token_attr}"
                    continue
            break
            
        if res_attr.status_code == 200:
            campos_objetivo = list(campos_personalizados.keys())
            for attr in res_attr.json():
                attr_id = str(attr.get("id"))
                if attr_id in campos_objetivo:
                    registro_plano[f"CAMPO_CUSTOM_{attr_id}"] = str(attr.get("value", "")).strip()
                    
    return registro_plano

def escanear_recursivo(project_id, folder_id, headers, lista_datos, longitud_maxima, ids_excluidos, campos_personalizados, ruta_actual="Raíz"):
    if folder_id.replace("%3A", ":") in ids_excluidos:
        return

    url_api = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/folders/{folder_id}/contents"
    while url_api:
        headers["x-ads-region"] = "US"
        res = requests.get(url_api, headers=headers)
        
        if res.status_code == 401:
            nuevo_token = obtener_token()
            if nuevo_token:
                headers["Authorization"] = f"Bearer {nuevo_token}"
                continue
            else:
                break
        elif res.status_code == 429:
            time.sleep(5)
            continue
        elif res.status_code != 200:
            headers["x-ads-region"] = "EMEA"
            res = requests.get(url_api, headers=headers)
            if res.status_code != 200:
                break
        
        data = res.json()
        items_pdf = []
        
        for item in data.get("data", []):
            tipo = item.get("type")
            attributes = item.get("attributes", {})
            name = str(attributes.get("displayName", "Sin Nombre")).strip()

            if tipo == "folders":
                sub_urn_raw = item["id"].replace("%3A", ":")
                if sub_urn_raw in ids_excluidos:
                    continue
                if len(name) <= longitud_maxima:
                    sub_urn = urllib.parse.quote(item["id"], safe='')
                    escanear_recursivo(project_id, sub_urn, headers, lista_datos, longitud_maxima, ids_excluidos, campos_personalizados, f"{ruta_actual} > {name}")
            
            elif tipo == "items":
                if ".pdf" in name.lower():
                    items_pdf.append((item, name))

        if items_pdf:
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(procesar_archivo_pdf, i[0], project_id, headers, ruta_actual, i[1], campos_personalizados) for i in items_pdf]
                for future in as_completed(futures):
                    res_item = future.result()
                    if res_item:
                        lista_datos.append(res_item)

        enlaces = data.get("links", {})
        url_api = enlaces["next"].get("href") if "next" in enlaces else None

@st.cache_data(ttl=3600)
def extraer_datos_contrato_completo(contrato_key, info_contrato, campos_personalizados, id_campo_fecha, id_campo_pendiente):
    token = obtener_token()
    if not token:
        return {}
    
    headers = {"Authorization": f"Bearer {token}", "x-ads-region": "US"}
    urls_excluidas_raw = info_contrato.get("urls_excluidas", [])
    ids_excluidos = obtener_ids_excluidos(urls_excluidas_raw)
    carpetas_config = info_contrato.get("carpetas_config", [])
    
    dataframes_por_pestana = {}
    col_pendiente = campos_personalizados.get(id_campo_pendiente)
    orden_columnas_deseado = ["CODIGO", "REV", "DESCRIPCION", "FECHA", "ESTADO", "PENDIENTE POR", "DIAS"]

    for config_carpeta in carpetas_config:
        url_carpeta = config_carpeta.get("url")
        nombre_pestana = config_carpeta.get("nombre", "Sin_Nombre")
        longitud_maxima = config_carpeta.get("longitud_maxima", 10)
        
        if not url_carpeta:
            continue
            
        try:
            project_id_raw = url_carpeta.split("projects/")[1].split("?")[0].split("/")[0]
            project_id = project_id_raw if project_id_raw.startswith("b.") else "b." + project_id_raw
            
            if "folderUrn=" in url_carpeta:
                folder_id = url_carpeta.split("folderUrn=")[1].split("&")[0].replace("%3A", ":")
            elif "folders/" in url_carpeta:
                folder_id = url_carpeta.split("folders/")[1].split("/")[0].replace("%3A", ":")
            else:
                continue
            
            datos_filtrados = []
            escanear_recursivo(project_id, folder_id, headers, datos_filtrados, longitud_maxima, ids_excluidos, campos_personalizados)
            
            if not datos_filtrados:
                continue
                
            df_completo = pd.DataFrame(datos_filtrados)
            
            for campo_id in campos_personalizados.keys():
                col_name = f"CAMPO_CUSTOM_{campo_id}"
                if col_name not in df_completo.columns:
                    df_completo[col_name] = pd.NaT if campo_id == id_campo_fecha else ""

            fechas_convertidas = pd.to_datetime(df_completo[f"CAMPO_CUSTOM_{id_campo_fecha}"], errors='coerce')
            if fechas_convertidas.dt.tz is not None:
                fechas_convertidas = fechas_convertidas.dt.tz_localize(None)
            
            fecha_hoy = pd.Timestamp.now().normalize()
            df_completo["DIAS"] = (fecha_hoy - fechas_convertidas).dt.days
            df_completo[f"CAMPO_CUSTOM_{id_campo_fecha}"] = fechas_convertidas.dt.strftime('%d/%m/%Y').fillna("")
            df_completo["DIAS"] = df_completo["DIAS"].fillna("")

            columnas_finales = ["NOMBRE_DEL_ARCHIVO", "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] + [f"CAMPO_CUSTOM_{c}" for c in campos_personalizados.keys()] + ["DIAS"]
            df = df_completo[columnas_finales].copy()
            
            diccionario_renombres = {
                "NOMBRE_DEL_ARCHIVO": COL_CODIGO, 
                "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": COL_DESCRIPCION,
            }
            for campo_id, nombre_columna in campos_personalizados.items():
                diccionario_renombres[f"CAMPO_CUSTOM_{campo_id}"] = nombre_columna
                
            df = df.rename(columns=diccionario_renombres)
            columnas_existentes = [col for col in orden_columnas_deseado if col in df.columns]
            df = df[columnas_existentes]

            if col_pendiente in df.columns:
                df.loc[df[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0
            
            dataframes_por_pestana[nombre_pestana] = df

        except Exception as e:
            continue

    return dataframes_por_pestana

# ====================================================================
# --- 5. INTERFAZ WEB STREAMLIT ---
# ====================================================================
st.set_page_config(page_title="Control Documentario SHP", layout="wide", page_icon="📊")

# --- ENCABEZADO CON LOGO EN LA PARTE SUPERIOR DERECHA ---
col_head1, col_head2 = st.columns([5, 1])
with col_head1:
    st.title("📊 Dashboard de Control Documentario")
    st.markdown("**Institución:** Shougang Hierro Perú")

with col_head2:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Shougang_Group_logo.svg/1200px-Shougang_Group_logo.svg.png",
        width=130
    )

st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Shougang_Group_logo.svg/1200px-Shougang_Group_logo.svg.png", width=150)
st.sidebar.title("Estructura de Proyectos")

grupo_activo = st.sidebar.selectbox("1. Tipo de Alcance:", list(ESTRUCTURA_MAESTRA.keys()))

subproyectos_dict = ESTRUCTURA_MAESTRA.get(grupo_activo, {})
if not subproyectos_dict:
    st.warning(f"No hay subproyectos configurados en '{grupo_activo}'.")
    st.stop()

subproyecto_activo = st.sidebar.selectbox("2. Subproyecto:", list(subproyectos_dict.keys()))
subproyecto_data = subproyectos_dict.get(subproyecto_activo, {})

config_sub = subproyecto_data.get("config", {
    "campos_personalizados": {
        "6011837": "REV",
        "6011840": "PENDIENTE POR",
        "6011839": "FECHA",
        "6011838": "ESTADO"
    },
    "id_campo_fecha": "6011839",
    "id_campo_pendiente": "6011840"
})

campos_pers_activos = config_sub.get("campos_personalizados")
id_fec_activo = config_sub.get("id_campo_fecha")
id_pen_activo = config_sub.get("id_campo_pendiente")

tipos_obra_dict = {k: v for k, v in subproyecto_data.items() if k != "config"}
if not tipos_obra_dict:
    st.warning(f"No hay tipos de obra o contratos configurados en '{subproyecto_activo}'.")
    st.stop()

tipo_obra_activo = st.sidebar.selectbox("3. Tipo de Obra:", list(tipos_obra_dict.keys()))

contratos_dict = tipos_obra_dict.get(tipo_obra_activo, {})
if not contratos_dict:
    st.warning(f"No hay contratos configurados en '{tipo_obra_activo}'.")
    st.stop()

contrato_activo = st.sidebar.selectbox("4. Contrato / Componente:", list(contratos_dict.keys()))
info_contrato = contratos_dict[contrato_activo]

st.sidebar.markdown("---")
st.sidebar.info(f"**Selección Actual:**\n• {grupo_activo}\n• {subproyecto_activo}\n• {tipo_obra_activo}\n• Contrato: **{contrato_activo}**")

# --- PRECARGA AUTOMÁTICA AL INICIAR ---
with st.spinner(f'Inicializando y precargando metadatos para {contrato_activo}...'):
    datos_pestanas = extraer_datos_contrato_completo(contrato_activo, info_contrato, campos_pers_activos, id_fec_activo, id_pen_activo)

if not datos_pestanas:
    st.info(f"El contrato {contrato_activo} no contiene registros o sus carpetas están pendientes de enlace en el código.")
else:
    # --- TARJETA DE DATOS GENERALES Y TIPO DE CONTRATO ---
    tipo_contrato_val = info_contrato.get("tipo_contrato", "Con Supervisión")
    
    with st.expander("📌 Datos Generales del Contrato y Equipo Asignado", expanded=True):
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Modalidad", tipo_contrato_val)
        
        if tipo_contrato_val == "Tripartito":
            emp1 = info_contrato.get("nombre_empresa_1", "Empresa 1")
            emp2 = info_contrato.get("nombre_empresa_2", "Empresa 2")
            c2.metric("Empresa 1", emp1)
            c3.metric("Empresa 2", emp2)
        else:
            c2.metric("Supervisor", info_contrato.get("nombre_supervisor", "N/A"))
            c3.metric("Coordinador SHP", info_contrato.get("coordinador_shp", "N/A"))
            
        c4.metric("Admin. Contratos", info_contrato.get("administrador_contratos", "N/A"))
        c5.metric("Estado Caché", "Activo (1h)")

    col_pendiente = campos_pers_activos.get(id_pen_activo, "PENDIENTE POR")
    
    # --- DETERMINAR ROLES SEGÚN TIPO DE CONTRATO ---
    if tipo_contrato_val == "Tripartito":
        emp1 = info_contrato.get("nombre_empresa_1", "EMPRESA 1")
        emp2 = info_contrato.get("nombre_empresa_2", "EMPRESA 2")
        roles_pendientes = [emp1.upper(), emp2.upper(), "SHP"]
    elif tipo_contrato_val == "Contrato Directo":
        roles_pendientes = ["CONTRATISTA", "SHP"]
    else:
        roles_pendientes = ["CONTRATISTA", "SUPERVISION", "SHP"]

    # --- CONSOLIDAR MÉTRICAS PARA RESUMEN Y GRÁFICO ---
    total_cerrados = 0
    total_abiertos = 0
    resumen_lista = []

    for nombre_carpeta, df in datos_pestanas.items():
        fila_res = {"Carpeta": nombre_carpeta}
        cerrados_carp = 0
        abiertos_carp = 0
        
        if col_pendiente in df.columns:
            estados = df[col_pendiente].astype(str).str.strip().str.upper()
            cerrados_carp = (estados == "CERRADO").sum()
            fila_res["Cerrados"] = cerrados_carp
            
            for rol in roles_pendientes:
                cant_rol = (estados == rol).sum()
                fila_res[rol.title()] = cant_rol
                abiertos_carp += cant_rol
            
            # --- CÁLCULO DEL DOCUMENTO CON MAYOR RETRASO ---
            df_abiertos = df[estados != "CERRADO"].copy()
            if not df_abiertos.empty and 'DIAS' in df_abiertos.columns:
                df_abiertos['DIAS_NUM'] = pd.to_numeric(df_abiertos['DIAS'], errors='coerce').fillna(0)
                if df_abiertos['DIAS_NUM'].max() > 0:
                    max_idx = df_abiertos['DIAS_NUM'].idxmax()
                    doc_codigo = str(df_abiertos.loc[max_idx, COL_CODIGO])
                    dias_retraso = int(df_abiertos.loc[max_idx, 'DIAS_NUM'])
                    responsable = str(df_abiertos.loc[max_idx, col_pendiente]).strip()
                    fila_res["Mayor Retraso"] = f"{doc_codigo} ({dias_retraso} días) [{responsable}]"
                else:
                    fila_res["Mayor Retraso"] = "Sin retrasos"
            else:
                fila_res["Mayor Retraso"] = "Sin pendientes"
        else:
            fila_res["Cerrados"] = 0
            for rol in roles_pendientes:
                fila_res[rol.title()] = 0
            fila_res["Mayor Retraso"] = "N/A"
                
        total_cerrados += cerrados_carp
        total_abiertos += abiertos_carp
        fila_res["Total"] = cerrados_carp + abiertos_carp
        resumen_lista.append(fila_res)

    # --- MÉTRICAS PRINCIPALES ---
    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Documentos", total_cerrados + total_abiertos)
    m2.metric("Documentos Cerrados", total_cerrados)
    m3.metric("Documentos Pendientes", total_abiertos, delta="Requieren atención", delta_color="inverse")
    st.markdown("---")

    # --- TABLA RESUMEN GENERAL Y GRÁFICO DE BARRAS ---
    st.subheader(f"Resumen General por Carpeta - Contrato: {contrato_activo}")
    df_resumen = pd.DataFrame(resumen_lista)
    
    col_g1, col_g2 = st.columns([1.3, 1])
    with col_g1:
        st.dataframe(df_resumen, width='stretch', hide_index=True)
        # Botón de exportación a PDF para la tabla resumen
        if not df_resumen.empty:
            pdf_resumen = exportar_df_a_pdf(
                df_resumen, 
                f"Resumen General - {contrato_activo} ({datetime.now().strftime('%d/%m/%Y')})"
            )
            st.download_button(
                label="📄 Exportar Resumen a PDF",
                data=pdf_resumen,
                file_name=f"Resumen_{contrato_activo.replace(' ', '_')}.pdf",
                mime="application/pdf",
                key="btn_descarga_resumen"
            )
    with col_g2:
        if not df_resumen.empty:
            columnas_grafico = ["Cerrados"] + [c for c in df_resumen.columns if c not in ["Carpeta", "Cerrados", "Total", "Mayor Retraso"]]
            fig = px.bar(
                df_resumen, 
                x="Carpeta", 
                y=columnas_grafico, 
                title="Estado Documentario por Carpeta y Responsable",
                barmode="group"
            )
            fig.update_layout(height=300, margin=dict(t=30, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # --- VISTA DETALLADA POR CARPETA INDIVIDUAL, GRÁFICO Y FILTRO ---
    st.subheader("Detalle por Carpeta / Disciplina")
    nombres_carpetas = list(datos_pestanas.keys())
    pestana_seleccionada = st.selectbox("Seleccionar Carpeta:", nombres_carpetas)
    
    df_actual = datos_pestanas[pestana_seleccionada].copy()

    if col_pendiente in df_actual.columns:
        roles_con_cerrado = roles_pendientes + ["CERRADO"]
        estados_disponibles = df_actual[col_pendiente].astype(str).unique().tolist()
        
        st.markdown(f"**Modalidad de contrato activa:** `{tipo_contrato_val}` (Roles configurados: *{', '.join(roles_con_cerrado)}*)")
        filtro_estado = st.multiselect("Filtrar por Responsable:", estados_disponibles, default=estados_disponibles)
        
        df_filtrado = df_actual[df_actual[col_pendiente].isin(filtro_estado)]
        
        if not df_filtrado.empty:
            # Crear columnas para colocar la tabla junto al gráfico
            col_tabla, col_grafico = st.columns([1.5, 1])
            
            with col_tabla:
                st.dataframe(df_filtrado, width='stretch', height=400)
                
                # Botón de exportación a PDF para la tabla detallada filtrada
                pdf_detalle = exportar_df_a_pdf(
                    df_filtrado, 
                    f"Detalle: {pestana_seleccionada} - Contrato: {contrato_activo}"
                )
                st.download_button(
                    label="📄 Exportar Detalle a PDF",
                    data=pdf_detalle,
                    file_name=f"{pestana_seleccionada}_{contrato_activo.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    key="btn_descarga_detalle"
                )
                
            with col_grafico:
                # Generar el gráfico basado en la tabla filtrada
                df_counts = df_filtrado[col_pendiente].value_counts().reset_index()
                df_counts.columns = [col_pendiente, 'Cantidad']
                
                fig_detalle = px.pie(
                    df_counts, 
                    names=col_pendiente, 
                    values='Cantidad', 
                    hole=0.4,
                    title=f"Distribución de Responsables"
                )
                fig_detalle.update_layout(margin=dict(t=40, b=10, l=10, r=10))
                st.plotly_chart(fig_detalle, use_container_width=True)
        else:
            st.info("No hay documentos que coincidan con el filtro seleccionado.")
            
    else:
        # En caso no exista la columna de responsable
        st.dataframe(df_actual, width='stretch', height=400)
        
        if not df_actual.empty:
            # Botón de exportación a PDF para la tabla detallada sin filtro
            pdf_detalle = exportar_df_a_pdf(
                df_actual, 
                f"Detalle: {pestana_seleccionada} - Contrato: {contrato_activo}"
            )
            st.download_button(
                label="📄 Exportar Detalle a PDF",
                data=pdf_detalle,
                file_name=f"{pestana_seleccionada}_{contrato_activo.replace(' ', '_')}.pdf",
                mime="application/pdf",
                key="btn_descarga_detalle_alt"
            )

# ====================================================================
# --- 6. AUTORÍA EN LA ESQUINA INFERIOR DERECHA ---
# ====================================================================
st.markdown(
    """
    <style>
    .footer-autor {
        position: fixed;
        bottom: 12px;
        right: 25px;
        background-color: rgba(255, 255, 255, 0.92);
        color: #2D3748;
        padding: 6px 14px;
        border-radius: 8px;
        font-size: 13px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        box-shadow: 0px 2px 8px rgba(0,0,0,0.15);
        z-index: 999999;
        border: 1px solid #E2E8F0;
    }
    </style>
    <div class="footer-autor">
        Desarrollado por <b>Roberto Cano</b>
    </div>
    """,
    unsafe_allow_html=True
)

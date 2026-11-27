import os
import time
import io
import threading
import requests
import pandas as pd
import numpy as np
import urllib.parse
from datetime import datetime, timedelta
import pytz
from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage

# Conexión al módulo de reportes
from modulo_formulario_reportes import mostrar_formulario_reportes

# Librerías para exportación a PDF (ReportLab)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.platypus import Image as RLImage

# Librerías para exportación a Excel (OpenPyXL)
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment, PatternFill

# Conexión al editor de estructura
from modulo_estructura_proyectos import cargar_estructura_github, mostrar_editor_estructura

# ====================================================================
# --- 1. CONFIGURACIÓN DE LA JERARQUÍA MAESTRA Y METADATOS ---
# ====================================================================
NOMBRE_PROYECTO_GENERAL = "Control Documentario - Shougang Hierro Perú"
URL_DASHBOARD_WEB = "https://dashboard-control-documentario-tpzxfmbgpfnz5amgqzsff2.streamlit.app/"

ESTRUCTURA_MAESTRA = {
    "Proyectos Mayores": {
        "994440 - Tercera Línea": {
            "config": {
                "campos_personalizados": {
                    "6011837": "REV",
                    "6011840": "PENDIENTE POR",
                    "6011839": "FECHA",
                    "6011838": "ESTADO",
                    "6011848": "DISCIPLINA",
                    "6011862": "TIPO DOC",
                    "6011845": "FASE",
                    "6011844": "AREA"
                },
                "id_campo_fecha": "6011839",
                "id_campo_pendiente": "6011840"
            },
            "Ingenieria de detalle": {
                "Ing. Detalle 10002116": {
                    "nombre_proyecto": "994440 Ingenieria de Detalle - 10002116 - NETC",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "Javier Brañez",
                    "coordinador_shp": "Jesus Cornejo",
                    "administrador_contratos": "Edison Chocce",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.P6ERQDdUQ9yDm6BWGwpyZw&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.uqQZrtVbSSGiYEXjWmQGyQ&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.9j3mWJsQTKOewx9UmgEWwA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.NreYzLusR2OWbKVy_jeVOw&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Documentos", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.2uZs4tFoT025Uktc38rzdg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "Planos", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.ZPFHsRxcR-m5wCEbmqmsLQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                    ]
                },
            },
            "Procura": {
                "Almacenes - CGI-SHP-036-2025": {
                    "nombre_proyecto": "994440 Almacenes - C536 - AGNAV",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "",
                    "coordinador_shp": "Jorcy Aliaga",
                    "administrador_contratos": "Roger Inga",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.QZdZPMk-SsaVXq7NY59Izw&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.jMQvRbehR96-nwRmxqdevQ&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.oWWntn6KRbuvoBAfWP8v5g&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.tyQuw6S4RUuHVDyE-xRtyQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.GrNZxGb_TqGiGQswIRLiKA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1Oe1cJhUQbWoA9UUBq7asg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.xUwIJDFKQ3ewMQovEEFHwg&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
            },
            "Movimiento de Tierras": {
                "Mov Tierras - CGI-SHP-023-2025": {
                    "nombre_proyecto": "994440 Movimiento de Tierras - C523 - SGMC",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Luis Rojas",
                    "coordinador_shp": "Wilbert Salas",
                    "administrador_contratos": "Gonzalo Vargas",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.rUDsmNBhQsiEyb9o7yPuxA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.FkIYg4xtRQ6AEPaWbDQYyg&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.UzkUvdTHRAuA78DzJRIb3Q&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.9H02MxlNS1-UMkSXObElHQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.bcLpku8JR3qxB98b3wN3FA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.BdMVCgXuTICC-giaUnP31w&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.lr5qiMw_R5-XTTeeQ6q3wg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Rc9TxyDhQn-TWNhfmSrbXA&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                },
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
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.UzkUvdTHRAuA78DzJRIb3Q&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.9H02MxlNS1-UMkSXObElHQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1mZq3p5tTFmIqwiWR8mk5Q&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.dq0X5F6lT-e8xtT57iEn5Q&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.7MdcGeqmSfqPE5jDZxhoSA&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.IgD5_slUQmCzIDI3-N5_ag&viewModel=detail&moduleId=folders", "longitud_maxima": 30}
                    ]
                },
                "Paquete 02 - CGI-SHP-027-2025": {
                    "nombre_proyecto": "994440 Paquete 02 - C527 - CHEC",
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
                        {"nombre": "AS BUILT", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Cm_uBP0wRHyvXG3hRMJnGg&viewModel=detail&moduleId=folders", "longitud_maxima": 50}
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
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.MtSekBToRhGs0xIKfvL9rg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
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
            },
            "Miscelaneos": {
                "Miscelaneos - CGI-SHP-018-2026": {
                    "nombre_proyecto": "994440 Miscelaneos - C618 - CVC",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "Wilbert Salas",
                    "coordinador_shp": "Wilbert Salas",
                    "administrador_contratos": "Ivan Lima",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.jHdVvMt5R3CqypB631rH-w&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.06WrpTqRT9mbATvnXIuS6w&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.HTWXT1_XSh28xnqRAjYE5A&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.J92le3HJRHuoQaAuYM-Rfg&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.1IBRbw97TDqVebH2M2FMLA&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.8FMvuHB7TlO90XEsZn1S1g&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.yyxKEKG3TGWscWS3MNsGkA&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
                "Miscelaneos - CGI-SHP-026-2026": {
                    "nombre_proyecto": "994440 Miscelaneos - C626 - TECGU",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "Wilbert Salas",
                    "coordinador_shp": "Wilbert Salas",
                    "administrador_contratos": "Ivan Lima",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco._swNh6iESsy6Kp8bCJ7TPA&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.SolQ4vcKQfWlsnXhZQH8xg&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.wFRHaVIPTYiuRemMpY66MQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.PMa7I9K6Reqctf1IXnLm1A&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.rYpVeh1jTOOHi7CgCENCog&viewModel=detail&moduleId=folders", "longitud_maxima": 6},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.FNCcWBXSSBaSobD-y5EV7w&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                        {"nombre": "RNCs", "url": "https://acc.autodesk.com/docs/files/projects/24914611-716e-4e2b-a8a2-bf28757efbe9?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.JSU9n5b-SduwlhrE00qBJg&viewModel=detail&moduleId=folders", "longitud_maxima": 7}
                    ]
                },
            }
        },
        "998316 Nuevo Muelle San Nicolas": {
            "config": {
                "campos_personalizados": {
                    "10295605": "REV",
                    "10295608": "PENDIENTE POR",
                    "10295607": "FECHA",
                    "10295613": "TIPO DOC",
                    "10295610": "AREA",
                    "10295611": "FASE",
                    "10295612": "DISCIPLINA",
                    "10295606": "ESTADO"
                },
                "id_campo_fecha": "10295607",
                "id_campo_pendiente": "10295608"
            },
            "EPC Nuevo Muelle": {
                "CGI-SHP-066-2025 Construccion Nuevo Muelle SN": {
                    "nombre_proyecto": "EPC Nuevo Muelle SN",
                    "tipo_contrato": "Con Supervisión",
                    "nombre_supervisor": "Sergio Saenz",
                    "coordinador_shp": "Melchor Almeyda / Jose Herrera",
                    "administrador_contratos": "Sergio Casado",
                    "urls_excluidas": [],
                    "carpetas_config": [
                        {
                            "nombre": "Planos As Built",
                            "url": "https://acc.autodesk.com/docs/files/projects/3fe40740-4483-4ed9-895b-9cb579cd7b1d?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.JYeXkA13Quiif-P-f5Ea5g&viewModel=detail&moduleId=folders",
                            "longitud_maxima": 30
                        }
                    ]
                }
            }
        },
        "I25G50 Subestacion Mina 02": {
            "config": {
                "campos_personalizados": {
                    "6805049": "REV",
                    "6805036": "PENDIENTE POR",
                    "6805048": "FECHA",
                    "6805037": "ESTADO"
                },
                "id_campo_fecha": "6805048",
                "id_campo_pendiente": "6805036"
            },
            "Ingenieria": {
                "Ing. Detalle CGI-SHP-016-2026": {
                    "nombre_proyecto": "I25G50 Ingenieria de Detalle - C616 - CESEL",
                    "tipo_contrato": "Tripartito",
                    "nombre_supervisor": "",
                    "coordinador_shp": "M. Yaranga / W. Benites",
                    "administrador_contratos": "Marco Escobedo",
                    "urls_excluidas": [
                        "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.kdWoi4hPRVW1MROh01b5dg&viewModel=detail&moduleId=folders",
                        "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Z6yRZ6zMTyyDSNy6UTk1hg&viewModel=detail&moduleId=folders",
                    ],
                    "carpetas_config": [
                        {"nombre": "Ingenieria", "url": "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.LCO0kFATToqAkxt0AcK1BQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.6h13bnP8Qx6upoXYPgmA5g&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Procura", "url": "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.hWIkzFIbTQyvdevWWgu_jA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.pVHQHFJuS62xO1Iv8haaMQ&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/18ec0b1a-6335-472d-95bb-854a9259704e?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.i94W9j7mS4e57y4YYrmrAg&viewModel=detail&moduleId=folders", "longitud_maxima": 7},
                    ]
                }
            },
        },
        "I25G40 Mina 09 y 10": {
            "config": {
                "campos_personalizados": {
                    "12464893": "REV",
                    "12464891": "PENDIENTE POR",
                    "12464892": "FECHA",
                    "12464894": "ESTADO",
                    "12464886": "TIPO DOC",
                    "12464889": "AREA",
                    "12464888": "FASE",
                    "12464890": "API",
                    "12464880": "EMPRESA",
                    "12464887": "DISCIPLINA"
                },
                "id_campo_fecha": "12464892",
                "id_campo_pendiente": "12464891"
            },
            "Estudios": {
                "CGI-SHP-030-2026 Estudio Geotecnico y Topografico": {
                    "nombre_proyecto": "Estudio Geotecnico y Topografico Mina 09 y 10",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "",
                    "coordinador_shp": "Jesus Cornejo",
                    "administrador_contratos": "Marco Escobedo",
                    "urls_excluidas": [],
                    "carpetas_config": [
                        {"nombre": "Estudio", "url": "https://acc.autodesk.com/docs/files/projects/4b6267f3-3668-46b1-ad43-7736ac60d7c1?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.uG9xlgHiS7K_YXJCsemrwQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "Cronograma", "url": "https://acc.autodesk.com/docs/files/projects/4b6267f3-3668-46b1-ad43-7736ac60d7c1?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.XrfMm5RpRG-UYRf7hGUAzQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "Cartas", "url": "https://acc.autodesk.com/docs/files/projects/4b6267f3-3668-46b1-ad43-7736ac60d7c1?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.fhjB6fE6R6-deJ-UKjnv6A&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "Valorizaciones", "url": "https://acc.autodesk.com/docs/files/projects/4b6267f3-3668-46b1-ad43-7736ac60d7c1?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.SIWEjMAOThqbni1Qx7FE0w&viewModel=detail&moduleId=folders", "longitud_maxima": 10}
                    ]
                }
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
                    "11981551": "ESTADO",
                    "11981546": "AREA",
                    "11981543": "TIPO DOC",
                    "11981544": "DISCIPLINA",
                    "11981545": "FASE"
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
                    "5737742": "ESTADO",
                    "5737729": "DISCIPLINA",
                    "5737728": "TIPO DOC",
                    "5737738": "AREA",
                    "5737737": "FASE"
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
        },
        "Paquete 07 Lima": {
            "config": {
                "campos_personalizados": {
                    "5737741": "FECHA",
                    "5737740": "PENDIENTE POR",
                    "5737742": "ESTADO",
                    "5737728": "TIPO DOC",
                    "5737738": "AREA",
                    "5737737": "FASE",
                    "5737729": "DISCIPLINA"
                },
                "id_campo_fecha": "5737741",
                "id_campo_pendiente": "5737740"
            },
            "Lima": {
                "CGI-SHP-013-2026": {
                    "nombre_proyecto": "Plan de Conservacion PQ07",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "",
                    "coordinador_shp": "R. Gamarra M. Wang",
                    "administrador_contratos": "Marco Escobedo",
                    "urls_excluidas": [],
                    "carpetas_config": [
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.XLSVu_D3QlSU-lV6QdJ5Gg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFI", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.effHI1b5QICsDAQBcUrS9Q&viewModel=detail&moduleId=folders", "longitud_maxima": 3},
                        {"nombre": "RNC", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.8DIKwX-bTyq0l2pX9a4_Fg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "VALORIZACIONES", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.Wk_us_7_TtCP1-kNSPpXTA&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "CARTAS", "url": "https://acc.autodesk.com/docs/files/projects/d11143b8-25a1-4eaa-a49c-d491c738f06c?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.czQUOAzcTgS5APfjX_oUlg&viewModel=detail&moduleId=folders", "longitud_maxima": 10}
                    ]
                }
            }
        }
    },
    "Proyecto con Operaciones": {
        "Sistema de Compensacion": {
            "config": {
                "campos_personalizados": {
                    "12102611": "REV",
                    "12102609": "PENDIENTE POR",
                    "12102610": "FECHA",
                    "12102612": "ESTADO"
                },
                "id_campo_fecha": "12102610",
                "id_campo_pendiente": "12102609"
            },
            "EPC": {
                "CGI-SHP-039-2026": {
                    "nombre_proyecto": "I26A40 - SISTEMA DE COMPENSACION DE TOLVA DE 4,500",
                    "tipo_contrato": "Contrato Directo",
                    "nombre_supervisor": "Manuel Loyola",
                    "coordinador_shp": "Jesus Cornejo",
                    "administrador_contratos": "Marco Escobedo",
                    "urls_excluidas": [],
                    "carpetas_config": [
                        {"nombre": "Ing Basica", "url": "https://acc.autodesk.com/docs/files/projects/cced61e4-8a0f-47a5-81f2-1aa6e85dafe5?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.zjxfTbUhRxG-yQ7Qa46MgQ&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "Construccion", "url": "https://acc.autodesk.com/docs/files/projects/cced61e4-8a0f-47a5-81f2-1aa6e85dafe5?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.8uO3tWdkT8u_K9Aj_VNCIg&viewModel=detail&moduleId=folders", "longitud_maxima": 10},
                        {"nombre": "RFIs", "url": "https://acc.autodesk.com/docs/files/projects/cced61e4-8a0f-47a5-81f2-1aa6e85dafe5?folderUrn=urn%3Aadsk.wipprod%3Afs.folder%3Aco.87jPmfFASbeZHlDGpTpx4w&viewModel=detail&moduleId=folders", "longitud_maxima": 10}
                    ]
                }
            }
        }
    }
}

ESTRUCTURA_MAESTRA, SHA_ESTRUCTURA = cargar_estructura_github(ESTRUCTURA_MAESTRA)
COL_CODIGO = "CODIGO"
COL_DESCRIPCION = "DESCRIPCION"

# ====================================================================
# --- 2. CREDENCIALES (SEGURAS NUBE / LOCAL) ---
# ====================================================================
try:
    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
    EMAIL_USER = st.secrets.get("EMAIL_USER", "")
    EMAIL_PASS = st.secrets.get("EMAIL_PASSWORD", "")
except Exception:
    CLIENT_ID = os.environ.get("CLIENT_ID", "")
    CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
    EMAIL_USER = os.environ.get("EMAIL_USER", "")
    EMAIL_PASS = os.environ.get("EMAIL_PASSWORD", "")

# ====================================================================
# --- 3. FUNCIONES DE EXPORTACIÓN A PDF, EXCEL Y CORREO ---
# ====================================================================
def obtener_colores_estado(estado):
    estado = str(estado).strip().upper()
    if estado == "CERRADO":
        return "#d4edda", "#155724"
    elif "CONTRATISTA" in estado or "EMPRESA" in estado:
        return "#fff3cd", "#856404"
    elif estado == "SUPERVISION":
        return "#f8d7da", "#721c24"
    else:
        return "#cce5ff", "#004085"

def generar_tabla_html(df):
    """Genera una tabla HTML con estilos en línea a partir de un DataFrame de Pandas"""
    html = df.to_html(index=False)
    html = html.replace('<table border="1" class="dataframe">', '<table cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; text-align: center; border: 1px solid #ddd;">')
    html = html.replace('<th>', '<th style="background-color: #1A365D; color: white; padding: 8px; border: 1px solid #ddd;">')
    html = html.replace('<td>', '<td style="padding: 5px; border: 1px solid #ddd;">')
    html = html.replace('<tr>', '<tr style="border: 1px solid #ddd;">')
    return html

def enviar_correo_con_pdf(destinatario, asunto, cuerpo_texto, cuerpo_html, pdf_bytes, nombre_archivo):
    try:
        msg = EmailMessage()
        msg['Subject'] = asunto
        msg['From'] = EMAIL_USER
        msg['To'] = destinatario

        msg.set_content(cuerpo_texto)
        msg.add_alternative(cuerpo_html, subtype='html')

        msg.add_attachment(
            pdf_bytes.getvalue(),
            maintype='application',
            subtype='pdf',
            filename=nombre_archivo
        )

        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(msg)

        return True, "¡Correo enviado exitosamente!"
    except Exception as e:
        return False, f"Error al enviar correo: {str(e)}"

def exportar_df_a_excel(df: pd.DataFrame, titulo: str, metadatos: dict, ruta_logo: str, col_pendiente: str = None) -> io.BytesIO:
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Documentos"

    if os.path.exists(ruta_logo):
        img = ExcelImage(ruta_logo)
        img.height = 40
        img.width = 120
        ws.add_image(img, 'H1')

    ws['A1'] = titulo
    ws['A1'].font = Font(size=14, bold=True)

    ws['A3'] = "Fecha de Exportación:"
    ws['B3'] = metadatos['fecha']
    ws['A4'] = "Alcance:"
    ws['B4'] = metadatos['alcance']
    ws['A5'] = "Subproyecto:"
    ws['B5'] = metadatos['subproyecto']

    if "tipo_obra" in metadatos:
        ws['A6'] = "Tipo de Obra:"
        ws['B6'] = metadatos['tipo_obra']
    if "contrato" in metadatos:
        ws['A7'] = "Contrato:"
        ws['B7'] = metadatos['contrato']

    for row in range(3, 8):
        ws[f'A{row}'].font = Font(bold=True)

    fila_inicio = 9
    for col_num, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=fila_inicio, column=col_num, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = 22

    col_idx_pendiente = None
    if col_pendiente and col_pendiente in df.columns:
        col_idx_pendiente = df.columns.tolist().index(col_pendiente)

    for r_idx, row_data in enumerate(df.values, fila_inicio + 1):
        bg_hex, font_hex = "FFFFFF", "000000"

        if col_idx_pendiente is not None:
            estado = row_data[col_idx_pendiente]
            bg, text = obtener_colores_estado(estado)
            bg_hex = bg.replace("#", "")
            font_hex = text.replace("#", "")

        for c_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value="" if pd.isna(value) else str(value))
            if col_idx_pendiente is not None:
                cell.fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
                cell.font = Font(color=font_hex)

    fila_pie = fila_inicio + len(df) + 2
    pie_cell = ws.cell(row=fila_pie, column=1, value="Desarrollado por Roberto Cano - Shougang Hierro Perú")
    pie_cell.font = Font(italic=True, color="4A5568")

    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generar_pdf_subproyecto(metadatos, df_resumen_global, ruta_logo, resumenes_por_contrato):
    buffer = io.BytesIO()

    def agregar_pie_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(colors.HexColor("#4A5568"))
        canvas.drawString(doc.leftMargin, 15, "Desarrollado por Roberto Cano - Shougang Hierro Perú")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(letter),
        leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=35
    )
    ancho_util = 742
    elements = []
    styles = getSampleStyleSheet()

    cell_style = ParagraphStyle(name="CellText", parent=styles["Normal"], fontSize=7, leading=9)
    header_style = ParagraphStyle(name="HeaderCellText", parent=styles["Normal"], fontSize=7, leading=9, textColor=colors.whitesmoke, fontName="Helvetica-Bold")
    link_style = ParagraphStyle(name="Link", parent=styles["Normal"], fontSize=8, textColor=colors.blue)

    bloque_global = []

    if os.path.exists(ruta_logo):
        logo = RLImage(ruta_logo, width=100, height=30, kind='proportional')
    else:
        logo = Paragraph("")

    meta_text = (f"<b><font size=12>Reporte Consolidado del Subproyecto</font></b><br/><br/>"
                 f"<b>Fecha de Emisión:</b> {metadatos['fecha']}<br/>"
                 f"<b>Alcance:</b> {metadatos['alcance']}<br/>"
                 f"<b>Subproyecto:</b> {metadatos['subproyecto']}")

    tabla_cabecera = Table([[logo, Paragraph(meta_text, styles["Normal"])]], colWidths=[120, 600])
    tabla_cabecera.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'LEFT'),
    ]))

    bloque_global.append(tabla_cabecera)
    bloque_global.append(Spacer(1, 15))

    bloque_global.append(Paragraph("<b>1. Resumen General por Contrato</b>", styles["Heading3"]))
    bloque_global.append(Spacer(1, 5))

    if not df_resumen_global.empty:
        encabezados_g = [Paragraph(str(col), header_style) for col in df_resumen_global.columns]
        data_matrix_g = [encabezados_g]
        for _, row in df_resumen_global.iterrows():
            fila = [Paragraph("" if pd.isna(val) else str(val), cell_style) for val in row]
            data_matrix_g.append(fila)

        anchos_resumen_g = []
        for col in df_resumen_global.columns:
            if col == "Tipo de Obra": anchos_resumen_g.append(0.15 * ancho_util)
            elif col == "Contrato": anchos_resumen_g.append(0.25 * ancho_util)
            elif col == "Total": anchos_resumen_g.append(0.10 * ancho_util)
            elif col == "Avance (%)": anchos_resumen_g.append(0.10 * ancho_util)
            else: anchos_resumen_g.append((0.40 * ancho_util) / max(1, (len(df_resumen_global.columns) - 4)))

        tabla_resumen_g = Table(data_matrix_g, repeatRows=1, colWidths=anchos_resumen_g)
        tabla_resumen_g.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
        ]))
        bloque_global.append(tabla_resumen_g)

    bloque_global.append(Spacer(1, 15))

    if not df_resumen_global.empty and "Contrato" in df_resumen_global.columns:
        columnas_roles = [c for c in df_resumen_global.columns if c not in ["Tipo de Obra", "Contrato", "Total", "Avance (%)"]]

        fig, ax = plt.subplots(figsize=(8, 2.8))
        x = np.arange(len(df_resumen_global["Contrato"]))
        ancho_barra = 0.8 / max(1, len(columnas_roles))
        colores = ['#0052cc', '#66b2ff', '#ff3333', '#ff9999', '#33cc33', '#99ff99']

        for i, col in enumerate(columnas_roles):
            desplazamiento = x + (i - len(columnas_roles)/2 + 0.5) * ancho_barra
            ax.bar(desplazamiento, df_resumen_global[col], ancho_barra, label=col, color=colores[i % len(colores)])

        ax.set_title("Estado Documentario por Contrato y Responsable", fontsize=10, fontweight='bold')
        ax.set_xticks(x)
        etiquetas = [c[:20] + "..." if len(c) > 20 else c for c in df_resumen_global["Contrato"]]
        ax.set_xticklabels(etiquetas, rotation=15, ha='right', fontsize=8)
        ax.legend(fontsize=8, bbox_to_anchor=(1.01, 1), loc='upper left')
        plt.tight_layout()

        img_data_global = io.BytesIO()
        plt.savefig(img_data_global, format='png', dpi=200, bbox_inches='tight')
        plt.close()
        img_data_global.seek(0)
        bloque_global.append(RLImage(img_data_global, width=450, height=155))

    elements.append(KeepTogether(bloque_global))

    if resumenes_por_contrato:
        nota_enlace = Paragraph(
            f'<i>Nota: Para mayor detalle de la documentación, revisar el dashboard en la web: <a href="{URL_DASHBOARD_WEB}" color="blue">{URL_DASHBOARD_WEB}</a></i>',
            link_style
        )

        for contrato_nombre, df_res_contrato in resumenes_por_contrato.items():
            if df_res_contrato.empty:
                continue

            elements.append(PageBreak())
            bloque_contrato = []

            bloque_contrato.append(nota_enlace)
            bloque_contrato.append(Spacer(1, 15))
            bloque_contrato.append(Paragraph(f"<b>Resumen General por Carpeta - Contrato: {contrato_nombre}</b>", styles["Heading3"]))
            bloque_contrato.append(Spacer(1, 10))

            encabezados_c = [Paragraph(str(col), header_style) for col in df_res_contrato.columns]
            data_matrix_c = [encabezados_c]
            for _, row in df_res_contrato.iterrows():
                fila = [Paragraph("" if pd.isna(val) else str(val), cell_style) for val in row]
                data_matrix_c.append(fila)

            anchos_c = []
            for col in df_res_contrato.columns:
                if col == "Carpeta": anchos_c.append(0.15 * ancho_util)
                elif col == "Mayor Retraso": anchos_c.append(0.35 * ancho_util)
                else: anchos_c.append((0.50 * ancho_util) / max(1, (len(df_res_contrato.columns) - 2)))

            tabla_c = Table(data_matrix_c, repeatRows=1, colWidths=anchos_c)
            tabla_c.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2B6CB0")),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
            ]))
            bloque_contrato.append(tabla_c)
            bloque_contrato.append(Spacer(1, 20))

            columnas_grafico_c = ["Cerrados"] + [c for c in df_res_contrato.columns if c not in ["Carpeta", "Cerrados", "Total", "Mayor Retraso"]]
            fig, ax = plt.subplots(figsize=(8, 3.2))
            x = np.arange(len(df_res_contrato["Carpeta"]))
            ancho_barra = 0.8 / max(1, len(columnas_grafico_c))
            colores = ['#0052cc', '#66b2ff', '#ff3333', '#ff9999', '#33cc33', '#99ff99']

            for i, col in enumerate(columnas_grafico_c):
                desplazamiento = x + (i - len(columnas_grafico_c)/2 + 0.5) * ancho_barra
                ax.bar(desplazamiento, df_res_contrato[col], ancho_barra, label=col, color=colores[i % len(colores)])

            ax.set_title("Estado Documentario por Carpeta y Responsable", fontsize=10, fontweight='bold')
            ax.set_xticks(x)
            ax.set_xticklabels(df_res_contrato["Carpeta"], rotation=15, ha='right', fontsize=8)
            ax.legend(fontsize=8, bbox_to_anchor=(1.01, 1), loc='upper left')
            plt.tight_layout()

            img_data_c = io.BytesIO()
            plt.savefig(img_data_c, format='png', dpi=200, bbox_inches='tight')
            plt.close()
            img_data_c.seek(0)
            bloque_contrato.append(RLImage(img_data_c, width=450, height=170))

            elements.append(KeepTogether(bloque_contrato))

    doc.build(elements, onFirstPage=agregar_pie_pagina, onLaterPages=agregar_pie_pagina)
    buffer.seek(0)
    return buffer

def exportar_df_a_pdf(df: pd.DataFrame, titulo: str, metadatos: dict, ruta_logo: str, col_pendiente: str = None) -> io.BytesIO:
    buffer = io.BytesIO()

    def agregar_pie_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(colors.HexColor("#4A5568"))
        canvas.drawString(doc.leftMargin, 15, "Desarrollado por Roberto Cano - Shougang Hierro Perú")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(letter),
        leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=35
    )
    ancho_util = 742
    elements = []
    styles = getSampleStyleSheet()

    cell_style = ParagraphStyle(name="CellText", parent=styles["Normal"], fontSize=6.5, leading=8.5)
    header_style = ParagraphStyle(name="HeaderCellText", parent=styles["Normal"], fontSize=6.5, leading=8.5, textColor=colors.whitesmoke, fontName="Helvetica-Bold")

    if os.path.exists(ruta_logo):
        logo = RLImage(ruta_logo, width=100, height=30, kind='proportional')
    else:
        logo = Paragraph("")

    meta_text = (f"<b><font size=12>{titulo}</font></b><br/><br/>"
                 f"<b>Fecha de Emisión:</b> {metadatos['fecha']}<br/>"
                 f"<b>Alcance:</b> {metadatos['alcance']}<br/>"
                 f"<b>Subproyecto:</b> {metadatos['subproyecto']}")

    if "tipo_obra" in metadatos:
        meta_text += f"<br/><b>Tipo de Obra:</b> {metadatos['tipo_obra']}"
    if "contrato" in metadatos:
        meta_text += f"<br/><b>Contrato:</b> {metadatos['contrato']}"

    tabla_cabecera = Table([[logo, Paragraph(meta_text, styles["Normal"])]], colWidths=[120, 600])
    tabla_cabecera.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'LEFT'),
    ]))

    elements.append(tabla_cabecera)
    elements.append(Spacer(1, 15))

    encabezados = [Paragraph(str(col), header_style) for col in df.columns]
    data_matrix = [encabezados]

    estilos_tabla = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
    ]

    for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
        fila = []
        if col_pendiente and col_pendiente in df.columns:
            bg_color, text_color = obtener_colores_estado(row[col_pendiente])
        else:
            bg_color = "#ffffff" if row_idx % 2 != 0 else "#f9fafb"
            text_color = "#000000"

        estilos_tabla.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor(bg_color)))

        for col_name in df.columns:
            val = row[col_name]
            text_str = ("" if pd.isna(val) else str(val)).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if col_pendiente and col_pendiente in df.columns:
                text_str = f'<font color="{text_color}">{text_str}</font>'
            fila.append(Paragraph(text_str, cell_style))
        data_matrix.append(fila)

    anchos_columnas = []
    for col in df.columns:
        col_u = str(col).upper()
        if col_u == "CARPETA": anchos_columnas.append(0.15 * ancho_util)
        elif col_u == "MAYOR RETRASO": anchos_columnas.append(0.35 * ancho_util)
        elif col_u == "CODIGO": anchos_columnas.append(0.18 * ancho_util)
        elif col_u == "DESCRIPCION": anchos_columnas.append(0.35 * ancho_util)
        elif col_u == "TIPO DE OBRA": anchos_columnas.append(0.15 * ancho_util)
        elif col_u == "CONTRATO": anchos_columnas.append(0.20 * ancho_util)
        elif col_u == "DOCUMENTOS": anchos_columnas.append(0.10 * ancho_util)
        elif col_u == "AVANCE (%)": anchos_columnas.append(0.08 * ancho_util)
        elif col_u in ["REV", "DIAS"]: anchos_columnas.append(0.06 * ancho_util)
        elif col_u == "FECHA": anchos_columnas.append(0.09 * ancho_util)
        elif col_u == "ESTADO": anchos_columnas.append(0.11 * ancho_util)
        elif col_u == "PENDIENTE POR": anchos_columnas.append(0.15 * ancho_util)
        else: anchos_columnas.append(0.10 * ancho_util)

    factor_norm = ancho_util / max(1e-6, sum(anchos_columnas))
    anchos_columnas = [w * factor_norm for w in anchos_columnas]

    tabla = Table(data_matrix, repeatRows=1, colWidths=anchos_columnas)
    tabla.setStyle(TableStyle(estilos_tabla))
    elements.append(tabla)

    doc.build(elements, onFirstPage=agregar_pie_pagina, onLaterPages=agregar_pie_pagina)
    buffer.seek(0)
    return buffer


# ====================================================================
# --- 4. LÓGICA DE API AUTODESK (SIN POOLS ANIDADOS + TIMEOUTS) ---
# ====================================================================
_TOKEN_LOCK = threading.Lock()
_TOKEN_CACHE = {"value": None, "expires_at": 0.0}
_LOCAL_HTTP = threading.local()


def _get_session() -> requests.Session:
    sess = getattr(_LOCAL_HTTP, "session", None)
    if sess is None:
        sess = requests.Session()
        adapter = HTTPAdapter(pool_connections=12, pool_maxsize=12, max_retries=1)
        sess.mount("https://", adapter)
        sess.mount("http://", adapter)
        _LOCAL_HTTP.session = sess
    return sess


def obtener_periodo_actual():
    tz_peru = pytz.timezone('America/Lima')
    ahora = datetime.now(tz_peru)
    fecha_corte = (ahora - timedelta(hours=4)).strftime("%Y-%m-%d")
    return f"{fecha_corte}_04AM"


def obtener_ids_excluidos(urls):
    ids = set()
    for url in urls:
        if "folderUrn=" in url:
            ids.add(urllib.parse.unquote(url.split("folderUrn=")[1].split("&")[0]).replace("%3A", ":"))
        elif "folders/" in url:
            ids.add(urllib.parse.unquote(url.split("folders/")[1].split("/")[0].split("?")[0]).replace("%3A", ":"))
    return ids


def obtener_token(forzar_renovacion: bool = False):
    ahora = time.time()
    with _TOKEN_LOCK:
        if not forzar_renovacion and _TOKEN_CACHE["value"] and ahora < _TOKEN_CACHE["expires_at"]:
            return _TOKEN_CACHE["value"]

        url = "https://developer.api.autodesk.com/authentication/v2/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
        data = {"grant_type": "client_credentials", "scope": "data:read"}
        try:
            response = _get_session().post(url, headers=headers, data=data, auth=(CLIENT_ID, CLIENT_SECRET), timeout=15)
            if response.status_code == 200:
                payload = response.json()
                tok = payload.get("access_token")
                ttl = float(payload.get("expires_in", 3599))
                _TOKEN_CACHE["value"] = tok
                _TOKEN_CACHE["expires_at"] = time.time() + max(60.0, ttl - 120.0)
                return tok
        except Exception:
            pass
        return None


def _request_ads_get(url: str, region_inicial: str = "US", timeout: int = 15):
    session = _get_session()
    token = obtener_token()
    region = region_inicial
    resp = None

    for intento in range(4):
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "x-ads-region": region
        }
        try:
            resp = session.get(url, headers=headers, timeout=timeout)
        except Exception:
            time.sleep(1.5)
            continue

        if resp.status_code == 429:
            retry_after = float(resp.headers.get("Retry-After", 2 * (intento + 1)))
            time.sleep(min(retry_after, 6.0))
            continue
        if resp.status_code == 401:
            token = obtener_token(forzar_renovacion=True)
            continue
        if resp.status_code != 200 and region == "US" and intento == 0:
            region = "EMEA"
            continue
        return resp
    return resp


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
    url = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/items/{urllib.parse.quote(clean_item_id, safe='')}/tip"
    response = _request_ads_get(url, region_inicial="US", timeout=15)
    if response is not None and response.status_code == 200:
        return response.json().get("data", {})
    return {}


def procesar_archivo_pdf(item, project_id, headers, ruta_actual, name, campos_personalizados):
    item_id = item.get("id", "")
    token_val = headers.get("Authorization", "").replace("Bearer ", "") or obtener_token()

    registro_plano = {
        "NOMBRE_DEL_ARCHIVO": name,
        "RUTA_EN_AUTODESK": ruta_actual,
        "ITEM_ID": item_id
    }
    deconstruir_diccionario(item, "ITEM", registro_plano)

    tip_version = obtener_tip_version_metadata(token=token_val, project_id=project_id, item_id=item_id)
    if tip_version:
        deconstruir_diccionario(tip_version, "VERSION", registro_plano)

    vid = (
        (tip_version or {}).get("id", "")
        or item.get("relationships", {}).get("tip", {}).get("data", {}).get("id", "")
    )
    if vid:
        vid_encoded = urllib.parse.quote(vid, safe='')
        url_attr = f"https://developer.api.autodesk.com/bim360/docs/v1/projects/{project_id.replace('b.', '')}/versions/{vid_encoded}/custom-attributes"
        res_attr = _request_ads_get(url_attr, region_inicial="US", timeout=15)

        if res_attr is not None and res_attr.status_code == 200:
            try:
                attrs = res_attr.json()
            except Exception:
                attrs = []
            for attr in attrs:
                attr_id = str(attr.get("id"))
                attr_name = str(attr.get("name", "")).strip()
                attr_val = str(attr.get("value", "")).strip()
                registro_plano[f"CAMPO_CUSTOM_{attr_id}"] = attr_val
                if attr_name:
                    registro_plano[f"ATTR_{attr_name.upper()}"] = attr_val

    url_web = item.get("links", {}).get("webView", {}).get("href", "")
    if not url_web and tip_version:
        url_web = tip_version.get("links", {}).get("webView", {}).get("href", "")
    if not url_web:
        parent_id = item.get("relationships", {}).get("parent", {}).get("data", {}).get("id", "")
        pid_limpio = project_id.replace("b.", "")
        url_web = (
            f"https://acc.autodesk.com/docs/files/projects/{pid_limpio}"
            f"?folderUrn={urllib.parse.quote(parent_id, safe='')}"
            f"&entityId={urllib.parse.quote(item_id, safe='')}"
            f"&viewModel=detail&moduleId=folders"
        )
    registro_plano["URL_DOC"] = url_web

    return registro_plano


def _recolectar_pdfs_recursivo(project_id, folder_id, longitud_maxima, ids_excluidos, ruta_actual, cola_pdfs):
    """Recorre carpetas de forma ligera y acumula los PDFs sin crear sub-pools por carpeta."""
    if folder_id.replace("%3A", ":") in ids_excluidos:
        return

    encoded_fid = urllib.parse.quote(urllib.parse.unquote(folder_id), safe="")
    url_api = f"https://developer.api.autodesk.com/data/v1/projects/{project_id}/folders/{encoded_fid}/contents"

    while url_api:
        res = _request_ads_get(url_api, region_inicial="US", timeout=15)
        if res is None or res.status_code != 200:
            break

        try:
            data = res.json()
        except Exception:
            break

        for item in data.get("data", []):
            tipo = item.get("type")
            attributes = item.get("attributes", {})
            name = str(attributes.get("displayName", "Sin Nombre")).strip()

            if tipo == "folders":
                sub_urn_raw = urllib.parse.unquote(item["id"]).replace("%3A", ":")
                if sub_urn_raw in ids_excluidos:
                    continue
                if len(name) <= longitud_maxima:
                    _recolectar_pdfs_recursivo(
                        project_id, sub_urn_raw, longitud_maxima,
                        ids_excluidos, f"{ruta_actual} > {name}", cola_pdfs
                    )
            elif tipo == "items":
                if ".pdf" in name.lower():
                    cola_pdfs.append((item, project_id, ruta_actual, name))

        enlaces = data.get("links", {})
        url_api = enlaces["next"].get("href") if "next" in enlaces else None


def escanear_recursivo(project_id, folder_id, headers, lista_datos, longitud_maxima, ids_excluidos, campos_personalizados, ruta_actual="Raíz"):
    cola_pdfs = []
    _recolectar_pdfs_recursivo(project_id, folder_id, longitud_maxima, ids_excluidos, ruta_actual, cola_pdfs)

    if not cola_pdfs:
        return

    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = [
            executor.submit(procesar_archivo_pdf, itm, pid, headers, r_act, nom, campos_personalizados)
            for itm, pid, r_act, nom in cola_pdfs
        ]
        for future in as_completed(futures):
            try:
                res_item = future.result()
                if res_item:
                    lista_datos.append(res_item)
            except Exception:
                continue


def _procesar_carpeta_config(config_carpeta, ids_excluidos, campos_personalizados, id_campo_fecha, id_campo_pendiente, col_pendiente, orden_columnas_deseado):
    url_carpeta = config_carpeta.get("url")
    nombre_pestana = config_carpeta.get("nombre", "Sin_Nombre")
    longitud_maxima = config_carpeta.get("longitud_maxima", 10)

    if not url_carpeta:
        return nombre_pestana, None, None

    try:
        project_id_raw = url_carpeta.split("projects/")[1].split("?")[0].split("/")[0]
        project_id = project_id_raw if project_id_raw.startswith("b.") else "b." + project_id_raw

        if "folderUrn=" in url_carpeta:
            folder_id = urllib.parse.unquote(url_carpeta.split("folderUrn=")[1].split("&")[0]).replace("%3A", ":")
        elif "folders/" in url_carpeta:
            folder_id = urllib.parse.unquote(url_carpeta.split("folders/")[1].split("/")[0].split("?")[0]).replace("%3A", ":")
        else:
            return nombre_pestana, None, None

        datos_filtrados = []
        headers = {"Authorization": f"Bearer {obtener_token() or ''}", "x-ads-region": "US"}
        escanear_recursivo(project_id, folder_id, headers, datos_filtrados, longitud_maxima, ids_excluidos, campos_personalizados, nombre_pestana)

        if not datos_filtrados:
            return nombre_pestana, None, None

        df_completo = pd.DataFrame(datos_filtrados)

        if "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION" not in df_completo.columns:
            df_completo["ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"] = ""
        if "URL_DOC" not in df_completo.columns:
            df_completo["URL_DOC"] = ""

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

        diccionario_renombres = {
            "NOMBRE_DEL_ARCHIVO": COL_CODIGO,
            "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION": COL_DESCRIPCION,
        }
        for campo_id, nombre_columna in campos_personalizados.items():
            diccionario_renombres[f"CAMPO_CUSTOM_{campo_id}"] = nombre_columna

        df_todo = df_completo.rename(columns=diccionario_renombres).copy()
        if col_pendiente in df_todo.columns:
            df_todo.loc[df_todo[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0
        df_todo.insert(0, "CARPETA", nombre_pestana)

        columnas_finales = (
            ["NOMBRE_DEL_ARCHIVO", "ITEM_ATTRIBUTES_EXTENSION_DATA_DESCRIPTION"]
            + [f"CAMPO_CUSTOM_{c}" for c in campos_personalizados.keys()]
            + ["DIAS", "URL_DOC"]
        )
        df = df_completo[columnas_finales].copy()
        df = df.rename(columns=diccionario_renombres)
        columnas_existentes = [col for col in orden_columnas_deseado if col in df.columns]
        df = df[columnas_existentes]

        if col_pendiente in df.columns:
            df.loc[df[col_pendiente].astype(str).str.strip().str.upper() == "CERRADO", "DIAS"] = 0

        return nombre_pestana, df, df_todo
    except Exception:
        return nombre_pestana, None, None


@st.cache_data(show_spinner=False)
def extraer_datos_contrato_completo(contrato_key, info_contrato, campos_personalizados, id_campo_fecha, id_campo_pendiente, periodo_cache):
    token = obtener_token()
    if not token:
        return {"datos": {}, "datos_completos": {}, "hora_actualizacion": "Sin token"}

    urls_excluidas_raw = info_contrato.get("urls_excluidas", [])
    ids_excluidos = obtener_ids_excluidos(urls_excluidas_raw)
    carpetas_config = info_contrato.get("carpetas_config", [])

    dataframes_por_pestana = {}
    dataframes_completos_por_pestana = {}
    col_pendiente = campos_personalizados.get(id_campo_pendiente)
    orden_columnas_deseado = ["CODIGO", "REV", "DESCRIPCION", "FECHA", "ESTADO", "PENDIENTE POR", "DIAS", "URL_DOC"]

    # Procesamiento secuencial limpio de carpetas (cada carpeta ya paraleliza sus PDFs con 6 workers)
    for cfg_c in carpetas_config:
        pest, df_p, df_t = _procesar_carpeta_config(
            cfg_c, ids_excluidos, campos_personalizados,
            id_campo_fecha, id_campo_pendiente, col_pendiente, orden_columnas_deseado
        )
        if df_p is not None:
            dataframes_por_pestana[pest] = df_p
        if df_t is not None:
            dataframes_completos_por_pestana[pest] = df_t

    tz_pe = pytz.timezone('America/Lima')
    hora_cache = datetime.now(tz_pe).strftime('%d/%m/%Y %H:%M')

    return {
        "datos": dataframes_por_pestana,
        "datos_completos": dataframes_completos_por_pestana,
        "hora_actualizacion": hora_cache
    }


def _formatear_hms(segundos: float) -> str:
    total_seg = max(0, int(round(segundos)))
    hh = total_seg // 3600
    mm = (total_seg % 3600) // 60
    ss = total_seg % 60
    return f"{hh:02d}h {mm:02d}m {ss:02d}s"


def precargar_todos_los_contratos(periodo, progress_bar=None, status_text=None):
    obtener_token(forzar_renovacion=True)
    tareas = []

    for alcance, sub_dict in ESTRUCTURA_MAESTRA.items():
        if not isinstance(sub_dict, dict):
            continue
        for sub_nom, sub_data in sub_dict.items():
            if sub_nom == "config" or not isinstance(sub_data, dict):
                continue

            config_sub = sub_data.get("config", {})
            campos = config_sub.get("campos_personalizados", {})
            id_fec = config_sub.get("id_campo_fecha", "")
            id_pen = config_sub.get("id_campo_pendiente", "")

            for tipo_obra, contratos_dict in sub_data.items():
                if tipo_obra == "config" or not isinstance(contratos_dict, dict):
                    continue

                for cont_nom, info_cont in contratos_dict.items():
                    if isinstance(info_cont, dict) and info_cont.get("carpetas_config"):
                        tareas.append((cont_nom, info_cont, campos, id_fec, id_pen, periodo))

    total_tareas = max(1, len(tareas))
    inicio_global = time.perf_counter()

    for idx, t in enumerate(tareas, start=1):
        nombre_c = t[0]
        inicio_c = time.perf_counter()
        if status_text:
            status_text.caption(
                f"Cargando {idx}/{total_tareas}: {nombre_c} | "
                f"Acumulado: {_formatear_hms(time.perf_counter() - inicio_global)}"
            )
        try:
            extraer_datos_contrato_completo(*t)
        except Exception:
            pass

        dur_c = time.perf_counter() - inicio_c
        dur_global = time.perf_counter() - inicio_global
        if progress_bar:
            progress_bar.progress(min(1.0, idx / total_tareas))
        if status_text:
            status_text.caption(
                f"Listo {idx}/{total_tareas}: {nombre_c} "
                f"(Tardó: {_formatear_hms(dur_c)} | Total: {_formatear_hms(dur_global)})"
            )


def preparar_df_con_links(df_in: pd.DataFrame) -> pd.DataFrame:
    """Incrusta el enlace de ACC en la columna CODIGO para mostrarlo como LinkColumn."""
    df_vista = df_in.copy()
    if "URL_DOC" in df_vista.columns and COL_CODIGO in df_vista.columns:
        df_vista[COL_CODIGO] = df_vista.apply(
            lambda r: f"{r['URL_DOC']}#doc={str(r[COL_CODIGO]).replace(' ', '\u00A0')}"
            if pd.notna(r["URL_DOC"]) and str(r["URL_DOC"]).startswith("http")
            else str(r[COL_CODIGO]),
            axis=1
        )
    return df_vista


# ====================================================================
# --- 5. INTERFAZ WEB STREAMLIT ---
# ====================================================================
st.set_page_config(page_title="Control Documentario SHP", layout="wide", page_icon="📊")

st.markdown(
    """
    <style>
        .bloque-guia {
            background-color: #f8fafc;
            border-left: 4px solid #1A365D;
            padding: 10px 14px;
            border-radius: 4px;
            margin-bottom: 12px;
            font-size: 13.5px;
            color: #334155;
        }
        .barra-contexto {
            background-color: #eef2f6;
            padding: 8px 14px;
            border-radius: 6px;
            margin-bottom: 14px;
            font-size: 13px;
            color: #1e293b;
        }
    </style>
    """,
    unsafe_allow_html=True
)

periodo_actual = obtener_periodo_actual()
ruta_logo = "logoshp.jpeg"

col_head1, col_head2 = st.columns([5, 1])
with col_head1:
    st.title("📊 Dashboard de Control Documentario")
    st.markdown("**Institución:** Shougang Hierro Perú · Seguimiento y reportes desde Autodesk Construction Cloud")

with col_head2:
    try:
        st.image(ruta_logo, width=130)
    except Exception:
        pass

col_espacio1, col_logo, col_espacio2 = st.sidebar.columns([1, 2.5, 1])
with col_logo:
    try:
        st.image(ruta_logo, use_container_width=True)
    except Exception:
        pass

st.sidebar.markdown(
    """
    <div style="text-align: center; font-size: 13.5px; color: #4A5568; margin-top: -10px; margin-bottom: 20px;">
        Desarrollado por<br>
        <b>Roberto Cano</b>
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.title("1. Estructura de Proyectos")

grupo_activo = st.sidebar.selectbox("Tipo de Alcance:", list(ESTRUCTURA_MAESTRA.keys()))

subproyectos_dict = ESTRUCTURA_MAESTRA.get(grupo_activo, {})
if not subproyectos_dict:
    st.warning(f"No hay subproyectos configurados en '{grupo_activo}'.")
    st.stop()

subproyecto_activo = st.sidebar.selectbox("Subproyecto:", list(subproyectos_dict.keys()))
subproyecto_data = subproyectos_dict.get(subproyecto_activo, {})

config_sub = subproyecto_data.get("config", {
    "campos_personalizados": {"6011837": "REV", "6011840": "PENDIENTE POR", "6011839": "FECHA", "6011838": "ESTADO"},
    "id_campo_fecha": "6011839",
    "id_campo_pendiente": "6011840"
})

campos_pers_activos = config_sub.get("campos_personalizados", {})
id_fec_activo = config_sub.get("id_campo_fecha", "")
id_pen_activo = config_sub.get("id_campo_pendiente", "")

tipos_obra_dict = {k: v for k, v in subproyecto_data.items() if k != "config"}
if not tipos_obra_dict:
    st.warning(f"No hay tipos de obra o contratos configurados en '{subproyecto_activo}'.")
    st.stop()

tipo_obra_activo = st.sidebar.selectbox("Tipo de Obra:", list(tipos_obra_dict.keys()))

contratos_dict = tipos_obra_dict.get(tipo_obra_activo, {})
if not contratos_dict:
    st.warning(f"No hay contratos configurados en '{tipo_obra_activo}'.")
    st.stop()

contrato_activo = st.sidebar.selectbox("Contrato / Componente:", list(contratos_dict.keys()))
info_contrato = contratos_dict[contrato_activo]

# --- SELECTOR DE MODO DE VISTA Y ACCESO ADMINISTRADOR ---
st.sidebar.markdown("---")
st.sidebar.subheader("2. Navegación")

if "es_admin" not in st.session_state:
    st.session_state["es_admin"] = False

if st.session_state["es_admin"]:
    st.sidebar.success("🔓 Modo Administrador activo")
    seccion = st.sidebar.radio(
        "Selecciona una sección:",
        ["📊 Reportes", "🛠️ Herramientas", "⚙️ Configuración"]
    )

    if seccion == "📊 Reportes":
        vista_modo = st.sidebar.radio(
            "Tipo de reporte:",
            ["Detalle de Contrato", "Consolidado de Subproyecto"]
        )
    elif seccion == "🛠️ Herramientas":
        vista_modo = "📝 Programar Reportes Diarios"
    elif seccion == "⚙️ Configuración":
        vista_modo = "🏗️ Administrar Estructura"

    if st.sidebar.button("🔒 Cerrar Sesión Admin", use_container_width=True):
        st.session_state["es_admin"] = False
        st.rerun()
else:
    vista_modo = st.sidebar.radio(
        "Tipo de reporte:",
        ["Detalle de Contrato", "Consolidado de Subproyecto"]
    )

    with st.sidebar.expander("🔐 Acceso Administrador"):
        with st.form("form_login_admin"):
            usuario_ing = st.text_input("Usuario:")
            clave_ing = st.text_input("Contraseña:", type="password")
            btn_login = st.form_submit_button("Ingresar", use_container_width=True)

            if btn_login:
                admin_user = st.secrets.get("ADMIN_USER", "rcano")
                admin_pass = st.secrets.get("ADMIN_PASS", "")
                if admin_pass and usuario_ing.strip() == admin_user and clave_ing == admin_pass:
                    st.session_state["es_admin"] = True
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")

st.sidebar.markdown("---")

if st.sidebar.button("🔄 Actualizar Todos los Contratos", use_container_width=True):
    extraer_datos_contrato_completo.clear()
    st.sidebar.warning("Descargando data global en paralelo (Autodesk API)...")
    barra_p = st.sidebar.progress(0.0)
    texto_p = st.sidebar.empty()
    precargar_todos_los_contratos(periodo_actual, progress_bar=barra_p, status_text=texto_p)
    st.sidebar.success("¡Data global actualizada exitosamente!")

st.sidebar.caption(f"Corte activo: {periodo_actual.split('_')[-1].upper()} · {periodo_actual.split('_')[0]}")

tz_pe = pytz.timezone('America/Lima')
metadata_export = {
    "fecha": datetime.now(tz_pe).strftime('%d/%m/%Y %H:%M'),
    "alcance": grupo_activo,
    "subproyecto": subproyecto_activo,
}

if "preload" in st.query_params and st.query_params["preload"] == "true":
    st.title("🤖 Modo Bot: Actualizando Caché Global")
    st.warning("Descargando toda la data de Autodesk en paralelo...")
    barra_bot = st.progress(0.0)
    texto_bot = st.empty()
    precargar_todos_los_contratos(periodo_actual, progress_bar=barra_bot, status_text=texto_bot)
    st.success("¡Data global precargada exitosamente!")
    st.stop()

st.markdown(
    f'<div class="barra-contexto"><b>Ruta activa:</b> {grupo_activo} › <b>{subproyecto_activo}</b> › {tipo_obra_activo} › <b>{contrato_activo}</b></div>',
    unsafe_allow_html=True
)

col_pendiente = campos_pers_activos.get(id_pen_activo, "PENDIENTE POR")

if vista_modo == "📝 Programar Reportes Diarios":
    mostrar_formulario_reportes(ESTRUCTURA_MAESTRA)

elif vista_modo == "🏗️ Administrar Estructura":
    mostrar_editor_estructura(ESTRUCTURA_MAESTRA, SHA_ESTRUCTURA)

elif vista_modo == "Consolidado de Subproyecto":
    st.subheader(f"📊 Resumen Consolidado: {subproyecto_activo}")
    st.markdown('<div class="bloque-guia">Vista global de todos los contratos del subproyecto. Compara cierres, pendientes y avance porcentual por contrato.</div>', unsafe_allow_html=True)

    with st.spinner("Procesando y agrupando todos los contratos del subproyecto..."):
        datos_subproyecto = []
        resumenes_por_contrato = {}

        for tipo_obra_nombre, contratos_dict_sub in tipos_obra_dict.items():
            for contrato_nombre, info_contrato_sub in contratos_dict_sub.items():
                res_cache = extraer_datos_contrato_completo(
                    contrato_nombre, info_contrato_sub, campos_pers_activos, id_fec_activo, id_pen_activo, periodo_actual
                )
                datos_pestanas_sub = res_cache.get("datos", {})

                conteo_roles = {}
                total_docs = 0
                for df_sub in datos_pestanas_sub.values():
                    total_docs += len(df_sub)
                    if col_pendiente in df_sub.columns:
                        estados_sub = df_sub[col_pendiente].astype(str).str.strip().str.upper()
                        for est in estados_sub:
                            est_clean = str(est).strip().upper()
                            if est_clean and est_clean not in ["NAN", "NAT", "NONE", "NULL", ""]:
                                conteo_roles[est_clean] = conteo_roles.get(est_clean, 0) + 1

                cerrados_docs = conteo_roles.get("CERRADO", 0)
                fila_sub = {
                    "Tipo de Obra": tipo_obra_nombre,
                    "Contrato": contrato_nombre,
                    "Cerrados": cerrados_docs
                }
                for est_clean, count in conteo_roles.items():
                    if est_clean != "CERRADO":
                        fila_sub[est_clean.title()] = count
                fila_sub["Total"] = total_docs
                avance_pct = (cerrados_docs / total_docs * 100) if total_docs > 0 else 0
                fila_sub["Avance (%)"] = f"{avance_pct:.1f}%"
                datos_subproyecto.append(fila_sub)

                tipo_contrato_val_sub = info_contrato_sub.get("tipo_contrato", "Con Supervisión")
                if tipo_contrato_val_sub == "Tripartito":
                    emp1 = info_contrato_sub.get("nombre_empresa_1", "EMPRESA 1")
                    emp2 = info_contrato_sub.get("nombre_empresa_2", "EMPRESA 2")
                    roles_pendientes_sub = [emp1.upper(), emp2.upper(), "SHP"]
                elif tipo_contrato_val_sub == "Contrato Directo":
                    roles_pendientes_sub = ["CONTRATISTA", "SHP"]
                else:
                    roles_pendientes_sub = ["CONTRATISTA", "SUPERVISION", "SHP"]

                resumen_lista_contrato = []
                for nombre_carpeta, df_sub in datos_pestanas_sub.items():
                    fila_res_c = {"Carpeta": nombre_carpeta}
                    cerrados_carp_c = 0
                    abiertos_carp_c = 0

                    if col_pendiente in df_sub.columns:
                        estados_c = df_sub[col_pendiente].astype(str).str.strip().str.upper()
                        cerrados_carp_c = (estados_c == "CERRADO").sum()
                        fila_res_c["Cerrados"] = cerrados_carp_c
                        for rol in roles_pendientes_sub:
                            cant_rol = (estados_c == rol).sum()
                            fila_res_c[rol.title()] = cant_rol
                            abiertos_carp_c += cant_rol

                        df_abiertos_c = df_sub[estados_c != "CERRADO"].copy()
                        if not df_abiertos_c.empty and 'DIAS' in df_abiertos_c.columns:
                            df_abiertos_c['DIAS_NUM'] = pd.to_numeric(df_abiertos_c['DIAS'], errors='coerce').fillna(0)
                            if df_abiertos_c['DIAS_NUM'].max() > 0:
                                max_idx_c = df_abiertos_c['DIAS_NUM'].idxmax()
                                doc_codigo_c = str(df_abiertos_c.loc[max_idx_c, COL_CODIGO])
                                dias_retraso_c = int(df_abiertos_c.loc[max_idx_c, 'DIAS_NUM'])
                                resp_c = str(df_abiertos_c.loc[max_idx_c, col_pendiente]).strip()
                                fila_res_c["Mayor Retraso"] = f"{doc_codigo_c} ({dias_retraso_c} días) [{resp_c}]"
                            else:
                                fila_res_c["Mayor Retraso"] = "Sin retrasos"
                        else:
                            fila_res_c["Mayor Retraso"] = "Sin pendientes"
                    else:
                        fila_res_c["Cerrados"] = 0
                        for rol in roles_pendientes_sub:
                            fila_res_c[rol.title()] = 0
                        fila_res_c["Mayor Retraso"] = "N/A"

                    fila_res_c["Total"] = cerrados_carp_c + abiertos_carp_c
                    resumen_lista_contrato.append(fila_res_c)

                resumenes_por_contrato[contrato_nombre] = pd.DataFrame(resumen_lista_contrato)

        df_consolidado_sub = pd.DataFrame(datos_subproyecto)
        if not df_consolidado_sub.empty:
            cols_numericas = [c for c in df_consolidado_sub.columns if c not in ["Tipo de Obra", "Contrato", "Avance (%)"]]
            df_consolidado_sub[cols_numericas] = df_consolidado_sub[cols_numericas].fillna(0).astype(int)

            roles_cols = [c for c in df_consolidado_sub.columns if c not in ["Tipo de Obra", "Contrato", "Cerrados", "Total", "Avance (%)"]]
            orden_cols = ["Tipo de Obra", "Contrato", "Cerrados"] + roles_cols + ["Total", "Avance (%)"]
            df_consolidado_sub = df_consolidado_sub[orden_cols]
        else:
            roles_cols = []

    col_s1, col_s2 = st.columns([1.3, 1])
    with col_s1:
        st.dataframe(df_consolidado_sub, use_container_width=True, hide_index=True)

        if not df_consolidado_sub.empty:
            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                pdf_sub = generar_pdf_subproyecto(metadata_export, df_consolidado_sub, ruta_logo, resumenes_por_contrato)
                st.download_button("📄 Exportar Consolidado (PDF)", data=pdf_sub, file_name=f"Subproyecto_{subproyecto_activo.replace(' ','_')}.pdf", mime="application/pdf", use_container_width=True)
            with c_btn2:
                excel_sub = exportar_df_a_excel(df_consolidado_sub, "Consolidado de Subproyecto", metadata_export, ruta_logo)
                st.download_button("📊 Exportar Consolidado (Excel)", data=excel_sub, file_name=f"Subproyecto_{subproyecto_activo.replace(' ','_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with col_s2:
        if not df_consolidado_sub.empty:
            columnas_graf = ["Cerrados"] + roles_cols
            fig_sub = px.bar(
                df_consolidado_sub,
                x="Contrato",
                y=columnas_graf,
                title="Estado General por Contrato y Responsable",
                barmode="stack",
            )
            fig_sub.update_layout(height=300, margin=dict(t=30, b=10, l=10, r=10))
            st.plotly_chart(fig_sub, use_container_width=True)

    st.markdown("---")
    with st.expander("📧 Enviar Consolidado de Subproyecto por Correo Electrónico"):
        st.markdown("Envía el resumen global por contrato y el PDF consolidado directamente desde esta pantalla.")
        correo_destino_sub = st.text_input("Destinatarios (separa múltiples correos con coma):", key="email_sub")

        if st.button("🚀 Enviar Consolidado Ahora"):
            if not correo_destino_sub.strip():
                st.error("Por favor, ingresa al menos un correo electrónico válido.")
            else:
                with st.spinner("Generando reporte consolidado del subproyecto y enviando correo..."):
                    pdf_buffer_sub = generar_pdf_subproyecto(
                        metadata_export, df_consolidado_sub, ruta_logo, resumenes_por_contrato
                    )
                    nombre_pdf_sub = f"Consolidado_Subproyecto_{subproyecto_activo.replace(' ', '_')}.pdf"

                    cuerpo_texto_sub = (
                        f"Estimados:\n\n"
                        f"Adjunto encontrará el reporte consolidado global actualizado al {metadata_export['fecha']} "
                        f"correspondiente al subproyecto: {subproyecto_activo}.\n\n"
                        f"Para mayor detalle de los documentos, acceda a nuestro dashboard interactivo: {URL_DASHBOARD_WEB}\n\n"
                        f"Atentamente,\nRoberto Cano\nControl Documentario - Shougang Hierro Perú"
                    )

                    cuerpo_html_sub = f"""
                    <html>
                    <body style="font-family: Arial, sans-serif; color: #333;">
                        <p>Estimados:</p>
                        <p>Adjunto encontrará el reporte consolidado global actualizado al <b>{metadata_export['fecha']}</b> correspondiente al subproyecto: <b>{subproyecto_activo}</b>.</p>
                        <p>A continuación, se muestra el resumen general por contrato:</p>
                        {generar_tabla_html(df_consolidado_sub)}
                        <p>Para mayor detalle, acceda a nuestro dashboard interactivo: <a href="{URL_DASHBOARD_WEB}" style="color: #1A365D; font-weight: bold;">Control Documentario SHP</a></p>
                        <br>
                        <p>Atentamente,<br>
                        <b>Roberto Cano</b><br>
                        Control Documentario - Shougang Hierro Perú</p>
                    </body>
                    </html>
                    """

                    exito, mensaje = enviar_correo_con_pdf(
                        destinatario=correo_destino_sub,
                        asunto=f"Reporte Consolidado - {subproyecto_activo} (Shougang)",
                        cuerpo_texto=cuerpo_texto_sub,
                        cuerpo_html=cuerpo_html_sub,
                        pdf_bytes=pdf_buffer_sub,
                        nombre_archivo=nombre_pdf_sub
                    )

                    if exito:
                        st.success(mensaje)
                    else:
                        st.error(mensaje)

else:
    metadata_export["tipo_obra"] = tipo_obra_activo
    metadata_export["contrato"] = contrato_activo

    with st.spinner(f"Cargando datos para el contrato {contrato_activo}..."):
        resultado_cache = extraer_datos_contrato_completo(
            contrato_activo, info_contrato, campos_pers_activos, id_fec_activo, id_pen_activo, periodo_actual
        )

    datos_pestanas = resultado_cache.get("datos", {})
    datos_completos_pestanas = resultado_cache.get("datos_completos", {})
    hora_actualizacion = resultado_cache.get("hora_actualizacion", "Desconocida")

    if not datos_pestanas:
        st.info(f"El contrato {contrato_activo} no contiene registros o sus carpetas están pendientes de enlace en la estructura.")
    else:
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
                c2.metric("Supervisor", info_contrato.get("nombre_supervisor") or "N/A")
                c3.metric("Coordinador SHP", info_contrato.get("coordinador_shp") or "N/A")

            c4.metric("Admin. Contratos", info_contrato.get("administrador_contratos") or "N/A")
            c5.metric("Última Actualiz.", hora_actualizacion)

        if tipo_contrato_val == "Tripartito":
            emp1 = info_contrato.get("nombre_empresa_1", "EMPRESA 1")
            emp2 = info_contrato.get("nombre_empresa_2", "EMPRESA 2")
            roles_pendientes = [emp1.upper(), emp2.upper(), "SHP"]
        elif tipo_contrato_val == "Contrato Directo":
            roles_pendientes = ["CONTRATISTA", "SHP"]
        else:
            roles_pendientes = ["CONTRATISTA", "SUPERVISION", "SHP"]

        total_cerrados = 0
        total_abiertos = 0
        resumen_lista = []

        for nombre_carpeta, df in datos_pestanas.items():
            fila_res = {"Carpeta": nombre_carpeta}
            cerrados_carp = 0
            abiertos_carp = 0

            if col_pendiente in df.columns:
                estados = df[col_pendiente].astype(str).str.strip().str.upper()
                cerrados_carp = int((estados == "CERRADO").sum())
                fila_res["Cerrados"] = cerrados_carp

                for rol in roles_pendientes:
                    cant_rol = int((estados == rol).sum())
                    fila_res[rol.title()] = cant_rol
                    abiertos_carp += cant_rol

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

        avance_contrato = (total_cerrados / (total_cerrados + total_abiertos) * 100) if (total_cerrados + total_abiertos) > 0 else 0.0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total de Documentos", total_cerrados + total_abiertos)
        m2.metric("Documentos Cerrados", total_cerrados)
        m3.metric("Documentos Pendientes", total_abiertos, delta="Requieren atención" if total_abiertos > 0 else "Al día", delta_color="inverse")
        m4.metric("Avance de Cierre", f"{avance_contrato:.1f}%")
        st.markdown("---")

        # ====================================================================
        # --- RESUMEN 1: POR RESPONSABLE (PENDIENTE POR) ---
        # ====================================================================
        st.subheader(f"1. Resumen por Responsable — {contrato_activo}")
        st.markdown('<div class="bloque-guia">Muestra cuántos documentos están cerrados o pendientes por cada parte involucrada y cuál documento acumula mayor retraso en cada carpeta.</div>', unsafe_allow_html=True)
        df_resumen = pd.DataFrame(resumen_lista)

        col_g1, col_g2 = st.columns([1.3, 1])
        with col_g1:
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
            if not df_resumen.empty:
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    pdf_resumen = exportar_df_a_pdf(df_resumen, "Resumen por Responsable", metadata_export, ruta_logo)
                    st.download_button(label="📄 Exportar Responsables (PDF)", data=pdf_resumen, file_name=f"Resumen_Responsables_{contrato_activo.replace(' ', '_')}.pdf", mime="application/pdf", key="btn_descarga_resumen_pdf", use_container_width=True)
                with c_btn2:
                    excel_resumen = exportar_df_a_excel(df_resumen, "Resumen por Responsable", metadata_export, ruta_logo)
                    st.download_button(label="📊 Exportar Responsables (Excel)", data=excel_resumen, file_name=f"Resumen_Responsables_{contrato_activo.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_descarga_resumen_excel", use_container_width=True)

        with col_g2:
            if not df_resumen.empty:
                columnas_grafico = ["Cerrados"] + [c for c in df_resumen.columns if c not in ["Carpeta", "Cerrados", "Total", "Mayor Retraso"]]
                fig = px.bar(
                    df_resumen,
                    x="Carpeta",
                    y=columnas_grafico,
                    title="Distribución por Carpeta y Responsable",
                    barmode="group"
                )
                fig.update_layout(height=300, margin=dict(t=30, b=10, l=10, r=10))
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")

        # ====================================================================
        # --- RESUMEN 2: POR ESTADO DE LOS DOCUMENTOS ---
        # ====================================================================
        st.subheader(f"2. Resumen por Estado Documentario — {contrato_activo}")
        st.markdown('<div class="bloque-guia">Agrupa los documentos según el valor registrado en su campo <b>ESTADO</b> en Autodesk ACC (por carpeta y total).</div>', unsafe_allow_html=True)

        todos_estados_unicos = set()
        for df_e in datos_pestanas.values():
            if "ESTADO" in df_e.columns:
                vals_est = df_e["ESTADO"].fillna("").astype(str).str.strip().str.upper()
                vals_est = vals_est.replace({"": "SIN ESTADO", "NAN": "SIN ESTADO", "NONE": "SIN ESTADO", "NULL": "SIN ESTADO"})
                todos_estados_unicos.update(vals_est.unique())

        lista_estados_ord = sorted(list(todos_estados_unicos))
        resumen_estado_lista = []

        for nombre_carpeta, df_e in datos_pestanas.items():
            fila_est = {"Carpeta": nombre_carpeta}
            total_carp_est = len(df_e)
            if "ESTADO" in df_e.columns:
                vals_est = df_e["ESTADO"].fillna("").astype(str).str.strip().str.upper()
                vals_est = vals_est.replace({"": "SIN ESTADO", "NAN": "SIN ESTADO", "NONE": "SIN ESTADO", "NULL": "SIN ESTADO"})
                conteo_e = vals_est.value_counts().to_dict()
                for est_key in lista_estados_ord:
                    fila_est[est_key.title()] = int(conteo_e.get(est_key, 0))
            else:
                for est_key in lista_estados_ord:
                    fila_est[est_key.title()] = 0
            fila_est["Total"] = total_carp_est
            resumen_estado_lista.append(fila_est)

        df_resumen_estado = pd.DataFrame(resumen_estado_lista)

        col_e1, col_e2 = st.columns([1.3, 1])
        with col_e1:
            st.dataframe(df_resumen_estado, use_container_width=True, hide_index=True)
            if not df_resumen_estado.empty:
                c_ebtn1, c_ebtn2 = st.columns(2)
                with c_ebtn1:
                    pdf_res_est = exportar_df_a_pdf(df_resumen_estado, "Resumen por Estado de Documentos", metadata_export, ruta_logo)
                    st.download_button(label="📄 Exportar por Estado (PDF)", data=pdf_res_est, file_name=f"Resumen_Estados_{contrato_activo.replace(' ', '_')}.pdf", mime="application/pdf", key="btn_descarga_estado_pdf", use_container_width=True)
                with c_ebtn2:
                    excel_res_est = exportar_df_a_excel(df_resumen_estado, "Resumen por Estado de Documentos", metadata_export, ruta_logo)
                    st.download_button(label="📊 Exportar por Estado (Excel)", data=excel_res_est, file_name=f"Resumen_Estados_{contrato_activo.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_descarga_estado_excel", use_container_width=True)

        with col_e2:
            if not df_resumen_estado.empty and lista_estados_ord:
                cols_graf_estado = [c for c in df_resumen_estado.columns if c not in ["Carpeta", "Total"]]
                fig_estado = px.bar(
                    df_resumen_estado,
                    x="Carpeta",
                    y=cols_graf_estado,
                    title="Distribución por Carpeta y Estado",
                    barmode="stack"
                )
                fig_estado.update_layout(height=300, margin=dict(t=30, b=10, l=10, r=10))
                st.plotly_chart(fig_estado, use_container_width=True)

        st.markdown("---")
        with st.expander("📧 Enviar Reporte de Contrato por Correo Electrónico"):
            st.markdown("Incluye ambos resúmenes (por responsable y por estado) en el cuerpo del correo y adjunta el PDF del contrato.")
            correo_destino = st.text_input("Destinatarios (separa múltiples correos con coma):", key="email_contrato")

            if st.button("🚀 Enviar Reporte Ahora"):
                if not correo_destino.strip():
                    st.error("Por favor, ingresa al menos un correo electrónico válido.")
                else:
                    with st.spinner("Generando reporte con detalle y enviando correo..."):
                        pdf_buffer = exportar_df_a_pdf(df_resumen, "Reporte de Contrato", metadata_export, ruta_logo, col_pendiente)
                        nombre_pdf = f"Reporte_Contrato_{contrato_activo.replace(' ', '_')}.pdf"

                        cuerpo_texto = (
                            f"Estimados:\n\n"
                            f"Adjunto encontrará el reporte documentario detallado actualizado al {metadata_export['fecha']} "
                            f"para el contrato: {contrato_activo}.\n\n"
                            f"Resumen de estado:\n"
                            f"- Total de documentos: {total_cerrados + total_abiertos}\n"
                            f"- Documentos cerrados: {total_cerrados}\n"
                            f"- Documentos pendientes: {total_abiertos}\n\n"
                            f"Para revisar en línea: {URL_DASHBOARD_WEB}\n\n"
                            f"Atentamente,\nRoberto Cano\nControl Documentario - Shougang Hierro Perú"
                        )

                        cuerpo_html = f"""
                        <html>
                        <body style="font-family: Arial, sans-serif; color: #333;">
                            <p>Estimados:</p>
                            <p>Adjunto encontrará el reporte documentario detallado actualizado al <b>{metadata_export['fecha']}</b> para el contrato: <b>{contrato_activo}</b>.</p>
                            <p><b>1. Resumen por responsable:</b></p>
                            {generar_tabla_html(df_resumen)}
                            <p><b>2. Resumen por estado del documento:</b></p>
                            {generar_tabla_html(df_resumen_estado)}
                            <p><b>Resumen de estado global:</b></p>
                            <ul>
                                <li>Total de documentos: {total_cerrados + total_abiertos}</li>
                                <li>Documentos cerrados: {total_cerrados}</li>
                                <li>Documentos pendientes: {total_abiertos}</li>
                            </ul>
                            <p>Para revisar en mayor detalle, acceda a nuestro dashboard interactivo: <a href="{URL_DASHBOARD_WEB}" style="color: #1A365D; font-weight: bold;">Control Documentario SHP</a></p>
                            <br>
                            <p>Atentamente,<br>
                            <b>Roberto Cano</b><br>
                            Control Documentario - Shougang Hierro Perú</p>
                        </body>
                        </html>
                        """

                        exito, mensaje = enviar_correo_con_pdf(
                            destinatario=correo_destino,
                            asunto=f"Reporte de Contrato - {contrato_activo} (Shougang)",
                            cuerpo_texto=cuerpo_texto,
                            cuerpo_html=cuerpo_html,
                            pdf_bytes=pdf_buffer,
                            nombre_archivo=nombre_pdf
                        )

                        if exito:
                            st.success(mensaje)
                        else:
                            st.error(mensaje)
        st.markdown("---")

        # ====================================================================
        # --- DETALLE POR CARPETA / DISCIPLINA ---
        # ====================================================================
        st.subheader("3. Detalle por Carpeta / Disciplina")
        st.markdown('<div class="bloque-guia">Selecciona una carpeta para revisar documento por documento, filtrar por responsable y abrir el enlace directo en Autodesk ACC.</div>', unsafe_allow_html=True)

        nombres_carpetas = list(datos_pestanas.keys())
        pestana_seleccionada = st.selectbox("Seleccionar Carpeta:", nombres_carpetas)

        df_actual = datos_pestanas[pestana_seleccionada].copy()

        config_columnas_detalle = {
            COL_CODIGO: st.column_config.LinkColumn(
                COL_CODIGO,
                display_text=r"#doc=(.*)"
            ),
            "URL_DOC": None
        }

        if col_pendiente in df_actual.columns:
            roles_con_cerrado = roles_pendientes + ["CERRADO"]
            estados_disponibles = df_actual[col_pendiente].astype(str).unique().tolist()

            st.caption(f"Modalidad: **{tipo_contrato_val}** · Roles configurados: {', '.join(roles_con_cerrado)}")
            col_filtros, col_grafico = st.columns([1, 1])

            with col_filtros:
                filtro_estado = st.multiselect("Filtrar por Responsable:", estados_disponibles, default=estados_disponibles)

            df_filtrado = df_actual[df_actual[col_pendiente].isin(filtro_estado)]

            def pintar_filas_streamlit(row):
                estado = str(row[col_pendiente]).strip().upper()
                bg_color, text_color = obtener_colores_estado(estado)
                return [f'background-color: {bg_color}; color: {text_color}'] * len(row)

            with col_grafico:
                if not df_filtrado.empty:
                    df_counts = df_filtrado[col_pendiente].value_counts().reset_index()
                    df_counts.columns = [col_pendiente, 'Cantidad']

                    fig_detalle = px.bar(
                        df_counts,
                        x=col_pendiente,
                        y='Cantidad',
                        color=col_pendiente,
                        title="Distribución de Responsables",
                        text='Cantidad'
                    )
                    fig_detalle.update_layout(
                        height=250,
                        margin=dict(t=30, b=0, l=10, r=10),
                        showlegend=False,
                        xaxis_title="",
                        yaxis_title="Cantidad"
                    )
                    st.plotly_chart(fig_detalle, use_container_width=True)

            if not df_filtrado.empty:
                df_vista = preparar_df_con_links(df_filtrado)
                st.dataframe(
                    df_vista.style.apply(pintar_filas_streamlit, axis=1),
                    use_container_width=True,
                    height=400,
                    hide_index=True,
                    column_config=config_columnas_detalle
                )

                df_export_det = df_filtrado.drop(columns=["URL_DOC"], errors="ignore")
                c_det1, c_det2 = st.columns(2)
                with c_det1:
                    pdf_detalle = exportar_df_a_pdf(df_export_det, f"Detalle de Documentos: {pestana_seleccionada}", metadata_export, ruta_logo, col_pendiente)
                    st.download_button(label="📄 Exportar Detalle (PDF)", data=pdf_detalle, file_name=f"{pestana_seleccionada}_{contrato_activo.replace(' ', '_')}.pdf", mime="application/pdf", key="btn_descarga_detalle_pdf", use_container_width=True)
                with c_det2:
                    excel_detalle = exportar_df_a_excel(df_export_det, f"Detalle de Documentos: {pestana_seleccionada}", metadata_export, ruta_logo, col_pendiente)
                    st.download_button(label="📊 Exportar Detalle (Excel)", data=excel_detalle, file_name=f"{pestana_seleccionada}_{contrato_activo.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_descarga_detalle_excel", use_container_width=True)
            else:
                st.info("No hay documentos que coincidan con el filtro seleccionado.")

        else:
            df_vista = preparar_df_con_links(df_actual)
            st.dataframe(
                df_vista,
                use_container_width=True,
                height=400,
                hide_index=True,
                column_config=config_columnas_detalle
            )

            if not df_actual.empty:
                df_export_det = df_actual.drop(columns=["URL_DOC"], errors="ignore")
                c_det1, c_det2 = st.columns(2)
                with c_det1:
                    pdf_detalle = exportar_df_a_pdf(df_export_det, f"Detalle de Documentos: {pestana_seleccionada}", metadata_export, ruta_logo)
                    st.download_button(label="📄 Exportar Detalle (PDF)", data=pdf_detalle, file_name=f"{pestana_seleccionada}_{contrato_activo.replace(' ', '_')}.pdf", mime="application/pdf", key="btn_descarga_detalle_alt_pdf", use_container_width=True)
                with c_det2:
                    excel_detalle = exportar_df_a_excel(df_export_det, f"Detalle de Documentos: {pestana_seleccionada}", metadata_export, ruta_logo)
                    st.download_button(label="📊 Exportar Detalle (Excel)", data=excel_detalle, file_name=f"{pestana_seleccionada}_{contrato_activo.replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_descarga_detalle_alt_excel", use_container_width=True)

        # ====================================================================
        # --- 4. CONSTRUCTOR DINÁMICO DE REPORTE (TODOS LOS CAMPOS Y DATOS) ---
        # ====================================================================
        st.markdown("---")
        st.subheader("4. Constructor Dinámico de Reporte Personalizado")
        st.markdown('<div class="bloque-guia">Arma un reporte a tu medida combinando carpetas, cualquier campo disponible en la data de Autodesk ACC y filtros exactos.</div>', unsafe_allow_html=True)

        with st.expander("🛠️ Abrir Constructor Dinámico", expanded=True):
            if datos_completos_pestanas:
                carpetas_disp_pers = list(datos_completos_pestanas.keys())
                carpetas_sel_pers = st.multiselect(
                    "Paso 1 · Selecciona las carpetas a consolidar:",
                    options=carpetas_disp_pers,
                    default=carpetas_disp_pers,
                    key="pers_carpetas"
                )

                if carpetas_sel_pers:
                    df_full_unido = pd.concat(
                        [datos_completos_pestanas[c] for c in carpetas_sel_pers if c in datos_completos_pestanas],
                        ignore_index=True
                    )

                    cols_principales = [
                        c for c in [
                            "CARPETA", "CODIGO", "REV", "DESCRIPCION", "FECHA",
                            "ESTADO", "PENDIENTE POR", "DIAS", "RUTA_EN_AUTODESK"
                        ] if c in df_full_unido.columns
                    ]
                    cols_restantes = sorted([c for c in df_full_unido.columns if c not in cols_principales])
                    todas_las_columnas = cols_principales + cols_restantes

                    if "pers_cols_state" not in st.session_state:
                        st.session_state["pers_cols_state"] = cols_principales

                    st.session_state["pers_cols_state"] = [
                        c for c in st.session_state["pers_cols_state"] if c in todas_las_columnas
                    ]
                    if not st.session_state["pers_cols_state"]:
                        st.session_state["pers_cols_state"] = cols_principales

                    b_sel1, b_sel2, b_sel3 = st.columns([1, 1, 2])
                    with b_sel1:
                        if st.button("✅ Todos los campos", key="btn_todos_campos", use_container_width=True):
                            st.session_state["pers_cols_state"] = todas_las_columnas
                            st.rerun()
                    with b_sel2:
                        if st.button("↩️ Campos esenciales", key="btn_esenciales_campos", use_container_width=True):
                            st.session_state["pers_cols_state"] = cols_principales
                            st.rerun()

                    columnas_sel_pers = st.multiselect(
                        f"Paso 2 · Elige los campos a incluir ({len(todas_las_columnas)} campos disponibles en toda la data):",
                        options=todas_las_columnas,
                        default=st.session_state["pers_cols_state"],
                        key="pers_columnas"
                    )
                    st.session_state["pers_cols_state"] = columnas_sel_pers

                    col_f_a, col_f_b = st.columns([1, 2])
                    with col_f_a:
                        col_filtro_pers = st.selectbox(
                            "Paso 3 · Filtrar por campo (opcional):",
                            options=["(Sin filtro)"] + todas_las_columnas,
                            key="pers_col_filtro"
                        )

                    df_pers_resultado = df_full_unido.copy()

                    with col_f_b:
                        if col_filtro_pers != "(Sin filtro)":
                            valores_unicos_filtro = sorted(df_pers_resultado[col_filtro_pers].fillna("").astype(str).unique().tolist())
                            valores_sel_pers = st.multiselect(
                                f"Valores permitidos en '{col_filtro_pers}':",
                                options=valores_unicos_filtro,
                                default=valores_unicos_filtro,
                                key="pers_val_filtro"
                            )
                            df_pers_resultado = df_pers_resultado[df_pers_resultado[col_filtro_pers].fillna("").astype(str).isin(valores_sel_pers)]
                        else:
                            st.caption("Sin filtro adicional: se mostrarán todos los registros de las carpetas elegidas.")

                    if columnas_sel_pers:
                        df_pers_final = df_pers_resultado[columnas_sel_pers].copy()
                        st.markdown(f"**Paso 4 · Vista previa:** **{len(df_pers_final)}** documentos · **{len(columnas_sel_pers)}** campos")
                        st.dataframe(df_pers_final, use_container_width=True, height=380, hide_index=True)

                        c_pbtn1, c_pbtn2 = st.columns(2)
                        with c_pbtn1:
                            excel_pers = exportar_df_a_excel(
                                df_pers_final,
                                f"Reporte Dinámico - {contrato_activo}",
                                metadata_export,
                                ruta_logo,
                                col_pendiente if col_pendiente in df_pers_final.columns else None
                            )
                            st.download_button(
                                label="📊 Descargar Reporte Dinámico (Excel)",
                                data=excel_pers,
                                file_name=f"Reporte_Dinamico_{contrato_activo.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="btn_descarga_pers_excel",
                                use_container_width=True
                            )
                        with c_pbtn2:
                            pdf_pers = exportar_df_a_pdf(
                                df_pers_final,
                                f"Reporte Dinámico - {contrato_activo}",
                                metadata_export,
                                ruta_logo,
                                col_pendiente if col_pendiente in df_pers_final.columns else None
                            )
                            st.download_button(
                                label="📄 Descargar Reporte Dinámico (PDF)",
                                data=pdf_pers,
                                file_name=f"Reporte_Dinamico_{contrato_activo.replace(' ', '_')}.pdf",
                                mime="application/pdf",
                                key="btn_descarga_pers_pdf",
                                use_container_width=True
                            )
                    else:
                        st.warning("Selecciona al menos un campo para generar la vista previa y exportar.")
                else:
                    st.info("Selecciona al menos una carpeta para explorar los campos disponibles.")
            else:
                st.info("No hay datos extendidos disponibles para este contrato.")

import streamlit as st
import pandas as pd
import io
import os
import base64
import requests
from datetime import datetime, date
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table as RLTable, TableStyle, Paragraph,
    Spacer, HRFlowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT

# ─── Page Config ───
# v2.1 - PDF redesign, simplified machine list, updated filters
st.set_page_config(page_title="SE PM Campaign", page_icon="SE", layout="wide", initial_sidebar_state="expanded")

# ─── Brand ───
SEC_RED = "#C8102E"
SEC_DARK = "#1A1A1A"
SEC_GRAY = "#4A4A4A"

# ─── Branches (17 locations, matching Parts Campaign) ───
BRANCHES = {
    1: "Cambridge", 2: "North Canton", 3: "Gallipolis", 4: "Dublin",
    5: "Monroe", 6: "Burlington", 7: "Perrysburg", 9: "Brunswick",
    11: "Mentor", 12: "Fort Wayne", 13: "Indianapolis", 14: "Mansfield",
    15: "Heath", 16: "Marietta", 17: "Evansville", 19: "Holt", 20: "Novi",
}
BRANCH_NAMES = sorted(BRANCHES.values())

REGIONS = {
    "SE Region": [1, 3, 4, 15, 16],       # Cambridge, Gallipolis, Dublin, Heath, Marietta
    "NE Region": [2, 7, 9, 11, 14],       # North Canton, Perrysburg, Brunswick, Mentor, Mansfield
    "West Region": [5, 6, 12, 13, 17, 19, 20],  # Monroe, Burlington, Fort Wayne, Indianapolis, Evansville, Holt, Novi
}

ADMIN_PASSWORD = "SEpm2026"

# ─── PM Dealsheet Pricing (from PM Contract Dealsheet Rev 3.0) ───
# Each model: {brand, cost_i, cost_1, cost_2, cost_3, cost_s, hr_i, hr_1, hr_2, hr_3, hr_s}
PM_DEALSHEET = {
    # === CASE (44 models) ===
    "CX12D": {"brand":"Case","cost_i":518,"cost_1":0,"cost_2":1207,"cost_3":2166,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX19D Cab": {"brand":"Case","cost_i":518,"cost_1":0,"cost_2":1207,"cost_3":2166,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX30D Cab": {"brand":"Case","cost_i":760,"cost_1":485,"cost_2":1175,"cost_3":2040,"cost_s":0,"hr_i":100,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX34D Cab": {"brand":"Case","cost_i":760,"cost_1":485,"cost_2":1175,"cost_3":2040,"cost_s":0,"hr_i":100,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX38D Cab": {"brand":"Case","cost_i":825,"cost_1":550,"cost_2":1235,"cost_3":2105,"cost_s":0,"hr_i":100,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX42D Cab": {"brand":"Case","cost_i":890,"cost_1":0,"cost_2":1360,"cost_3":2280,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX50D Cab": {"brand":"Case","cost_i":820,"cost_1":0,"cost_2":1395,"cost_3":2800,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX60D": {"brand":"Case","cost_i":855,"cost_1":0,"cost_2":1420,"cost_3":2815,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX70E": {"brand":"Case","cost_i":825,"cost_1":0,"cost_2":1175,"cost_3":2580,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX85E": {"brand":"Case","cost_i":825,"cost_1":0,"cost_2":1175,"cost_3":2580,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "CX90E": {"brand":"Case","cost_i":825,"cost_1":0,"cost_2":1175,"cost_3":2580,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "TR270B": {"brand":"Case","cost_i":215,"cost_1":0,"cost_2":1025,"cost_3":1550,"cost_s":2595,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "TR310B": {"brand":"Case","cost_i":215,"cost_1":0,"cost_2":1025,"cost_3":1550,"cost_s":2575,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "TR340B": {"brand":"Case","cost_i":215,"cost_1":0,"cost_2":1045,"cost_3":1650,"cost_s":2694,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "TV370B": {"brand":"Case","cost_i":145,"cost_1":0,"cost_2":946,"cost_3":1540,"cost_s":2590,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "TV450B": {"brand":"Case","cost_i":105,"cost_1":0,"cost_2":982,"cost_3":2105,"cost_s":0,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "TV620B": {"brand":"Case","cost_i":175,"cost_1":0,"cost_2":1625,"cost_3":2640,"cost_s":2860,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "DL550": {"brand":"Case","cost_i":175,"cost_1":0,"cost_2":1905,"cost_3":2815,"cost_s":4000,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "SR175B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":885,"cost_3":2640,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SR210B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":895,"cost_3":1550,"cost_s":2545,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "SR240B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":870,"cost_3":1685,"cost_s":2680,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "SR270B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":870,"cost_3":1770,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SV270B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":870,"cost_3":1770,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SV185B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":700,"cost_3":2480,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SV280B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":870,"cost_3":1685,"cost_s":2680,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "SV340B": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":960,"cost_3":1890,"cost_s":2880,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "575N": {"brand":"Case","cost_i":1005,"cost_1":0,"cost_2":585,"cost_3":2680,"cost_s":2780,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "580SV": {"brand":"Case","cost_i":0,"cost_1":0,"cost_2":570,"cost_3":2495,"cost_s":2550,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "580SN": {"brand":"Case","cost_i":670,"cost_1":0,"cost_2":905,"cost_3":3280,"cost_s":3410,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "590SN": {"brand":"Case","cost_i":670,"cost_1":0,"cost_2":905,"cost_3":3280,"cost_s":3410,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "586H": {"brand":"Case","cost_i":590,"cost_1":0,"cost_2":580,"cost_3":2630,"cost_s":2795,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "588H": {"brand":"Case","cost_i":590,"cost_1":0,"cost_2":580,"cost_3":2630,"cost_s":2795,"hr_i":100,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    "SL12": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL12 TR": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL15": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL23": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL27": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL27 TR": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL35 TR": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SL50 TR": {"brand":"Case","cost_i":950,"cost_1":575,"cost_2":2835,"cost_3":6750,"cost_s":0,"hr_i":50,"hr_1":250,"hr_2":500,"hr_3":1000,"hr_s":0},
    "21F": {"brand":"Case","cost_i":470,"cost_1":0,"cost_2":566,"cost_3":2090,"cost_s":2437,"hr_i":150,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":2000},
    "221F": {"brand":"Case","cost_i":638,"cost_1":0,"cost_2":1062,"cost_3":2620,"cost_s":2970,"hr_i":150,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":2000},
    "321F": {"brand":"Case","cost_i":638,"cost_1":0,"cost_2":1062,"cost_3":2620,"cost_s":2970,"hr_i":150,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":2000},
    "421F": {"brand":"Case","cost_i":758,"cost_1":0,"cost_2":1080,"cost_3":2820,"cost_s":3625,"hr_i":150,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":1500},
    # === KOBELCO (19 models) ===
    "SK17SR-6E": {"brand":"Kobelco","cost_i":885,"cost_1":0,"cost_2":930,"cost_3":1260,"cost_s":0,"hr_i":250,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK26SR-7": {"brand":"Kobelco","cost_i":745,"cost_1":0,"cost_2":700,"cost_3":1070,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK35SR-7": {"brand":"Kobelco","cost_i":550,"cost_1":0,"cost_2":755,"cost_3":1200,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK45SRX-7": {"brand":"Kobelco","cost_i":600,"cost_1":0,"cost_2":685,"cost_3":1270,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK55SRX-7": {"brand":"Kobelco","cost_i":600,"cost_1":0,"cost_2":800,"cost_3":1270,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK75SR-7": {"brand":"Kobelco","cost_i":720,"cost_1":0,"cost_2":925,"cost_3":1700,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK85CS-7": {"brand":"Kobelco","cost_i":620,"cost_1":0,"cost_2":1050,"cost_3":1700,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK130LC-11": {"brand":"Kobelco","cost_i":650,"cost_1":0,"cost_2":1050,"cost_3":2100,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK140SR-7": {"brand":"Kobelco","cost_i":840,"cost_1":0,"cost_2":1675,"cost_3":2080,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "ED160-7": {"brand":"Kobelco","cost_i":900,"cost_1":0,"cost_2":1170,"cost_3":2010,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK170LC-11": {"brand":"Kobelco","cost_i":990,"cost_1":0,"cost_2":2290,"cost_3":2085,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK210LC-11": {"brand":"Kobelco","cost_i":1010,"cost_1":0,"cost_2":1500,"cost_3":2250,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK230SR-7": {"brand":"Kobelco","cost_i":1010,"cost_1":0,"cost_2":1700,"cost_3":2250,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK260LC-11": {"brand":"Kobelco","cost_i":1010,"cost_1":0,"cost_2":1520,"cost_3":2250,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK270SR-7": {"brand":"Kobelco","cost_i":930,"cost_1":0,"cost_2":1730,"cost_3":2250,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK300LC-11": {"brand":"Kobelco","cost_i":1060,"cost_1":0,"cost_2":1360,"cost_3":2050,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK350LC-11": {"brand":"Kobelco","cost_i":1060,"cost_1":0,"cost_2":1550,"cost_3":2110,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK380SRLC-7": {"brand":"Kobelco","cost_i":1385,"cost_1":0,"cost_2":1950,"cost_3":2370,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "SK520LC-11": {"brand":"Kobelco","cost_i":2390,"cost_1":0,"cost_2":3050,"cost_3":3120,"cost_s":0,"hr_i":50,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    # === DEVELON (12 models) ===
    "DA45": {"brand":"Develon","cost_i":2895,"cost_1":1575,"cost_2":5495,"cost_3":5495,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DD100": {"brand":"Develon","cost_i":470,"cost_1":1250,"cost_2":2625,"cost_3":2625,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DD130": {"brand":"Develon","cost_i":470,"cost_1":1250,"cost_2":2625,"cost_3":2625,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX35Z-7": {"brand":"Develon","cost_i":450,"cost_1":1300,"cost_2":1925,"cost_3":1925,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX50Z-7": {"brand":"Develon","cost_i":450,"cost_1":1425,"cost_2":1965,"cost_3":1965,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX63-7": {"brand":"Develon","cost_i":530,"cost_1":1425,"cost_2":1965,"cost_3":1965,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX89R-7": {"brand":"Develon","cost_i":450,"cost_1":1350,"cost_2":1965,"cost_3":1965,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX140LC-7": {"brand":"Develon","cost_i":450,"cost_1":1425,"cost_2":2125,"cost_3":2125,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX140LCR-7": {"brand":"Develon","cost_i":450,"cost_1":1425,"cost_2":2125,"cost_3":2125,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DX170LC-5": {"brand":"Develon","cost_i":450,"cost_1":1425,"cost_2":2175,"cost_3":2175,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DL220-7": {"brand":"Develon","cost_i":525,"cost_1":1175,"cost_2":3265,"cost_3":3265,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    "DL250-7": {"brand":"Develon","cost_i":525,"cost_1":1175,"cost_2":3265,"cost_3":3265,"cost_s":0,"hr_i":50,"hr_1":500,"hr_2":1000,"hr_3":2000,"hr_s":0},
    # === BOMAG (12 models) ===
    "BMP8500": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":550,"cost_3":1065,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW11RH": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":780,"cost_3":1585,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW120AD": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":490,"cost_3":1470,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW138AD": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":490,"cost_3":1375,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW141": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":490,"cost_3":1525,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW177": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":740,"cost_3":1670,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW190": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":740,"cost_3":1760,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW211": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":895,"cost_3":2180,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW900": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":340,"cost_3":895,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BW90AD": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":340,"cost_3":895,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BF200C": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":635,"cost_3":2130,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
    "BF300C": {"brand":"Bomag","cost_i":0,"cost_1":0,"cost_2":635,"cost_3":2130,"cost_s":0,"hr_i":0,"hr_1":0,"hr_2":500,"hr_3":1000,"hr_s":0},
}

# Build brand -> model list for dropdowns
PM_BRANDS = {}
for model, info in PM_DEALSHEET.items():
    brand = info["brand"]
    if brand not in PM_BRANDS:
        PM_BRANDS[brand] = []
    PM_BRANDS[brand].append(model)

# Approximate annual PM values by category (for lead scoring only, not quoting)
def match_model_to_dealsheet(alert_model):
    """Match a CASE alert model name to a PM_DEALSHEET key.
    Handles suffixes like EVOLUTION, HS, EP, WT, and Cab variants.
    Returns the dealsheet key or None if no match.
    """
    if not alert_model or str(alert_model).strip() in ("", "Total"):
        return None
    m = str(alert_model).strip()
    # Exact match
    if m in PM_DEALSHEET:
        return m
    # Strip common suffixes
    base = m.replace(" EVOLUTION", "").replace(" HS", "").replace(" EP", "").replace(" WT", "").strip()
    if base in PM_DEALSHEET:
        return base
    # Try adding " Cab" (CX mini excavators)
    cab = base + " Cab"
    if cab in PM_DEALSHEET:
        return cab
    return None

SERVICE_TYPES = ["Field", "Shop"]

# ─── Session State ───
for key in ["quotes", "leads_df", "procare_vins"]:
    if key not in st.session_state:
        st.session_state[key] = [] if key in ["quotes", "procare_vins"] else None

# Login state
if "page" not in st.session_state:
    st.session_state.page = "login"
if "branch" not in st.session_state:
    st.session_state.branch = None
if "branch_name" not in st.session_state:
    st.session_state.branch_name = None
if "login_month" not in st.session_state:
    st.session_state.login_month = None

# ─── Google Sheets ───
def get_gsheet_connection():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets.get("gcp_service_account", {}))
        if not creds_dict:
            return None
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        sheet_url = st.secrets.get("spreadsheet_url", "")
        if not sheet_url or sheet_url == "PASTE_GOOGLE_SHEET_URL_HERE":
            return None
        return client.open_by_url(sheet_url).sheet1
    except Exception:
        return None

def save_quote_to_sheet(quote_data):
    sheet = get_gsheet_connection()
    if sheet is None:
        st.session_state.quotes.append(quote_data)
        return False
    try:
        row = [
            quote_data.get("date", ""), quote_data.get("customer_name", ""),
            quote_data.get("branch", ""), quote_data.get("rep", ""),
            quote_data.get("service_type", ""), quote_data.get("make", ""),
            quote_data.get("model", ""), quote_data.get("serial", ""),
            quote_data.get("hours_requested", 0), quote_data.get("travel_time", 0),
            quote_data.get("travel_cost", 0), quote_data.get("total_cost", 0),
            quote_data.get("annual_pm_price", 0), quote_data.get("notes", ""),
        ]
        sheet.append_row(row, value_input_option="USER_ENTERED")
        st.session_state.quotes.append(quote_data)
        return True
    except Exception:
        st.session_state.quotes.append(quote_data)
        return False

def load_quotes_from_sheet():
    sheet = get_gsheet_connection()
    if sheet is None:
        return pd.DataFrame(st.session_state.quotes) if st.session_state.quotes else pd.DataFrame()
    try:
        data = sheet.get_all_records()
        return pd.DataFrame(data) if data else pd.DataFrame()
    except Exception:
        return pd.DataFrame(st.session_state.quotes) if st.session_state.quotes else pd.DataFrame()


# ═══════════════════════════════════════════════════════════
# BUNDLED DATA LOADER
# ═══════════════════════════════════════════════════════════
DATA_DIR = Path(__file__).parent / "data"

@st.cache_data(ttl=3600)
def load_bundled_alerts():
    """Load maintenance alerts from the bundled data file."""
    # Try short name first, then original name pattern
    for pattern in ["alerts.xlsx", "Southeastern Equipment Maintenance Alerts*.xlsx"]:
        files = sorted(DATA_DIR.glob(pattern))
        if files:
            return parse_maintenance_alerts(files[-1])
    return pd.DataFrame()

@st.cache_data(ttl=3600)
def load_bundled_procare():
    """Load ProCare stops from the bundled data file."""
    for pattern in ["procare.xlsx", "Southeastern ProCare Stops*.xlsx"]:
        files = sorted(DATA_DIR.glob(pattern))
        if files:
            return parse_procare_stops(files[-1])
    return set()

# Make code mapping for equipment report
EQUIP_MAKE_MAP = {
    "UT": "Case", "EX": "Case", "CE": "Case",
    "KB": "Kobelco", "BO": "Bomag", "DE": "Develon",
}
EQUIP_TARGET_MAKES = set(EQUIP_MAKE_MAP.keys())

@st.cache_data(ttl=3600)
def load_equipment_report():
    """Load the 5-year equipment sold report, filtered to our 4 dealsheet brands."""
    for pattern in ["equipment_report.xlsx", "KJ*EQUIPMENT*REPORT*.xlsx"]:
        files = sorted(DATA_DIR.glob(pattern))
        if files:
            try:
                df = pd.read_excel(files[-1], sheet_name="Sheet2", header=0)
                # Filter to our 4 brands
                df = df[df["EM2_MAKE"].isin(EQUIP_TARGET_MAKES)].copy()
                # Clean up
                df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
                df["EM2_MODEL"] = df["EM2_MODEL"].astype(str).str.strip()
                df["Brand"] = df["EM2_MAKE"].map(EQUIP_MAKE_MAP)
                df["EM_METER"] = pd.to_numeric(df["EM_METER"], errors="coerce").fillna(0).astype(int)
                df["Sell Price"] = pd.to_numeric(df["Sell Price"], errors="coerce").fillna(0)
                df["Parts and Service $"] = pd.to_numeric(df["Parts and Service $"], errors="coerce").fillna(0)
                return df
            except Exception:
                return pd.DataFrame()
    return pd.DataFrame()

def build_equipment_report_leads(equip_df, existing_customers):
    """Build lead rows from the equipment report for customers not already in other sources."""
    if equip_df.empty:
        return pd.DataFrame()

    existing_upper = {str(c).strip().upper() for c in existing_customers}

    # Group equipment by customer
    grouped = equip_df.groupby("Customer Name")
    rows = []
    for cust_name, grp in grouped:
        cust_upper = str(cust_name).strip().upper()
        if not cust_upper or cust_upper in ("NAN", "TOTAL", ""):
            continue
        # Skip internal SEC machines
        if "SOUTHEASTERN" in cust_upper or "RENTAL" in cust_upper:
            continue
        # Skip if already in CASE alerts or HubSpot
        already = False
        for ex in existing_upper:
            if ex and cust_upper and (ex == cust_upper or ex in cust_upper or cust_upper in ex):
                already = True
                break
        if already:
            continue

        # Aggregate this customer's equipment
        machines = len(grp)
        brands = grp["Brand"].unique().tolist()
        models = grp["EM2_MODEL"].unique().tolist()
        total_sell = grp["Sell Price"].sum()
        total_ps = grp["Parts and Service $"].sum()
        max_meter = grp["EM_METER"].max()
        latest_sale = grp["EM_SOLD_DATE"].max()
        top_model = grp["EM2_MODEL"].value_counts().index[0] if len(grp) > 0 else ""

        # Match the top model to dealsheet for PM value estimate
        ds_value, ds_model = get_dealsheet_pm_value(top_model, max_meter)
        if ds_value == 0:
            # Try other models
            for m in models:
                v, dm = get_dealsheet_pm_value(m, max_meter)
                if v > 0:
                    ds_value, ds_model = v, dm
                    break

        # Score this lead
        score = 0.0

        # Fleet size (machines bought from SEC)
        if machines >= 10:
            score += 30
        elif machines >= 5:
            score += 22
        elif machines >= 3:
            score += 15
        elif machines >= 2:
            score += 10
        else:
            score += 5

        # Total equipment value purchased
        if total_sell >= 500000:
            score += 20
        elif total_sell >= 200000:
            score += 15
        elif total_sell >= 100000:
            score += 10
        elif total_sell >= 50000:
            score += 7
        elif total_sell > 0:
            score += 3

        # Parts and service percentage (low % = not servicing with SEC = PM opportunity)
        ps_pct = grp["% of Machine in parts/service"].mean() if "% of Machine in parts/service" in grp.columns else 0
        ps_pct = float(ps_pct) if pd.notna(ps_pct) else 0
        if ps_pct == 0:
            score += 15  # No parts/service = prime PM target
        elif ps_pct < 5:
            score += 10
        elif ps_pct < 15:
            score += 5

        # Engine hours (active machine)
        if 500 <= max_meter <= 3000:
            score += 15
        elif max_meter > 3000:
            score += 10
        elif max_meter > 100:
            score += 5

        # Recency of purchase
        if pd.notna(latest_sale):
            try:
                days_ago = (datetime.now() - pd.Timestamp(latest_sale)).days
                if days_ago < 365:
                    score += 10
                elif days_ago < 730:
                    score += 7
                elif days_ago < 1095:
                    score += 4
                else:
                    score += 2
            except Exception:
                score += 2

        # Multiple brands = diversified fleet
        if len(brands) >= 3:
            score += 5
        elif len(brands) >= 2:
            score += 3

        score = max(0, min(100, round(score, 1)))

        # Tier
        if score >= 65:
            tier = "Top"
        elif score >= 50:
            tier = "High"
        elif score >= 35:
            tier = "Medium"
        else:
            tier = "Low"

        # PM value estimate
        est_pm_value = ds_value if ds_value > 0 else (machines * 2400)

        lead_cat = "Equipment Buyer (No Service)" if total_ps == 0 else "Equipment Buyer"

        rows.append({
            "Customer": cust_name.strip().title(),
            "Lead Score": score,
            "Tier": tier,
            "Source": "Equipment Report",
            "Location": "",
            "Model": top_model,
            "Dealsheet Model": ds_model,
            "VIN": "",
            "Eng Hrs": max_meter,
            "Stop": "",
            "Parts Value": total_ps,
            "Labor Hrs": 0,
            "Annual PM Value": est_pm_value,
            "In HubSpot": False,
            "HubSpot Deals": 0,
            "Lifecycle": "",
            "CASE Class": "",
            "Fleet": str(machines),
            "Has PM": False,
            "Lead Category": lead_cat,
            "Service Status": "",
            "YTD Parts": 0,
            "YTD Service": 0,
            "Total Spend": total_ps,
            "has_procare": False,
            "is_internal": False,
            "Equip Machines": machines,
            "Equip Brands": ", ".join(brands),
            "Equip Models": ", ".join(models[:5]),
            "Equip Total Sold": total_sell,
        })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    return df.sort_values("Lead Score", ascending=False)


# ═══════════════════════════════════════════════════════════
# HUBSPOT ENRICHMENT
# ═══════════════════════════════════════════════════════════
HUBSPOT_TOKEN = st.secrets.get("hubspot_token", "")

@st.cache_data(ttl=1800, show_spinner="Pulling HubSpot companies...")
def fetch_hubspot_companies():
    """Pull company records from HubSpot with rich scoring data."""
    if not HUBSPOT_TOKEN:
        return {}
    headers = {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}
    companies = {}
    url = "https://api.hubapi.com/crm/v3/objects/companies"
    params = {
        "limit": 100,
        "properties": ",".join([
            "name", "city", "state", "lifecyclestage", "num_associated_deals",
            "case_customer_classification", "case_ucc_prospect_classification",
            "fleet_size__c", "account_stage__c", "annualrevenue", "description",
            "eda_last_purchase_date", "hs_lastmodifieddate",
            "parts___service_engagement", "last_service_purchase",
            "last_parts_purchase", "last_parts_invoice_date__c",
            "sa_ytd_charges__c", "oe_ytd_charges__c",
        ]),
    }
    try:
        after = None
        for _ in range(40):  # max 4000 companies
            if after:
                params["after"] = after
            resp = requests.get(url, headers=headers, params=params, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            for c in data.get("results", []):
                props = c.get("properties", {})
                name = (props.get("name") or "").strip().upper()
                if name:
                    companies[name] = {
                        "hs_id": c["id"],
                        "city": props.get("city", ""),
                        "state": props.get("state", ""),
                        "lifecycle": props.get("lifecyclestage", ""),
                        "deals": props.get("num_associated_deals", 0),
                        "case_class": props.get("case_customer_classification", ""),
                        "prospect_class": props.get("case_ucc_prospect_classification", ""),
                        "fleet_size": props.get("fleet_size__c", ""),
                        "account_stage": props.get("account_stage__c", ""),
                        "annual_revenue": props.get("annualrevenue", ""),
                        "last_purchase": props.get("eda_last_purchase_date", ""),
                        "last_modified": props.get("hs_lastmodifieddate", ""),
                        "ps_engagement": props.get("parts___service_engagement", ""),
                        "last_service": props.get("last_service_purchase", ""),
                        "last_parts": props.get("last_parts_purchase", ""),
                        "last_parts_date": props.get("last_parts_invoice_date__c", ""),
                        "ytd_service": float(props.get("sa_ytd_charges__c") or 0),
                        "ytd_parts": float(props.get("oe_ytd_charges__c") or 0),
                        "description": props.get("description", ""),
                    }
            paging = data.get("paging", {}).get("next", {})
            after = paging.get("after")
            if not after:
                break
    except Exception:
        pass
    return companies

@st.cache_data(ttl=1800, show_spinner="Pulling HubSpot deal history...")
def fetch_hubspot_deals():
    """Pull deal records to build win/loss history and detect existing PM contracts."""
    if not HUBSPOT_TOKEN:
        return {}, set()
    headers = {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}
    url = "https://api.hubapi.com/crm/v3/objects/deals/search"

    # Track deals per company and PM-active companies
    deals_by_company = {}  # company_name -> {won, lost, total_won_amount, last_close, warranty_years, warranty_close}
    pm_active_companies = set()  # company names with active PM deals

    try:
        # Pull closed won deals with associations
        for stage, label in [("closedwon", "won"), ("closedlost", "lost")]:
            after_val = 0
            for _ in range(30):  # max 3000 per stage
                payload = {
                    "filterGroups": [{"filters": [{"propertyName": "dealstage", "operator": "EQ", "value": stage}]}],
                    "properties": ["dealname", "amount", "closedate", "primary_company_name", "pm_eligible", "pm_status", "warranty_type__c", "warranty_information__c"],
                    "limit": 100,
                    "after": str(after_val),
                }
                resp = requests.post(url, headers=headers, json=payload, timeout=15)
                if resp.status_code != 200:
                    break
                data = resp.json()
                for d in data.get("results", []):
                    props = d.get("properties", {})
                    co_name = (props.get("primary_company_name") or "").strip().upper()
                    if not co_name:
                        # Try to extract from deal name (format: "COMPANY - Location - Date")
                        dn = props.get("dealname", "")
                        if " - " in dn:
                            co_name = dn.split(" - ")[0].strip().upper()
                    if co_name:
                        if co_name not in deals_by_company:
                            deals_by_company[co_name] = {"won": 0, "lost": 0, "total_won_amount": 0, "last_close": "", "warranty_years": 0, "warranty_close": "", "deal_names": ""}
                        deals_by_company[co_name][label] += 1
                        dn = props.get("dealname", "")
                        if dn:
                            deals_by_company[co_name]["deal_names"] += f" {dn}"
                        amt = float(props.get("amount") or 0)
                        if label == "won":
                            deals_by_company[co_name]["total_won_amount"] += amt
                        cd = props.get("closedate", "")
                        if cd > deals_by_company[co_name]["last_close"]:
                            deals_by_company[co_name]["last_close"] = cd
                        # Parse warranty duration from warranty fields
                        wtype = (props.get("warranty_type__c") or "").lower()
                        winfo = (props.get("warranty_information__c") or "").upper()
                        if label == "won" and wtype and "no warranty" not in wtype and "as is" not in wtype and "n/a" not in wtype:
                            import re
                            wyears = 0
                            # Try to extract years: "3 YEAR", "3 YR", "36 MONTH", "48 MONTHS", "5 YEAR"
                            yr_match = re.search(r'(\d+)\s*(?:YEAR|YR)', winfo)
                            mo_match = re.search(r'(\d+)\s*(?:MONTH|MO)', winfo)
                            if yr_match:
                                wyears = int(yr_match.group(1))
                            elif mo_match:
                                wyears = int(mo_match.group(1)) / 12
                            elif "STANDARD" in winfo or "STD" in winfo or "FACTORY" in winfo:
                                wyears = 3  # Default CASE/Kobelco standard is 3yr/3000hr
                            elif winfo and winfo not in ("", "N/A"):
                                wyears = 2  # Reasonable default for other warranties
                            # Keep the longest warranty per company
                            if wyears > deals_by_company[co_name]["warranty_years"]:
                                deals_by_company[co_name]["warranty_years"] = wyears
                                deals_by_company[co_name]["warranty_close"] = cd

                paging = data.get("paging", {}).get("next", {})
                next_after = paging.get("after")
                if not next_after:
                    break
                after_val = int(next_after)

        # Pull PM-active deals (any pm_eligible deal that isn't closed lost)
        payload = {
            "filterGroups": [{"filters": [{"propertyName": "pm_eligible", "operator": "EQ", "value": "true"}]}],
            "properties": ["primary_company_name", "pm_status", "pm_contract_value", "dealstage", "dealname"],
            "limit": 100,
        }
        resp = requests.post(url, headers=headers, json=payload, timeout=15)
        if resp.status_code == 200:
            for d in resp.json().get("results", []):
                props = d.get("properties", {})
                pm_st = (props.get("pm_status") or "").lower()
                deal_st = (props.get("dealstage") or "").lower()
                # Active if not closed lost
                if deal_st != "closedlost" and pm_st not in ("closed lost", ""):
                    co_name = (props.get("primary_company_name") or "").strip().upper()
                    if not co_name:
                        dn = props.get("dealname", "")
                        if " - " in dn:
                            co_name = dn.split(" - ")[0].strip().upper()
                    if co_name:
                        pm_active_companies.add(co_name)

    except Exception:
        pass
    return deals_by_company, pm_active_companies

def _match_hubspot_company(cust_upper, hs_companies):
    """Match a customer name to HubSpot, trying exact then partial."""
    match = hs_companies.get(cust_upper)
    if match:
        return match
    for hs_name, hs_data in hs_companies.items():
        if cust_upper in hs_name or hs_name in cust_upper:
            return hs_data
    return None


# CASE classification scoring: how warm is this customer relationship?
CASE_CLASS_BOOST = {
    "Champion": 0.12,           # Best customers, lock them in with PM
    "Loyal Customer": 0.10,     # Strong relationship, easy sell
    "Potential Loyalist": 0.08, # Growing relationship, PM solidifies it
    "New Customer": 0.06,       # Just started buying, PM builds stickiness
    "Promising": 0.05,          # Showing interest
    "Needs Attention": 0.15,    # HIGHEST boost: drifting away, PM re-engages
    "About to Sleep": 0.12,     # At risk of leaving, PM creates recurring touchpoint
    "Can't Lose Them": 0.14,    # High value at risk, PM is retention play
    "At Risk": 0.13,            # Slipping away, PM contract locks in relationship
    "Hibernating": 0.04,        # Gone quiet, still worth a shot but lower priority
}

# Fleet size scoring: bigger fleets = bigger contracts
FLEET_SIZE_SCORE = {
    "0 - Rent Only": 0,     # Renters don't need PM
    "1-3": 0.02,
    "4-10": 0.04,
    "11-25": 0.06,
    "26+": 0.07,
    "26-50": 0.07,
    "51-100": 0.09,
    "101-250": 0.10,
    "251-500": 0.12,
}


def _classify_lead_category(match, deal_info, has_procare_expired, has_pm, cust_upper):
    """
    Assign a lead category based on Nick's priority tiers:
    1. No ProCare Coverage - machine has alerts but no ProCare (base case from scoring)
    2. Warranty Expiring Soon - warranty ends within 6 months (hottest warranty lead)
    3. Warranty Expired - had warranty, now expired (no coverage, needs PM)
    4. Regular Maintenance - comes in for service but no PM contract
    5. Parts Only, No Service - buys parts but handles own service (pitch PM)
    6. Unknown - in alerts but no HubSpot match
    """
    if not match:
        return "No ProCare (New Lead)"

    ps_engagement = match.get("ps_engagement", "")
    last_service = match.get("last_service", "")
    last_parts = match.get("last_parts", "")
    ytd_service = match.get("ytd_service", 0)
    ytd_parts = match.get("ytd_parts", 0)

    # Already has active PM, still show but tag it
    if has_pm:
        return "Active PM (Upsell)"

    # Warranty status: expiring soon vs expired vs still active
    warranty_years = deal_info.get("warranty_years", 0) if deal_info else 0
    warranty_close = deal_info.get("warranty_close", "") if deal_info else ""
    if warranty_years > 0 and warranty_close:
        try:
            close_dt = datetime.strptime(warranty_close[:10], "%Y-%m-%d")
            from dateutil.relativedelta import relativedelta
            expiry_dt = close_dt + relativedelta(years=int(warranty_years), months=int((warranty_years % 1) * 12))
            days_until = (expiry_dt - datetime.now()).days
            if days_until < 0:
                return "Warranty Expired"
            elif days_until <= 180:
                return "Warranty Expiring"
            # else: warranty still active 6+ months out, don't tag as warranty lead
        except Exception:
            pass

    # Category 4: Parts buyer, no service - "let us handle the service"
    # This is the key one Nick called out
    if ps_engagement == "Customer purchases parts from SEC, but mostly manages their own service":
        return "Parts Only, No Service"
    # Fallback: has parts activity but no service activity
    if last_parts and last_parts != "No Purchase" and (last_service == "No Purchase" or not last_service):
        return "Parts Only, No Service"
    if ytd_parts > 0 and ytd_service == 0:
        return "Parts Only, No Service"

    # Category 3: Regular maintenance customers - already bring machines to SEC for service
    if ps_engagement == "Customer wants SEC to manage Parts and Service for their fleet":
        return "Full Service (Lock In)"
    if last_service in ("Hot (0-3 Months)", "Warm (3-6 Months)"):
        return "Active Service Customer"
    if last_service in ("Cool (6-12 Months)", "Cold (12-18 Months)"):
        return "Lapsed Service"
    if ytd_service > 0:
        return "Active Service Customer"

    # Category 1 with CRM context
    return "No ProCare (In CRM)"


def enrich_leads_with_hubspot(scored_df, hs_companies, deal_history=None, pm_active_companies=None):
    """
    Enrich leads with HubSpot intelligence and assign Lead Categories.

    Lead Categories (Nick's priority tiers):
    1. No ProCare - machines throwing alerts without coverage
    2. Warranty Expiring - warranty ends within 6 months (hottest)
    3. Warranty Expired - had warranty, now expired (needs PM)
    4. Parts Only, No Service - buys parts, does own service (pitch PM)
    5. Active/Lapsed Service - already uses SEC service (lock them in with PM)
    6. Active PM - already has PM (upsell or skip)

    Scoring boosts based on:
    - CASE classification (relationship health)
    - Fleet size (bigger = more contract value)
    - Deal history (repeat buyers are warmer)
    - Parts/service engagement (parts-only = prime target)
    - Recency of last purchase
    """
    if not hs_companies or scored_df.empty:
        scored_df["In HubSpot"] = False
        scored_df["HubSpot Deals"] = 0
        scored_df["Lifecycle"] = ""
        scored_df["CASE Class"] = ""
        scored_df["Fleet"] = ""
        scored_df["Has PM"] = False
        scored_df["Lead Category"] = "No ProCare (New Lead)"
        scored_df["Service Status"] = ""
        return scored_df

    deal_history = deal_history or {}
    pm_active_companies = pm_active_companies or set()

    df = scored_df.copy()
    df["cust_upper"] = df["Customer"].str.strip().str.upper()

    # Match each row to HubSpot
    hs_match = []
    hs_deals = []
    hs_lifecycle = []
    hs_case_class = []
    hs_fleet = []
    hs_has_pm = []
    hs_category = []
    hs_service_status = []
    hs_boost = []

    for _, row in df.iterrows():
        cust = row["cust_upper"]
        match = _match_hubspot_company(cust, hs_companies)

        hs_match.append(bool(match))
        hs_deals.append(int(match["deals"]) if match and match.get("deals") else 0)
        hs_lifecycle.append(match["lifecycle"] if match else "")
        case_class = match["case_class"] if match else ""
        hs_case_class.append(case_class)
        fleet = match["fleet_size"] if match else ""
        hs_fleet.append(fleet)

        has_pm = cust in pm_active_companies
        hs_has_pm.append(has_pm)

        # Service status from HubSpot
        svc = match.get("last_service", "") if match else ""
        hs_service_status.append(svc)

        # Lead category
        deal_info = deal_history.get(cust, {})
        category = _classify_lead_category(match, deal_info, False, has_pm, cust)
        hs_category.append(category)

        # Calculate composite HubSpot boost
        boost = 0.0

        if match:
            boost += 0.03  # In CRM

            # CASE classification boost
            boost += CASE_CLASS_BOOST.get(case_class, 0)

            # Fleet size boost
            boost += FLEET_SIZE_SCORE.get(fleet, 0)

            # Deal history
            won = deal_info.get("won", 0)
            lost = deal_info.get("lost", 0)
            if won > 0:
                win_rate = won / max(won + lost, 1)
                boost += min(win_rate * 0.08, 0.08)
                if won >= 10:
                    boost += 0.05
                elif won >= 5:
                    boost += 0.03
                elif won >= 2:
                    boost += 0.01

            # Recency
            last_purchase = match.get("last_purchase", "")
            if last_purchase:
                try:
                    lp_date = datetime.strptime(last_purchase[:10], "%Y-%m-%d")
                    days_ago = (datetime.now() - lp_date).days
                    if days_ago < 90:
                        boost += 0.06
                    elif days_ago < 180:
                        boost += 0.04
                    elif days_ago < 365:
                        boost += 0.02
                except (ValueError, TypeError):
                    pass

            # Category-specific boosts
            if category == "Parts Only, No Service":
                boost += 0.12  # Prime targets: they trust SEC for parts, pitch service+PM
            elif category == "Warranty Expiring":
                boost += 0.15  # Hottest warranty lead: coverage ending soon, perfect PM pitch
            elif category == "Warranty Expired":
                boost += 0.12  # No coverage right now, needs PM
            elif category == "Active Service Customer":
                boost += 0.08  # Already using SEC service, PM is easy sell
            elif category == "Full Service (Lock In)":
                boost += 0.10  # They want full service, PM formalizes it
            elif category == "Lapsed Service":
                boost += 0.06  # Used to come in, PM brings them back

            # Penalty: rent-only
            if fleet == "0 - Rent Only":
                boost -= 0.10

            # Penalty: already has PM
            if has_pm:
                boost -= 0.15

        hs_boost.append(boost)

    df["In HubSpot"] = hs_match
    df["HubSpot Deals"] = hs_deals
    df["Lifecycle"] = hs_lifecycle
    df["CASE Class"] = hs_case_class
    df["Fleet"] = hs_fleet
    df["Has PM"] = hs_has_pm
    df["Lead Category"] = hs_category
    df["Service Status"] = hs_service_status

    # Apply the composite boost
    df["HS Boost"] = hs_boost
    df["Lead Score"] = (df["Lead Score"] * (1 + df["HS Boost"])).clip(0, 100).round(1)

    # Re-tier
    df["Tier"] = pd.cut(
        df["Lead Score"],
        bins=[0, 35, 50, 65, 100],
        labels=["Low", "Medium", "High", "Top"],
        include_lowest=True,
    )

    # Add spend columns for CASE leads (from HubSpot match data)
    ytd_parts_vals = []
    ytd_service_vals = []
    for _, row in df.iterrows():
        cust = row["cust_upper"]
        match = _match_hubspot_company(cust, hs_companies)
        if match:
            ytd_parts_vals.append(float(match.get("ytd_parts", 0) or 0))
            ytd_service_vals.append(float(match.get("ytd_service", 0) or 0))
        else:
            ytd_parts_vals.append(0.0)
            ytd_service_vals.append(0.0)
    df["YTD Parts"] = ytd_parts_vals
    df["YTD Service"] = ytd_service_vals
    df["Total Spend"] = df["YTD Parts"] + df["YTD Service"]

    df.drop(columns=["cust_upper", "HS Boost"], inplace=True)
    return df


def build_hubspot_only_leads(hs_companies, deal_history, pm_active_companies, existing_customers):
    """
    Build lead rows from HubSpot companies that are NOT already in the CASE alerts.
    These are customers with purchase/service history who could benefit from PM contracts.
    Score them based on spend, fleet size, deal history, and relationship status.
    """
    if not hs_companies:
        return pd.DataFrame()

    deal_history = deal_history or {}
    pm_active_companies = pm_active_companies or set()

    # Normalize existing customer names for matching
    existing_upper = set()
    for c in existing_customers:
        existing_upper.add(str(c).strip().upper())

    rows = []
    for hs_name, data in hs_companies.items():
        # Skip if already in CASE alerts
        if hs_name in existing_upper:
            continue
        # Skip partial matches too
        skip = False
        for ex in existing_upper:
            if ex and hs_name and (ex in hs_name or hs_name in ex):
                skip = True
                break
        if skip:
            continue

        # Skip rent-only accounts
        fleet = data.get("fleet_size", "")
        if fleet == "0 - Rent Only":
            continue

        # Skip if already has active PM
        has_pm = hs_name in pm_active_companies

        # Calculate spend signals
        ytd_parts = float(data.get("ytd_parts", 0) or 0)
        ytd_service = float(data.get("ytd_service", 0) or 0)
        total_spend = ytd_parts + ytd_service

        deal_info = deal_history.get(hs_name, {})
        deals_won = deal_info.get("won", 0)
        total_won_amount = deal_info.get("total_won_amount", 0)
        warranty_years = deal_info.get("warranty_years", 0)
        warranty_close = deal_info.get("warranty_close", "")

        # Calculate warranty status
        warranty_status = ""  # "", "expiring", "expired", "active"
        if warranty_years > 0 and warranty_close:
            try:
                close_dt = datetime.strptime(warranty_close[:10], "%Y-%m-%d")
                from dateutil.relativedelta import relativedelta
                expiry_dt = close_dt + relativedelta(years=int(warranty_years), months=int((warranty_years % 1) * 12))
                days_until = (expiry_dt - datetime.now()).days
                if days_until < 0:
                    warranty_status = "expired"
                elif days_until <= 180:
                    warranty_status = "expiring"
                else:
                    warranty_status = "active"
            except Exception:
                pass

        # Determine if this customer has enough signal to be worth showing
        # Need at least SOME indicator: spend, deals, fleet info, or classification
        case_class = data.get("case_class", "")
        prospect_class = data.get("prospect_class", "")
        account_stage = data.get("account_stage", "")
        lifecycle = data.get("lifecycle", "")
        annual_rev = data.get("annual_revenue", "")

        has_signal = (
            total_spend > 0 or
            deals_won > 0 or
            fleet not in ("", "0 - Rent Only") or
            case_class != "" or
            warranty_status in ("expiring", "expired")
        )

        if not has_signal:
            continue

        # ── Score this lead (0-100) ──
        score = 0.0

        # Spend score (biggest weight): parts + service YTD
        # $50K+ = 100, $20K = 70, $10K = 55, $5K = 40, $1K = 20
        if total_spend >= 50000:
            score += 35
        elif total_spend >= 20000:
            score += 30
        elif total_spend >= 10000:
            score += 25
        elif total_spend >= 5000:
            score += 20
        elif total_spend >= 1000:
            score += 12
        elif total_spend > 0:
            score += 6

        # Parts-only bonus (buying parts but no service = doing their own maintenance)
        if ytd_parts > 0 and ytd_service == 0:
            score += 10  # Prime PM conversion target

        # Fleet size score
        fleet_points = {
            "1-3": 5, "4-10": 10, "11-25": 18,
            "26+": 22, "26-50": 22, "51-100": 28,
            "101-250": 32, "251-500": 35,
        }
        score += fleet_points.get(fleet, 0)

        # Deal history: repeat buyers who trust SEC
        if deals_won >= 10:
            score += 15
        elif deals_won >= 5:
            score += 12
        elif deals_won >= 2:
            score += 8
        elif deals_won >= 1:
            score += 4

        # Won deal amount (bigger deals = bigger customer)
        if total_won_amount >= 500000:
            score += 10
        elif total_won_amount >= 100000:
            score += 7
        elif total_won_amount >= 50000:
            score += 4

        # CASE classification boost
        case_boost = CASE_CLASS_BOOST.get(case_class, 0) * 100
        score += case_boost

        # Warranty status scoring: expiring soon is the hottest
        if warranty_status == "expiring":
            score += 15  # Warranty ending soon, perfect PM timing
        elif warranty_status == "expired":
            score += 10  # No coverage now, needs PM

        # Recency: recent activity = warmer lead
        last_purchase = data.get("last_purchase", "")
        if last_purchase:
            try:
                lp_date = datetime.strptime(last_purchase[:10], "%Y-%m-%d")
                days_ago = (datetime.now() - lp_date).days
                if days_ago < 90:
                    score += 8
                elif days_ago < 180:
                    score += 5
                elif days_ago < 365:
                    score += 3
            except (ValueError, TypeError):
                pass

        # Penalty for active PM (they're already covered, just upsell)
        if has_pm:
            score -= 20

        score = max(0, min(100, round(score, 1)))

        # Assign lead category
        ps_engagement = data.get("ps_engagement", "")
        last_service = data.get("last_service", "")
        last_parts = data.get("last_parts", "")

        if has_pm:
            lead_cat = "Active PM (Upsell)"
        elif warranty_status == "expiring":
            lead_cat = "Warranty Expiring"
        elif warranty_status == "expired":
            lead_cat = "Warranty Expired"
        elif ps_engagement == "Customer purchases parts from SEC, but mostly manages their own service":
            lead_cat = "Parts Only, No Service"
        elif ytd_parts > 0 and ytd_service == 0:
            lead_cat = "Parts Only, No Service"
        elif last_parts and last_parts != "No Purchase" and (last_service == "No Purchase" or not last_service):
            lead_cat = "Parts Only, No Service"
        elif ps_engagement == "Customer wants SEC to manage Parts and Service for their fleet":
            lead_cat = "Full Service (Lock In)"
        elif last_service in ("Hot (0-3 Months)", "Warm (3-6 Months)"):
            lead_cat = "Active Service Customer"
        elif last_service in ("Cool (6-12 Months)", "Cold (12-18 Months)"):
            lead_cat = "Lapsed Service"
        elif ytd_service > 0:
            lead_cat = "Active Service Customer"
        elif deals_won > 0:
            lead_cat = "Equipment Buyer (No Service)"
        else:
            lead_cat = "In CRM (Prospect)"

        # Tier
        if score >= 65:
            tier = "Top"
        elif score >= 50:
            tier = "High"
        elif score >= 35:
            tier = "Medium"
        else:
            tier = "Low"

        # Estimate annual PM value based on fleet size
        fleet_pm_estimate = {
            "1-3": 4800, "4-10": 16800, "11-25": 43200,
            "26+": 62400, "26-50": 62400, "51-100": 96000,
            "101-250": 144000, "251-500": 240000,
        }
        est_pm_value = fleet_pm_estimate.get(fleet, 3600)

        rows.append({
            "Customer": hs_name.title(),
            "Lead Score": score,
            "Tier": tier,
            "Source": "HubSpot",
            "Location": data.get("city", "") or "",
            "Model": "",
            "Dealsheet Model": None,
            "VIN": "",
            "Eng Hrs": 0,
            "Stop": "",
            "Parts Value": ytd_parts,
            "Labor Hrs": ytd_service / 150 if ytd_service > 0 else 0,  # rough estimate
            "Annual PM Value": est_pm_value,
            "In HubSpot": True,
            "HubSpot Deals": deals_won,
            "Lifecycle": lifecycle,
            "CASE Class": case_class,
            "Fleet": fleet,
            "Has PM": has_pm,
            "Lead Category": lead_cat,
            "Service Status": last_service,
            "YTD Parts": ytd_parts,
            "YTD Service": ytd_service,
            "Total Spend": total_spend,
            "has_procare": False,
            "is_internal": False,
        })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    return df.sort_values("Lead Score", ascending=False)


# ═══════════════════════════════════════════════════════════
# LEAD SCORING ALGORITHM
# ═══════════════════════════════════════════════════════════
def parse_maintenance_alerts(file):
    """Parse Case maintenance alerts Excel into machine-level rows."""
    df = pd.read_excel(file)
    machines = df[
        (df["Customer"].notna()) & (df["Customer"] != "Total") &
        (df["VIN"].notna()) & (df["VIN"] != "Total") &
        (df["Model"].notna()) & (df["Model"] != "Total") &
        (df["Location"].notna()) & (df["Location"] != "Total")
    ].copy()
    machines["Parts Value"] = pd.to_numeric(machines["Parts Value"], errors="coerce").fillna(0)
    machines["Labor Hrs"] = pd.to_numeric(machines["Labor Hrs"], errors="coerce").fillna(0)
    machines["Eng Hrs"] = pd.to_numeric(machines["Eng Hrs"], errors="coerce").fillna(0)
    return machines

def parse_procare_stops(file):
    """Parse ProCare stops Excel, return set of VINs with active ProCare."""
    df = pd.read_excel(file)
    machines = df[
        (df["VinHrs"].notna()) & (df["VinHrs"] != "Total") &
        (df["Model"].notna()) & (df["Model"] != "Total")
    ].copy()
    machines["VIN"] = machines["VinHrs"].apply(lambda x: str(x).split(" - ")[0].strip())
    return set(machines["VIN"].unique())

def get_dealsheet_pm_value(alert_model, eng_hours):
    """Calculate real PM value for a machine using dealsheet pricing and actual hours."""
    ds_key = match_model_to_dealsheet(alert_model)
    if not ds_key:
        return 0.0, None
    hrs = max(int(eng_hours or 0), 500)  # Minimum 500 hrs for estimate
    result = calculate_pm_cost(ds_key, hrs)
    if result:
        return float(result["total_cost"]), ds_key
    return 0.0, None

def score_leads(alerts_df, procare_vins):
    """
    Score each machine/customer combination.
    Higher score = better PM lead opportunity.

    Scoring factors:
    - No ProCare coverage (required, or excluded)
    - Parts value (higher = more PM revenue opportunity)
    - Labor hours (higher = more complex service = more value)
    - Engine hours (sweet spot: 500-3000 hrs = active machine needing PM)
    - Service stop type (higher hour stops = overdue maintenance)
    - Multiple machines per customer (fleet deals)
    - Not an internal SEC machine
    """
    df = alerts_df.copy()

    # Exclude machines with ProCare
    df["has_procare"] = df["VIN"].isin(procare_vins)
    df = df[~df["has_procare"]].copy()

    # Exclude internal SEC machines
    df["is_internal"] = (
        df["Customer"].str.contains("Southeastern Equipment", case=False, na=False) |
        df["Customer"].str.contains("RENTAL", case=False, na=False)
    )
    df = df[~df["is_internal"]].copy()

    if df.empty:
        return df

    # Match models to dealsheet and calculate real PM values
    ds_results = df.apply(lambda r: get_dealsheet_pm_value(r["Model"], r["Eng Hrs"]), axis=1)
    df["Annual PM Value"] = ds_results.apply(lambda x: x[0])
    df["Dealsheet Model"] = ds_results.apply(lambda x: x[1])

    # Filter to only machines that match the dealsheet (4 target brands)
    df = df[df["Dealsheet Model"].notna()].copy()
    if df.empty:
        return df

    # ── Score components (0-100 scale each) ──

    # Parts value score: higher parts = more opportunity
    max_parts = df["Parts Value"].quantile(0.95) if len(df) > 10 else df["Parts Value"].max()
    max_parts = max(max_parts, 1)
    df["parts_score"] = (df["Parts Value"] / max_parts * 100).clip(0, 100)

    # Labor hours score
    max_labor = df["Labor Hrs"].quantile(0.95) if len(df) > 10 else df["Labor Hrs"].max()
    max_labor = max(max_labor, 1)
    df["labor_score"] = (df["Labor Hrs"] / max_labor * 100).clip(0, 100)

    # Engine hours score: sweet spot 500-3000, penalize very low (not using it)
    def hours_score(h):
        if h < 100:
            return 15
        elif h < 500:
            return 40
        elif h < 1500:
            return 85
        elif h < 3000:
            return 100
        elif h < 5000:
            return 75
        else:
            return 50  # very high hours, machine may be near end of life
    df["hours_score"] = df["Eng Hrs"].apply(hours_score)

    # Stop type score: higher service stops = more overdue = hotter lead
    stop_scores = {
        "50 Hr Stop": 20, "100 Hr Stop": 30, "150 Hr Stop": 35,
        "250 Hr Stop": 45, "500 Hr Stop": 60, "1000 Hr Stop": 80,
        "1500 Hr Stop": 85, "2000 Hr Stop": 90, "2500 Hr Stop": 92,
        "3000 Hr Stop": 95, "3500 Hr Stop": 97, "4000 Hr Stop": 98,
        "4500 Hr Stop": 99, "5000 Hr Stop": 100, "Other": 50,
    }
    df["stop_score"] = df["Stop"].map(stop_scores).fillna(50)

    # Fleet multiplier: even 1 machine is worth calling, scale from there
    fleet_counts = df.groupby("Customer")["VIN"].transform("nunique")
    df["fleet_count"] = fleet_counts
    # 1 machine = 40, 3 = 60, 7+ = 80, 15+ = 100
    df["fleet_score"] = fleet_counts.apply(
        lambda n: min(100, 40 + (n - 1) * 10)
    ).clip(0, 100)

    # Annual PM value score: real dealsheet value relative to the max in this set
    max_annual = df["Annual PM Value"].quantile(0.95) if len(df) > 10 else df["Annual PM Value"].max()
    max_annual = max(max_annual, 1)
    df["value_score"] = (df["Annual PM Value"] / max_annual * 100).clip(0, 100)

    # ── Weighted composite score ──
    df["Lead Score"] = (
        df["parts_score"]  * 0.15 +
        df["labor_score"]  * 0.10 +
        df["hours_score"]  * 0.20 +
        df["stop_score"]   * 0.20 +
        df["fleet_score"]  * 0.10 +
        df["value_score"]  * 0.25
    ).round(1)

    # Tier assignment
    df["Tier"] = pd.cut(
        df["Lead Score"],
        bins=[0, 35, 50, 65, 100],
        labels=["Low", "Medium", "High", "Top"],
        include_lowest=True,
    )

    df["Source"] = "CASE Alert"
    return df.sort_values("Lead Score", ascending=False)

def aggregate_customer_leads(scored_df):
    """Roll up machine-level scores to customer level."""
    if scored_df.empty:
        return pd.DataFrame()

    agg_dict = {
        "machines": ("VIN", lambda x: x.nunique() if x.notna().any() and (x != "").any() else 0),
        "total_parts_value": ("Parts Value", "sum"),
        "total_labor_hrs": ("Labor Hrs", "sum"),
        "avg_hours": ("Eng Hrs", "mean"),
        "avg_score": ("Lead Score", "mean"),
        "max_score": ("Lead Score", "max"),
        "total_annual_pm": ("Annual PM Value", "sum"),
        "location": ("Location", "first"),
        "models": ("Model", lambda x: ", ".join(sorted(set(m for m in x if m))) if any(m for m in x) else ""),
        "dealsheet_models": ("Dealsheet Model", lambda x: ", ".join(sorted(set(str(m) for m in x if m and str(m) != "None"))) if any(m for m in x) else ""),
    }
    # Carry through HubSpot fields if present
    if "CASE Class" in scored_df.columns:
        agg_dict["case_class"] = ("CASE Class", "first")
    if "Fleet" in scored_df.columns:
        agg_dict["fleet"] = ("Fleet", "first")
    if "In HubSpot" in scored_df.columns:
        agg_dict["in_hubspot"] = ("In HubSpot", "first")
    if "Has PM" in scored_df.columns:
        agg_dict["has_pm"] = ("Has PM", "first")
    if "HubSpot Deals" in scored_df.columns:
        agg_dict["hs_deals"] = ("HubSpot Deals", "first")
    if "Lead Category" in scored_df.columns:
        agg_dict["lead_category"] = ("Lead Category", "first")
    if "Service Status" in scored_df.columns:
        agg_dict["service_status"] = ("Service Status", "first")
    if "Source" in scored_df.columns:
        agg_dict["source"] = ("Source", "first")
    if "YTD Parts" in scored_df.columns:
        agg_dict["ytd_parts"] = ("YTD Parts", "first")
    if "YTD Service" in scored_df.columns:
        agg_dict["ytd_service"] = ("YTD Service", "first")
    if "Total Spend" in scored_df.columns:
        agg_dict["total_spend"] = ("Total Spend", "first")

    agg = scored_df.groupby("Customer").agg(**agg_dict).reset_index()

    # Customer-level score: weighted by fleet size and total opportunity
    max_pm = agg["total_annual_pm"].max() if agg["total_annual_pm"].max() > 0 else 1
    agg["Customer Score"] = (
        agg["avg_score"] * 0.5 +
        (agg["machines"].clip(0, 30) / max(30, 1) * 100) * 0.25 +
        (agg["total_annual_pm"] / max_pm * 100) * 0.25
    ).round(1)

    agg["Tier"] = pd.cut(
        agg["Customer Score"],
        bins=[0, 35, 50, 65, 100],
        labels=["Low", "Medium", "High", "Top"],
        include_lowest=True,
    )

    return agg.sort_values("Customer Score", ascending=False)


def build_lead_explanation(row):
    """Build a plain English explanation of why this customer is a lead and what the angle is."""
    reasons = []
    angle = ""

    cat = row.get("lead_category", "")
    fleet = row.get("fleet", "")
    machines = int(row.get("machines", 0))
    ytd_parts = float(row.get("ytd_parts", 0) or 0)
    ytd_service = float(row.get("ytd_service", 0) or 0)
    total_spend = float(row.get("total_spend", 0) or 0)
    deals = int(row.get("hs_deals", 0) or 0)
    case_class = row.get("case_class", "")
    source = row.get("source", "")
    models = row.get("models", "")
    svc_status = row.get("service_status", "")

    # Why they're on the list
    if source == "CASE Alert" and machines > 0:
        reasons.append(f"{machines} machine(s) with maintenance alerts ({models})" if models else f"{machines} machine(s) with maintenance alerts")
    elif source == "HubSpot" and fleet:
        reasons.append(f"Fleet size: {fleet}")

    if total_spend > 0:
        parts_str = f"${ytd_parts:,.0f} parts" if ytd_parts > 0 else ""
        svc_str = f"${ytd_service:,.0f} service" if ytd_service > 0 else ""
        spend_parts = [p for p in [parts_str, svc_str] if p]
        reasons.append(f"Spending ${total_spend:,.0f} YTD ({', '.join(spend_parts)})" if spend_parts else f"Spending ${total_spend:,.0f} YTD")
    elif ytd_parts > 0:
        reasons.append(f"Buying ${ytd_parts:,.0f} in parts YTD")

    if deals > 0:
        reasons.append(f"{deals} deal(s) in CRM history")

    if case_class:
        reasons.append(f"CASE classification: {case_class}")

    # The angle based on lead category
    if cat == "Warranty Expiring":
        angle = "Their warranty is expiring within 6 months. Perfect timing to transition them into a PM contract before they lose coverage."
    elif cat == "Warranty Expired":
        angle = "Their warranty already expired. They have no coverage right now. A PM contract fills that gap and keeps their machines running."
    elif cat == "Parts Only, No Service":
        angle = "They buy parts from us but do their own service work. Pitch the value of having SEC handle maintenance so they can focus on their jobs."
    elif cat == "Active Service Customer":
        angle = "They already bring machines to us for service. A PM contract locks in that relationship and gives them predictable costs."
    elif cat == "Full Service (Lock In)":
        angle = "They want us managing everything. A PM contract formalizes that and guarantees recurring revenue."
    elif cat == "Lapsed Service":
        angle = "Used to bring machines in for service but stopped. A PM contract is a reason to re-engage and bring them back."
    elif cat == "Equipment Buyer (No Service)":
        angle = "Bought equipment from us but never used our service department. Introduce PM as part of owning the machine."
    elif cat == "Active PM (Upsell)":
        angle = "Already has a PM contract. Look for upsell opportunities on additional machines or upgraded coverage."
    elif cat == "No ProCare (New Lead)":
        angle = "Machines are throwing maintenance alerts with no ProCare coverage. They need a PM plan before something breaks."
    elif cat == "No ProCare (In CRM)":
        angle = "In our CRM with maintenance alerts but no PM coverage. Start the conversation about preventing downtime."
    else:
        angle = "Potential PM opportunity based on their equipment and relationship with SEC."

    return reasons, angle


# ═══════════════════════════════════════════════════════════
# PDF GENERATION
# ═══════════════════════════════════════════════════════════
SEC_SLATE = "#7A8B9C"  # Slate-blue from SEC invoice template

def _pdf_header_footer(canvas_obj, doc):
    """Draw the SEC invoice-style header banner and footer on every page."""
    canvas_obj.saveState()
    w, h = letter

    # ── Header banner (right side, slate-blue) ──
    banner_h = 70
    banner_x = w * 0.38
    canvas_obj.setFillColor(colors.HexColor(SEC_SLATE))
    canvas_obj.rect(banner_x, h - banner_h - 10, w - banner_x, banner_h, fill=1, stroke=0)
    # Company name in banner
    canvas_obj.setFillColor(colors.white)
    canvas_obj.setFont("Helvetica-Bold", 22)
    canvas_obj.drawString(banner_x + 20, h - 42, "Southeastern")
    canvas_obj.setFont("Helvetica", 11)
    canvas_obj.drawString(banner_x + 20, h - 58, "EQUIPMENT COMPANY")

    # Left side: "PREVENTIVE MAINTENANCE QUOTE" label
    canvas_obj.setFillColor(colors.HexColor(SEC_RED))
    canvas_obj.setFont("Helvetica-Bold", 13)
    canvas_obj.drawString(doc.leftMargin, h - 38, "PREVENTIVE")
    canvas_obj.drawString(doc.leftMargin, h - 53, "MAINTENANCE QUOTE")

    # Thin red line under header
    canvas_obj.setStrokeColor(colors.HexColor(SEC_RED))
    canvas_obj.setLineWidth(2)
    canvas_obj.line(doc.leftMargin, h - banner_h - 16, w - doc.rightMargin, h - banner_h - 16)

    # ── Footer ──
    # Remit payment line
    canvas_obj.setStrokeColor(colors.HexColor(SEC_SLATE))
    canvas_obj.setLineWidth(0.5)
    canvas_obj.line(doc.leftMargin, 58, w - doc.rightMargin, 58)

    canvas_obj.setFillColor(colors.HexColor(SEC_DARK))
    canvas_obj.setFont("Helvetica-Bold", 7)
    canvas_obj.drawString(doc.leftMargin + 40, 63, "REMIT PAYMENT TO:")
    canvas_obj.setFont("Helvetica", 7)
    canvas_obj.drawString(doc.leftMargin + 145, 63, "Southeastern Equipment Co., Inc., 10874 East Pike Rd., Cambridge, Ohio 43725")

    # Bottom line: website | terms | EIN
    canvas_obj.setStrokeColor(colors.HexColor("#CCCCCC"))
    canvas_obj.line(doc.leftMargin, 44, w - doc.rightMargin, 44)

    canvas_obj.setFillColor(colors.HexColor(SEC_GRAY))
    canvas_obj.setFont("Helvetica", 7)
    canvas_obj.drawString(doc.leftMargin, 34, "www.southeasternequip.com")
    canvas_obj.drawCentredString(w / 2, 34, "See Reverse for Terms and Conditions")
    canvas_obj.drawRightString(w - doc.rightMargin, 34, "EIN 34-1503254")

    canvas_obj.restoreState()


def generate_pdf(quote_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        topMargin=1.2*inch,   # Room for header banner
        bottomMargin=1.0*inch,  # Room for footer
        leftMargin=0.6*inch,
        rightMargin=0.6*inch,
    )
    usable_w = letter[0] - doc.leftMargin - doc.rightMargin

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="BoxLabel", fontSize=7, fontName="Helvetica-Bold", textColor=colors.HexColor(SEC_GRAY), leading=10))
    styles.add(ParagraphStyle(name="BoxValue", fontSize=10, fontName="Helvetica", textColor=colors.HexColor(SEC_DARK), leading=13))
    styles.add(ParagraphStyle(name="SectionHead", fontSize=10, fontName="Helvetica-Bold", textColor=colors.HexColor(SEC_DARK), spaceBefore=10, spaceAfter=4))
    styles.add(ParagraphStyle(name="SmallNote", fontSize=8, fontName="Helvetica", textColor=colors.HexColor(SEC_GRAY), leading=10))

    elements = []

    # ── Quote info centered box ──
    quote_num = f"Q-{datetime.now().strftime('%y%m%d%H%M')}"
    quote_info_data = [
        [Paragraph("<b>Quote #</b>", styles["BoxLabel"]), Paragraph("<b>Date</b>", styles["BoxLabel"]), Paragraph("<b>Branch</b>", styles["BoxLabel"])],
        [Paragraph(quote_num, styles["BoxValue"]), Paragraph(quote_data.get("date", ""), styles["BoxValue"]), Paragraph(quote_data.get("branch", ""), styles["BoxValue"])],
    ]
    t = RLTable(quote_info_data, colWidths=[usable_w * 0.33] * 3)
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor(SEC_SLATE)),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 10))

    # ── Two side-by-side boxes: Customer (left) and Machine Info (right) ──
    half_w = usable_w * 0.48
    cust = quote_data.get("customer_name", "")
    rep = quote_data.get("rep", "")
    stype = quote_data.get("service_type", "")

    left_data = [
        [Paragraph("<b>CUSTOMER</b>", styles["BoxLabel"])],
        [Paragraph(cust if cust else " ", styles["BoxValue"])],
        [Paragraph(" ", styles["SmallNote"])],
        [Paragraph(f"<b>Service Rep:</b> {rep}" if rep else " ", styles["SmallNote"])],
        [Paragraph(f"<b>Service Type:</b> {stype}" if stype else " ", styles["SmallNote"])],
    ]
    left_t = RLTable(left_data, colWidths=[half_w])
    left_t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor(SEC_SLATE)),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))

    make = quote_data.get("make", "")
    model = quote_data.get("model", "")
    serial = quote_data.get("serial", "")
    hrs_req = quote_data.get("hours_requested", 0)

    right_data = [
        [Paragraph("<b>MACHINE DETAILS</b>", styles["BoxLabel"])],
        [Paragraph(f"{make} {model}" if make else " ", styles["BoxValue"])],
        [Paragraph(f"<b>Serial:</b> {serial}" if serial else " ", styles["SmallNote"])],
        [Paragraph(f"<b>Hours Requested:</b> {hrs_req:,}", styles["SmallNote"])],
        [Paragraph(" ", styles["SmallNote"])],
    ]
    right_t = RLTable(right_data, colWidths=[half_w])
    right_t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor(SEC_SLATE)),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))

    # Wrap both in an outer table for side-by-side
    outer = RLTable([[left_t, right_t]], colWidths=[half_w + 6, half_w + 6])
    outer.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(outer)
    elements.append(Spacer(1, 16))

    # ── Service interval pricing table ──
    intervals = quote_data.get("intervals", [])
    travel_cost = quote_data.get("travel_cost", 0)
    annual = quote_data.get("annual_pm_price", 0)

    price_data = [["Service", "Hour Interval", "Qty", "Cost (Per)", "Subtotal"]]
    for iv in intervals:
        price_data.append([
            iv["name"], f"{iv['hours']:,} hr", str(iv["qty"]),
            f"${iv['cost_per']:,.0f}", f"${iv['subtotal']:,.0f}",
        ])
    if travel_cost > 0:
        price_data.append(["Travel", "", "", "", f"${travel_cost:,.0f}"])
    price_data.append(["", "", "", "", ""])
    price_data.append(["PM Contract Total", "", "", "", f"${annual:,.0f}"])

    col_widths = [usable_w * 0.30, usable_w * 0.18, usable_w * 0.10, usable_w * 0.18, usable_w * 0.24]
    t = RLTable(price_data, colWidths=col_widths)
    t.setStyle(TableStyle([
        # Header row
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(SEC_SLATE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        # Body rows
        ("FONTNAME", (0, 1), (-1, -2), "Helvetica"), ("FONTSIZE", (0, 1), (-1, -2), 10),
        ("TOPPADDING", (0, 1), (-1, -1), 6), ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("LINEBELOW", (0, 1), (-1, -3), 0.25, colors.HexColor("#DDDDDD")),
        # Separator before total
        ("LINEBELOW", (0, -2), (-1, -2), 0.75, colors.HexColor(SEC_SLATE)),
        # Total row (highlight)
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("FONTSIZE", (0, -1), (-1, -1), 12),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor(SEC_RED)),
        ("LINEBELOW", (0, -1), (-1, -1), 1.5, colors.HexColor(SEC_RED)),
        ("TOPPADDING", (0, -1), (-1, -1), 8),
    ]))
    elements.append(t)

    # ── Notes ──
    notes = quote_data.get("notes", "")
    if notes:
        elements.append(Spacer(1, 14))
        elements.append(Paragraph("<b>Notes</b>", styles["SectionHead"]))
        elements.append(Paragraph(notes, styles["Normal"]))

    doc.build(elements, onFirstPage=_pdf_header_footer, onLaterPages=_pdf_header_footer)
    buffer.seek(0)
    return buffer


# ─── Helpers ───
def get_models_for_brand(brand):
    """Get sorted model list for a brand from the dealsheet."""
    return sorted(PM_BRANDS.get(brand, []))

def calculate_pm_cost(model_key, hours_requested):
    """Calculate PM contract cost using real dealsheet interval pricing.
    Returns breakdown of each service interval, quantities, and total.
    """
    ds = PM_DEALSHEET.get(model_key)
    if not ds:
        return None

    intervals = []
    total = 0

    # Initial service (one-time)
    if ds["hr_i"] and ds["cost_i"]:
        intervals.append({"name": "Initial Service", "hours": ds["hr_i"], "qty": 1, "cost_per": ds["cost_i"], "subtotal": ds["cost_i"]})
        total += ds["cost_i"]

    # Interval 1 (e.g. 250hr) - only some models have this
    if ds["hr_1"] and ds["cost_1"]:
        # Count how many Interval 1 services occur, minus those covered by Interval 2
        n1_total = (hours_requested - 1) // ds["hr_1"] if ds["hr_1"] > 0 else 0
        n2_total = (hours_requested - 1) // ds["hr_2"] if ds["hr_2"] > 0 else 0
        n1_net = max(n1_total - n2_total, 0)
        if n1_net > 0:
            intervals.append({"name": f"Interval 1 ({ds['hr_1']}hr)", "hours": ds["hr_1"], "qty": n1_net, "cost_per": ds["cost_1"], "subtotal": n1_net * ds["cost_1"]})
            total += n1_net * ds["cost_1"]

    # Interval 2 (typically 500hr)
    if ds["hr_2"] and ds["cost_2"]:
        n2_total = (hours_requested - 1) // ds["hr_2"] if ds["hr_2"] > 0 else 0
        n3_total = (hours_requested - 1) // ds["hr_3"] if ds["hr_3"] > 0 else 0
        n2_net = max(n2_total - n3_total, 0)
        if n2_net > 0:
            intervals.append({"name": f"Interval 2 ({ds['hr_2']}hr)", "hours": ds["hr_2"], "qty": n2_net, "cost_per": ds["cost_2"], "subtotal": n2_net * ds["cost_2"]})
            total += n2_net * ds["cost_2"]

    # Interval 3 (typically 1000hr)
    if ds["hr_3"] and ds["cost_3"]:
        n3_total = (hours_requested - 1) // ds["hr_3"] if ds["hr_3"] > 0 else 0
        if n3_total > 0:
            intervals.append({"name": f"Interval 3 ({ds['hr_3']}hr)", "hours": ds["hr_3"], "qty": n3_total, "cost_per": ds["cost_3"], "subtotal": n3_total * ds["cost_3"]})
            total += n3_total * ds["cost_3"]

    # Specialty service (one-time at specific hour mark)
    if ds["hr_s"] and ds["cost_s"]:
        n_s = (hours_requested - 1) // ds["hr_s"] if ds["hr_s"] > 0 else 0
        if n_s > 0:
            intervals.append({"name": f"Specialty ({ds['hr_s']}hr)", "hours": ds["hr_s"], "qty": n_s, "cost_per": ds["cost_s"], "subtotal": n_s * ds["cost_s"]})
            total += n_s * ds["cost_s"]

    return {
        "intervals": intervals,
        "total_cost": total,
        "brand": ds["brand"],
        "model": model_key,
    }


# ═══════════════════════════════════════════════════════════
# CUSTOM CSS
# ═══════════════════════════════════════════════════════════
st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #C8102E 0%, #8B0000 100%); padding: 20px 30px; border-radius: 10px; margin-bottom: 24px; }
    .main-header h1 { color: white !important; font-size: 28px !important; margin: 0 !important; }
    .main-header p { color: rgba(255,255,255,0.85) !important; font-size: 14px !important; margin: 4px 0 0 0 !important; }
    .metric-card { background: white; border: 1px solid #E5E7EB; border-radius: 8px; padding: 16px; text-align: center; border-top: 3px solid #C8102E; }
    .metric-card .label { font-size: 12px; color: #6B7280; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-card .value { font-size: 24px; font-weight: 700; color: #1A1A1A; margin-top: 4px; }
    .lead-top { border-left: 4px solid #C8102E; padding-left: 12px; margin-bottom: 8px; }
    .lead-high { border-left: 4px solid #F59E0B; padding-left: 12px; margin-bottom: 8px; }
    .lead-med { border-left: 4px solid #3B82F6; padding-left: 12px; margin-bottom: 8px; }
    /* lead cards are fully inline-styled for Streamlit compatibility */
    .login-container { text-align: center; padding: 60px 20px; }
    .login-title { color: #C8102E; font-size: 36px; font-weight: 600; margin-bottom: 8px; }
    .login-subtitle { color: #666; font-size: 18px; margin-bottom: 8px; }
    .login-detail { color: #888; font-size: 13px; margin-bottom: 40px; }
    .header-bar { background: linear-gradient(135deg, #C8102E 0%, #8B0000 100%); color: white; padding: 14px 24px; border-radius: 8px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center; }
    .header-bar .branch-name { font-size: 20px; font-weight: 600; }
    .header-bar .rep-info { font-size: 14px; opacity: 0.85; }
    .region-header { font-size: 16px; font-weight: 600; margin: 16px 0 8px 0; padding: 8px 12px; background: #F3F4F6; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
# LOGIN PAGE
# ═══════════════════════════════════════════════════════════
def show_login():
    st.markdown("""
    <div class="login-container">
        <div class="login-title">PM Campaign</div>
        <div class="login-subtitle">2026</div>
        <div class="login-detail">Lead Discovery &bull; PM Quoting &bull; Branch Tracking</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("#### Select Your Branch")
        options = []
        branch_map = {}
        for num, name in sorted(BRANCHES.items()):
            label = f"{num} - {name}"
            options.append(label)
            branch_map[label] = num
        selected = st.selectbox("Branch", options, label_visibility="collapsed")

        st.markdown("#### Month")
        month_options = ["January", "February", "March", "April", "May", "June",
                         "July", "August", "September", "October", "November", "December"]
        current_month_idx = datetime.now().month - 1
        selected_month = st.selectbox("Month", month_options, index=current_month_idx, label_visibility="collapsed")

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("Start", use_container_width=True, type="primary"):
            if selected:
                branch_id = branch_map[selected]
                st.session_state.branch = branch_id
                st.session_state.branch_name = BRANCHES[branch_id]
                st.session_state.login_month = selected_month
                st.session_state.page = "dashboard"
                st.rerun()
            else:
                st.warning("Pick your branch to continue.")

        st.markdown("<br><br>", unsafe_allow_html=True)
        if st.button("Admin Dashboard", use_container_width=True):
            st.session_state.page = "admin_login"
            st.rerun()

    st.markdown("""
    <div style="text-align:center; color:#999; font-size:12px; margin-top:60px;">
        Southeastern Equipment Co. &bull; PM Program 2026
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════
# ADMIN LOGIN
# ═══════════════════════════════════════════════════════════
def show_admin_login():
    st.markdown("""
    <div class="login-container">
        <div class="login-title">Admin Dashboard</div>
        <div class="login-subtitle">PM Program Performance</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        password = st.text_input("Password", type="password")

        st.markdown("<br>", unsafe_allow_html=True)

        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Back", use_container_width=True):
                st.session_state.page = "login"
                st.rerun()
        with col_b:
            if st.button("Login", use_container_width=True, type="primary"):
                if password == ADMIN_PASSWORD:
                    st.session_state.page = "admin"
                    st.session_state.branch_name = "All Branches"
                    st.session_state.login_month = "All"
                    st.rerun()
                else:
                    st.error("Incorrect password")


# ═══════════════════════════════════════════════════════════
# ADMIN DASHBOARD
# ═══════════════════════════════════════════════════════════
def show_admin_dashboard():
    import plotly.express as px

    st.markdown(f"""
    <div class="header-bar">
        <span class="branch-name">Admin Dashboard &mdash; PM Program</span>
        <span class="rep-info">All Regions &bull; 2026</span>
    </div>
    """, unsafe_allow_html=True)

    col_l, col_r = st.columns([6, 1])
    with col_r:
        if st.button("Logout"):
            st.session_state.page = "login"
            st.session_state.branch = None
            st.session_state.branch_name = None
            st.session_state.login_month = None
            st.rerun()

    # Load tracking data from Google Sheets
    sheet = get_gsheet_connection()
    tracking_data = []
    if sheet:
        try:
            ws_list = sheet.spreadsheet.worksheets()
            tracking_ws = None
            for ws in ws_list:
                if ws.title == "Tracking":
                    tracking_ws = ws
                    break
            if tracking_ws:
                records = tracking_ws.get_all_records()
                if records:
                    tracking_data = records
        except Exception:
            pass

    tracking_df = pd.DataFrame(tracking_data) if tracking_data else pd.DataFrame()

    # Month filter for admin
    month_options = ["All", "January", "February", "March", "April", "May", "June",
                     "July", "August", "September", "October", "November", "December"]
    admin_month = st.selectbox("Filter by Month", month_options, index=0)

    filtered_df = tracking_df.copy()
    if admin_month != "All" and not filtered_df.empty and "Month" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Month"] == admin_month]

    # Summary metrics
    st.markdown("### Program Overview")
    col1, col2, col3, col4 = st.columns(4)

    if not filtered_df.empty and "Status" in filtered_df.columns:
        total_tracked = len(filtered_df)
        called = len(filtered_df[filtered_df["Status"].isin(["Called", "Quoted", "Sold", "In Progress"])])
        quoted = len(filtered_df[filtered_df["Status"].isin(["Quoted", "Sold"])])
        sold = len(filtered_df[filtered_df["Status"] == "Sold"])
    else:
        total_tracked = called = quoted = sold = 0

    with col1:
        st.markdown(f'<div class="metric-card"><div class="label">Total Tracked</div><div class="value">{total_tracked}</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><div class="label">Calls Made</div><div class="value">{called}</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><div class="label">Quotes Sent</div><div class="value">{quoted}</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="metric-card"><div class="label">PMs Sold</div><div class="value">{sold}</div></div>', unsafe_allow_html=True)

    st.markdown("---")

    # Region breakdown
    st.markdown("### Performance by Region")
    for region_name, branch_ids in REGIONS.items():
        region_branches = [BRANCHES[bid] for bid in branch_ids if bid in BRANCHES]

        if not filtered_df.empty and "Branch" in filtered_df.columns:
            region_df = filtered_df[filtered_df["Branch"].isin(region_branches)]
            r_total = len(region_df)
            r_called = len(region_df[region_df["Status"].isin(["Called", "Quoted", "Sold", "In Progress"])]) if "Status" in region_df.columns else 0
            r_quoted = len(region_df[region_df["Status"].isin(["Quoted", "Sold"])]) if "Status" in region_df.columns else 0
            r_sold = len(region_df[region_df["Status"] == "Sold"]) if "Status" in region_df.columns else 0
        else:
            region_df = pd.DataFrame()
            r_total = r_called = r_quoted = r_sold = 0

        st.markdown(f'<div class="region-header">{region_name} &mdash; {r_called} Calls / {r_quoted} Quotes / {r_sold} Sold</div>', unsafe_allow_html=True)

        branch_rows = []
        for bid in branch_ids:
            bname = BRANCHES.get(bid, "")
            if not filtered_df.empty and "Branch" in filtered_df.columns:
                bdf = filtered_df[filtered_df["Branch"] == bname]
                b_called = len(bdf[bdf["Status"].isin(["Called", "Quoted", "Sold", "In Progress"])]) if "Status" in bdf.columns else 0
                b_quoted = len(bdf[bdf["Status"].isin(["Quoted", "Sold"])]) if "Status" in bdf.columns else 0
                b_sold = len(bdf[bdf["Status"] == "Sold"]) if "Status" in bdf.columns else 0
            else:
                b_called = b_quoted = b_sold = 0
            branch_rows.append({"Branch": bname, "Calls": b_called, "Quotes": b_quoted, "Sold": b_sold})

        branch_perf_df = pd.DataFrame(branch_rows)
        st.dataframe(branch_perf_df, use_container_width=True, hide_index=True)

    st.markdown("---")

    # Monthly breakdown
    st.markdown("### Activity by Month")
    if not tracking_df.empty and "Month" in tracking_df.columns and "Status" in tracking_df.columns:
        month_stats = tracking_df.groupby("Month").agg(
            Calls=("Status", lambda x: x.isin(["Called", "Quoted", "Sold", "In Progress"]).sum()),
            Quotes=("Status", lambda x: x.isin(["Quoted", "Sold"]).sum()),
            Sold=("Status", lambda x: (x == "Sold").sum()),
        ).reset_index().sort_values("Sold", ascending=False)
        st.dataframe(month_stats, use_container_width=True, hide_index=True)
    else:
        st.info("No tracking data yet. Activity will appear here once branches start logging.")

    # Quote history from main sheet
    st.markdown("---")
    st.markdown("### Recent Quotes (All Branches)")
    quotes_df = load_quotes_from_sheet()
    if not quotes_df.empty:
        st.dataframe(quotes_df.tail(20), use_container_width=True, hide_index=True)
    else:
        st.info("No quotes saved yet.")


# ═══════════════════════════════════════════════════════════
# TRACKING HELPERS (save rep activity to Sheets)
# ═══════════════════════════════════════════════════════════
def get_tracking_sheet():
    """Get or create the Tracking worksheet."""
    sheet = get_gsheet_connection()
    if sheet is None:
        return None
    try:
        ws_list = sheet.spreadsheet.worksheets()
        for ws in ws_list:
            if ws.title == "Tracking":
                return ws
        # Create it if missing
        tracking_ws = sheet.spreadsheet.add_worksheet(title="Tracking", rows=5000, cols=10)
        tracking_ws.append_row(["Date", "Month", "Branch", "Customer", "Status", "Notes", "PM Value"], value_input_option="USER_ENTERED")
        return tracking_ws
    except Exception:
        return None

def save_tracking_entry(customer_name, status, notes="", pm_value=0):
    """Save a lead tracking entry."""
    ws = get_tracking_sheet()
    if ws is None:
        return False
    try:
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            st.session_state.login_month or "",
            st.session_state.branch_name or "",
            customer_name,
            status,
            notes,
            pm_value,
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════
# PM TRACKER (dedicated deal lifecycle tracking)
# ═══════════════════════════════════════════════════════════
PM_TRACKER_HEADERS = [
    "Date", "Customer", "Branch", "Rep", "Make", "Model", "Serial",
    "Eng Hours at Deal", "PM Interval (hrs)", "Contract Value",
    "Status", "Notes", "HubSpot Deal ID", "Next PM Due (hrs)",
    "Last Contact Date", "Hours Updated",
]

def get_pm_tracker_sheet():
    """Get or create the PM Tracker worksheet (separate from Tracking)."""
    sheet = get_gsheet_connection()
    if sheet is None:
        return None
    try:
        ws_list = sheet.spreadsheet.worksheets()
        for ws in ws_list:
            if ws.title == "PM Tracker":
                return ws
        pm_ws = sheet.spreadsheet.add_worksheet(title="PM Tracker", rows=5000, cols=len(PM_TRACKER_HEADERS))
        pm_ws.append_row(PM_TRACKER_HEADERS, value_input_option="USER_ENTERED")
        return pm_ws
    except Exception:
        return None

def save_pm_tracker_entry(data):
    """Save a PM deal lifecycle entry to the PM Tracker tab."""
    ws = get_pm_tracker_sheet()
    if ws is None:
        return False
    try:
        # Determine the first PM interval for this model (smallest interval hours)
        pm_interval = 500  # default
        ds = PM_DEALSHEET.get(data.get("model", ""))
        if ds:
            intervals = [ds.get(f"hr_{k}") for k in ["1", "2", "3"] if ds.get(f"hr_{k}")]
            if intervals:
                pm_interval = min(intervals)
        eng_hours = int(data.get("eng_hours", 0))
        # Calculate next PM due: next multiple of the interval above current hours
        if eng_hours > 0 and pm_interval > 0:
            next_pm = ((eng_hours // pm_interval) + 1) * pm_interval
        else:
            next_pm = pm_interval
        row = [
            data.get("date", datetime.now().strftime("%m/%d/%Y")),
            data.get("customer", ""),
            data.get("branch", ""),
            data.get("rep", ""),
            data.get("make", ""),
            data.get("model", ""),
            data.get("serial", ""),
            eng_hours,
            pm_interval,
            data.get("contract_value", 0),
            data.get("status", "Quoted"),
            data.get("notes", ""),
            data.get("hs_deal_id", ""),
            next_pm,
            datetime.now().strftime("%m/%d/%Y"),
            datetime.now().strftime("%m/%d/%Y"),
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception:
        return False

def load_pm_tracker():
    """Load all PM Tracker entries as a DataFrame."""
    ws = get_pm_tracker_sheet()
    if ws is None:
        return pd.DataFrame()
    try:
        data = ws.get_all_records()
        return pd.DataFrame(data) if data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()


# ═══════════════════════════════════════════════════════════
# HUBSPOT WRITE-BACK (PM deals only, separate from reads)
# ═══════════════════════════════════════════════════════════
def hubspot_create_or_update_pm_deal(deal_data):
    """Create or update a PM deal in HubSpot. Returns deal ID or None."""
    if not HUBSPOT_TOKEN:
        return None
    headers = {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}

    # Search for existing deal by customer + model + serial
    customer = deal_data.get("customer", "")
    model = deal_data.get("model", "")
    serial = deal_data.get("serial", "")
    deal_name = f"{customer} - PM {model}" + (f" ({serial})" if serial else "")

    try:
        # Try to find an existing PM deal for this customer+model
        search_payload = {
            "filterGroups": [{
                "filters": [
                    {"propertyName": "dealname", "operator": "CONTAINS_TOKEN", "value": customer},
                    {"propertyName": "pm_eligible", "operator": "EQ", "value": "true"},
                ]
            }],
            "properties": ["dealname", "pm_status", "pm_contract_value", "dealstage"],
            "limit": 10,
        }
        resp = requests.post(
            "https://api.hubapi.com/crm/v3/objects/deals/search",
            headers=headers, json=search_payload, timeout=15
        )
        existing_deal_id = None
        if resp.status_code == 200:
            for d in resp.json().get("results", []):
                dn = (d.get("properties", {}).get("dealname") or "").upper()
                if model.upper() in dn and customer.upper() in dn:
                    existing_deal_id = d["id"]
                    break

        # Build properties
        status = deal_data.get("status", "Quoted")
        stage_map = {
            "Quoted": "presentationscheduled",
            "Called": "qualifiedtobuy",
            "In Progress": "contractsent",
            "Sold": "closedwon",
            "Not Interested": "closedlost",
        }
        properties = {
            "dealname": deal_name,
            "pm_eligible": "true",
            "pm_status": status.lower(),
            "amount": str(deal_data.get("contract_value", 0)),
            "dealstage": stage_map.get(status, "presentationscheduled"),
        }
        # Add PM-specific custom properties if they exist in HubSpot
        pm_props = {
            "pm_contract_value": str(deal_data.get("contract_value", 0)),
            "pm_machine_model": model,
            "pm_machine_serial": serial,
            "pm_engine_hours": str(deal_data.get("eng_hours", 0)),
            "pm_interval_hours": str(deal_data.get("pm_interval", 500)),
            "pm_next_service_hours": str(deal_data.get("next_pm_hours", 0)),
        }
        # Only include PM props that HubSpot accepts (skip if they cause errors)
        properties.update(pm_props)

        if existing_deal_id:
            # Update existing deal
            resp = requests.patch(
                f"https://api.hubapi.com/crm/v3/objects/deals/{existing_deal_id}",
                headers=headers, json={"properties": properties}, timeout=15
            )
            if resp.status_code == 200:
                return existing_deal_id
            # If custom props failed, retry with just standard props
            std_props = {k: v for k, v in properties.items() if not k.startswith("pm_") or k in ("pm_eligible", "pm_status")}
            resp = requests.patch(
                f"https://api.hubapi.com/crm/v3/objects/deals/{existing_deal_id}",
                headers=headers, json={"properties": std_props}, timeout=15
            )
            return existing_deal_id if resp.status_code == 200 else None
        else:
            # Create new deal
            resp = requests.post(
                "https://api.hubapi.com/crm/v3/objects/deals",
                headers=headers, json={"properties": properties}, timeout=15
            )
            if resp.status_code == 201:
                new_id = resp.json().get("id")
                # Try to associate with matching company
                _associate_deal_to_company(new_id, customer, headers)
                return new_id
            # Retry with just standard props if custom ones failed
            std_props = {k: v for k, v in properties.items() if not k.startswith("pm_") or k in ("pm_eligible", "pm_status")}
            resp = requests.post(
                "https://api.hubapi.com/crm/v3/objects/deals",
                headers=headers, json={"properties": std_props}, timeout=15
            )
            if resp.status_code == 201:
                new_id = resp.json().get("id")
                _associate_deal_to_company(new_id, customer, headers)
                return new_id
            return None
    except Exception:
        return None

def _associate_deal_to_company(deal_id, customer_name, headers):
    """Associate a deal with a matching HubSpot company."""
    try:
        # Search for the company
        search = {
            "filterGroups": [{"filters": [{"propertyName": "name", "operator": "CONTAINS_TOKEN", "value": customer_name}]}],
            "properties": ["name"],
            "limit": 1,
        }
        resp = requests.post(
            "https://api.hubapi.com/crm/v3/objects/companies/search",
            headers=headers, json=search, timeout=10
        )
        if resp.status_code == 200:
            results = resp.json().get("results", [])
            if results:
                company_id = results[0]["id"]
                requests.put(
                    f"https://api.hubapi.com/crm/v3/objects/deals/{deal_id}/associations/companies/{company_id}/deal_to_company",
                    headers=headers, timeout=10
                )
    except Exception:
        pass

def hubspot_update_pm_alert(deal_id, alert_type, message):
    """Update a HubSpot deal with an alert flag for workflow triggers."""
    if not HUBSPOT_TOKEN or not deal_id:
        return False
    headers = {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}
    try:
        properties = {
            "pm_alert_type": alert_type,  # "hours_approaching" or "followup_overdue"
            "pm_alert_message": message,
            "pm_alert_date": datetime.now().strftime("%Y-%m-%d"),
        }
        resp = requests.patch(
            f"https://api.hubapi.com/crm/v3/objects/deals/{deal_id}",
            headers=headers, json={"properties": properties}, timeout=15
        )
        # If custom alert props don't exist yet, that's okay, it'll fail silently
        return resp.status_code == 200
    except Exception:
        return False


# ═══════════════════════════════════════════════════════════
# PM ALERT ENGINE (hours threshold + follow-up tracking)
# ═══════════════════════════════════════════════════════════
def check_pm_alerts(pm_tracker_df, current_fleet_df=None):
    """
    Compare PM Tracker entries against current fleet data to find:
    1. Machines approaching next PM interval (10% buffer)
    2. Customers with no contact in 30+ days
    Returns list of alert dicts.
    """
    alerts = []
    if pm_tracker_df.empty:
        return alerts

    today = datetime.now()

    for _, row in pm_tracker_df.iterrows():
        customer = str(row.get("Customer", ""))
        model = str(row.get("Model", ""))
        serial = str(row.get("Serial", ""))
        status = str(row.get("Status", "")).lower()
        next_pm = int(row.get("Next PM Due (hrs)", 0) or 0)
        pm_interval = int(row.get("PM Interval (hrs)", 500) or 500)
        hs_deal_id = str(row.get("HubSpot Deal ID", ""))
        last_contact = str(row.get("Last Contact Date", ""))

        # Skip closed/lost deals
        if status in ("not interested", "closed", "lost"):
            continue

        # 1. Check hours threshold (if we have current fleet data)
        if current_fleet_df is not None and not current_fleet_df.empty and next_pm > 0:
            # Find this machine in fleet data by model match
            fleet_match = None
            if serial:
                fleet_match = current_fleet_df[current_fleet_df["VIN"].astype(str).str.upper() == serial.upper()]
            if fleet_match is None or fleet_match.empty:
                fleet_match = current_fleet_df[
                    (current_fleet_df["Customer"].astype(str).str.upper() == customer.upper()) &
                    (current_fleet_df["Model"].astype(str).str.contains(model.split()[0] if model else "ZZZZZ", case=False, na=False))
                ]
            if fleet_match is not None and not fleet_match.empty:
                current_hours = int(fleet_match.iloc[0].get("Eng Hrs", 0) or 0)
                buffer = max(int(pm_interval * 0.10), 25)  # 10% of interval, minimum 25 hrs
                hours_remaining = next_pm - current_hours

                if 0 < hours_remaining <= buffer:
                    alerts.append({
                        "type": "hours_approaching",
                        "severity": "high" if hours_remaining <= buffer // 2 else "medium",
                        "customer": customer,
                        "model": model,
                        "serial": serial,
                        "current_hours": current_hours,
                        "next_pm": next_pm,
                        "hours_remaining": hours_remaining,
                        "hs_deal_id": hs_deal_id,
                        "message": f"{model} at {current_hours:,} hrs, next PM at {next_pm:,} hrs ({hours_remaining} hrs away)",
                    })
                elif current_hours >= next_pm:
                    alerts.append({
                        "type": "hours_overdue",
                        "severity": "critical",
                        "customer": customer,
                        "model": model,
                        "serial": serial,
                        "current_hours": current_hours,
                        "next_pm": next_pm,
                        "hours_remaining": hours_remaining,
                        "hs_deal_id": hs_deal_id,
                        "message": f"{model} at {current_hours:,} hrs, OVERDUE for PM at {next_pm:,} hrs",
                    })

        # 2. Check follow-up staleness
        if last_contact and status not in ("sold",):
            try:
                from dateutil.parser import parse as parse_date
                last_dt = parse_date(last_contact)
                days_since = (today - last_dt).days
                if days_since >= 30:
                    alerts.append({
                        "type": "followup_overdue",
                        "severity": "medium" if days_since < 60 else "high",
                        "customer": customer,
                        "model": model,
                        "serial": serial,
                        "days_since_contact": days_since,
                        "hs_deal_id": hs_deal_id,
                        "message": f"No contact in {days_since} days for {model}",
                    })
            except Exception:
                pass

    return alerts

def push_alerts_to_hubspot(alerts):
    """Push alert flags to HubSpot deals for workflow triggers."""
    pushed = 0
    for alert in alerts:
        deal_id = alert.get("hs_deal_id", "")
        if deal_id and deal_id != "nan":
            success = hubspot_update_pm_alert(deal_id, alert["type"], alert["message"])
            if success:
                pushed += 1
    return pushed


# ═══════════════════════════════════════════════════════════
# PAGE ROUTING
# ═══════════════════════════════════════════════════════════
if st.session_state.page == "login":
    show_login()
    st.stop()
elif st.session_state.page == "admin_login":
    show_admin_login()
    st.stop()
elif st.session_state.page == "admin":
    show_admin_dashboard()
    st.stop()

# ═══════════════════════════════════════════════════════════
# MAIN DASHBOARD (logged-in branch reps)
# ═══════════════════════════════════════════════════════════
# ─── Header ───
st.markdown(f"""
<div class="header-bar">
    <span class="branch-name">{st.session_state.branch_name} &mdash; PM Tool</span>
    <span class="rep-info">{st.session_state.login_month or ""}</span>
</div>
""", unsafe_allow_html=True)

col_head_l, col_head_r = st.columns([6, 1])
with col_head_r:
    if st.button("Logout", key="main_logout"):
        st.session_state.page = "login"
        st.session_state.branch = None
        st.session_state.branch_name = None
        st.session_state.login_month = None
        st.rerun()

# ═══════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════
tab_leads, tab_history = st.tabs(["Lead Discovery", "Quote History"])

# ═══════════════════════════════════════════════════════════
# TAB 1: LEAD DISCOVERY
# ═══════════════════════════════════════════════════════════
with tab_leads:
    st.subheader("PM Lead Discovery")
    st.caption("Scores customers with Case, Kobelco, Develon, and Bomag machines from CASE alerts and HubSpot CRM. Only customers tied to our 4 dealsheet brands are shown.")

    # Auto-load bundled data
    alerts_df = load_bundled_alerts()
    procare_vins = load_bundled_procare()

    # Score CASE alert leads
    scored = pd.DataFrame()
    if alerts_df is not None and not alerts_df.empty:
        scored = score_leads(alerts_df, procare_vins)

    # HubSpot enrichment + filtered HubSpot-only leads (4 dealsheet brands only)
    hs_companies = fetch_hubspot_companies()
    deal_history, pm_active = {}, set()
    hs_only_leads = pd.DataFrame()

    if hs_companies:
        deal_history, pm_active = fetch_hubspot_deals()

        # Enrich CASE leads with HubSpot data (adds spend, deal history, warranty info)
        if not scored.empty:
            scored = enrich_leads_with_hubspot(scored, hs_companies, deal_history, pm_active)

        # Build HubSpot-only leads, filtered to 4 dealsheet brands
        existing_customers = set(scored["Customer"].str.strip().str.upper()) if not scored.empty else set()
        hs_only_leads = build_hubspot_only_leads(hs_companies, deal_history, pm_active, existing_customers)

        # Filter HubSpot-only leads to customers tied to our 4 brands
        if not hs_only_leads.empty:
            target_brands = {"case", "kobelco", "develon", "bomag"}
            # Model prefixes that are specific enough to not false-match
            brand_keywords = [
                "kobelco", "develon", "bomag", "doosan",
                "case ce", "case construction",
                # CASE model prefixes (with enough context to avoid false matches)
                "cx80", "cx130", "cx160", "cx210", "cx220", "cx240", "cx250",
                "cx300", "cx350", "cx370", "cx490", "cx500", "cx750", "cx800",
                "dx55", "dx60", "dx140", "dx225", "dx235", "dx255", "dx300", "dx350", "dx380", "dx420", "dx490", "dx530",
                "sv185", "sv208", "sv215", "sv250", "sv280", "sv340",
                "580ev", "580n", "580sn", "590sn",
                "321f", "521g", "621g", "721g", "821g", "921g",
                "650m", "750m", "850m",
                "tv450", "tr270", "tr310", "tr320", "tr340",
                # Kobelco models
                "sk55", "sk75", "sk140", "sk170", "sk210", "sk230", "sk260",
                "sk300", "sk350", "sk390", "sk490", "sk500", "sk850",
                # Bomag models
                "bw120", "bw145", "bw177", "bw190", "bw206", "bw211", "bw213",
                "bw219", "bw226", "bw900",
                "bf200", "bf300", "bf600", "bf700", "bf800",
            ]
            def _is_target_brand_customer(hs_name):
                """Check if this HubSpot customer is associated with our 4 dealsheet brands."""
                data = hs_companies.get(hs_name.upper(), {})
                # Check CASE classification (if set, they're tagged for CASE/Kobelco/etc)
                case_class = (data.get("case_class", "") or "").lower()
                if case_class and case_class not in ("", "competitor", "competitive"):
                    return True
                # Check prospect classification for any of the 4 brands
                prospect_class = (data.get("prospect_class", "") or "").lower()
                if prospect_class:
                    for brand in ("case", "kobelco", "develon", "bomag", "doosan"):
                        if brand in prospect_class:
                            return True
                # Won deals = bought equipment from SEC (SEC only sells these 4 brands)
                deal_info = deal_history.get(hs_name.upper(), {})
                if deal_info.get("won", 0) > 0:
                    return True
                # Check deal names for brand/model keywords (e.g. "CX350 - SMITH CONST")
                deal_names = (deal_info.get("deal_names", "") or "").lower()
                desc = (data.get("description", "") or "").lower()
                combined_text = f"{deal_names} {desc}"
                for kw in brand_keywords:
                    if kw in combined_text:
                        return True
                return False

            mask = hs_only_leads["Customer"].apply(lambda c: _is_target_brand_customer(str(c).strip()))
            hs_only_leads = hs_only_leads[mask].copy()

    # Load equipment report leads (customers who bought our 4 brands but aren't in alerts/HubSpot)
    equip_df = load_equipment_report()
    equip_leads = pd.DataFrame()
    if not equip_df.empty:
        # Build set of customers already covered by CASE alerts and HubSpot
        existing_equip = set()
        if not scored.empty:
            existing_equip.update(scored["Customer"].str.strip().str.upper())
        if not hs_only_leads.empty:
            existing_equip.update(hs_only_leads["Customer"].str.strip().str.upper())
        equip_leads = build_equipment_report_leads(equip_df, existing_equip)

    # Merge all three sources
    sources = []
    if not scored.empty:
        if "Source" not in scored.columns:
            scored["Source"] = "CASE Alert"
        sources.append(scored)
    if not hs_only_leads.empty:
        sources.append(hs_only_leads)
    if not equip_leads.empty:
        sources.append(equip_leads)

    if sources:
        # Align columns across all sources
        all_cols = set()
        for s in sources:
            all_cols.update(s.columns)
        for s in sources:
            for col in all_cols:
                if col not in s.columns:
                    s[col] = None
        all_leads = pd.concat(sources, ignore_index=True)
    else:
        all_leads = pd.DataFrame()

    st.session_state.leads_df = all_leads
    st.session_state.procare_vins = procare_vins

    # Load PM Tracker and check for alerts
    pm_tracker = load_pm_tracker()
    pm_alerts = []
    if not pm_tracker.empty:
        fleet_for_alerts = all_leads if not all_leads.empty else None
        pm_alerts = check_pm_alerts(pm_tracker, fleet_for_alerts)
        # Build lookup: customer -> list of alerts
    pm_alerts_by_customer = {}
    for a in pm_alerts:
        cust = a.get("customer", "").upper()
        if cust:
            pm_alerts_by_customer.setdefault(cust, []).append(a)

    if all_leads.empty:
        st.warning("No leads found. Check that data files are loaded or HubSpot is connected.")
    else:
        # Summary metrics
        case_leads = all_leads[all_leads["Source"] == "CASE Alert"] if "Source" in all_leads.columns else all_leads
        hs_leads = all_leads[all_leads["Source"] == "HubSpot"] if "Source" in all_leads.columns else pd.DataFrame()

        n_alerts = len(pm_alerts)
        if n_alerts > 0:
            col1, col2, col3, col4 = st.columns(4)
        else:
            col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Leads", f"{all_leads['Customer'].nunique():,}")
        with col2:
            top_high = all_leads[all_leads["Tier"].isin(["Top", "High"])]["Customer"].nunique()
            st.metric("Top + High", top_high)
        with col3:
            # Only count YTD spend from HubSpot (accurate YTD), not equipment report lifetime spend
            if "YTD Parts" in all_leads.columns and "YTD Service" in all_leads.columns:
                ytd_by_cust = all_leads.groupby("Customer").agg({"YTD Parts": "first", "YTD Service": "first"})
                ytd_total = ytd_by_cust["YTD Parts"].fillna(0).sum() + ytd_by_cust["YTD Service"].fillna(0).sum()
                st.metric("Known Spend (YTD)", f"${ytd_total:,.0f}")
            elif "Total Spend" in all_leads.columns:
                total_spend = all_leads.groupby("Customer")["Total Spend"].first().sum()
                st.metric("Known Spend (YTD)", f"${total_spend:,.0f}")
            else:
                st.metric("Data Files", "Loaded")
        if n_alerts > 0:
            with col4:
                st.metric("Active Alerts", n_alerts)

        # Show alert banner if there are critical/high alerts
        critical_alerts = [a for a in pm_alerts if a.get("severity") in ("critical", "high")]
        if critical_alerts:
            alert_msgs = [f"**{a['customer']}**: {a['message']}" for a in critical_alerts[:5]]
            st.error(f"Action needed on {len(critical_alerts)} machine{'s' if len(critical_alerts) > 1 else ''}:  \n" + "  \n".join(alert_msgs))

        st.divider()

        # Filters
        col1, col2 = st.columns(2)
        with col1:
            filter_tier = st.multiselect("Tier", ["Top", "High", "Medium", "Low"], default=["Top", "High"])
        with col2:
            pm_filter = st.radio("PM Status", ["No Active PM", "All", "Has Active PM"], horizontal=True) if "Has PM" in all_leads.columns else "All"

        # Apply filters
        display = all_leads.copy()
        if filter_tier:
            display = display[display["Tier"].isin(filter_tier)]
        if pm_filter == "No Active PM" and "Has PM" in display.columns:
            display = display[~display["Has PM"]]
        elif pm_filter == "Has Active PM" and "Has PM" in display.columns:
            display = display[display["Has PM"]]

        # View toggle
        view = st.radio("View", ["By Customer", "By Machine/Lead"], horizontal=True)

        if view == "By Customer":
            cust_display = aggregate_customer_leads(display)
            if cust_display.empty:
                st.info("No customers match filters.")
            else:
                st.caption(f"Showing {len(cust_display)} customers")
                for _, row in cust_display.iterrows():
                    tier = row["Tier"]
                    cust_name = row["Customer"]
                    tier_class = {"Top": "top", "High": "high", "Medium": "med"}.get(tier, "low")
                    tier_css = {"Top": "tier-top", "High": "tier-high", "Medium": "tier-med"}.get(tier, "")

                    # Build machine list
                    machines_str = ""
                    cust_machines = display[display["Customer"] == cust_name]
                    if not cust_machines.empty and "Model" in cust_machines.columns:
                        models = sorted(set(str(m) for m in cust_machines["Model"] if m and str(m).strip()))
                        if models:
                            machines_str = ", ".join(models)
                    elif "fleet" in row and row.get("fleet"):
                        machines_str = f"Fleet: {row['fleet']}"

                    # Values for card
                    parts_opp = float(row.get("total_parts_value", 0) or 0)
                    pm_value = float(row.get("total_annual_pm", 0) or 0)
                    score = row["Customer Score"]
                    loc = row.get("location", "") or ""
                    cat_label = row.get("lead_category", "") or ""

                    # Tier accent color
                    accent = {"Top": "#C8102E", "High": "#F59E0B", "Medium": "#3B82F6"}.get(tier, "#9CA3AF")

                    # Machines pill HTML
                    machines_html = ""
                    if machines_str:
                        machine_list = [m.strip() for m in machines_str.replace("Fleet: ", "").split(",") if m.strip()]
                        pills = "".join(
                            f'<span style="display:inline-block;background:#F3F4F6;color:#374151;font-size:12px;padding:3px 10px;border-radius:12px;margin:2px 4px 2px 0;font-weight:500;">{m}</span>'
                            for m in machine_list[:8]
                        )
                        if len(machine_list) > 8:
                            pills += f'<span style="display:inline-block;color:#9CA3AF;font-size:12px;padding:3px 4px;">+{len(machine_list)-8} more</span>'
                        machines_html = f'<div style="margin:10px 0 0 0;">{pills}</div>'

                    # Value pills
                    value_pills = ""
                    if pm_value > 0:
                        value_pills += f'<div style="display:inline-block;margin-right:16px;"><span style="font-size:11px;color:#6B7280;text-transform:uppercase;letter-spacing:0.5px;">Est PM Value</span><br><span style="font-size:18px;font-weight:700;color:#C8102E;">${pm_value:,.0f}</span></div>'
                    if parts_opp > 0:
                        value_pills += f'<div style="display:inline-block;"><span style="font-size:11px;color:#6B7280;text-transform:uppercase;letter-spacing:0.5px;">Parts Opp</span><br><span style="font-size:18px;font-weight:700;color:#1A1A1A;">${parts_opp:,.0f}</span></div>'

                    # Location + category subtitle
                    subtitle_parts = []
                    if loc:
                        subtitle_parts.append(loc)
                    if cat_label:
                        subtitle_parts.append(cat_label)
                    subtitle = " · ".join(subtitle_parts)
                    subtitle_html = f'<div style="font-size:12px;color:#9CA3AF;margin-top:2px;">{subtitle}</div>' if subtitle else ""

                    # Build alert badges if this customer has active alerts
                    cust_alerts = pm_alerts_by_customer.get(cust_name.strip().upper(), [])
                    alert_html = ""
                    if cust_alerts:
                        alert_pills = []
                        for ca in cust_alerts[:3]:  # Show max 3 alerts
                            if ca["type"] == "hours_overdue":
                                a_color = "#DC2626"
                                a_icon = "OVERDUE"
                            elif ca["type"] == "hours_approaching":
                                a_color = "#F59E0B"
                                a_icon = "PM DUE SOON"
                            else:
                                a_color = "#6B7280"
                                a_icon = "FOLLOW UP"
                            a_msg = ca.get("message", "")
                            alert_pills.append(
                                f'<span style="display:inline-block;background:{a_color};color:white;font-size:10px;font-weight:700;'
                                f'padding:3px 8px;border-radius:4px;margin-right:4px;" title="{a_msg}">{a_icon}</span>'
                            )
                        alert_html = f'<div style="margin-top:6px;">{"".join(alert_pills)}</div>'

                    # Render card with all inline styles
                    card_html = f'''<div style="background:#FFFFFF;border:1px solid #E5E7EB;border-left:4px solid {accent};border-radius:10px;padding:18px 22px 14px 22px;margin-bottom:12px;">
                        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
                            <div>
                                <span style="font-size:16px;font-weight:700;color:#1A1A1A;">{cust_name}</span>
                                {subtitle_html}
                                {alert_html}
                            </div>
                            <div style="text-align:right;min-width:60px;">
                                <div style="font-size:26px;font-weight:700;color:#1A1A1A;line-height:1;">{score:.0f}</div>
                                <div style="font-size:11px;font-weight:600;color:{accent};margin-top:2px;">{tier}</div>
                            </div>
                        </div>
                        {machines_html}
                        <div style="margin-top:12px;">{value_pills}</div>
                    </div>'''
                    st.markdown(card_html, unsafe_allow_html=True)

                    cust_key = cust_name.replace(" ", "_")[:20]

                    # Two action buttons side by side
                    btn_col1, btn_col2, _ = st.columns([1, 1, 3])
                    with btn_col1:
                        quote_open = st.toggle("Quote", key=f"qt_{cust_key}", value=False)
                    with btn_col2:
                        log_open = st.toggle("Log Activity", key=f"la_{cust_key}", value=False)

                    # Inline PM Calculator
                    if quote_open:
                        with st.container():
                            st.markdown(f'<div style="background:#FAFBFC;border:1px solid #E5E7EB;border-radius:8px;padding:16px 20px;margin:4px 0 12px 0;">', unsafe_allow_html=True)
                            st.caption(f"PM Quote for {cust_name}")
                            qc1, qc2 = st.columns(2)
                            with qc1:
                                q_service = st.selectbox("Field or Shop", SERVICE_TYPES, key=f"qs_{cust_key}")
                            with qc2:
                                q_travel = st.number_input("Travel (min, one way)", min_value=0, max_value=480, value=0, step=15, key=f"qtr_{cust_key}")
                            qc1, qc2 = st.columns(2)
                            with qc1:
                                q_make = st.selectbox("Make", [""] + sorted(PM_BRANDS.keys()), key=f"qm_{cust_key}")
                            with qc2:
                                if q_make and q_make in PM_BRANDS:
                                    q_model = st.selectbox("Model", [""] + get_models_for_brand(q_make), key=f"qmd_{cust_key}")
                                else:
                                    q_model = ""
                            qc1, qc2, qc3 = st.columns(3)
                            with qc1:
                                q_serial = st.text_input("Serial #", key=f"qsr_{cust_key}")
                            with qc2:
                                q_hours = st.selectbox("Hours Requested", [500, 1000, 1500, 2000, 2500, 3000, 3500], index=3, key=f"qh_{cust_key}")
                            with qc3:
                                q_rep = st.text_input("Rep", value=st.session_state.get("rep_name", ""), key=f"qr_{cust_key}")
                            q_notes = st.text_input("Notes", key=f"qn_{cust_key}", placeholder="Machine condition, special requirements...")

                            can_calc = bool(q_make and q_model and q_model in PM_DEALSHEET)
                            if st.button("Calculate PM Price", type="primary", use_container_width=True, disabled=not can_calc, key=f"qcalc_{cust_key}"):
                                result = calculate_pm_cost(q_model, q_hours)
                                if result:
                                    travel_cost = round((q_travel / 60) * 225 * 2, 2) if q_service == "Field" and q_travel > 0 else 0
                                    st.session_state[f"quote_{cust_key}"] = {
                                        "date": datetime.now().strftime("%m/%d/%Y"),
                                        "customer_name": cust_name, "branch": st.session_state.get("branch_name", ""),
                                        "rep": q_rep, "service_type": q_service, "make": q_make,
                                        "model": q_model, "serial": q_serial,
                                        "hours_requested": q_hours,
                                        "travel_time": q_travel if q_service == "Field" else 0,
                                        "travel_cost": travel_cost, "notes": q_notes,
                                        "intervals": result["intervals"],
                                        "total_cost": result["total_cost"],
                                        "annual_pm_price": result["total_cost"] + travel_cost,
                                    }

                            # Show results if calculated
                            quote_key = f"quote_{cust_key}"
                            if quote_key in st.session_state and st.session_state[quote_key]:
                                q = st.session_state[quote_key]
                                st.divider()
                                if "intervals" in q and q["intervals"]:
                                    interval_rows = []
                                    for iv in q["intervals"]:
                                        interval_rows.append({
                                            "Service": iv["name"],
                                            "Hour Interval": f"{iv['hours']:,} hr",
                                            "Qty": iv["qty"],
                                            "Cost (Per)": f"${iv['cost_per']:,.0f}",
                                            "Subtotal": f"${iv['subtotal']:,.0f}",
                                        })
                                    st.dataframe(pd.DataFrame(interval_rows), use_container_width=True, hide_index=True)
                                t_cost = q.get("travel_cost", 0)
                                if t_cost > 0:
                                    st.markdown(f"**Travel:** ${t_cost:,.0f}")
                                st.markdown(f"""<div style="background:white;border:1px solid #E5E7EB;border-radius:8px;padding:14px;text-align:center;border-top:3px solid #C8102E;margin-top:8px;">
                                    <div style="font-size:12px;color:#6B7280;text-transform:uppercase;">Total PM Contract Price ({q.get('hours_requested',0):,} hrs)</div>
                                    <div style="font-size:24px;font-weight:700;color:#C8102E;margin-top:4px;">${q['annual_pm_price']:,.0f}</div>
                                </div>""", unsafe_allow_html=True)

                                rc1, rc2, rc3 = st.columns(3)
                                with rc1:
                                    pdf_buf = generate_pdf(q)
                                    safe = cust_name.replace(" ", "_").replace("/", "-")
                                    date_str = datetime.now().strftime("%m-%d-%Y")
                                    st.download_button("Download PDF", data=pdf_buf, file_name=f"SEC_PM_Quote_{safe}_{date_str}.pdf", mime="application/pdf", use_container_width=True, key=f"qpdf_{cust_key}")
                                with rc2:
                                    if st.button("Save Quote", use_container_width=True, type="secondary", key=f"qsave_{cust_key}"):
                                        saved = save_quote_to_sheet(q)
                                        if cust_name:
                                            save_tracking_entry(cust_name, "Quoted", f"{q.get('make','')} {q.get('model','')}", q.get("annual_pm_price", 0))
                                        # Write to PM Tracker
                                        pm_entry = {
                                            "customer": cust_name,
                                            "branch": st.session_state.get("branch_name", ""),
                                            "rep": q.get("rep", ""),
                                            "make": q.get("make", ""),
                                            "model": q.get("model", ""),
                                            "serial": q.get("serial", ""),
                                            "eng_hours": q.get("hours_requested", 0),
                                            "contract_value": q.get("annual_pm_price", 0),
                                            "status": "Quoted",
                                            "notes": q.get("notes", ""),
                                        }
                                        save_pm_tracker_entry(pm_entry)
                                        # Push to HubSpot
                                        hs_deal_id = hubspot_create_or_update_pm_deal(pm_entry)
                                        if hs_deal_id:
                                            # Update PM Tracker with HubSpot deal ID
                                            pm_entry["hs_deal_id"] = hs_deal_id
                                        st.success("Quote saved" if saved else "Saved locally (Sheets not connected)")
                                with rc3:
                                    if st.button("Clear Quote", use_container_width=True, key=f"qclr_{cust_key}"):
                                        st.session_state[quote_key] = {}
                                        st.rerun()

                            st.markdown('</div>', unsafe_allow_html=True)

                    # Log Activity
                    if log_open:
                        tc1, tc2, tc3 = st.columns([1, 1, 2])
                        with tc1:
                            track_status = st.selectbox("Status", ["Called", "Quoted", "In Progress", "Sold", "Not Interested"], key=f"ts_{cust_key}")
                        with tc2:
                            track_pm_val = st.number_input("PM Value ($)", min_value=0, value=int(pm_value), key=f"tv_{cust_key}")
                        with tc3:
                            track_notes = st.text_input("Notes", key=f"tn_{cust_key}")
                        if st.button("Save", key=f"tb_{cust_key}", type="primary"):
                            if save_tracking_entry(cust_name, track_status, track_notes, track_pm_val):
                                st.success(f"Logged: {cust_name} marked as {track_status}")
                            else:
                                st.warning("Could not save to Google Sheets. Check connection.")
                            # Also log to PM Tracker for deal lifecycle
                            # Get machine info from the lead data
                            lead_model = ""
                            lead_make = ""
                            cust_machines_log = display[display["Customer"] == cust_name] if not display.empty else pd.DataFrame()
                            if not cust_machines_log.empty:
                                lead_model = str(cust_machines_log.iloc[0].get("Dealsheet Model", "") or cust_machines_log.iloc[0].get("Model", "") or "")
                                lead_make = str(cust_machines_log.iloc[0].get("Make", "") or "")
                            pm_log = {
                                "customer": cust_name,
                                "branch": st.session_state.get("branch_name", ""),
                                "rep": st.session_state.get("rep_name", ""),
                                "make": lead_make,
                                "model": lead_model,
                                "contract_value": track_pm_val,
                                "status": track_status,
                                "notes": track_notes,
                            }
                            save_pm_tracker_entry(pm_log)
                            # Push status update to HubSpot
                            hs_id = hubspot_create_or_update_pm_deal(pm_log)
                            if hs_id:
                                pm_log["hs_deal_id"] = hs_id

        else:  # By Machine/Lead
            st.caption(f"Showing {len(display)} leads")
            show_cols = ["Lead Score", "Tier", "Source", "Customer", "Lead Category", "CASE Class", "Fleet", "Location"]
            # Add CASE-specific columns
            show_cols += ["Model", "Dealsheet Model", "Eng Hrs", "Stop", "Parts Value", "Labor Hrs", "Annual PM Value"]
            # Add spend columns if present
            if "Total Spend" in display.columns:
                show_cols.insert(show_cols.index("Parts Value"), "Total Spend")
            show_cols += ["Has PM", "VIN"]
            show_cols = [c for c in show_cols if c in display.columns]
            st.dataframe(
                display[show_cols].reset_index(drop=True),
                use_container_width=True, hide_index=True,
                column_config={
                    "Lead Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.0f"),
                    "Parts Value": st.column_config.NumberColumn("Parts $", format="$%.0f"),
                    "Total Spend": st.column_config.NumberColumn("YTD Spend", format="$%.0f"),
                    "Annual PM Value": st.column_config.NumberColumn("Annual PM", format="$%.0f"),
                    "Eng Hrs": st.column_config.NumberColumn("Hours", format="%.0f"),
                    "Labor Hrs": st.column_config.NumberColumn("Labor Hrs", format="%.1f"),
                    "Has PM": st.column_config.CheckboxColumn("Has PM", default=False),
                    "Lead Category": st.column_config.TextColumn("Lead Category", width="medium"),
                    "CASE Class": st.column_config.TextColumn("CASE Class", width="small"),
                    "Fleet": st.column_config.TextColumn("Fleet Size", width="small"),
                    "Source": st.column_config.TextColumn("Source", width="small"),
                },
            )

            # Export leads
            st.divider()
            col_exp1, col_exp2, _ = st.columns([1, 1, 3])
            with col_exp1:
                csv_data = display.to_csv(index=False).encode("utf-8")
                st.download_button("Export Leads (CSV)", data=csv_data, file_name=f"SEC_PM_Leads_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", use_container_width=True)
            with col_exp2:
                cust_csv = aggregate_customer_leads(display)
                if not cust_csv.empty:
                    cust_csv_data = cust_csv.to_csv(index=False).encode("utf-8")
                    st.download_button("Export by Customer (CSV)", data=cust_csv_data, file_name=f"SEC_PM_Leads_Customers_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", use_container_width=True)

            # Charts
            st.divider()
            import plotly.express as px

            # Lead Category breakdown
            if "Lead Category" in display.columns:
                cat_colors = {
                    "Parts Only, No Service": "#C8102E",
                    "Warranty Expiring": "#DC2626",
                    "Warranty Expired": "#E8601C",
                    "Active Service Customer": "#2F5496",
                    "Full Service (Lock In)": "#1B7340",
                    "Lapsed Service": "#F59E0B",
                    "No ProCare (In CRM)": "#6B7280",
                    "No ProCare (New Lead)": "#9CA3AF",
                    "Active PM (Upsell)": "#8B5CF6",
                    "Equipment Buyer (No Service)": "#D97706",
                    "In CRM (Prospect)": "#94A3B8",
                }
                lc_data = display.groupby("Lead Category").agg(
                    customers=("Customer", "nunique"),
                    pm_value=("Annual PM Value", "sum"),
                ).reset_index().sort_values("pm_value", ascending=True)
                fig = px.bar(lc_data, x="pm_value", y="Lead Category",
                             orientation="h", title="PM Opportunity by Lead Category",
                             labels={"pm_value": "Annual PM Value ($)", "Lead Category": ""},
                             hover_data=["customers"],
                             color="Lead Category", color_discrete_map=cat_colors)
                fig.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig, use_container_width=True)

            # Source breakdown
            if "Source" in display.columns:
                col1, col2 = st.columns(2)
                with col1:
                    src_data = display.groupby("Source").agg(
                        customers=("Customer", "nunique"),
                        pm_value=("Annual PM Value", "sum"),
                    ).reset_index()
                    fig = px.bar(src_data, x="Source", y="customers",
                                 title="Leads by Source",
                                 labels={"customers": "Unique Customers", "Source": ""},
                                 color="Source", color_discrete_map={"CASE Alert": SEC_RED, "HubSpot": "#2F5496"})
                    fig.update_layout(showlegend=False, height=350)
                    st.plotly_chart(fig, use_container_width=True)

                with col2:
                    fig = px.bar(src_data, x="Source", y="pm_value",
                                 title="PM Opportunity by Source",
                                 labels={"pm_value": "Annual PM Value ($)", "Source": ""},
                                 color="Source", color_discrete_map={"CASE Alert": SEC_RED, "HubSpot": "#2F5496"})
                    fig.update_layout(showlegend=False, height=350)
                    st.plotly_chart(fig, use_container_width=True)

            # Location and Category charts (filter out blanks from HubSpot leads)
            col1, col2 = st.columns(2)
            with col1:
                loc_filtered = display[display["Location"].notna() & (display["Location"] != "")]
                if not loc_filtered.empty:
                    loc_data = loc_filtered.groupby("Location").agg(
                        customers=("Customer", "nunique"),
                        pm_value=("Annual PM Value", "sum"),
                    ).reset_index()
                    fig = px.bar(loc_data.sort_values("pm_value", ascending=True), x="pm_value", y="Location",
                                 orientation="h", title="PM Opportunity by Location",
                                 labels={"pm_value": "Annual PM Value ($)", "Location": ""},
                                 color_discrete_sequence=[SEC_RED])
                    fig.update_layout(showlegend=False, height=400)
                    st.plotly_chart(fig, use_container_width=True)

            with col2:
                ds_filtered = display[display["Dealsheet Model"].notna() & (display["Dealsheet Model"] != "")]
                if not ds_filtered.empty:
                    ds_data = ds_filtered.groupby("Dealsheet Model").agg(
                        customers=("Customer", "nunique"),
                        pm_value=("Annual PM Value", "sum"),
                    ).reset_index().sort_values("pm_value", ascending=True).tail(15)
                    fig = px.bar(ds_data, x="pm_value", y="Dealsheet Model",
                                 orientation="h", title="PM Opportunity by Model",
                                 labels={"pm_value": "Annual PM Value ($)", "Dealsheet Model": ""},
                                 color_discrete_sequence=["#2F5496"])
                    fig.update_layout(showlegend=False, height=400)
                    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
# TAB 2: QUOTE HISTORY
# ═══════════════════════════════════════════════════════════
with tab_history:
    st.subheader("Quote History & Tracking")
    df = load_quotes_from_sheet()
    if df.empty:
        st.info("No quotes saved yet. Use the calculator to create and save quotes.")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            fb = st.multiselect("Branch", sorted(df["branch"].unique()) if "branch" in df.columns else [])
        with col2:
            fr = st.multiselect("Rep", sorted(df["rep"].unique()) if "rep" in df.columns else [])
        with col3:
            fm = st.multiselect("Make", sorted(df["make"].unique()) if "make" in df.columns else [])

        filt = df.copy()
        if fb: filt = filt[filt["branch"].isin(fb)]
        if fr: filt = filt[filt["rep"].isin(fr)]
        if fm: filt = filt[filt["make"].isin(fm)]

        if not filt.empty and "annual_pm_price" in filt.columns:
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Quotes", len(filt))
            with c2: st.metric("Total Value", f"${filt['annual_pm_price'].sum():,.0f}")
            with c3: st.metric("Avg Value", f"${filt['annual_pm_price'].mean():,.0f}")

        st.dataframe(filt, use_container_width=True, hide_index=True)

        if len(filt) > 1:
            import plotly.express as px
            col1, col2 = st.columns(2)
            with col1:
                if "branch" in filt.columns:
                    bd = filt.groupby("branch")["annual_pm_price"].sum().reset_index()
                    fig = px.bar(bd, x="branch", y="annual_pm_price", title="By Branch", color_discrete_sequence=[SEC_RED])
                    fig.update_layout(showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            with col2:
                if "make" in filt.columns:
                    md = filt.groupby("make")["annual_pm_price"].sum().reset_index()
                    fig = px.bar(md, x="make", y="annual_pm_price", title="By Make", color_discrete_sequence=["#2F5496"])
                    fig.update_layout(showlegend=False, xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
# TAB 4: PRICING REFERENCE
# ═══════════════════════════════════════════════════════════

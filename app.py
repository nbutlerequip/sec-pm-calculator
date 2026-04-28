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
st.set_page_config(page_title="SEC PM Tool", page_icon="🔧", layout="wide", initial_sidebar_state="expanded")

# ─── Brand ───
SEC_RED = "#C8102E"
SEC_DARK = "#1A1A1A"
SEC_GRAY = "#4A4A4A"

# ─── Branches ───
BRANCHES = [
    "Brunswick", "Burlington", "Cambridge", "Columbus", "Dayton", "Dublin",
    "Gallipolis", "Heath", "Hebron", "Holt", "Indianapolis", "Louisville",
    "Mansfield", "Marietta", "Mentor", "Monroe", "Nashville",
    "North Canton", "Novi", "Painesville", "Perrysburg", "Richfield",
    "Tri-Cities", "Wooster",
]

# ─── Makes & Models ───
# MAKES_MODELS used only for Lead Discovery tab (alert file parsing / category mapping).
# The PM Calculator tab uses DEALSHEET_MAKES directly from the dealsheet data above.
MAKES_MODELS = {
    "Case": {
        "Compact Track Loader": ["TR270B", "TR310B", "TR340B", "TV370B", "TV450B", "TV620B", "DL550",
                                 "SL12", "SL12 TR", "SL15", "SL23", "SL27", "SL27 TR", "SL35 TR", "SL50 TR"],
        "Skid Steer": ["SR175B", "SR210B", "SR240B", "SR270B", "SV185B", "SV270B", "SV280B", "SV340B"],
        "Excavator (Mini)": ["CX12D", "CX19D Cab", "CX30D Cab", "CX34D Cab", "CX38D Cab", "CX42D Cab", "CX50D Cab", "CX60D"],
        "Excavator (Standard)": ["CX70E", "CX85E", "CX90E"],
        "Backhoe": ["575N", "580SV", "580SN 4WD", "590SN 4WD", "586H 4WD", "588H 4WD"],
        "Wheel Loader": ["21F", "221F", "321F", "421F"],
    },
    "Kobelco": {
        "Excavator (Mini)": ["SK17SR-6E", "SK26SR-7", "SK35SR-7", "SK45SRX-7", "SK55SRX-7"],
        "Excavator (Standard)": ["SK75SR-7", "SK85CS-7", "SK130LC-11", "SK140SR-7", "ED160-7",
                                  "SK170LC-11", "SK210LC-11", "SK230SR-7", "SK260LC-11", "SK270SR-7",
                                  "SK300LC-11", "SK350LC-11", "SK380SRLC-7", "SK520LC-11"],
    },
    "Develon": {
        "Excavator (Mini)": ["DX35Z-7", "DX50Z-7", "DX63-7", "DX89R-7"],
        "Excavator (Standard)": ["DX140LC-7", "DX140LCR-7", "DX170LC-5"],
        "Articulated Dump Truck": ["DA45"],
        "Dozer": ["DD100", "DD130"],
        "Wheel Loader": ["DL220-7", "DL250-7"],
    },
    "Bomag": {
        "Roller/Compaction": ["BMP8500", "BW11RH", "BW120AD/SL", "BW138AD", "BW141/151/161",
                              "BW177", "BW190", "BW211", "BW900", "BW90AD"],
        "Paver": ["BF200C", "BF300C"],
    },
}

# ─── PM Pricing (will be replaced when Jarred provides averages) ───
# ─── PM Dealsheet Rev 3.0 — per-model service interval pricing (verified by Jarred & service team) ───
# Keys: (brand, model) → {cost_i, cost_1, cost_2, cost_3, cost_s, hr_i, hr_1, hr_2, hr_3, hr_s}
DEALSHEET_MODELS = {
    # ── Case Excavators (Mini / Standard) ──
    ("Case", "CX12D"):            {"cost_i": 518, "cost_1": 0,   "cost_2": 1207, "cost_3": 2166, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX19D Cab"):        {"cost_i": 518, "cost_1": 0,   "cost_2": 1207, "cost_3": 2166, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX30D Cab"):        {"cost_i": 760, "cost_1": 485, "cost_2": 1175, "cost_3": 2040, "cost_s": 0,    "hr_i": 100, "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX34D Cab"):        {"cost_i": 760, "cost_1": 485, "cost_2": 1175, "cost_3": 2040, "cost_s": 0,    "hr_i": 100, "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX38D Cab"):        {"cost_i": 825, "cost_1": 550, "cost_2": 1235, "cost_3": 2105, "cost_s": 0,    "hr_i": 100, "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX42D Cab"):        {"cost_i": 890, "cost_1": 0,   "cost_2": 1360, "cost_3": 2280, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX50D Cab"):        {"cost_i": 820, "cost_1": 0,   "cost_2": 1395, "cost_3": 2800, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX60D"):            {"cost_i": 855, "cost_1": 0,   "cost_2": 1420, "cost_3": 2815, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX70E"):            {"cost_i": 825, "cost_1": 0,   "cost_2": 1175, "cost_3": 2580, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX85E"):            {"cost_i": 825, "cost_1": 0,   "cost_2": 1175, "cost_3": 2580, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "CX90E"):            {"cost_i": 825, "cost_1": 0,   "cost_2": 1175, "cost_3": 2580, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    # ── Case Compact Track Loaders ──
    ("Case", "TR270B"):            {"cost_i": 215, "cost_1": 0,   "cost_2": 1025, "cost_3": 1550, "cost_s": 2595, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "TR310B"):            {"cost_i": 215, "cost_1": 0,   "cost_2": 1025, "cost_3": 1550, "cost_s": 2575, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "TR340B"):            {"cost_i": 215, "cost_1": 0,   "cost_2": 1045, "cost_3": 1650, "cost_s": 2694, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "TV370B"):            {"cost_i": 145, "cost_1": 0,   "cost_2": 946,  "cost_3": 1540, "cost_s": 2590, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "TV450B"):            {"cost_i": 105, "cost_1": 0,   "cost_2": 982,  "cost_3": 2105, "cost_s": 0,    "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "TV620B"):            {"cost_i": 175, "cost_1": 0,   "cost_2": 1625, "cost_3": 2640, "cost_s": 2860, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "DL550"):             {"cost_i": 175, "cost_1": 0,   "cost_2": 1905, "cost_3": 2815, "cost_s": 4000, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    # ── Case Skid Steers ──
    ("Case", "SR175B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 885,  "cost_3": 2640, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SR210B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 895,  "cost_3": 1550, "cost_s": 2545, "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "SR240B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 870,  "cost_3": 1685, "cost_s": 2680, "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "SR270B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 870,  "cost_3": 1770, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SV270B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 870,  "cost_3": 1770, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SV185B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 700,  "cost_3": 2480, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SV280B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 870,  "cost_3": 1685, "cost_s": 2680, "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "SV340B"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 960,  "cost_3": 1890, "cost_s": 2880, "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    # ── Case Backhoes ──
    ("Case", "575N"):              {"cost_i": 1005,"cost_1": 0,   "cost_2": 585,  "cost_3": 2680, "cost_s": 2780, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "580SV"):             {"cost_i": 0,   "cost_1": 0,   "cost_2": 570,  "cost_3": 2495, "cost_s": 2550, "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "580SN 4WD"):         {"cost_i": 670, "cost_1": 0,   "cost_2": 905,  "cost_3": 3280, "cost_s": 3410, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "590SN 4WD"):         {"cost_i": 670, "cost_1": 0,   "cost_2": 905,  "cost_3": 3280, "cost_s": 3410, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "586H 4WD"):          {"cost_i": 590, "cost_1": 0,   "cost_2": 580,  "cost_3": 2630, "cost_s": 2795, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    ("Case", "588H 4WD"):          {"cost_i": 590, "cost_1": 0,   "cost_2": 580,  "cost_3": 2630, "cost_s": 2795, "hr_i": 100, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    # ── Case SL Series (Compact Track Loader) ──
    ("Case", "SL12"):              {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL12 TR"):           {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL15"):              {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL23"):              {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL27"):              {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL27 TR"):           {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL35 TR"):           {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Case", "SL50 TR"):           {"cost_i": 950, "cost_1": 575, "cost_2": 2835, "cost_3": 6750, "cost_s": 0,    "hr_i": 50,  "hr_1": 250, "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    # ── Case Wheel Loaders ──
    ("Case", "21F"):               {"cost_i": 470, "cost_1": 0,   "cost_2": 566,  "cost_3": 2090, "cost_s": 2437, "hr_i": 150, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 2000},
    ("Case", "221F"):              {"cost_i": 638, "cost_1": 0,   "cost_2": 1062, "cost_3": 2620, "cost_s": 2970, "hr_i": 150, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 2000},
    ("Case", "321F"):              {"cost_i": 638, "cost_1": 0,   "cost_2": 1062, "cost_3": 2620, "cost_s": 2970, "hr_i": 150, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 2000},
    ("Case", "421F"):              {"cost_i": 758, "cost_1": 0,   "cost_2": 1080, "cost_3": 2820, "cost_s": 3625, "hr_i": 150, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 1500},
    # ── Kobelco Excavators ──
    ("Kobelco", "SK17SR-6E"):      {"cost_i": 885, "cost_1": 0,   "cost_2": 930,  "cost_3": 1260, "cost_s": 0,    "hr_i": 250, "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK26SR-7"):       {"cost_i": 745, "cost_1": 0,   "cost_2": 700,  "cost_3": 1070, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK35SR-7"):       {"cost_i": 550, "cost_1": 0,   "cost_2": 755,  "cost_3": 1200, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK45SRX-7"):      {"cost_i": 600, "cost_1": 0,   "cost_2": 685,  "cost_3": 1270, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK55SRX-7"):      {"cost_i": 600, "cost_1": 0,   "cost_2": 800,  "cost_3": 1270, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK75SR-7"):       {"cost_i": 720, "cost_1": 0,   "cost_2": 925,  "cost_3": 1700, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK85CS-7"):       {"cost_i": 620, "cost_1": 0,   "cost_2": 1050, "cost_3": 1700, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK130LC-11"):     {"cost_i": 650, "cost_1": 0,   "cost_2": 1050, "cost_3": 2100, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK140SR-7"):      {"cost_i": 840, "cost_1": 0,   "cost_2": 1675, "cost_3": 2080, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "ED160-7"):        {"cost_i": 900, "cost_1": 0,   "cost_2": 1170, "cost_3": 2010, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK170LC-11"):     {"cost_i": 990, "cost_1": 0,   "cost_2": 2290, "cost_3": 2085, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK210LC-11"):     {"cost_i": 1010,"cost_1": 0,   "cost_2": 1500, "cost_3": 2250, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK230SR-7"):      {"cost_i": 1010,"cost_1": 0,   "cost_2": 1700, "cost_3": 2250, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK260LC-11"):     {"cost_i": 1010,"cost_1": 0,   "cost_2": 1520, "cost_3": 2250, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK270SR-7"):      {"cost_i": 930, "cost_1": 0,   "cost_2": 1730, "cost_3": 2250, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK300LC-11"):     {"cost_i": 1060,"cost_1": 0,   "cost_2": 1360, "cost_3": 2050, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK350LC-11"):     {"cost_i": 1060,"cost_1": 0,   "cost_2": 1550, "cost_3": 2110, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK380SRLC-7"):    {"cost_i": 1385,"cost_1": 0,   "cost_2": 1950, "cost_3": 2370, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Kobelco", "SK520LC-11"):     {"cost_i": 2390,"cost_1": 0,   "cost_2": 3050, "cost_3": 3120, "cost_s": 0,    "hr_i": 50,  "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    # ── Develon ──
    ("Develon", "DA45"):           {"cost_i": 2895,"cost_1": 1575,"cost_2": 5495, "cost_3": 5495, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DD100"):          {"cost_i": 470, "cost_1": 1250,"cost_2": 2625, "cost_3": 2625, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DD130"):          {"cost_i": 470, "cost_1": 1250,"cost_2": 2625, "cost_3": 2625, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX35Z-7"):        {"cost_i": 450, "cost_1": 1300,"cost_2": 1925, "cost_3": 1925, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX50Z-7"):        {"cost_i": 450, "cost_1": 1425,"cost_2": 1965, "cost_3": 1965, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX63-7"):         {"cost_i": 530, "cost_1": 1425,"cost_2": 1965, "cost_3": 1965, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX89R-7"):        {"cost_i": 450, "cost_1": 1350,"cost_2": 1965, "cost_3": 1965, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX140LC-7"):      {"cost_i": 450, "cost_1": 1425,"cost_2": 2125, "cost_3": 2125, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX140LCR-7"):     {"cost_i": 450, "cost_1": 1425,"cost_2": 2125, "cost_3": 2125, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DX170LC-5"):      {"cost_i": 450, "cost_1": 1425,"cost_2": 2175, "cost_3": 2175, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DL220-7"):        {"cost_i": 525, "cost_1": 1175,"cost_2": 3265, "cost_3": 3265, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    ("Develon", "DL250-7"):        {"cost_i": 525, "cost_1": 1175,"cost_2": 3265, "cost_3": 3265, "cost_s": 0,    "hr_i": 50,  "hr_1": 500, "hr_2": 1000,"hr_3": 2000, "hr_s": 0},
    # ── Bomag ──
    ("Bomag", "BMP8500"):          {"cost_i": 0,   "cost_1": 0,   "cost_2": 550,  "cost_3": 1065, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW11RH"):           {"cost_i": 0,   "cost_1": 0,   "cost_2": 780,  "cost_3": 1585, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW120AD/SL"):       {"cost_i": 0,   "cost_1": 0,   "cost_2": 490,  "cost_3": 1470, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW138AD"):          {"cost_i": 0,   "cost_1": 0,   "cost_2": 490,  "cost_3": 1375, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW141/151/161"):    {"cost_i": 0,   "cost_1": 0,   "cost_2": 490,  "cost_3": 1525, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW177"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 740,  "cost_3": 1670, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW190"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 740,  "cost_3": 1760, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW211"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 895,  "cost_3": 2180, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW900"):            {"cost_i": 0,   "cost_1": 0,   "cost_2": 340,  "cost_3": 895,  "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BW90AD"):           {"cost_i": 0,   "cost_1": 0,   "cost_2": 340,  "cost_3": 895,  "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BF200C"):           {"cost_i": 0,   "cost_1": 0,   "cost_2": 635,  "cost_3": 2130, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
    ("Bomag", "BF300C"):           {"cost_i": 0,   "cost_1": 0,   "cost_2": 635,  "cost_3": 2130, "cost_s": 0,    "hr_i": 0,   "hr_1": 0,   "hr_2": 500, "hr_3": 1000, "hr_s": 0},
}

# Hours Requested options (from dealsheet Data for Formula sheet)
HOURS_REQUESTED_OPTIONS = [500, 1000, 1500, 2000, 2500, 3000, 3500]

# Legacy category-based pricing — used ONLY by the Lead Discovery tab for scoring alerts.
# The PM Calculator tab uses DEALSHEET_MODELS above (exact per-model pricing).
PM_PRICING = {
    "Compact Track Loader": {"annual": 2900}, "Skid Steer": {"annual": 2700},
    "Excavator (Mini)": {"annual": 2400}, "Excavator (Standard)": {"annual": 4800},
    "Backhoe": {"annual": 3200}, "Wheel Loader": {"annual": 4200},
    "Dozer": {"annual": 5200}, "Grader": {"annual": 5500},
    "Roller/Compaction": {"annual": 2800}, "Paver": {"annual": 6500},
    "Milling Machine": {"annual": 7200}, "Telehandler": {"annual": 2200},
    "Forklift": {"annual": 1800}, "Dump Truck": {"annual": 3800},
    "Aerial Lift": {"annual": 2000}, "Generator": {"annual": 1500},
    "Trencher": {"annual": 3200}, "Scissor Lift": {"annual": 1600},
    "Articulated Dump Truck": {"annual": 5000},
    "Other": {"annual": 3200},
}

# Build DEALSHEET_MAKES → {brand: [models]} for the dropdown
DEALSHEET_MAKES = {}
for (brand, model) in DEALSHEET_MODELS:
    DEALSHEET_MAKES.setdefault(brand, []).append(model)
for brand in DEALSHEET_MAKES:
    DEALSHEET_MAKES[brand].sort()

# Map Case model names from alerts to PM categories
MODEL_TO_CATEGORY = {
    "580SN": "Backhoe", "580SN WT": "Backhoe", "580N": "Backhoe", "590SN": "Backhoe",
    "TV370B": "Compact Track Loader", "TV450B": "Compact Track Loader", "TV620B": "Compact Track Loader",
    "TR270B": "Compact Track Loader", "TR310B": "Compact Track Loader", "TR340B": "Compact Track Loader",
    "TL100": "Compact Track Loader",
    "221F EVOLUTION": "Wheel Loader", "321F EVOLUTION": "Wheel Loader",
    "521G": "Wheel Loader", "621G": "Wheel Loader", "721G": "Wheel Loader", "821G": "Wheel Loader", "921G": "Wheel Loader", "1021G": "Wheel Loader",
    "SV280B": "Skid Steer", "SV340B": "Skid Steer",
    "CX42D": "Excavator (Mini)", "CX50D": "Excavator (Mini)", "CX37D": "Excavator (Mini)",
    "CX145D SR": "Excavator (Standard)", "CX210D": "Excavator (Standard)", "CX220E": "Excavator (Standard)",
    "CX260E": "Excavator (Standard)", "CX350D": "Excavator (Standard)", "CX190E": "Excavator (Standard)",
    "CX170E": "Excavator (Standard)", "CX140E": "Excavator (Standard)",
    "DL550": "Dozer", "650M": "Dozer", "850M": "Dozer", "750M": "Dozer", "1650M": "Dozer", "1150M": "Dozer",
    "SL35 TR": "Compact Track Loader", "651G": "Wheel Loader",
}

SERVICE_TYPES = ["Field", "Shop"]

# ─── Session State ───
for key in ["quotes", "current_quote", "leads_df", "procare_vins"]:
    if key not in st.session_state:
        st.session_state[key] = [] if key in ["quotes", "procare_vins"] else ({} if key == "current_quote" else None)

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
            quote_data.get("model", ""), quote_data.get("category", ""),
            quote_data.get("serial", ""), quote_data.get("hours_requested", 0),
            quote_data.get("travel_miles", 0),
            quote_data.get("pm_total", 0), quote_data.get("travel_cost", 0),
            quote_data.get("grand_total", 0),
            quote_data.get("notes", ""),
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


# ═══════════════════════════════════════════════════════════
# HUBSPOT ENRICHMENT
# ═══════════════════════════════════════════════════════════
HUBSPOT_TOKEN = st.secrets.get("hubspot_token", "")

@st.cache_data(ttl=1800, show_spinner="Pulling HubSpot data...")
def fetch_hubspot_companies():
    """Pull company records from HubSpot to cross-reference with leads."""
    if not HUBSPOT_TOKEN:
        return {}
    headers = {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}
    companies = {}
    url = "https://api.hubapi.com/crm/v3/objects/companies"
    params = {
        "limit": 100,
        "properties": "name,city,state,industry,hs_lastmodifieddate,lifecyclestage,num_associated_deals",
    }
    try:
        after = None
        for _ in range(20):  # max 2000 companies
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
                        "last_modified": props.get("hs_lastmodifieddate", ""),
                    }
            paging = data.get("paging", {}).get("next", {})
            after = paging.get("after")
            if not after:
                break
    except Exception:
        pass
    return companies

@st.cache_data(ttl=1800, show_spinner="Pulling HubSpot service tickets...")
def fetch_hubspot_tickets():
    """Pull service tickets from HubSpot for service spend enrichment."""
    if not HUBSPOT_TOKEN:
        return {}
    headers = {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}
    tickets_by_company = {}
    url = "https://api.hubapi.com/crm/v3/objects/tickets"
    params = {
        "limit": 100,
        "properties": "subject,hs_pipeline_stage,createdate,hs_ticket_priority,content",
    }
    try:
        after = None
        for _ in range(10):  # max 1000 tickets
            if after:
                params["after"] = after
            resp = requests.get(url, headers=headers, params=params, timeout=15)
            if resp.status_code != 200:
                break
            data = resp.json()
            for t in data.get("results", []):
                props = t.get("properties", {})
                subj = (props.get("subject") or "").upper()
                tickets_by_company[t["id"]] = {
                    "subject": props.get("subject", ""),
                    "created": props.get("createdate", ""),
                    "priority": props.get("hs_ticket_priority", ""),
                }
            paging = data.get("paging", {}).get("next", {})
            after = paging.get("after")
            if not after:
                break
    except Exception:
        pass
    return tickets_by_company

def enrich_leads_with_hubspot(scored_df, hs_companies):
    """Add HubSpot data columns to scored leads."""
    if not hs_companies or scored_df.empty:
        scored_df["In HubSpot"] = False
        scored_df["HubSpot Deals"] = 0
        scored_df["Lifecycle"] = ""
        return scored_df

    df = scored_df.copy()
    df["cust_upper"] = df["Customer"].str.strip().str.upper()

    hs_match = []
    hs_deals = []
    hs_lifecycle = []
    for _, row in df.iterrows():
        cust = row["cust_upper"]
        # Try exact match first, then partial
        match = hs_companies.get(cust)
        if not match:
            for hs_name, hs_data in hs_companies.items():
                if cust in hs_name or hs_name in cust:
                    match = hs_data
                    break
        hs_match.append(bool(match))
        hs_deals.append(int(match["deals"]) if match and match.get("deals") else 0)
        hs_lifecycle.append(match["lifecycle"] if match else "")

    df["In HubSpot"] = hs_match
    df["HubSpot Deals"] = hs_deals
    df["Lifecycle"] = hs_lifecycle
    df.drop(columns=["cust_upper"], inplace=True)

    # Boost score for active HubSpot customers (warm relationship)
    df.loc[df["In HubSpot"], "Lead Score"] = (df.loc[df["In HubSpot"], "Lead Score"] * 1.05).clip(0, 100).round(1)
    # Re-tier after boost
    df["Tier"] = pd.cut(
        df["Lead Score"],
        bins=[0, 35, 50, 65, 100],
        labels=["Low", "Medium", "High", "Top"],
        include_lowest=True,
    )

    return df


# ═══════════════════════════════════════════════════════════
# HUBSPOT DEAL CREATION (PM Contracts Pipeline)
# ═══════════════════════════════════════════════════════════
PM_PIPELINE_ID = "894132043"  # PM Contracts pipeline

def _hs_headers():
    return {"Authorization": f"Bearer {HUBSPOT_TOKEN}", "Content-Type": "application/json"}

def _get_pm_stage_id(stage_name="Qualification"):
    """Look up the stage ID for a given stage name in the PM Contracts pipeline."""
    if not HUBSPOT_TOKEN:
        return None
    try:
        resp = requests.get(
            f"https://api.hubapi.com/crm/v3/pipelines/deals/{PM_PIPELINE_ID}",
            headers=_hs_headers(), timeout=10,
        )
        if resp.status_code == 200:
            for stage in resp.json().get("stages", []):
                if stage["label"].lower() == stage_name.lower():
                    return stage["id"]
    except Exception:
        pass
    return None

def create_pm_deal(quote_data, stage_name="Qualification"):
    """Create a deal in the PM Contracts pipeline and attach a line item for the machine.
    Returns (deal_id, error_msg). Only touches the PM Contracts pipeline.
    """
    if not HUBSPOT_TOKEN:
        return None, "HubSpot token not configured"

    stage_id = _get_pm_stage_id(stage_name)
    if not stage_id:
        return None, f"Could not find '{stage_name}' stage in PM Contracts pipeline"

    customer = quote_data.get("customer_name", "Unknown")
    make = quote_data.get("make", "")
    model = quote_data.get("model", "")
    category = quote_data.get("category", "")
    grand_total = quote_data.get("grand_total", 0)
    branch = quote_data.get("branch", "")

    deal_name = f"PM Contract - {customer} ({make} {model})".strip()

    # Build deal properties — only PM-specific fields, won't affect other pipelines
    properties = {
        "dealname": deal_name,
        "pipeline": PM_PIPELINE_ID,
        "dealstage": stage_id,
        "amount": str(grand_total),
        "description": (
            f"Make: {make} | Model: {model} | Category: {category}\n"
            f"Serial: {quote_data.get('serial', 'N/A')} | "
            f"Hours Requested: {quote_data.get('hours_requested', 0)}\n"
            f"Service: {quote_data.get('service_type', 'N/A')} | "
            f"Rep: {quote_data.get('rep', 'N/A')}\n"
            f"PM Total: ${quote_data.get('pm_total', 0):,.0f} | "
            f"Travel: ${quote_data.get('travel_cost', 0):,.0f}\n"
            f"Notes: {quote_data.get('notes', '')}"
        ),
        "num_machines_needed": "1",
    }
    # Add PM custom properties if they have values
    if branch:
        properties["quoting_branch"] = branch

    try:
        # 1. Create the deal
        resp = requests.post(
            "https://api.hubapi.com/crm/v3/objects/deals",
            headers=_hs_headers(), json={"properties": properties}, timeout=15,
        )
        if resp.status_code not in (200, 201):
            return None, f"Failed to create deal: {resp.status_code} - {resp.text[:200]}"

        deal_id = resp.json()["id"]

        # 2. Create a line item for this machine
        line_item_props = {
            "name": f"{make} {model}".strip() or category,
            "description": (
                f"Category: {category} | Hours Requested: {quote_data.get('hours_requested', 0)} | "
                f"Serial: {quote_data.get('serial', 'N/A')} | "
                f"Service: {quote_data.get('service_type', 'N/A')} | "
                f"PM Services: ${quote_data.get('pm_total', 0):,.0f}"
            ),
            "price": str(grand_total),
            "quantity": "1",
            "hs_sku": quote_data.get("serial", ""),
        }
        li_resp = requests.post(
            "https://api.hubapi.com/crm/v3/objects/line_items",
            headers=_hs_headers(), json={"properties": line_item_props}, timeout=15,
        )

        if li_resp.status_code in (200, 201):
            li_id = li_resp.json()["id"]
            # 3. Associate line item with the deal
            requests.put(
                f"https://api.hubapi.com/crm/v3/objects/deals/{deal_id}/associations/line_items/{li_id}/deal_to_line_item",
                headers=_hs_headers(), timeout=10,
            )

        return deal_id, None

    except Exception as e:
        return None, f"Error: {str(e)}"


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

def get_pm_category(model_name):
    """Map a Case model to PM pricing category."""
    if not model_name:
        return "Other"
    model_upper = str(model_name).upper().strip()
    # Direct lookup
    for k, v in MODEL_TO_CATEGORY.items():
        if k.upper() == model_upper:
            return v
    # Fuzzy
    if any(x in model_upper for x in ["580", "590", "LB"]):
        return "Backhoe"
    if any(x in model_upper for x in ["TV", "TR", "SVL", "TL"]):
        return "Compact Track Loader"
    if any(x in model_upper for x in ["SR", "SV", "SSV"]):
        return "Skid Steer"
    if any(x in model_upper for x in ["21F", "21G", "21B", "21C", "21D", "21R", "21X"]):
        return "Wheel Loader"
    if any(x in model_upper for x in ["CX1", "CX2", "CX3", "CX5", "SK1", "SK2", "SK3"]):
        return "Excavator (Standard)"
    if any(x in model_upper for x in ["CX3", "CX4", "CX5", "CX6", "CX7", "SK17", "SK25", "SK35", "SK45", "SK55"]):
        return "Excavator (Mini)"
    if any(x in model_upper for x in ["50M", "50L", "50K", "DL", "D4", "D5", "D6", "D8", "1650"]):
        return "Dozer"
    if any(x in model_upper for x in ["BW", "BPR"]):
        return "Roller/Compaction"
    return "Other"

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

    # Add PM category
    df["Category"] = df["Model"].apply(get_pm_category)
    df["Annual PM Value"] = df["Category"].apply(lambda c: PM_PRICING.get(c, PM_PRICING["Other"])["annual"])

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

    # Annual PM value score: bigger machines = more contract value
    max_annual = max(PM_PRICING.values(), key=lambda x: x["annual"])["annual"]
    df["value_score"] = (df["Annual PM Value"] / max_annual * 100)

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

    return df.sort_values("Lead Score", ascending=False)

def aggregate_customer_leads(scored_df):
    """Roll up machine-level scores to customer level."""
    if scored_df.empty:
        return pd.DataFrame()

    agg = scored_df.groupby("Customer").agg(
        machines=("VIN", "nunique"),
        total_parts_value=("Parts Value", "sum"),
        total_labor_hrs=("Labor Hrs", "sum"),
        avg_hours=("Eng Hrs", "mean"),
        avg_score=("Lead Score", "mean"),
        max_score=("Lead Score", "max"),
        total_annual_pm=("Annual PM Value", "sum"),
        location=("Location", "first"),
        models=("Model", lambda x: ", ".join(sorted(set(x)))),
        categories=("Category", lambda x: ", ".join(sorted(set(x)))),
    ).reset_index()

    # Customer-level score: weighted by fleet size and total opportunity
    agg["Customer Score"] = (
        agg["avg_score"] * 0.5 +
        (agg["machines"].clip(1, 30) / 30 * 100) * 0.25 +
        (agg["total_annual_pm"] / agg["total_annual_pm"].max() * 100) * 0.25
    ).round(1)

    agg["Tier"] = pd.cut(
        agg["Customer Score"],
        bins=[0, 35, 50, 65, 100],
        labels=["Low", "Medium", "High", "Top"],
        include_lowest=True,
    )

    return agg.sort_values("Customer Score", ascending=False)


# ═══════════════════════════════════════════════════════════
# PDF GENERATION
# ═══════════════════════════════════════════════════════════
def generate_pdf(quote_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.75*inch, rightMargin=0.75*inch)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SECTitle", fontSize=22, fontName="Helvetica-Bold", textColor=colors.HexColor(SEC_RED), spaceAfter=4))
    styles.add(ParagraphStyle(name="SECSubtitle", fontSize=11, fontName="Helvetica", textColor=colors.HexColor(SEC_GRAY), spaceAfter=16))
    styles.add(ParagraphStyle(name="SectionHead", fontSize=13, fontName="Helvetica-Bold", textColor=colors.HexColor(SEC_DARK), spaceBefore=16, spaceAfter=8))
    styles.add(ParagraphStyle(name="FooterText", fontSize=8, fontName="Helvetica", textColor=colors.HexColor(SEC_GRAY), alignment=TA_CENTER))

    elements = []
    elements.append(Paragraph("SOUTHEASTERN EQUIPMENT CO.", styles["SECTitle"]))
    elements.append(Paragraph("Preventive Maintenance Quote", styles["SECSubtitle"]))
    elements.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor(SEC_RED)))
    elements.append(Spacer(1, 12))

    # Quote info
    info_data = [
        ["Quote Date", "Branch", "Service Rep", "Service Type"],
        [quote_data.get("date", ""), quote_data.get("branch", ""), quote_data.get("rep", ""), quote_data.get("service_type", "")],
    ]
    t = RLTable(info_data, colWidths=[1.7*inch]*4)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(SEC_GRAY)),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica"), ("FONTSIZE", (0, 1), (-1, 1), 11),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2), ("TOPPADDING", (0, 1), (-1, 1), 0),
        ("LINEBELOW", (0, 1), (-1, 1), 0.5, colors.HexColor("#DDDDDD")),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 8))

    # Customer
    cust = quote_data.get("customer_name", "")
    if cust:
        elements.append(Paragraph("Customer", styles["SectionHead"]))
        elements.append(Paragraph(cust, styles["Normal"]))
        elements.append(Spacer(1, 8))

    # Machine details
    elements.append(Paragraph("Machine Details", styles["SectionHead"]))
    mach_data = [
        ["Brand", "Model", "Category", "Serial Number"],
        [quote_data.get("make", ""), quote_data.get("model", ""), quote_data.get("category", ""), quote_data.get("serial", "")],
        ["Current Hours", "Hours Requested", "Service Type", "Mileage (one way)"],
        [f"{quote_data.get('current_hours', 0):,}", f"{quote_data.get('hours_requested', 0):,}", quote_data.get("service_type", ""), f"{quote_data.get('travel_miles', 0)} mi" if quote_data.get("travel_miles", 0) > 0 else "N/A"],
    ]
    t = RLTable(mach_data, colWidths=[1.7*inch]*4)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(SEC_GRAY)),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica"), ("FONTSIZE", (0, 1), (-1, 1), 11),
        ("FONTNAME", (0, 2), (-1, 2), "Helvetica-Bold"), ("FONTSIZE", (0, 2), (-1, 2), 8),
        ("TEXTCOLOR", (0, 2), (-1, 2), colors.HexColor(SEC_GRAY)),
        ("FONTNAME", (0, 3), (-1, 3), "Helvetica"), ("FONTSIZE", (0, 3), (-1, 3), 11),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2), ("BOTTOMPADDING", (0, 2), (-1, 2), 2),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 8))

    # Service Breakdown — only include lines with qty > 0
    cur_hr = quote_data.get("current_hours", 0)
    end_hr = quote_data.get("end_hours", cur_hr + quote_data.get("hours_requested", 0))
    range_label = f"(Current {cur_hr:,} hrs → {end_hr:,} hrs)" if cur_hr > 0 else ""
    elements.append(Paragraph(f"Service Breakdown {range_label}", styles["SectionHead"]))
    svc_data = [["Service", "Hour Interval", "Qty", "Cost (Per)", "Subtotal"]]
    if quote_data.get("initial_qty", 0) > 0:
        svc_data.append(["Initial *", str(quote_data["initial_hr"]), str(quote_data["initial_qty"]),
                         f"${quote_data['initial_cost']:,}", f"${quote_data['initial_total']:,}"])
    if quote_data.get("spec_qty", 0) > 0:
        svc_data.append(["Specialty", str(quote_data["spec_hr"]), str(quote_data["spec_qty"]),
                         f"${quote_data['spec_cost']:,}", f"${quote_data['spec_total']:,}"])
    if quote_data.get("int1_qty", 0) > 0:
        svc_data.append(["Interval 1", str(quote_data["int1_hr"]), str(quote_data["int1_qty"]),
                         f"${quote_data['int1_cost']:,}", f"${quote_data['int1_total']:,}"])
    if quote_data.get("int2_qty", 0) > 0:
        svc_data.append(["Interval 2", str(quote_data["int2_hr"]), str(quote_data["int2_qty"]),
                         f"${quote_data['int2_cost']:,}", f"${quote_data['int2_total']:,}"])
    if quote_data.get("int3_qty", 0) > 0:
        svc_data.append(["Interval 3", str(quote_data["int3_hr"]), str(quote_data["int3_qty"]),
                         f"${quote_data['int3_cost']:,}", f"${quote_data['int3_total']:,}"])

    travel_cost = quote_data.get("travel_cost", 0)
    pm_total = quote_data.get("pm_total", 0)
    grand_total = quote_data.get("grand_total", 0)

    svc_data.append(["", "", "", "", ""])
    if travel_cost > 0:
        visits = quote_data.get("travel_visits", 0)
        per_visit = quote_data.get("travel_cost_per_visit", 0)
        svc_data.append([f"Travel ({visits} visits × ${per_visit:,.0f})", "", "", "", f"${travel_cost:,.0f}"])
    svc_data.append(["PM Services Total", "", "", "", f"${pm_total:,}"])
    svc_data.append(["Grand Total", "", "", "", f"${grand_total:,}"])

    t = RLTable(svc_data, colWidths=[1.8*inch, 1.2*inch, 0.6*inch, 1.4*inch, 1.5*inch])
    row_count = len(svc_data)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5496")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("FONTNAME", (0, -2), (-1, -1), "Helvetica-Bold"), ("FONTSIZE", (0, -2), (-1, -1), 11),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor(SEC_RED)),
        ("TOPPADDING", (0, 1), (-1, -1), 5), ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("LINEBELOW", (0, -3), (-1, -3), 0.5, colors.HexColor("#DDDDDD")),
        ("LINEBELOW", (0, -1), (-1, -1), 1.5, colors.HexColor(SEC_RED)),
    ]))
    elements.append(t)
    if quote_data.get("initial_qty", 0) > 0:
        elements.append(Spacer(1, 4))
        elements.append(Paragraph("* Unit has an initial service that is only performed once.", styles["FooterText"]))

    # Notes
    notes = quote_data.get("notes", "")
    if notes:
        elements.append(Spacer(1, 12))
        elements.append(Paragraph("Notes", styles["SectionHead"]))
        elements.append(Paragraph(notes, styles["Normal"]))

    # Footer
    elements.append(Spacer(1, 24))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#DDDDDD")))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("Southeastern Equipment Co. | southeasternequip.com", styles["FooterText"]))
    elements.append(Paragraph(f"Generated {datetime.now().strftime('%m/%d/%Y %I:%M %p')}", styles["FooterText"]))

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ─── Helpers ───
def get_category_for_model(make, model):
    make_data = MAKES_MODELS.get(make, {})
    for category, models in make_data.items():
        if model in models:
            return category
    return "Other"

def get_models_for_make(make):
    make_data = MAKES_MODELS.get(make, {})
    all_models = []
    for models in make_data.values():
        all_models.extend(models)
    return sorted(set(all_models))

TRAVEL_FREE_MILES    = 10    # under 10 miles from branch → no travel charge
TRAVEL_FLAT_FEE      = 300   # flat $300 for mileage + labor, 10-60 miles
TRAVEL_FLAT_MAX      = 60    # miles threshold for flat fee
TRAVEL_OVER_RATE     = 5.00  # $/mile beyond 60 miles

def _calc_travel(travel_miles, service_type):
    """Travel pricing: free ≤10mi, flat $300 for 10-60mi, $300 + $5/mi over 60."""
    if service_type != "Field" or travel_miles <= TRAVEL_FREE_MILES:
        return 0
    if travel_miles <= TRAVEL_FLAT_MAX:
        return TRAVEL_FLAT_FEE
    return round(TRAVEL_FLAT_FEE + (travel_miles - TRAVEL_FLAT_MAX) * TRAVEL_OVER_RATE, 2)

def calculate_pm_cost(brand, model, hours_requested, travel_miles=0, service_type="Shop", current_hours=0):
    """Calculate PM cost using the exact dealsheet Rev 3.0 algorithm.
    Matches the formulas on the Main Sheet of the PM Contract Dealsheet.
    current_hours: the machine's current hour meter — determines which service
    intervals actually fall within the contract window.
    """
    key = (brand, model)
    if key not in DEALSHEET_MODELS:
        return None  # Model not in dealsheet

    m = DEALSHEET_MODELS[key]
    cost_i = m["cost_i"]
    cost_1 = m["cost_1"]
    cost_2 = m["cost_2"]
    cost_3 = m["cost_3"]
    cost_s = m["cost_s"]
    hr_i   = m["hr_i"]
    hr_1   = m["hr_1"]
    hr_2   = m["hr_2"]
    hr_3   = m["hr_3"]
    hr_s   = m["hr_s"]

    # Contract window: from current_hours up to the next hours_requested milestone.
    # Per Jarred: "1,000 hours of PM" at 4,900 hrs means 4,900 → 5,000 (next 1,000-hr mark),
    # NOT 4,900 → 5,900.  If current_hours is 0 or already on a milestone, end = start + hours_requested.
    start_hr = current_hours
    if current_hours > 0 and hours_requested > 0:
        end_hr = ((current_hours // hours_requested) + 1) * hours_requested
    else:
        end_hr = current_hours + hours_requested

    # ── Helper: count how many times an interval hits inside (start_hr, end_hr] ──
    def _hits(interval):
        """How many multiples of `interval` land in (start_hr, end_hr]."""
        if not interval or interval <= 0:
            return 0
        first = (start_hr // interval) * interval + interval  # next multiple after start
        if first > end_hr:
            return 0
        return int((end_hr - first) // interval) + 1

    # Raw hit counts per interval
    raw_s  = _hits(hr_s) if hr_s and hr_s > 0 else 0
    raw_3  = _hits(hr_3) if hr_3 and hr_3 > 0 else 0
    raw_2  = _hits(hr_2) if hr_2 and hr_2 > 0 else 0
    raw_1  = _hits(hr_1) if hr_1 and hr_1 > 0 else 0

    # Higher intervals supersede lower ones (same dealsheet logic)
    spec_qty = raw_s
    int3_qty = raw_3 - spec_qty if raw_3 > spec_qty else 0
    int2_qty = max(raw_2 - raw_3, 0)
    int1_qty = max(raw_1 - raw_2, 0)

    # Initial — only if the machine hasn't yet hit the initial service hour
    initial_qty = 1 if (hr_i and hr_i > 0 and start_hr < hr_i <= end_hr) else 0

    # ── Cost totals ──
    initial_total  = initial_qty * cost_i
    spec_total     = spec_qty * cost_s
    int1_total     = int1_qty * cost_1
    int2_total     = int2_qty * cost_2
    int3_total     = int3_qty * cost_3
    pm_total       = initial_total + spec_total + int1_total + int2_total + int3_total

    # Travel cost — free ≤10mi, flat $300 for 10-60mi, $300 + $5/mi over 60
    travel_cost = _calc_travel(travel_miles, service_type)
    # Multiply travel by total number of service visits
    total_visits = initial_qty + spec_qty + int3_qty + int2_qty + int1_qty
    travel_total = travel_cost * total_visits if total_visits > 0 else travel_cost
    grand_total = pm_total + travel_total

    return {
        "pm_total": pm_total,
        "travel_cost_per_visit": travel_cost,
        "travel_visits": total_visits,
        "travel_cost": travel_total,
        "grand_total": grand_total,
        "hours_requested": hours_requested,
        "current_hours": current_hours,
        "end_hours": end_hr,
        # Service breakdown for display
        "initial_qty": initial_qty, "initial_cost": cost_i, "initial_total": initial_total, "initial_hr": hr_i,
        "int1_qty": int1_qty, "int1_cost": cost_1, "int1_total": int1_total, "int1_hr": hr_1,
        "int2_qty": int2_qty, "int2_cost": cost_2, "int2_total": int2_total, "int2_hr": hr_2,
        "int3_qty": int3_qty, "int3_cost": cost_3, "int3_total": int3_total, "int3_hr": hr_3,
        "spec_qty": spec_qty, "spec_cost": cost_s, "spec_total": spec_total, "spec_hr": hr_s,
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
</style>
""", unsafe_allow_html=True)

# ─── Header ───
st.markdown('<div class="main-header"><h1>Southeastern Equipment Co.</h1><p>PM Lead Discovery & Calculator</p></div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════
tab_leads, tab_calc, tab_history, tab_pricing = st.tabs(["Lead Discovery", "PM Calculator", "Quote History", "Pricing Reference"])

# ═══════════════════════════════════════════════════════════
# TAB 1: LEAD DISCOVERY
# ═══════════════════════════════════════════════════════════
with tab_leads:
    st.subheader("PM Lead Discovery")
    st.caption("Scores every machine from Case maintenance alerts that does NOT have ProCare coverage. Data auto-loads from the latest files. Upload new files below to refresh.")

    # Auto-load bundled data
    alerts_df = load_bundled_alerts()
    procare_vins = load_bundled_procare()

    # Let users upload newer files to replace
    with st.expander("Upload Updated Data Files", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            alerts_file = st.file_uploader("Maintenance Alerts (.xlsx)", type=["xlsx"], key="alerts_upload")
        with col2:
            procare_file = st.file_uploader("ProCare Stops (.xlsx)", type=["xlsx"], key="procare_upload")
        if alerts_file:
            alerts_df = parse_maintenance_alerts(alerts_file)
        if procare_file:
            procare_vins = parse_procare_stops(procare_file)

    if alerts_df is not None and not alerts_df.empty:
        scored = score_leads(alerts_df, procare_vins)

        # HubSpot enrichment
        hs_companies = fetch_hubspot_companies()
        if hs_companies:
            scored = enrich_leads_with_hubspot(scored, hs_companies)

        st.session_state.leads_df = scored
        st.session_state.procare_vins = procare_vins

        if scored.empty:
            st.warning("No external leads found after filtering out ProCare and internal machines.")
        else:
            # Summary metrics
            cust_agg = aggregate_customer_leads(scored)

            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Total Machines", len(scored))
            with col2:
                st.metric("Unique Customers", scored["Customer"].nunique())
            with col3:
                top_count = len(scored[scored["Tier"] == "Top"])
                st.metric("Top Tier Leads", top_count)
            with col4:
                total_pm_opp = scored["Annual PM Value"].sum()
                st.metric("Total PM Opportunity", f"${total_pm_opp:,.0f}")
            with col5:
                if "In HubSpot" in scored.columns:
                    hs_count = scored[scored["In HubSpot"]]["Customer"].nunique()
                    st.metric("In HubSpot", f"{hs_count} customers")
                else:
                    st.metric("ProCare Excluded", len(procare_vins))

            st.divider()

            # Filters
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                filter_tier = st.multiselect("Tier", ["Top", "High", "Medium", "Low"], default=["Top", "High"])
            with col2:
                filter_location = st.multiselect("Branch", sorted(scored["Location"].unique()))
            with col3:
                filter_category = st.multiselect("Machine Category", sorted(scored["Category"].unique()))
            with col4:
                min_score = st.slider("Min Lead Score", 0, 100, 50)

            # HubSpot filter if available
            if "In HubSpot" in scored.columns:
                hs_filter = st.radio("HubSpot Status", ["All Leads", "In HubSpot Only", "NOT in HubSpot"], horizontal=True)
            else:
                hs_filter = "All Leads"

            # Apply filters
            display = scored.copy()
            if filter_tier:
                display = display[display["Tier"].isin(filter_tier)]
            if filter_location:
                display = display[display["Location"].isin(filter_location)]
            if filter_category:
                display = display[display["Category"].isin(filter_category)]
            display = display[display["Lead Score"] >= min_score]
            if hs_filter == "In HubSpot Only" and "In HubSpot" in display.columns:
                display = display[display["In HubSpot"]]
            elif hs_filter == "NOT in HubSpot" and "In HubSpot" in display.columns:
                display = display[~display["In HubSpot"]]

            # View toggle
            view = st.radio("View", ["By Customer", "By Machine"], horizontal=True)

            if view == "By Customer":
                cust_display = aggregate_customer_leads(display)
                if cust_display.empty:
                    st.info("No customers match filters.")
                else:
                    st.caption(f"Showing {len(cust_display)} customers")
                    for _, row in cust_display.iterrows():
                        tier = row["Tier"]

                        with st.container():
                            c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 1])
                            with c1:
                                st.markdown(f"**{row['Customer']}**")
                                st.caption(f"{row['location']} | {row['models']}")
                            with c2:
                                st.metric("Machines", int(row["machines"]))
                            with c3:
                                st.metric("PM Value", f"${row['total_annual_pm']:,.0f}")
                            with c4:
                                st.metric("Parts Opp", f"${row['total_parts_value']:,.0f}")
                            with c5:
                                st.metric("Score", f"{row['Customer Score']:.0f}", delta=tier)
                            st.divider()

            else:  # By Machine
                st.caption(f"Showing {len(display)} machines")
                show_cols = ["Lead Score", "Tier", "Customer", "Model", "Category", "Location", "Eng Hrs", "Stop", "Parts Value", "Labor Hrs", "Annual PM Value", "VIN"]
                if "In HubSpot" in display.columns:
                    show_cols.insert(3, "In HubSpot")
                show_cols = [c for c in show_cols if c in display.columns]
                st.dataframe(
                    display[show_cols].reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                    column_config={
                        "Lead Score": st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%.0f"),
                        "Parts Value": st.column_config.NumberColumn("Parts $", format="$%.0f"),
                        "Annual PM Value": st.column_config.NumberColumn("Annual PM", format="$%.0f"),
                        "Eng Hrs": st.column_config.NumberColumn("Hours", format="%.0f"),
                        "Labor Hrs": st.column_config.NumberColumn("Labor Hrs", format="%.1f"),
                        "In HubSpot": st.column_config.CheckboxColumn("In HS", default=False),
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
            col1, col2 = st.columns(2)
            with col1:
                import plotly.express as px
                loc_data = display.groupby("Location").agg(
                    machines=("VIN", "nunique"),
                    pm_value=("Annual PM Value", "sum"),
                ).reset_index()
                fig = px.bar(loc_data.sort_values("pm_value", ascending=True), x="pm_value", y="Location",
                             orientation="h", title="PM Opportunity by Branch",
                             labels={"pm_value": "Annual PM Value ($)", "Location": ""},
                             color_discrete_sequence=[SEC_RED])
                fig.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig, use_container_width=True)

            with col2:
                cat_data = display.groupby("Category").agg(
                    machines=("VIN", "nunique"),
                    pm_value=("Annual PM Value", "sum"),
                ).reset_index()
                fig = px.bar(cat_data.sort_values("pm_value", ascending=True), x="pm_value", y="Category",
                             orientation="h", title="PM Opportunity by Machine Category",
                             labels={"pm_value": "Annual PM Value ($)", "Category": ""},
                             color_discrete_sequence=["#2F5496"])
                fig.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No data files found. Upload maintenance alerts above or add files to the data/ folder in the repo.")


# ═══════════════════════════════════════════════════════════
# TAB 2: PM CALCULATOR  (Dealsheet Rev 3.0 algorithm)
# ═══════════════════════════════════════════════════════════
with tab_calc:
    st.subheader("Customer & Job Info")

    col1, col2, col3 = st.columns(3)
    with col1:
        customer_name = st.text_input("Customer Name")
    with col2:
        branch = st.selectbox("Branch", [""] + BRANCHES)
    with col3:
        rep = st.text_input("Service Rep")

    col1, col2 = st.columns(2)
    with col1:
        service_type = st.selectbox("Field or Shop", SERVICE_TYPES)
    with col2:
        travel_miles = st.number_input("Mileage from Branch (one way)", min_value=0, max_value=500, value=0, step=5,
                                       help="Free ≤10 mi · Flat $300 for 10–60 mi · $300 + $5/mi over 60 mi")

    st.divider()
    st.subheader("Machine Info")

    col1, col2 = st.columns(2)
    with col1:
        make = st.selectbox("Make", [""] + sorted(MAKES_MODELS.keys()) + ["Other"])
    if make and make != "Other" and make in MAKES_MODELS:
        with col2:
            model = st.selectbox("Model", [""] + get_models_for_make(make) + ["Other / Not Listed"])
    else:
        with col2:
            model = st.text_input("Model (type in)")

    auto_category = None
    if make and make != "Other" and model and model != "Other / Not Listed":
        auto_category = get_category_for_model(make, model)

    col1, col2, col3 = st.columns(3)
    with col1:
        if auto_category and auto_category != "Other":
            st.text_input("Category (auto)", value=auto_category, disabled=True)
            category = auto_category
        else:
            category = st.selectbox("Machine Category", [""] + sorted(set(
                cat for cats in MAKES_MODELS.values() for cat in cats.keys()
            )))
    with col2:
        serial = st.text_input("Serial Number")
    with col3:
        hours_requested = st.selectbox("Hours Requested", HOURS_REQUESTED_OPTIONS, index=3,
                                       help="Total contract hours — matches the dealsheet")

    col1, col2 = st.columns(2)
    with col1:
        machine_age = st.number_input("Machine Age (years)", min_value=0, max_value=50, value=0, step=1)
    with col2:
        machine_hours = st.number_input("Current Hours", min_value=0, max_value=100000, value=0, step=100)

    st.divider()
    notes = st.text_area("Notes", placeholder="Machine condition, special requirements, etc.", height=80)

    st.divider()
    # Check if model is in dealsheet for accurate pricing
    in_dealsheet = bool(make and model and model != "Other / Not Listed" and (make, model) in DEALSHEET_MODELS)
    can_calc = bool(make and model and model != "Other / Not Listed" and in_dealsheet)

    if make and model and model != "Other / Not Listed" and not in_dealsheet:
        st.warning(f"'{make} {model}' is not in the PM Dealsheet Rev 3.0 yet. Contact Jarred to get pricing added.")

    if st.button("Calculate PM Estimate", type="primary", use_container_width=True, disabled=not can_calc):
        result = calculate_pm_cost(
            brand=make, model=model, hours_requested=hours_requested,
            travel_miles=travel_miles if service_type == "Field" else 0,
            service_type=service_type,
            current_hours=machine_hours,
        )
        if result:
            st.session_state.current_quote = {
                "date": datetime.now().strftime("%m/%d/%Y"),
                "customer_name": customer_name, "branch": branch, "rep": rep,
                "service_type": service_type, "make": make,
                "model": model, "category": category, "serial": serial,
                "machine_age": machine_age, "machine_hours": machine_hours,
                "hours_requested": hours_requested,
                "travel_miles": travel_miles if service_type == "Field" else 0,
                "notes": notes, **result,
            }

    if st.session_state.current_quote:
        q = st.session_state.current_quote
        st.divider()
        st.subheader("PM Estimate")

        # Top-line totals
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f'<div class="metric-card"><div class="label">PM Services Total</div><div class="value">${q["pm_total"]:,.0f}</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><div class="label">Travel</div><div class="value">${q["travel_cost"]:,.0f}</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card"><div class="label">Grand Total</div><div class="value" style="color:#C8102E;">${q["grand_total"]:,.0f}</div></div>', unsafe_allow_html=True)

        cur_hr = q.get("current_hours", 0)
        end_hr = q.get("end_hours", cur_hr + q.get("hours_requested", 0))
        hr_range = f"Current {cur_hr:,} hrs → {end_hr:,} hrs" if cur_hr > 0 else f'{q.get("hours_requested",0):,} hours requested'
        st.caption(f'{q.get("make","")} {q.get("model","")} | {hr_range} | Dealsheet Rev 3.0')

        # Service breakdown table — ONLY show rows with qty > 0
        st.markdown("**Service Breakdown**")
        breakdown_rows = []
        if q.get("initial_hr") and q["initial_hr"] > 0 and q.get("initial_qty", 0) > 0:
            breakdown_rows.append({"Service": "Initial *", "Hour Interval": q["initial_hr"], "Qty": q["initial_qty"], "Cost (Per)": f"${q['initial_cost']:,}", "Subtotal": f"${q['initial_total']:,}"})
        if q.get("spec_hr") and q["spec_hr"] > 0 and q.get("spec_qty", 0) > 0:
            breakdown_rows.append({"Service": "Specialty", "Hour Interval": q["spec_hr"], "Qty": q["spec_qty"], "Cost (Per)": f"${q['spec_cost']:,}", "Subtotal": f"${q['spec_total']:,}"})
        if q.get("int1_hr") and q["int1_hr"] > 0 and q.get("int1_qty", 0) > 0:
            breakdown_rows.append({"Service": "Interval 1", "Hour Interval": q["int1_hr"], "Qty": q["int1_qty"], "Cost (Per)": f"${q['int1_cost']:,}", "Subtotal": f"${q['int1_total']:,}"})
        if q.get("int2_hr") and q["int2_hr"] > 0 and q.get("int2_qty", 0) > 0:
            breakdown_rows.append({"Service": "Interval 2", "Hour Interval": q["int2_hr"], "Qty": q["int2_qty"], "Cost (Per)": f"${q['int2_cost']:,}", "Subtotal": f"${q['int2_total']:,}"})
        if q.get("int3_hr") and q["int3_hr"] > 0 and q.get("int3_qty", 0) > 0:
            breakdown_rows.append({"Service": "Interval 3", "Hour Interval": q["int3_hr"], "Qty": q["int3_qty"], "Cost (Per)": f"${q['int3_cost']:,}", "Subtotal": f"${q['int3_total']:,}"})
        if breakdown_rows:
            st.dataframe(pd.DataFrame(breakdown_rows), use_container_width=True, hide_index=True)
        else:
            st.info("No service intervals fall within this hour range.")
        if q.get("initial_qty", 0) > 0:
            st.caption("\\* Unit has an initial service that is only performed once.")
        # Show travel detail if applicable
        if q.get("travel_cost", 0) > 0:
            st.caption(f'Travel: ${q.get("travel_cost_per_visit", 0):,.0f}/visit × {q.get("travel_visits", 0)} visits = ${q["travel_cost"]:,.0f}')

        # Action buttons
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            pdf_buf = generate_pdf(q)
            safe = (customer_name or "quote").replace(" ", "_")
            st.download_button("Download PDF Quote", data=pdf_buf, file_name=f"SEC_PM_{safe}_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
        with col2:
            if st.button("Save Quote", use_container_width=True, type="secondary"):
                saved = save_quote_to_sheet(q)
                st.success("Quote saved to Google Sheets" if saved else "Quote saved locally (Sheets not connected)")
        with col3:
            if st.button("Log to HubSpot", use_container_width=True, type="secondary"):
                with st.spinner("Creating deal in PM Contracts pipeline..."):
                    deal_id, err = create_pm_deal(q)
                if deal_id:
                    st.success(f"Deal created in PM Contracts pipeline (ID: {deal_id})")
                else:
                    st.error(f"Could not log to HubSpot: {err}")
        with col4:
            if st.button("Clear / New Quote", use_container_width=True):
                st.session_state.current_quote = {}
                st.rerun()


# ═══════════════════════════════════════════════════════════
# TAB 3: QUOTE HISTORY
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
            fc = st.multiselect("Category", sorted(df["category"].unique()) if "category" in df.columns else [])

        filt = df.copy()
        if fb: filt = filt[filt["branch"].isin(fb)]
        if fr: filt = filt[filt["rep"].isin(fr)]
        if fc: filt = filt[filt["category"].isin(fc)]

        # Use grand_total if available, fall back to annual_pm_price for older quotes
        val_col = "grand_total" if "grand_total" in filt.columns else ("annual_pm_price" if "annual_pm_price" in filt.columns else None)
        if not filt.empty and val_col:
            filt[val_col] = pd.to_numeric(filt[val_col], errors="coerce").fillna(0)
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Quotes", len(filt))
            with c2: st.metric("Total Value", f"${filt[val_col].sum():,.0f}")
            with c3: st.metric("Avg Value", f"${filt[val_col].mean():,.0f}")

        st.dataframe(filt, use_container_width=True, hide_index=True)

        # Bulk push existing quotes to HubSpot
        if not filt.empty and HUBSPOT_TOKEN:
            if st.button("Push All Quotes to HubSpot PM Pipeline", type="primary", use_container_width=False):
                success_count = 0
                fail_count = 0
                progress = st.progress(0, text="Pushing quotes to HubSpot...")
                for idx, row in filt.iterrows():
                    # Build quote_data dict from the dataframe row
                    q = {
                        "customer_name": row.get("customer_name", "Unknown"),
                        "branch": row.get("branch", ""),
                        "rep": row.get("rep", ""),
                        "service_type": row.get("service_type", ""),
                        "make": row.get("make", ""),
                        "model": row.get("model", ""),
                        "category": row.get("category", ""),
                        "serial": row.get("serial", ""),
                        "hours_requested": row.get("hours_requested", 0),
                        "travel_miles": row.get("travel_miles", 0),
                        "pm_total": row.get("pm_total", row.get("annual_pm_price", 0)),
                        "travel_cost": row.get("travel_cost", 0),
                        "grand_total": row.get("grand_total", row.get("annual_pm_price", 0)),
                        "notes": row.get("notes", ""),
                    }
                    deal_id, err = create_pm_deal(q, stage_name="Quote Sent")
                    if deal_id:
                        success_count += 1
                    else:
                        fail_count += 1
                    progress.progress((idx + 1) / len(filt), text=f"Pushed {success_count + fail_count} of {len(filt)}...")
                progress.empty()
                if success_count:
                    st.success(f"Pushed {success_count} quotes to PM Contracts pipeline!")
                if fail_count:
                    st.warning(f"{fail_count} quotes failed to push (check HubSpot token / connection)")

        if not filt.empty and val_col and len(filt) > 1:
            import plotly.express as px
            col1, col2 = st.columns(2)
            with col1:
                if "branch" in filt.columns:
                    bd = filt.groupby("branch")[val_col].sum().reset_index()
                    fig = px.bar(bd, x="branch", y=val_col, title="By Branch", color_discrete_sequence=[SEC_RED])
                    fig.update_layout(showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            with col2:
                if "category" in filt.columns:
                    cd = filt.groupby("category")[val_col].sum().reset_index()
                    fig = px.bar(cd, x="category", y=val_col, title="By Category", color_discrete_sequence=["#2F5496"])
                    fig.update_layout(showlegend=False, xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
# TAB 4: PRICING REFERENCE (Dealsheet Rev 3.0)
# ═══════════════════════════════════════════════════════════
with tab_pricing:
    st.subheader("PM Dealsheet Rev 3.0 — Per-Model Pricing")
    st.caption("Verified by Jarred & service team. Service interval costs by model.")

    # Build a display table from DEALSHEET_MODELS
    ref_rows = []
    for (brand, model), m in sorted(DEALSHEET_MODELS.items()):
        ref_rows.append({
            "Brand": brand, "Model": model,
            "Initial": f"${m['cost_i']:,}" if m['cost_i'] else "—",
            "Int 1": f"${m['cost_1']:,}" if m['cost_1'] else "—",
            "Int 2 ({0}hr)".format(m['hr_2']): f"${m['cost_2']:,}",
            "Int 3 ({0}hr)".format(m['hr_3']): f"${m['cost_3']:,}",
            "Specialty": f"${m['cost_s']:,}" if m['cost_s'] else "—",
        })
    st.dataframe(pd.DataFrame(ref_rows), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Models by Brand")
    for brand in sorted(DEALSHEET_MAKES.keys()):
        models = DEALSHEET_MAKES[brand]
        with st.expander(f"{brand} — {len(models)} models"):
            st.markdown(", ".join(models))

    st.caption("Source: PM Contract Dealsheet Rev 3.0. Contact Jarred to add new models.")

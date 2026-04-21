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
MAKES_MODELS = {
    "Case": {
        "Compact Track Loader": ["TR270", "TR270B", "TR310", "TR310B", "TR320", "TR340", "TR340B", "TV370", "TV370B", "TV380", "TV450", "TV450B", "TV620", "TV620B"],
        "Skid Steer": ["SR130", "SR160B", "SR175", "SR210", "SR210B", "SR220", "SR240", "SR240B", "SR250", "SR270", "SR270B", "SV185", "SV185B", "SV250", "SV280", "SV280B", "SV300", "SV340"],
        "Excavator (Mini)": ["CX17C", "CX26", "CX26C", "CX30", "CX30C", "CX37", "CX37C", "CX42", "CX50B", "CX55B", "CX57", "CX60", "CX60C", "CX75C"],
        "Excavator (Standard)": ["CX130", "CX130D", "CX145", "CX145C", "CX145D", "CX160", "CX160D", "CX170", "CX210", "CX210C", "CX210D", "CX220", "CX350", "CX500"],
        "Backhoe": ["580", "580K", "580N", "580SK", "580SL", "580SM", "580SMII", "580SN", "580SNW", "580SNWT", "580SV", "590", "590N", "590SM", "590SN"],
        "Wheel Loader": ["221F", "221FHS", "321", "321F", "321FHS", "521D", "521EXT", "521F", "521G", "521GXT", "621", "621B", "621C", "621D", "621EXR", "621F", "621FXT", "621G", "621R", "621XR", "721", "721B", "721FXT", "721G", "821", "821F", "821G", "921G"],
        "Dozer": ["650", "650L", "650M", "750M", "850", "850M", "1150", "1150K", "1150M"],
        "Grader": ["670A", "720A", "836C"],
        "Trencher": ["660"],
    },
    "Kobelco": {
        "Excavator (Mini)": ["SK17SR", "SK25", "SK35", "SK35SR", "SK45", "SK55", "SK55SRX"],
        "Excavator (Standard)": ["SK80", "SK85", "SK85CS", "SK85SR", "SK130", "SK140", "SK140SR", "SK170", "SK210", "SK210D", "SK210LC", "SK230", "SK260LC", "SK300", "SK300LC", "SK330", "SK350"],
    },
    "Bomag": {
        "Roller/Compaction": ["BPR45/45", "BW120", "BW120AD", "BW120SL", "BW135AD-5", "BW138", "BW161", "BW177", "BW190", "BW190AD-5", "BW211", "BW211PD", "BW213", "BW266"],
    },
    "Kubota": {
        "Compact Track Loader": ["SVL65", "SVL75", "SVL95", "SVL97"],
        "Skid Steer": ["SSV75"],
        "Excavator (Mini)": ["KX033", "KX040", "KX057", "KX121"],
        "Excavator (Standard)": ["KX080"],
    },
    "Gradall": {"Excavator (Standard)": ["XL3100", "XL3100IV", "XL3100V"]},
    "Komatsu": {"Excavator (Standard)": ["PC128", "PC170"]},
    "Caterpillar": {"Dozer": ["D4G", "D8T"]},
    "New Holland": {"Backhoe": ["LB75", "LB90"]},
    "JLG": {"Aerial Lift": ["25AM", "660SJ"], "Telehandler": ["1255"]},
    "Ditch Witch": {"Trencher": ["TRX300"]},
    "KM International": {"Paver": ["KM4000", "KM8000TEDD"]},
    "Link-Belt": {"Excavator (Standard)": ["130X4", "145X4", "160X4", "210X4", "245X4", "300X4", "350X4"]},
    "John Deere": {
        "Dozer": ["450K", "550K", "650K", "700K", "850K"],
        "Excavator (Standard)": ["135G", "210G", "245G", "300G", "350G"],
        "Wheel Loader": ["344L", "444L", "524L", "544L", "624L", "644L", "744L", "844L"],
    },
    "Volvo": {
        "Excavator (Standard)": ["EC140", "EC160", "EC200", "EC220", "EC250", "EC300", "EC350"],
        "Wheel Loader": ["L60H", "L70H", "L90H", "L110H", "L120H", "L150H"],
    },
}

# ─── PM Pricing (will be replaced when Jarred provides averages) ───
PM_PRICING = {
    "Compact Track Loader": {"annual": 2900, "parts_avg": 1160, "labor_avg": 1740},
    "Skid Steer":           {"annual": 2700, "parts_avg": 1080, "labor_avg": 1620},
    "Excavator (Mini)":     {"annual": 2400, "parts_avg": 960,  "labor_avg": 1440},
    "Excavator (Standard)": {"annual": 4800, "parts_avg": 1920, "labor_avg": 2880},
    "Backhoe":              {"annual": 3200, "parts_avg": 1280, "labor_avg": 1920},
    "Wheel Loader":         {"annual": 4200, "parts_avg": 1680, "labor_avg": 2520},
    "Dozer":                {"annual": 5200, "parts_avg": 2080, "labor_avg": 3120},
    "Grader":               {"annual": 5500, "parts_avg": 2200, "labor_avg": 3300},
    "Roller/Compaction":    {"annual": 2800, "parts_avg": 1120, "labor_avg": 1680},
    "Paver":                {"annual": 6500, "parts_avg": 2600, "labor_avg": 3900},
    "Milling Machine":      {"annual": 7200, "parts_avg": 2880, "labor_avg": 4320},
    "Telehandler":          {"annual": 2200, "parts_avg": 880,  "labor_avg": 1320},
    "Forklift":             {"annual": 1800, "parts_avg": 720,  "labor_avg": 1080},
    "Dump Truck":           {"annual": 3800, "parts_avg": 1520, "labor_avg": 2280},
    "Aerial Lift":          {"annual": 2000, "parts_avg": 800,  "labor_avg": 1200},
    "Generator":            {"annual": 1500, "parts_avg": 600,  "labor_avg": 900},
    "Trencher":             {"annual": 3200, "parts_avg": 1280, "labor_avg": 1920},
    "Scissor Lift":         {"annual": 1600, "parts_avg": 640,  "labor_avg": 960},
    "Other":                {"annual": 3200, "parts_avg": 1280, "labor_avg": 1920},
}

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
            quote_data.get("serial", ""), quote_data.get("machine_age", 0),
            quote_data.get("machine_hours", 0), quote_data.get("travel_time", 0),
            quote_data.get("parts_cost", 0), quote_data.get("labor_cost", 0),
            quote_data.get("travel_cost", 0), quote_data.get("total_cost", 0),
            quote_data.get("annual_pm_price", 0), quote_data.get("margin_pct", 0),
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
            "fleet_size__c", "account_stage__c", "annualrevenue",
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
    deals_by_company = {}  # company_name -> {won, lost, total_won_amount, last_close, has_warranty, warranty_expired}
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
                            deals_by_company[co_name] = {"won": 0, "lost": 0, "total_won_amount": 0, "last_close": "", "has_warranty": False}
                        deals_by_company[co_name][label] += 1
                        amt = float(props.get("amount") or 0)
                        if label == "won":
                            deals_by_company[co_name]["total_won_amount"] += amt
                        cd = props.get("closedate", "")
                        if cd > deals_by_company[co_name]["last_close"]:
                            deals_by_company[co_name]["last_close"] = cd
                        # Track warranty status
                        wtype = (props.get("warranty_type__c") or "").lower()
                        if label == "won" and wtype and "no warranty" not in wtype and "n/a" not in wtype:
                            deals_by_company[co_name]["has_warranty"] = True

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
    2. Warranty Expiring/Expired - bought equipment with warranty that's aging out
    3. Regular Maintenance - comes in for service but no PM contract
    4. Parts Only, No Service - buys parts but handles own service (pitch PM)
    5. Unknown - in alerts but no HubSpot match
    """
    if not match:
        return "No ProCare (New Lead)"

    ps_engagement = match.get("ps_engagement", "")
    last_service = match.get("last_service", "")
    last_parts = match.get("last_parts", "")
    ytd_service = match.get("ytd_service", 0)
    ytd_parts = match.get("ytd_parts", 0)
    has_warranty = deal_info.get("has_warranty", False) if deal_info else False

    # Already has active PM, still show but tag it
    if has_pm:
        return "Active PM (Upsell)"

    # Category 2: Warranty customers - they bought with warranty, good PM targets
    if has_warranty:
        return "Warranty Customer"

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
    2. Warranty Customer - bought with warranty, natural PM transition
    3. Parts Only, No Service - buys parts, does own service (pitch PM)
    4. Active/Lapsed Service - already uses SEC service (lock them in with PM)
    5. Active PM - already has PM (upsell or skip)

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
            elif category == "Warranty Customer":
                boost += 0.10  # Natural transition from warranty to PM
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

    df.drop(columns=["cust_upper", "HS Boost"], inplace=True)
    return df


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

    agg_dict = {
        "machines": ("VIN", "nunique"),
        "total_parts_value": ("Parts Value", "sum"),
        "total_labor_hrs": ("Labor Hrs", "sum"),
        "avg_hours": ("Eng Hrs", "mean"),
        "avg_score": ("Lead Score", "mean"),
        "max_score": ("Lead Score", "max"),
        "total_annual_pm": ("Annual PM Value", "sum"),
        "location": ("Location", "first"),
        "models": ("Model", lambda x: ", ".join(sorted(set(x)))),
        "categories": ("Category", lambda x: ", ".join(sorted(set(x)))),
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

    agg = scored_df.groupby("Customer").agg(**agg_dict).reset_index()

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
        ["Make", "Model", "Category", "Serial Number"],
        [quote_data.get("make", ""), quote_data.get("model", ""), quote_data.get("category", ""), quote_data.get("serial", "")],
        ["Machine Age (Years)", "Current Hours", "Travel Time (min)", ""],
        [str(quote_data.get("machine_age", "")), f"{quote_data.get('machine_hours', 0):,}", str(quote_data.get("travel_time", "")), ""],
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

    # Pricing
    elements.append(Paragraph("PM Pricing Breakdown", styles["SectionHead"]))
    parts_cost = quote_data.get("parts_cost", 0)
    labor_cost = quote_data.get("labor_cost", 0)
    travel_cost = quote_data.get("travel_cost", 0)
    total = quote_data.get("total_cost", 0)
    annual = quote_data.get("annual_pm_price", 0)

    price_data = [
        ["Description", "Amount"],
        ["Estimated Parts", f"${parts_cost:,.2f}"],
        ["Estimated Labor / Service", f"${labor_cost:,.2f}"],
        ["Travel", f"${travel_cost:,.2f}"],
        ["", ""],
        ["Total Estimated Cost", f"${total:,.2f}"],
        ["Annual PM Contract Price", f"${annual:,.2f}"],
    ]
    t = RLTable(price_data, colWidths=[4.5*inch, 2.3*inch])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5496")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTNAME", (0, -2), (-1, -1), "Helvetica-Bold"), ("FONTSIZE", (0, -2), (-1, -1), 12),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor(SEC_RED)),
        ("TOPPADDING", (0, 1), (-1, -1), 6), ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("LINEBELOW", (0, -3), (-1, -3), 0.5, colors.HexColor("#DDDDDD")),
        ("LINEBELOW", (0, -1), (-1, -1), 1.5, colors.HexColor(SEC_RED)),
    ]))
    elements.append(t)

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

def calculate_pm_cost(category, hours, age, travel_minutes, service_type):
    pricing = PM_PRICING.get(category, PM_PRICING["Other"])
    base_parts = pricing["parts_avg"]
    base_labor = pricing["labor_avg"]

    hours_per_year = hours / max(age, 1)
    if hours_per_year > 2000: hours_mult = 1.4
    elif hours_per_year > 1500: hours_mult = 1.2
    elif hours_per_year > 1000: hours_mult = 1.0
    elif hours_per_year > 500: hours_mult = 0.85
    else: hours_mult = 0.7

    if age > 15: age_mult = 1.25
    elif age > 10: age_mult = 1.15
    elif age > 5: age_mult = 1.05
    else: age_mult = 1.0

    parts_cost = round(base_parts * hours_mult * age_mult, 2)
    labor_cost = round(base_labor * hours_mult, 2)
    travel_cost = round((travel_minutes / 60) * 225 * 2, 2) if service_type == "Field" and travel_minutes > 0 else 0

    total_cost = parts_cost + labor_cost + travel_cost
    annual_price = round(total_cost * 1.15, 2)

    return {
        "parts_cost": parts_cost, "labor_cost": labor_cost, "travel_cost": travel_cost,
        "total_cost": total_cost, "annual_pm_price": annual_price,
        "margin_pct": round((1 - total_cost / annual_price) * 100, 1) if annual_price > 0 else 0,
        "hours_mult": hours_mult, "age_mult": age_mult,
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
            deal_history, pm_active = fetch_hubspot_deals()
            scored = enrich_leads_with_hubspot(scored, hs_companies, deal_history, pm_active)

        st.session_state.leads_df = scored
        st.session_state.procare_vins = procare_vins

        if scored.empty:
            st.warning("No external leads found after filtering out ProCare and internal machines.")
        else:
            # Summary metrics
            cust_agg = aggregate_customer_leads(scored)

            col1, col2, col3, col4, col5, col6 = st.columns(6)
            with col1:
                st.metric("Total Machines", len(scored))
            with col2:
                st.metric("Unique Customers", scored["Customer"].nunique())
            with col3:
                top_count = len(scored[scored["Tier"] == "Top"]) + len(scored[scored["Tier"] == "High"])
                st.metric("Top + High Leads", top_count)
            with col4:
                total_pm_opp = scored["Annual PM Value"].sum()
                st.metric("Total PM Opportunity", f"${total_pm_opp:,.0f}")
            with col5:
                if "In HubSpot" in scored.columns:
                    hs_count = scored[scored["In HubSpot"]]["Customer"].nunique()
                    st.metric("In HubSpot", f"{hs_count} customers")
                else:
                    st.metric("ProCare Excluded", len(procare_vins))
            with col6:
                if "CASE Class" in scored.columns:
                    attention_count = scored[scored["CASE Class"].isin(["Needs Attention", "At Risk", "Can't Lose Them", "About to Sleep"])]["Customer"].nunique()
                    st.metric("Need Attention", f"{attention_count} customers")
                else:
                    st.metric("Data Files", "Loaded")

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

            # HubSpot filters if available
            has_hs = "In HubSpot" in scored.columns
            if has_hs:
                fc1, fc2, fc3, fc4 = st.columns(4)
                with fc1:
                    lead_cats = sorted([c for c in scored["Lead Category"].unique() if c]) if "Lead Category" in scored.columns else []
                    filter_lead_cat = st.multiselect("Lead Category", lead_cats) if lead_cats else []
                with fc2:
                    hs_filter = st.radio("HubSpot Status", ["All Leads", "In HubSpot Only", "NOT in HubSpot"], horizontal=True)
                with fc3:
                    case_classes = sorted([c for c in scored["CASE Class"].unique() if c]) if "CASE Class" in scored.columns else []
                    filter_case = st.multiselect("CASE Classification", case_classes) if case_classes else []
                with fc4:
                    pm_filter = st.radio("PM Status", ["All", "No Active PM", "Has Active PM"], horizontal=True) if "Has PM" in scored.columns else "All"
            else:
                hs_filter = "All Leads"
                filter_case = []
                pm_filter = "All"
                filter_lead_cat = []

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
            if filter_lead_cat and "Lead Category" in display.columns:
                display = display[display["Lead Category"].isin(filter_lead_cat)]
            if filter_case and "CASE Class" in display.columns:
                display = display[display["CASE Class"].isin(filter_case)]
            if pm_filter == "No Active PM" and "Has PM" in display.columns:
                display = display[~display["Has PM"]]
            elif pm_filter == "Has Active PM" and "Has PM" in display.columns:
                display = display[display["Has PM"]]

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
                        # Build subtitle with HubSpot info
                        subtitle_parts = [row['location'], row['models']]
                        if "case_class" in row and row.get("case_class"):
                            subtitle_parts.insert(0, f"CASE: {row['case_class']}")
                        if "fleet" in row and row.get("fleet"):
                            subtitle_parts.insert(1, f"Fleet: {row['fleet']}")

                        with st.container():
                            c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 1])
                            with c1:
                                label = f"**{row['Customer']}**"
                                if "lead_category" in row and row.get("lead_category"):
                                    label += f"  &nbsp; `{row['lead_category']}`"
                                st.markdown(label, unsafe_allow_html=True)
                                st.caption(" | ".join(subtitle_parts))
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
                # Insert HubSpot columns after Customer
                hs_insert_cols = ["Lead Category", "CASE Class", "Fleet", "In HubSpot", "Has PM"]
                for i, col in enumerate(hs_insert_cols):
                    if col in display.columns:
                        show_cols.insert(3 + i, col)
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
                        "Has PM": st.column_config.CheckboxColumn("Has PM", default=False),
                        "Lead Category": st.column_config.TextColumn("Lead Category", width="medium"),
                        "CASE Class": st.column_config.TextColumn("CASE Class", width="small"),
                        "Fleet": st.column_config.TextColumn("Fleet Size", width="small"),
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

            # Lead Category breakdown (if HubSpot connected)
            if "Lead Category" in display.columns:
                cat_colors = {
                    "Parts Only, No Service": "#C8102E",
                    "Warranty Customer": "#E8601C",
                    "Active Service Customer": "#2F5496",
                    "Full Service (Lock In)": "#1B7340",
                    "Lapsed Service": "#F59E0B",
                    "No ProCare (In CRM)": "#6B7280",
                    "No ProCare (New Lead)": "#9CA3AF",
                    "Active PM (Upsell)": "#8B5CF6",
                }
                lc_data = display.groupby("Lead Category").agg(
                    customers=("Customer", "nunique"),
                    machines=("VIN", "nunique"),
                    pm_value=("Annual PM Value", "sum"),
                ).reset_index().sort_values("pm_value", ascending=True)
                fig = px.bar(lc_data, x="pm_value", y="Lead Category",
                             orientation="h", title="PM Opportunity by Lead Category",
                             labels={"pm_value": "Annual PM Value ($)", "Lead Category": ""},
                             hover_data=["customers", "machines"],
                             color="Lead Category", color_discrete_map=cat_colors)
                fig.update_layout(showlegend=False, height=350)
                st.plotly_chart(fig, use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
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
# TAB 2: PM CALCULATOR
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
        travel_time = st.number_input("Travel Time (minutes, one way)", min_value=0, max_value=480, value=0, step=15)

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
            category = st.selectbox("Machine Category", [""] + sorted(PM_PRICING.keys()))
    with col2:
        serial = st.text_input("Serial Number")
    with col3:
        pass

    col1, col2 = st.columns(2)
    with col1:
        machine_age = st.number_input("Machine Age (years)", min_value=0, max_value=50, value=0, step=1)
    with col2:
        machine_hours = st.number_input("Current Hours", min_value=0, max_value=100000, value=0, step=100)

    st.divider()
    notes = st.text_area("Notes", placeholder="Machine condition, special requirements, etc.", height=80)

    st.divider()
    can_calc = bool(category and make)

    if st.button("Calculate PM Estimate", type="primary", use_container_width=True, disabled=not can_calc):
        result = calculate_pm_cost(
            category=category, hours=machine_hours,
            age=machine_age if machine_age > 0 else 1,
            travel_minutes=travel_time if service_type == "Field" else 0,
            service_type=service_type,
        )
        st.session_state.current_quote = {
            "date": datetime.now().strftime("%m/%d/%Y"),
            "customer_name": customer_name, "branch": branch, "rep": rep,
            "service_type": service_type, "make": make,
            "model": model if model != "Other / Not Listed" else "",
            "category": category, "serial": serial,
            "machine_age": machine_age, "machine_hours": machine_hours,
            "travel_time": travel_time if service_type == "Field" else 0,
            "notes": notes, **result,
        }

    if st.session_state.current_quote:
        q = st.session_state.current_quote
        st.divider()
        st.subheader("PM Estimate")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-card"><div class="label">Parts</div><div class="value">${q["parts_cost"]:,.0f}</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><div class="label">Labor</div><div class="value">${q["labor_cost"]:,.0f}</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-card"><div class="label">Travel</div><div class="value">${q["travel_cost"]:,.0f}</div></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card"><div class="label">Annual PM Price</div><div class="value" style="color:#C8102E;">${q["annual_pm_price"]:,.0f}</div></div>', unsafe_allow_html=True)

        st.caption(f"Cost: ${q['total_cost']:,.2f} | Margin: {q['margin_pct']}% | Hours mult: {q['hours_mult']}x | Age factor: {q['age_mult']}x")

        col1, col2, col3 = st.columns(3)
        with col1:
            pdf_buf = generate_pdf(q)
            safe = (customer_name or "quote").replace(" ", "_")
            st.download_button("Download PDF Quote", data=pdf_buf, file_name=f"SEC_PM_{safe}_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)
        with col2:
            if st.button("Save Quote", use_container_width=True, type="secondary"):
                saved = save_quote_to_sheet(q)
                st.success("Quote saved to Google Sheets" if saved else "Quote saved locally (Sheets not connected)")
        with col3:
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

        if not filt.empty and "annual_pm_price" in filt.columns:
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("Quotes", len(filt))
            with c2: st.metric("Total Value", f"${filt['annual_pm_price'].sum():,.0f}")
            with c3: st.metric("Avg Value", f"${filt['annual_pm_price'].mean():,.0f}")
            with c4:
                if "margin_pct" in filt.columns:
                    st.metric("Avg Margin", f"{filt['margin_pct'].mean():.1f}%")

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
                if "category" in filt.columns:
                    cd = filt.groupby("category")["annual_pm_price"].sum().reset_index()
                    fig = px.bar(cd, x="category", y="annual_pm_price", title="By Category", color_discrete_sequence=["#2F5496"])
                    fig.update_layout(showlegend=False, xaxis_tickangle=-45)
                    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════
# TAB 4: PRICING REFERENCE
# ═══════════════════════════════════════════════════════════
with tab_pricing:
    st.subheader("PM Pricing by Machine Category")
    st.caption("Base annual pricing at ~1,000 hours/year. Adjusts for actual hours and machine age.")

    rows = []
    for cat, p in sorted(PM_PRICING.items(), key=lambda x: -x[1]["annual"]):
        rows.append({"Category": cat, "Annual PM": f"${p['annual']:,}", "Parts (est)": f"${p['parts_avg']:,}", "Labor (est)": f"${p['labor_avg']:,}"})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Models by Category")
    for cat in sorted(PM_PRICING.keys()):
        if cat == "Other":
            continue
        makes_in = []
        for mk, mk_cats in MAKES_MODELS.items():
            if cat in mk_cats:
                makes_in.append(f"**{mk}**: {', '.join(mk_cats[cat])}")
        if makes_in:
            with st.expander(f"{cat} — ${PM_PRICING[cat]['annual']:,}/yr"):
                for line in makes_in:
                    st.markdown(line)

    st.caption("Pricing from SEC service history + PM contract data. Updates coming from Jarred's averages.")

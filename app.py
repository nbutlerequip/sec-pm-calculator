import streamlit as st
import pandas as pd
import io
import base64
from datetime import datetime, date
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
    st.caption("Upload the Case maintenance alerts and ProCare stops files. The algorithm scores every machine without ProCare coverage and ranks the best opportunities.")

    col1, col2 = st.columns(2)
    with col1:
        alerts_file = st.file_uploader("Maintenance Alerts (.xlsx)", type=["xlsx"], key="alerts_upload")
    with col2:
        procare_file = st.file_uploader("ProCare Stops (.xlsx)", type=["xlsx"], key="procare_upload")

    if alerts_file:
        alerts_df = parse_maintenance_alerts(alerts_file)
        procare_vins = set()
        if procare_file:
            procare_vins = parse_procare_stops(procare_file)
            st.session_state.procare_vins = procare_vins

        scored = score_leads(alerts_df, procare_vins)
        st.session_state.leads_df = scored

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
                if procare_file:
                    st.metric("ProCare Machines", len(procare_vins))
                else:
                    st.metric("ProCare File", "Not uploaded")

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

            # Apply filters
            display = scored.copy()
            if filter_tier:
                display = display[display["Tier"].isin(filter_tier)]
            if filter_location:
                display = display[display["Location"].isin(filter_location)]
            if filter_category:
                display = display[display["Category"].isin(filter_category)]
            display = display[display["Lead Score"] >= min_score]

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
                        css_class = "lead-top" if tier == "Top" else ("lead-high" if tier == "High" else "lead-med")
                        tier_color = "#C8102E" if tier == "Top" else ("#F59E0B" if tier == "High" else ("#3B82F6" if tier == "Medium" else "#6B7280"))

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
                    },
                )

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

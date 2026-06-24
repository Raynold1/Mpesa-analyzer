import io
import os
import sys
import pandas as pd
import streamlit as st
import altair as alt

# Prevent running with plain python; instruct to use `streamlit run`
if __name__ == "__main__" and "streamlit" not in " ".join(sys.argv):
    print("Please run this app with Streamlit:  streamlit run streamlit_app.py")
    sys.exit(0)

# Set page configuration
st.set_page_config(
    page_title="Mpesa Analyzer - Premium Insights",
    page_icon="📥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── MOBILE-FIRST CSS + TOGGLEABLE NAV BAR ──────────────────────────────────
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

    /* ── Reset & Base ── */
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        background: #f0f4f8;
        color: #1e293b;
        overflow-x: hidden;
    }

    /* Hide Streamlit chrome */
    [data-testid="stHeader"] { display: none !important; }
    [data-testid="stSidebar"] { display: none !important; }
    section[data-testid="stSidebar"] { display: none !important; }
    #MainMenu { display: none !important; }
    footer { display: none !important; }
    .stDeployButton { display: none !important; }

    /* Push content below the fixed navbar */
    [data-testid="stAppViewContainer"] > .main {
        padding-top: 72px !important;
    }
    .block-container {
        padding-top: 1.5rem !important;
        padding-left: 1.25rem !important;
        padding-right: 1.25rem !important;
        max-width: 1200px !important;
    }

    /* ── TOP NAV BAR ── */
    .mpesa-navbar {
        position: fixed;
        top: 0; left: 0; right: 0;
        z-index: 9999;
        background: linear-gradient(135deg, #0a2e1f 0%, #0f4c35 50%, #1a5c40 100%);
        box-shadow: 0 2px 20px rgba(0,0,0,0.25);
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 1.5rem;
        height: 64px;
    }

    .mpesa-brand {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        text-decoration: none;
    }
    .mpesa-brand-icon {
        font-size: 1.6rem;
        line-height: 1;
    }
    .mpesa-brand-text {
        display: flex;
        flex-direction: column;
    }
    .mpesa-brand-name {
        font-size: 1.1rem;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: -0.02em;
        line-height: 1.1;
    }
    .mpesa-brand-tagline {
        font-size: 0.68rem;
        color: #6ee7b7;
        font-weight: 500;
        letter-spacing: 0.04em;
    }

    /* Desktop nav links */
    .mpesa-nav-links {
        display: flex;
        align-items: center;
        gap: 0.25rem;
        list-style: none;
    }
    .mpesa-nav-links a {
        display: flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.45rem 0.9rem;
        border-radius: 8px;
        color: rgba(255,255,255,0.75);
        font-size: 0.85rem;
        font-weight: 500;
        text-decoration: none;
        transition: all 0.2s ease;
        white-space: nowrap;
        cursor: pointer;
        border: none;
        background: transparent;
    }
    .mpesa-nav-links a:hover {
        background: rgba(255,255,255,0.12);
        color: #ffffff;
    }
    .mpesa-nav-links a.active {
        background: rgba(16,185,129,0.25);
        color: #6ee7b7;
        font-weight: 600;
        border: 1px solid rgba(16,185,129,0.3);
    }

    /* Hamburger button */
    .mpesa-hamburger {
        display: none;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        gap: 5px;
        width: 40px;
        height: 40px;
        background: rgba(255,255,255,0.1);
        border: 1px solid rgba(255,255,255,0.2);
        border-radius: 8px;
        cursor: pointer;
        transition: background 0.2s;
        z-index: 10001;
        padding: 0;
    }
    .mpesa-hamburger:hover { background: rgba(255,255,255,0.2); }
    .mpesa-hamburger span {
        display: block;
        width: 20px;
        height: 2px;
        background: #ffffff;
        border-radius: 2px;
        transition: all 0.3s ease;
    }
    .mpesa-hamburger.open span:nth-child(1) { transform: translateY(7px) rotate(45deg); }
    .mpesa-hamburger.open span:nth-child(2) { opacity: 0; transform: scaleX(0); }
    .mpesa-hamburger.open span:nth-child(3) { transform: translateY(-7px) rotate(-45deg); }

    /* Mobile drawer overlay */
    .mpesa-overlay {
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.5);
        z-index: 9998;
        backdrop-filter: blur(2px);
    }
    .mpesa-overlay.open { display: block; }

    /* Mobile drawer */
    .mpesa-drawer {
        display: none;
        position: fixed;
        top: 0; right: 0;
        width: min(280px, 85vw);
        height: 100vh;
        background: linear-gradient(180deg, #0a2e1f 0%, #0f4c35 100%);
        z-index: 9999;
        padding: 5rem 1.25rem 2rem;
        transform: translateX(100%);
        transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: -8px 0 32px rgba(0,0,0,0.3);
        overflow-y: auto;
    }
    .mpesa-drawer.open { transform: translateX(0); }

    .mpesa-drawer-links {
        list-style: none;
        display: flex;
        flex-direction: column;
        gap: 0.5rem;
    }
    .mpesa-drawer-links a {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        padding: 0.85rem 1rem;
        border-radius: 12px;
        color: rgba(255,255,255,0.8);
        font-size: 0.95rem;
        font-weight: 500;
        text-decoration: none;
        transition: all 0.2s ease;
        border: 1px solid transparent;
    }
    .mpesa-drawer-links a:hover {
        background: rgba(255,255,255,0.1);
        color: #ffffff;
        border-color: rgba(255,255,255,0.1);
    }
    .mpesa-drawer-links a.active {
        background: rgba(16,185,129,0.2);
        color: #6ee7b7;
        font-weight: 600;
        border-color: rgba(16,185,129,0.35);
    }
    .mpesa-drawer-links .nav-icon { font-size: 1.2rem; }

    .mpesa-drawer-divider {
        border: none;
        border-top: 1px solid rgba(255,255,255,0.1);
        margin: 1.25rem 0;
    }
    .mpesa-drawer-label {
        font-size: 0.7rem;
        font-weight: 700;
        color: rgba(255,255,255,0.4);
        text-transform: uppercase;
        letter-spacing: 0.1em;
        padding: 0 1rem;
        margin-bottom: 0.5rem;
    }

    /* ── MAIN CONTENT CARD ── */
    .main-title {
        font-size: clamp(1.5rem, 4vw, 2.25rem);
        font-weight: 800;
        color: #0f172a;
        letter-spacing: -0.03em;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #047857 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.25rem;
        line-height: 1.2;
    }
    .main-subtitle {
        font-size: 1rem;
        color: #64748b;
        margin-bottom: 1.75rem;
        font-weight: 400;
    }

    /* ── KPI CARDS ── */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 1rem;
        margin-bottom: 1.75rem;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 1.15rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        transition: all 0.25s cubic-bezier(0.4,0,0.2,1);
        display: flex;
        flex-direction: column;
        min-height: 115px;
    }
    .kpi-card:hover { transform: translateY(-3px); box-shadow: 0 10px 24px rgba(0,0,0,0.08); }
    .kpi-card-indigo:hover { border-color: #6366f1; }
    .kpi-card-amber:hover  { border-color: #f59e0b; }
    .kpi-card-emerald:hover{ border-color: #10b981; }
    .kpi-card-rose:hover   { border-color: #f43f5e; }
    .kpi-card-teal:hover   { border-color: #0d9488; }
    .kpi-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.4rem; }
    .kpi-label  { font-size: 0.75rem; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-icon-container { width: 30px; height: 30px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 1rem; }
    .kpi-value  { font-size: clamp(1.2rem,2.5vw,1.55rem); font-weight: 700; color: #0f172a; line-height: 1.2; margin-top: 0.2rem; }
    .kpi-badge  { font-size: 0.72rem; font-weight: 600; padding: 0.15rem 0.5rem; border-radius: 9999px; display: inline-block; margin-top: 0.6rem; width: fit-content; }

    .bg-indigo-50 { background-color: #e0e7ff; color: #4338ca; }
    .bg-amber-50  { background-color: #fffbeb; color: #b45309; }
    .bg-emerald-50{ background-color: #ecfdf5; color: #047857; }
    .bg-rose-50   { background-color: #fff1f2; color: #b91c1c; }
    .bg-teal-50   { background-color: #f0fdfa; color: #0f766e; }
    .badge-indigo { background-color: #e0e7ff; color: #3730a3; }
    .badge-amber  { background-color: #fef3c7; color: #92400e; }
    .badge-emerald{ background-color: #d1fae5; color: #065f46; }
    .badge-rose   { background-color: #ffe4e6; color: #991b1b; }
    .badge-teal   { background-color: #ccfbf1; color: #115e59; }

    /* ── WELCOME CARD ── */
    .welcome-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 2rem;
        box-shadow: 0 4px 16px rgba(0,0,0,0.04);
        margin-bottom: 1.5rem;
    }
    .welcome-step {
        display: flex;
        align-items: flex-start;
        gap: 1rem;
        margin-bottom: 1.25rem;
        padding: 0.85rem;
        border-radius: 12px;
        transition: all 0.2s;
    }
    .welcome-step:hover { background: #f8fafc; transform: translateX(4px); }
    .welcome-icon {
        background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        color: #059669;
        border-radius: 10px;
        width: 34px; height: 34px;
        display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 1rem; flex-shrink: 0;
    }
    .welcome-text { font-size: 0.9rem; color: #475569; line-height: 1.5; }

    /* ── FILE UPLOADER ── */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 16px !important;
        padding: 1.5rem !important;
        background: #ffffff !important;
        transition: all 0.25s ease !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #10b981 !important;
        background: #f0fdf4 !important;
    }
    div[data-testid="stFileUploader"] button {
        background-color: #10b981 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }

    /* ── BUTTONS ── */
    div.stButton button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.6rem 1.5rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 6px -1px rgba(16,185,129,0.2) !important;
        transition: all 0.2s cubic-bezier(0.4,0,0.2,1) !important;
    }
    div.stButton button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 15px -3px rgba(16,185,129,0.25) !important;
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
    }
    div.stDownloadButton button {
        background: #ffffff !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
        width: 100% !important;
    }
    div.stDownloadButton button:hover {
        border-color: #10b981 !important;
        color: #059669 !important;
        transform: translateY(-1px) !important;
    }

    /* ── INPUTS ── */
    div[data-testid="stSelectbox"] > div { border-radius: 12px !important; }
    div[data-testid="stTextInput"] input {
        border-radius: 12px !important;
        border: 1px solid #cbd5e1 !important;
        padding: 0.5rem 0.75rem !important;
    }
    div[data-testid="stTextInput"] input:focus {
        border-color: #10b981 !important;
        box-shadow: 0 0 0 3px rgba(16,185,129,0.15) !important;
    }

    /* ── DATAFRAME ── */
    div[data-testid="stDataFrame"] {
        border-radius: 16px !important;
        overflow: hidden !important;
        border: 1px solid #e2e8f0 !important;
    }

    /* ── EXPANDER ── */
    div[data-testid="stExpander"] {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 16px !important;
        overflow: hidden !important;
    }

    /* ── ALERTS ── */
    .stAlert { border-radius: 14px !important; }

    /* ── SECTION HEADERS ── */
    h3 {
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        margin-top: 1.25rem !important;
        margin-bottom: 0.75rem !important;
        letter-spacing: -0.02em !important;
    }

    /* ── FOOTER ── */
    .app-footer {
        text-align: center;
        padding: 2rem 0;
        color: #94a3b8;
        font-size: 0.82rem;
        border-top: 1px solid #e2e8f0;
        margin-top: 4rem;
        font-weight: 500;
    }

    /* ── RESPONSIVE BREAKPOINTS ── */
    @media (max-width: 768px) {
        .mpesa-nav-links { display: none !important; }
        .mpesa-hamburger { display: flex !important; }
        .mpesa-drawer    { display: block !important; }

        .block-container {
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
        }
        .kpi-grid {
            grid-template-columns: repeat(2, 1fr) !important;
            gap: 0.75rem !important;
        }
        .welcome-card { padding: 1.25rem; }
    }
    @media (max-width: 480px) {
        .kpi-grid { grid-template-columns: 1fr 1fr !important; }
        .mpesa-navbar { padding: 0 1rem; }
    }
    </style>
""", unsafe_allow_html=True)

# ─── NAV BAR HTML + JS ───────────────────────────────────────────────────────
NAV_ITEMS = [
    ("\U0001f3e0", "Home & Upload",    "home"),
    ("\U0001f4c8", "Visual Dashboard", "dashboard"),
    ("\U0001f4c1", "Merged Statement", "merged"),
    ("\U0001f4ca", "Monthly Pivot",    "pivot"),
    ("\U0001f5a8", "Print Report",    "report"),
]

# ─── ROUTING & STATE MANAGEMENT ──────────────────────────────────────────────
# Read initial page from query parameters or session state
initial_page = "home"
try:
    qp = st.query_params.get("nav_page", None)
    if qp and qp in ["home", "dashboard", "merged", "pivot", "report"]:
        initial_page = qp
except Exception:
    pass

if "page" not in st.session_state:
    st.session_state["page"] = initial_page

# Render a hidden radio button in the sidebar to act as our reactive router.
# The sidebar is hidden via CSS, so this is completely invisible to users,
# but remains in the DOM where our custom JavaScript can click it.
with st.sidebar:
    if "hidden_nav_radio" not in st.session_state:
        st.session_state["hidden_nav_radio"] = st.session_state["page"]
        
    current_page = st.radio(
        "Navigation Router",
        options=["home", "dashboard", "merged", "pivot", "report"],
        key="hidden_nav_radio",
        label_visibility="collapsed"
    )
    # Synchronize the main page state
    st.session_state["page"] = current_page

# Build nav link HTML for desktop & drawer using data-nav-page attributes
def build_nav_link(icon, label, key, mobile=False):
    active_cls = "active" if current_page == key else ""
    if mobile:
        return (
            '<li>'
            f'<a class="{active_cls}" data-nav-page="{key}" href="#">'
            f'<span class="nav-icon">{icon}</span><span>{label}</span>'
            '</a></li>'
        )
    else:
        return (
            f'<a class="{active_cls}" data-nav-page="{key}" href="#">'
            f'<span class="nav-icon">{icon}</span><span>{label}</span>'
            '</a>'
        )

desktop_links = "".join(build_nav_link(i, l, k, False) for i, l, k in NAV_ITEMS)
drawer_links  = "".join(build_nav_link(i, l, k, True)  for i, l, k in NAV_ITEMS)

# HTML for top navbar, hamburger toggle, and mobile drawer
nav_html = (
    # Hidden checkbox — the "toggle state"
    '<input type="checkbox" id="navToggle" style="display:none;">'
    '<nav class="mpesa-navbar" id="mpesaNavbar">'
    '<div class="mpesa-brand">'
    '<span class="mpesa-brand-icon">\U0001f4e5</span>'
    '<div class="mpesa-brand-text">'
    '<span class="mpesa-brand-name">Mpesa Analyzer</span>'
    '<span class="mpesa-brand-tagline">Financial Insights</span>'
    '</div></div>'
    f'<ul class="mpesa-nav-links" id="desktopNav">{desktop_links}</ul>'
    # The hamburger label toggles #navToggle checkbox
    '<label for="navToggle" class="mpesa-hamburger" id="hamburgerBtn" aria-label="Toggle menu">'
    '<span></span><span></span><span></span>'
    '</label>'
    '</nav>'
    '<label for="navToggle" class="mpesa-overlay" id="navOverlay"></label>'
    '<div class="mpesa-drawer" id="navDrawer">'
    '<p class="mpesa-drawer-label">Navigation</p>'
    f'<ul class="mpesa-drawer-links">{drawer_links}</ul>'
    '<hr class="mpesa-drawer-divider">'
    '<p style="color:rgba(255,255,255,0.4);font-size:0.75rem;padding:0 1rem;">Mpesa Analyzer &copy; 2026</p>'
    '</div>'
)

st.markdown(nav_html, unsafe_allow_html=True)

# Inject client-side JS event delegation to handle clicks on custom navbar items
# and trigger the hidden st.radio button without reloading the page.
# Using event delegation on parentDoc.body ensures it persists across all Streamlit reruns,
# and navListenerBound ensures it's only bound once.
components_js = """
<script>
(function() {
    const parentDoc = window.parent ? window.parent.document : document;
    if (!parentDoc) return;
    
    if (!parentDoc.body.dataset.navListenerBound) {
        parentDoc.body.dataset.navListenerBound = "true";
        
        parentDoc.body.addEventListener('click', function(e) {
            const navLink = e.target.closest('[data-nav-page]');
            if (!navLink) return;
            
            e.preventDefault();
            const pageKey = navLink.getAttribute('data-nav-page');
            
            // Find the hidden radio button in the parent document
            let targetInput = null;
            const pages = ["home", "dashboard", "merged", "pivot", "report"];
            const pageIndex = pages.indexOf(pageKey);
            
            if (pageIndex !== -1) {
                // Try searching inside the sidebar first (highly precise)
                const sidebar = parentDoc.querySelector('[data-testid="stSidebar"]') || parentDoc.querySelector('section[data-testid="stSidebar"]');
                if (sidebar) {
                    const radio = sidebar.querySelector('div[data-testid="stRadio"]');
                    if (radio) {
                        const inputs = radio.querySelectorAll('input[type="radio"]');
                        if (inputs[pageIndex]) {
                            targetInput = inputs[pageIndex];
                        }
                    }
                }
                
                // Fallback: search all radio containers on the page
                if (!targetInput) {
                    const radioContainers = parentDoc.querySelectorAll('div[data-testid="stRadio"]');
                    for (const container of radioContainers) {
                        const inputs = container.querySelectorAll('input[type="radio"]');
                        if (inputs.length === 5 && inputs[pageIndex]) {
                            targetInput = inputs[pageIndex];
                            break;
                        }
                    }
                }
            }
            
            if (targetInput) {
                targetInput.click();
                
                // Update URL query parameters in the address bar without page reload
                const parentWin = window.parent || window;
                const newUrl = parentWin.location.protocol + "//" + parentWin.location.host + parentWin.location.pathname + "?nav_page=" + pageKey;
                parentWin.history.pushState({ path: newUrl }, '', newUrl);
                
                // Uncheck the hamburger checkbox to close mobile drawer on transition
                const navToggle = parentDoc.getElementById('navToggle');
                if (navToggle) {
                    navToggle.checked = false;
                }
            } else {
                console.error("Navigation target radio button not found for page: " + pageKey + " (index: " + pageIndex + ")");
            }
        });
    }
})();
</script>
"""
st.components.v1.html(components_js, height=0, width=0)

# CSS for checkbox-hack hamburger — uses :has() which works across DOM boundaries
st.markdown("""
<style>
/* Drawer slides in when navToggle checkbox is checked */
body:has(#navToggle:checked) .mpesa-drawer {
    transform: translateX(0) !important;
}
/* Overlay shows when drawer is open */
body:has(#navToggle:checked) .mpesa-overlay {
    display: block !important;
}
/* Hamburger animates when checked */
body:has(#navToggle:checked) .mpesa-hamburger span:nth-child(1) {
    transform: translateY(7px) rotate(45deg) !important;
}
body:has(#navToggle:checked) .mpesa-hamburger span:nth-child(2) {
    opacity: 0 !important;
    transform: scaleX(0) !important;
}
body:has(#navToggle:checked) .mpesa-hamburger span:nth-child(3) {
    transform: translateY(-7px) rotate(-45deg) !important;
}
</style>
""", unsafe_allow_html=True)

menu_page = st.session_state["page"]

# Map page key to display name for legacy code
PAGE_LABELS = {
    "home":      "🏠 Home & Upload",
    "dashboard": "📈 Visual Dashboard",
    "merged":    "📁 Merged Statement",
    "pivot":     "📊 Monthly Pivot",
    "report":    "🖨️ Printable Report",
}
menu = PAGE_LABELS.get(menu_page, "🏠 Home & Upload")


# ─── HELPER ──────────────────────────────────────────────────────────────────
def find_col(df, target):
    for c in df.columns:
        if c.lower() == target.lower():
            return c
    for c in df.columns:
        if target.lower() in c.lower():
            return c
    return None


# ─── SESSION STATE ────────────────────────────────────────────────────────────
if "uploaded_file_bytes" not in st.session_state:
    st.session_state["uploaded_file_bytes"] = None
if "uploaded_file_name" not in st.session_state:
    st.session_state["uploaded_file_name"] = None
if "required_cols" not in st.session_state:
    st.session_state["required_cols"] = "Paid In, Withdrawn, Balance"
if "case_insensitive" not in st.session_state:
    st.session_state["case_insensitive"] = True
if "date_col" not in st.session_state:
    st.session_state["date_col"] = ""


# ─── EXCEL PROCESSING ────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Processing statements…")
def process_statement(file_bytes, required_cols_str, case_insensitive_flag):
    try:
        sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
        required_columns = [c.strip() for c in required_cols_str.split(",") if c.strip()]
        if not required_columns:
            return None, "Please specify at least one required column in settings.", {}, sheets, []
        merged_dfs, included_sheets, skipped = [], [], {}
        req_lower = [c.lower() for c in required_columns]
        for sheet_name, df in sheets.items():
            try:
                cols = df.columns.tolist()
                if case_insensitive_flag:
                    cols_lower = [c.lower() for c in cols]
                    has_all = all(r in cols_lower for r in req_lower)
                else:
                    has_all = all(r in cols for r in required_columns)
                if has_all:
                    merged_dfs.append(df)
                    included_sheets.append(sheet_name)
                else:
                    missing = [r for r in (req_lower if case_insensitive_flag else required_columns)
                               if r not in (cols_lower if case_insensitive_flag else cols)]
                    skipped[sheet_name] = missing
            except Exception as e:
                skipped[sheet_name] = f"Read error: {e}"
        if not merged_dfs:
            return None, f"No sheets contained all required columns: {', '.join(required_columns)}", skipped, sheets, []
        return pd.concat(merged_dfs, ignore_index=True), None, skipped, sheets, included_sheets
    except Exception as exc:
        return None, f"Excel file read error: {exc}", {}, {}, []


# ─── RESOLVE DATA ─────────────────────────────────────────────────────────────
final_df = error_msg = paid_col = withdrawn_col = balance_col = None
skipped_sheets, all_sheets, included_sheets = {}, {}, []

if st.session_state["uploaded_file_bytes"] is not None:
    final_df, error_msg, skipped_sheets, all_sheets, included_sheets = process_statement(
        st.session_state["uploaded_file_bytes"],
        st.session_state["required_cols"],
        st.session_state["case_insensitive"]
    )
    if final_df is not None:
        paid_col      = find_col(final_df, "Paid In")
        withdrawn_col = find_col(final_df, "Withdrawn")
        balance_col   = find_col(final_df, "Balance")
        for col in [paid_col, withdrawn_col, balance_col]:
            if col:
                final_df[col] = pd.to_numeric(final_df[col], errors="coerce")

        # Date column picker — shown inline on non-home pages
        candidate_date_cols = [
            c for c in final_df.columns
            if any(k in c.lower() for k in ("date", "time", "completion", "timestamp"))
        ]
        saved_date = st.session_state["date_col"]
        if candidate_date_cols and saved_date not in candidate_date_cols:
            st.session_state["date_col"] = candidate_date_cols[0]


# ─── KPI HELPER ───────────────────────────────────────────────────────────────
def render_kpis():
    if final_df is None:
        return
    total_transactions = len(final_df)
    merged_count = len(included_sheets)
    total_sheets = len(all_sheets)
    sum_paid      = final_df[paid_col].sum()      if paid_col      else 0.0
    sum_withdrawn = final_df[withdrawn_col].sum() if withdrawn_col else 0.0
    net_flow      = sum_paid - sum_withdrawn

    sum_paid_str      = f"KES {sum_paid:,.0f}"
    sum_withdrawn_str = f"KES {sum_withdrawn:,.0f}"
    net_flow_str      = f"KES {abs(net_flow):,.0f}"
    if net_flow < 0:
        net_flow_str = f"- {net_flow_str}"

    net_bg    = "bg-teal-50"    if net_flow >= 0 else "bg-rose-50"
    net_badge = "badge-teal"    if net_flow >= 0 else "badge-rose"
    net_card  = "kpi-card-teal" if net_flow >= 0 else "kpi-card-rose"
    net_trend = "Positive Flow" if net_flow >= 0 else "Negative Flow"

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card kpi-card-indigo">
            <div class="kpi-header"><span class="kpi-label">Sheets Merged</span><div class="kpi-icon-container bg-indigo-50">📁</div></div>
            <div class="kpi-value">{merged_count}/{total_sheets}</div>
            <div class="kpi-badge badge-indigo">Workbook</div>
        </div>
        <div class="kpi-card kpi-card-amber">
            <div class="kpi-header"><span class="kpi-label">Transactions</span><div class="kpi-icon-container bg-amber-50">🔢</div></div>
            <div class="kpi-value">{total_transactions:,}</div>
            <div class="kpi-badge badge-amber">Total Rows</div>
        </div>
        <div class="kpi-card kpi-card-emerald">
            <div class="kpi-header"><span class="kpi-label">Total Deposits</span><div class="kpi-icon-container bg-emerald-50">📈</div></div>
            <div class="kpi-value">{sum_paid_str}</div>
            <div class="kpi-badge badge-emerald">Cash Inflow</div>
        </div>
        <div class="kpi-card kpi-card-rose">
            <div class="kpi-header"><span class="kpi-label">Total Withdrawn</span><div class="kpi-icon-container bg-rose-50">📉</div></div>
            <div class="kpi-value">{sum_withdrawn_str}</div>
            <div class="kpi-badge badge-rose">Cash Outflow</div>
        </div>
        <div class="kpi-card {net_card}">
            <div class="kpi-header"><span class="kpi-label">Net Cash Flow</span><div class="kpi-icon-container {net_bg}">💰</div></div>
            <div class="kpi-value">{net_flow_str}</div>
            <div class="kpi-badge {net_badge}">{net_trend}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# ─── GLOBAL HEADER (non-home pages) ──────────────────────────────────────────
def render_page_header(title, subtitle):
    st.markdown(f'<div class="main-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="main-subtitle">{subtitle}</div>', unsafe_allow_html=True)


# ─── DATE COLUMN PICKER (inline widget, shown when needed) ───────────────────
def render_date_picker(label="⚙️ Grouping: Date Column"):
    if final_df is None:
        return
    candidate_date_cols = [
        c for c in final_df.columns
        if any(k in c.lower() for k in ("date", "time", "completion", "timestamp"))
    ]
    if candidate_date_cols:
        saved = st.session_state["date_col"]
        if saved not in candidate_date_cols:
            saved = candidate_date_cols[0]
        sel = st.selectbox(label, candidate_date_cols,
                           index=candidate_date_cols.index(saved),
                           help="Column used to group transactions by month.")
        st.session_state["date_col"] = sel
    else:
        val = st.text_input("Specify Date Column name",
                            value=st.session_state["date_col"])
        st.session_state["date_col"] = val


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — HOME & UPLOAD
# ═══════════════════════════════════════════════════════════════════════════════
if menu_page == "home":
    render_page_header("📥 Mpesa Statement Analyzer",
                       "Combine, summarize, and visualize multi-sheet Mpesa data instantly.")

    if st.session_state["uploaded_file_bytes"] is None:
        uploaded_file = st.file_uploader(
            "Upload Excel Workbook (.xlsx, .xls)",
            type=["xlsx", "xls"],
            help="Upload the Excel file containing one or more client Mpesa sheets."
        )
        if uploaded_file is not None:
            st.session_state["uploaded_file_bytes"] = uploaded_file.getvalue()
            st.session_state["uploaded_file_name"]  = uploaded_file.name
            st.rerun()
    else:
        st.success(f"✅ **{st.session_state['uploaded_file_name']}** is loaded and ready.")
        st.info("👆 Use the navigation bar above to switch between dashboard views.")
        if st.button("🗑️ Clear & Upload Another File"):
            st.session_state["uploaded_file_bytes"] = None
            st.session_state["uploaded_file_name"]  = None
            st.session_state["date_col"] = ""
            st.rerun()

    st.markdown("""
    <div class="welcome-card">
        <h3 style="margin-top:0;margin-bottom:1rem;">Get Started in Seconds</h3>
        <div class="welcome-step">
            <div class="welcome-icon">1</div>
            <div class="welcome-text"><strong>Upload Statement:</strong> Drag &amp; drop your client's Excel workbook above.</div>
        </div>
        <div class="welcome-step">
            <div class="welcome-icon">2</div>
            <div class="welcome-text"><strong>Confirm Settings:</strong> Expand configuration below to customize column matching.</div>
        </div>
        <div class="welcome-step">
            <div class="welcome-icon">3</div>
            <div class="welcome-text"><strong>View Insights:</strong> Tap any item in the navigation bar to explore metrics, charts, and reports.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("⚙️ Configuration Settings", expanded=True):
        req_cols = st.text_input(
            "Required columns (comma-separated)",
            value=st.session_state["required_cols"],
            help="Only sheets containing ALL these columns will be merged."
        )
        if req_cols != st.session_state["required_cols"]:
            st.session_state["required_cols"] = req_cols
            st.rerun()
        case_ins = st.checkbox(
            "Case-insensitive column matching",
            value=st.session_state["case_insensitive"]
        )
        if case_ins != st.session_state["case_insensitive"]:
            st.session_state["case_insensitive"] = case_ins
            st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — VISUAL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
elif menu_page == "dashboard":
    if st.session_state["uploaded_file_bytes"] is None:
        render_page_header("📈 Visual Dashboard", "Upload a workbook first to see charts.")
        st.warning("⚠️ No workbook uploaded. Go to **Home & Upload** to get started.")
    elif error_msg:
        render_page_header("📈 Visual Dashboard", "")
        st.error(f"❌ {error_msg}")
    elif final_df is not None:
        render_page_header(f"📈 {st.session_state['uploaded_file_name']}",
                           "Parsed transaction metrics and summary insights")
        render_kpis()
        render_date_picker()
        date_col = st.session_state["date_col"]
        st.markdown("---")

        if not date_col:
            st.info("ℹ️ Select a date column above to compute monthly trends.")
        elif paid_col is None or withdrawn_col is None:
            st.warning("⚠️ 'Paid In' and 'Withdrawn' columns are required for charts.")
        else:
            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
            if final_df["_parsed_date"].notna().sum() == 0:
                st.error(f"❌ Could not parse dates from **'{date_col}'**.")
            else:
                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")
                final_df["_MonthSort"]  = final_df["_parsed_date"].dt.strftime("%Y-%m")
                dash_pivot = (
                    final_df.groupby(["_MonthSort", "_MonthLabel"])[[paid_col, withdrawn_col]]
                    .sum(min_count=1).reset_index()
                    .rename(columns={"_MonthLabel": "Month", paid_col: "Deposits", withdrawn_col: "Withdrawals"})
                )
                dash_pivot = dash_pivot.sort_values("_MonthSort").reset_index(drop=True)
                melt_df = dash_pivot.melt(
                    id_vars=["Month", "_MonthSort"],
                    value_vars=["Deposits", "Withdrawals"],
                    var_name="Transaction Type", value_name="Amount"
                )

                st.markdown("##### Monthly Cash Flow — Deposits vs Withdrawals")
                cash_chart = alt.Chart(melt_df).mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5).encode(
                    x=alt.X('Month:N', sort=alt.SortField(field='_MonthSort', order='ascending'), title='Month'),
                    y=alt.Y('Amount:Q', title='Amount (KES)'),
                    color=alt.Color('Transaction Type:N', scale=alt.Scale(
                        domain=['Deposits', 'Withdrawals'], range=['#10b981', '#f43f5e'])),
                    xOffset='Transaction Type:N',
                    tooltip=[alt.Tooltip('Month:N'), alt.Tooltip('Transaction Type:N'),
                             alt.Tooltip('Amount:Q', format=",.0f")]
                ).properties(height=320).configure_view(strokeOpacity=0)
                st.altair_chart(cash_chart, use_container_width=True)

                st.markdown("---")
                st.markdown("##### Net Monthly Savings")
                dash_pivot["Net Savings"] = dash_pivot["Deposits"].fillna(0) - dash_pivot["Withdrawals"].fillna(0)
                net_chart = alt.Chart(dash_pivot).mark_area(
                    line={'color': '#0d9488', 'width': 2.5},
                    color=alt.Gradient(gradient='linear',
                        stops=[alt.GradientStop(color='#ccfbf1', offset=0),
                               alt.GradientStop(color='#ffffff', offset=1)],
                        x1=1, y1=1, x2=1, y2=0)
                ).encode(
                    x=alt.X('Month:N', sort=alt.SortField(field='_MonthSort', order='ascending')),
                    y=alt.Y('Net Savings:Q', title='Net Amount (KES)'),
                    tooltip=[alt.Tooltip('Month:N'), alt.Tooltip('Net Savings:Q', format=",.0f")]
                ).properties(height=280)
                st.altair_chart(net_chart, use_container_width=True)

                st.markdown("---")
                st.markdown("##### Transaction Distribution")
                c_vol, c_val = st.columns(2)
                with c_vol:
                    cnt_df = pd.DataFrame({
                        "Category": ["Deposits", "Withdrawals"],
                        "Count": [final_df[paid_col].notna().sum(), final_df[withdrawn_col].notna().sum()]
                    })
                    donut1 = alt.Chart(cnt_df).mark_arc(innerRadius=55).encode(
                        theta=alt.Theta("Count:Q"),
                        color=alt.Color("Category:N", scale=alt.Scale(range=['#10b981', '#f43f5e'])),
                        tooltip=["Category", "Count"]
                    ).properties(height=230)
                    st.markdown("<p style='text-align:center;color:#64748b;'>Transaction Count</p>", unsafe_allow_html=True)
                    st.altair_chart(donut1, use_container_width=True)
                with c_val:
                    sum_paid_v = final_df[paid_col].sum() if paid_col else 0
                    sum_wd_v   = final_df[withdrawn_col].sum() if withdrawn_col else 0
                    val_df = pd.DataFrame({
                        "Category": ["Deposits", "Withdrawals"],
                        "Total Value": [sum_paid_v, sum_wd_v]
                    })
                    donut2 = alt.Chart(val_df).mark_arc(innerRadius=55).encode(
                        theta=alt.Theta("Total Value:Q"),
                        color=alt.Color("Category:N", scale=alt.Scale(range=['#10b981', '#f43f5e'])),
                        tooltip=["Category", alt.Tooltip("Total Value:Q", format=",.0f")]
                    ).properties(height=230)
                    st.markdown("<p style='text-align:center;color:#64748b;'>Transaction Volume (KES)</p>", unsafe_allow_html=True)
                    st.altair_chart(donut2, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — MERGED STATEMENT
# ═══════════════════════════════════════════════════════════════════════════════
elif menu_page == "merged":
    if st.session_state["uploaded_file_bytes"] is None:
        render_page_header("📁 Merged Statement", "Upload a workbook first.")
        st.warning("⚠️ No workbook uploaded. Go to **Home & Upload** to get started.")
    elif error_msg:
        render_page_header("📁 Merged Statement", "")
        st.error(f"❌ {error_msg}")
    elif final_df is not None:
        render_page_header(f"📁 {st.session_state['uploaded_file_name']}",
                           "Parsed transaction metrics and summary insights")
        render_kpis()
        st.markdown("---")
        st.subheader("Combined Transaction Records")
        st.write(f"Showing first 100 of **{len(final_df):,}** rows:")
        st.dataframe(final_df.head(100), use_container_width=True)

        st.markdown("##### Download Combined Statement")
        dc1, dc2, _ = st.columns([1, 1, 2])

        try:
            towrite = io.BytesIO()
            final_df.to_excel(towrite, index=False, engine="openpyxl")
            towrite.seek(0)
            dc1.download_button(
                "📥 Excel (.xlsx)", data=towrite,
                file_name=f"merged_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception:
            pass

        csv_bytes = final_df.to_csv(index=False).encode("utf-8")
        dc2.download_button(
            "📥 CSV (.csv)", data=csv_bytes,
            file_name=f"merged_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.csv",
            mime="text/csv"
        )

        st.markdown("---")
        st.subheader("Sheet Processing Details")
        c_inc, c_skp = st.columns(2)
        with c_inc:
            st.markdown(f"🟢 **Merged ({len(included_sheets)}):**")
            for s in included_sheets:
                st.write(f"- `{s}`")
        with c_skp:
            st.markdown(f"🟡 **Skipped ({len(skipped_sheets)}):**")
            if skipped_sheets:
                for sn, reason in skipped_sheets.items():
                    if isinstance(reason, list):
                        st.write(f"- `{sn}`: missing `{', '.join(reason)}`")
                    else:
                        st.write(f"- `{sn}`: {reason}")
            else:
                st.write("None — all sheets were included.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — MONTHLY PIVOT
# ═══════════════════════════════════════════════════════════════════════════════
elif menu_page == "pivot":
    if st.session_state["uploaded_file_bytes"] is None:
        render_page_header("📊 Monthly Pivot", "Upload a workbook first.")
        st.warning("⚠️ No workbook uploaded. Go to **Home & Upload** to get started.")
    elif error_msg:
        render_page_header("📊 Monthly Pivot", "")
        st.error(f"❌ {error_msg}")
    elif final_df is not None:
        render_page_header(f"📊 {st.session_state['uploaded_file_name']}",
                           "Parsed transaction metrics and summary insights")
        render_kpis()
        render_date_picker()
        date_col = st.session_state["date_col"]
        st.markdown("---")

        if not date_col:
            st.warning("⚠️ Select a date column above to build the pivot table.")
        elif paid_col is None or withdrawn_col is None:
            st.warning("⚠️ 'Paid In' and 'Withdrawn' columns are required.")
        else:
            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
            if final_df["_parsed_date"].isna().all():
                st.error(f"❌ Cannot parse dates from **'{date_col}'**.")
            else:
                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")
                final_df["_MonthSort"]  = final_df["_parsed_date"].dt.strftime("%Y-%m")
                pivot_df = (
                    final_df.groupby(["_MonthSort", "_MonthLabel"])[[paid_col, withdrawn_col]]
                    .sum(min_count=1).reset_index()
                    .rename(columns={"_MonthLabel": "Month",
                                     paid_col: "Sum Paid In",
                                     withdrawn_col: "Sum Withdrawn"})
                )
                pivot_df = pivot_df.sort_values("_MonthSort").reset_index(drop=True)
                pivot_df["Net Cash Flow"] = pivot_df["Sum Paid In"].fillna(0) - pivot_df["Sum Withdrawn"].fillna(0)
                display_pivot = pivot_df.drop(columns=["_MonthSort"])

                st.subheader("Monthly Transaction Pivot Summary")
                st.dataframe(display_pivot, use_container_width=True)

                st.markdown("##### Download Pivot Report")
                p1, p2, _ = st.columns([1, 1, 2])
                p1.download_button(
                    "📥 Pivot CSV", data=display_pivot.to_csv(index=False).encode("utf-8"),
                    file_name=f"pivot_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.csv",
                    mime="text/csv"
                )
                try:
                    piv_bytes = io.BytesIO()
                    display_pivot.to_excel(piv_bytes, index=False, engine="openpyxl")
                    piv_bytes.seek(0)
                    p2.download_button(
                        "📥 Pivot Excel", data=piv_bytes,
                        file_name=f"pivot_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception:
                    pass


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — PRINTABLE REPORT
# ═══════════════════════════════════════════════════════════════════════════════
elif menu_page == "report":
    if st.session_state["uploaded_file_bytes"] is None:
        render_page_header("🖨️ Printable Report", "Upload a workbook first.")
        st.warning("⚠️ No workbook uploaded. Go to **Home & Upload** to get started.")
    elif error_msg:
        render_page_header("🖨️ Printable Report", "")
        st.error(f"❌ {error_msg}")
    elif final_df is not None:
        render_page_header(f"🖨️ {st.session_state['uploaded_file_name']}",
                           "Print-ready monthly cash flow summary")
        render_kpis()
        render_date_picker()
        date_col = st.session_state["date_col"]
        st.markdown("---")

        if not date_col or paid_col is None or withdrawn_col is None:
            st.warning("⚠️ Configure date and column settings above to generate the report.")
        else:
            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
            if not final_df["_parsed_date"].isna().all():
                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")
                final_df["_MonthSort"]  = final_df["_parsed_date"].dt.strftime("%Y-%m")
                pivot_df = (
                    final_df.groupby(["_MonthSort", "_MonthLabel"])[[paid_col, withdrawn_col]]
                    .sum(min_count=1).reset_index()
                    .rename(columns={"_MonthLabel": "Month",
                                     paid_col: "Sum Paid In",
                                     withdrawn_col: "Sum Withdrawn"})
                )
                pivot_df = pivot_df.sort_values("_MonthSort").reset_index(drop=True)
                pivot_df["Net Cash Flow"] = pivot_df["Sum Paid In"].fillna(0) - pivot_df["Sum Withdrawn"].fillna(0)
                display_pivot = pivot_df.drop(columns=["_MonthSort"])

                print_df = display_pivot.copy()
                for c in ["Sum Paid In", "Sum Withdrawn", "Net Cash Flow"]:
                    print_df[c] = print_df[c].apply(lambda v: f"{v:,.2f}" if pd.notna(v) else "-")

                pivot_html_table = print_df.to_html(index=False, classes="pivot-table", border=0)
                total_tx = len(final_df)

                printable_html = f"""
                <html><head><meta charset="utf-8">
                <meta name="viewport" content="width=device-width,initial-scale=1">
                <style>
                  body {{ font-family:'Helvetica Neue',Helvetica,Arial,sans-serif; padding:20px; color:#334155; background:#fff; }}
                  .header {{ border-bottom:2px solid #1b8a5a; padding-bottom:12px; margin-bottom:20px; }}
                  .title  {{ font-size:20px; font-weight:bold; color:#0f172a; margin:0; }}
                  .meta   {{ font-size:12px; color:#64748b; margin-top:4px; }}
                  table.pivot-table {{ border-collapse:collapse; width:100%; margin-top:12px; }}
                  table.pivot-table th {{ background:#f8fafc; border-bottom:2px solid #cbd5e1; color:#475569; font-weight:600; padding:9px 10px; text-align:left; font-size:12px; text-transform:uppercase; }}
                  table.pivot-table td {{ border-bottom:1px solid #e2e8f0; padding:9px 10px; font-size:13px; color:#334155; }}
                  table.pivot-table tr:hover {{ background:#f8fafc; }}
                  .print-btn {{ display:inline-block; margin-bottom:14px; padding:9px 16px; background:#1b8a5a; color:#fff; border-radius:6px; cursor:pointer; text-decoration:none; font-weight:500; font-size:13px; border:none; }}
                  .print-btn:hover {{ background:#0d5c3a; }}
                  @media print {{ .print-btn {{ display:none; }} body {{ padding:0; }} }}
                  @media (max-width:600px) {{ table.pivot-table th, table.pivot-table td {{ padding:6px 6px; font-size:11px; }} }}
                </style></head>
                <body>
                  <button class="print-btn" onclick="window.print()">🖨️ Print Report</button>
                  <div class="header">
                    <div class="title">M-Pesa Cash Flow Statement Summary</div>
                    <div class="meta">Document: {st.session_state['uploaded_file_name']} &nbsp;|&nbsp; Total Transactions: {total_tx:,}</div>
                  </div>
                  {pivot_html_table}
                </body></html>
                """
                st.components.v1.html(printable_html, height=600, scrolling=True)


# ─── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-footer">
    Mpesa Statement Analyzer &copy; 2026 &nbsp;·&nbsp; Made with ♥ for Relationship Officers
</div>
""", unsafe_allow_html=True)

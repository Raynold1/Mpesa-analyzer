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

# Set page configuration with a nice favicon and layout
st.set_page_config(
    page_title="Mpesa Analyzer - Premium Insights",
    page_icon="📥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium styling
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=Outfit:wght@300;400;500;600;700&display=swap');
    
    /* Global Background and Typography */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', 'Inter', 'Outfit', -apple-system, BlinkMacSystemFont, sans-serif;
        background-color: #f8fafc;
        color: #1e293b;
    }
    
    /* Top Header */
    [data-testid="stHeader"] {
        background-color: rgba(248, 250, 252, 0.8);
        backdrop-filter: blur(8px);
    }
    
    /* Sidebar Overrides */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #f1f5f9 !important;
        box-shadow: 4px 0 24px rgba(15, 23, 42, 0.02) !important;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        font-size: 0.95rem;
    }
    
    /* Brand Design */
    .brand-title {
        font-size: 1.75rem;
        font-weight: 800;
        background: linear-gradient(135deg, #10b981 0%, #059669 50%, #0f766e 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.15rem;
        letter-spacing: -0.025em;
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }
    
    .brand-subtitle {
        font-size: 0.85rem;
        color: #64748b;
        margin-bottom: 1.75rem;
        border-bottom: 1px solid #f1f5f9;
        padding-bottom: 1rem;
        font-weight: 500;
    }
    
    .main-title {
        font-size: 2.25rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 0.15rem;
        letter-spacing: -0.03em;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #047857 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .main-subtitle {
        font-size: 1rem;
        color: #64748b;
        margin-bottom: 2rem;
        font-weight: 400;
    }
    
    /* Navigation Menu Styling */
    div[data-testid="stRadio"] [data-testid="stWidgetLabel"] {
        font-weight: 700 !important;
        color: #64748b !important;
        text-transform: uppercase !important;
        font-size: 0.75rem !important;
        letter-spacing: 0.05em !important;
        margin-bottom: 0.75rem !important;
    }
    
    div[data-testid="stRadio"] div[role="radiogroup"] {
        gap: 0 !important;
    }
    
    div[data-testid="stRadio"] div[role="radiogroup"] > label {
        display: flex !important;
        align-items: center !important;
        padding: 12px 16px !important;
        margin-bottom: 8px !important;
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
        background-color: #ffffff !important;
        color: #475569 !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        font-weight: 500 !important;
        cursor: pointer !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.02) !important;
    }
    
    div[data-testid="stRadio"] div[role="radiogroup"] > label:hover {
        background-color: #f8fafc !important;
        border-color: #cbd5e1 !important;
        color: #0f172a !important;
        transform: translateX(3px) !important;
    }
    
    div[data-testid="stRadio"] div[role="radiogroup"] > label:has(input[type="radio"]:checked) {
        background-color: #ecfdf4 !important;
        border-color: #10b981 !important;
        color: #065f46 !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px -1px rgba(16, 185, 129, 0.12), 0 2px 4px -2px rgba(16, 185, 129, 0.1) !important;
    }
    
    /* Hide default radio elements */
    div[data-testid="stRadio"] div[role="radiogroup"] > label input[type="radio"] {
        display: none !important;
    }
    div[data-testid="stRadio"] div[role="radiogroup"] > label > div:first-of-type {
        display: none !important;
    }
    div[data-testid="stRadio"] div[role="radiogroup"] > label div[data-testid="stMarkdownContainer"] {
        margin-left: 0 !important;
    }
    
    /* Subheader & Section Titles */
    h3 {
        font-size: 1.5rem !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        margin-top: 1.5rem !important;
        margin-bottom: 1rem !important;
        letter-spacing: -0.02em !important;
    }
    
    h5 {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        color: #1e293b !important;
        margin-bottom: 0.75rem !important;
    }
    
    /* File Uploader override */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 16px !important;
        padding: 1.75rem !important;
        background-color: #ffffff !important;
        transition: all 0.25s ease !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.02) !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #10b981 !important;
        background-color: #f0fdf4 !important;
        box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.05) !important;
    }
    
    div[data-testid="stFileUploader"] button {
        background-color: #10b981 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1.25rem !important;
        font-weight: 600 !important;
        transition: all 0.2s !important;
        box-shadow: 0 2px 4px rgba(16, 185, 129, 0.15) !important;
    }
    
    div[data-testid="stFileUploader"] button:hover {
        background-color: #059669 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 6px rgba(16, 185, 129, 0.25) !important;
    }
    
    /* Standard and Custom Buttons styling */
    div.stButton button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.6rem 1.5rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.2), 0 2px 4px -2px rgba(16, 185, 129, 0.1) !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }
    div.stButton button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.25), 0 4px 6px -2px rgba(16, 185, 129, 0.15) !important;
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
    }
    div.stButton button:active {
        transform: translateY(0) !important;
    }
    
    /* Clear action button */
    div.stButton button[help*="Clear"] {
        background: #ffffff !important;
        color: #ef4444 !important;
        border: 1px solid #fee2e2 !important;
        box-shadow: 0 1px 2px rgba(239, 68, 68, 0.05) !important;
    }
    div.stButton button[help*="Clear"]:hover {
        background: #fef2f2 !important;
        border-color: #fca5a5 !important;
        box-shadow: 0 4px 6px -1px rgba(239, 68, 68, 0.08) !important;
        color: #dc2626 !important;
    }
    
    /* Download Buttons */
    div.stDownloadButton button {
        background-color: #ffffff !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 12px !important;
        padding: 0.6rem 1.25rem !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        transition: all 0.2s ease-in-out !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05) !important;
        width: 100% !important;
    }
    div.stDownloadButton button:hover {
        background-color: #f8fafc !important;
        border-color: #10b981 !important;
        color: #059669 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05) !important;
    }
    
    /* Input and Select elements */
    div[data-testid="stSelectbox"] > div {
        border-radius: 12px !important;
    }
    div[data-testid="stTextInput"] input {
        border-radius: 12px !important;
        border: 1px solid #cbd5e1 !important;
        padding: 0.5rem 0.75rem !important;
        font-size: 0.95rem !important;
        transition: all 0.2s !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.02) !important;
    }
    div[data-testid="stTextInput"] input:focus {
        border-color: #10b981 !important;
        box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.15) !important;
    }
    
    /* Alert Messages style */
    .stAlert {
        border-radius: 16px !important;
        border: 1px solid #f1f5f9 !important;
        box-shadow: 0 4px 6px -1px rgba(15, 23, 42, 0.02) !important;
        background-color: #ffffff !important;
    }
    .stAlert [data-testid="stMarkdownContainer"] {
        color: #334155 !important;
        font-size: 0.95rem !important;
    }
    
    /* Welcome card / Guide panel */
    .welcome-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 2.5rem;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.02), 0 8px 10px -6px rgba(0, 0, 0, 0.02);
        margin-bottom: 2rem;
    }
    .welcome-step {
        display: flex;
        align-items: flex-start;
        gap: 1.25rem;
        margin-bottom: 1.5rem;
        padding: 1rem;
        border-radius: 12px;
        transition: all 0.2s;
    }
    .welcome-step:hover {
        background-color: #f8fafc;
        transform: translateX(4px);
    }
    .welcome-icon {
        background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        color: #059669;
        border-radius: 10px;
        width: 36px;
        height: 36px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 1.1rem;
        flex-shrink: 0;
        box-shadow: 0 2px 4px 0 rgba(16, 185, 129, 0.05);
    }
    .welcome-text {
        font-size: 0.95rem;
        color: #475569;
        line-height: 1.5;
    }
    
    /* Expander Container overrides */
    div[data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 16px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02) !important;
        overflow: hidden !important;
    }
    
    /* KPI Cards Styling Grid */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 1.25rem;
        margin-bottom: 2rem;
        width: 100%;
    }
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 1.25rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.03), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 125px;
    }
    .kpi-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 20px -3px rgba(0, 0, 0, 0.06), 0 4px 6px -2px rgba(0, 0, 0, 0.03);
    }
    .kpi-card-indigo:hover { border-color: #6366f1; }
    .kpi-card-amber:hover { border-color: #f59e0b; }
    .kpi-card-emerald:hover { border-color: #10b981; }
    .kpi-card-rose:hover { border-color: #f43f5e; }
    .kpi-card-teal:hover { border-color: #0d9488; }
    
    .kpi-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.5rem;
    }
    .kpi-label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .kpi-icon-container {
        width: 32px;
        height: 32px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.1rem;
    }
    .kpi-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.2;
        margin-top: 0.25rem;
    }
    .kpi-badge {
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.15rem 0.5rem;
        border-radius: 9999px;
        display: inline-block;
        margin-top: 0.75rem;
        width: fit-content;
    }
    
    /* Color utility classes for KPI Cards */
    .bg-indigo-50 { background-color: #e0e7ff; color: #4338ca; }
    .bg-amber-50 { background-color: #fffbeb; color: #b45309; }
    .bg-emerald-50 { background-color: #ecfdf5; color: #047857; }
    .bg-rose-50 { background-color: #fff1f2; color: #b91c1c; }
    .bg-teal-50 { background-color: #f0fdfa; color: #0f766e; }
    
    .badge-indigo { background-color: #e0e7ff; color: #3730a3; }
    .badge-amber { background-color: #fef3c7; color: #92400e; }
    .badge-emerald { background-color: #d1fae5; color: #065f46; }
    .badge-rose { background-color: #ffe4e6; color: #991b1b; }
    .badge-teal { background-color: #ccfbf1; color: #115e59; }
    
    /* Tables and Dataframe tweaks */
    div[data-testid="stDataFrame"] {
        border-radius: 16px !important;
        overflow: hidden !important;
        border: 1px solid #e2e8f0 !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02) !important;
    }
    
    /* Footer Styling */
    .app-footer {
        text-align: center;
        padding: 2.5rem 0;
        color: #94a3b8;
        font-size: 0.85rem;
        border-top: 1px solid #f1f5f9;
        margin-top: 5rem;
        font-weight: 500;
    }
    </style>
""", unsafe_allow_html=True)

# Helper function to find columns case-insensitively/partially
def find_col(df, target):
    for c in df.columns:
        if c.lower() == target.lower():
            return c
    for c in df.columns:
        if target.lower() in c.lower():
            return c
    return None

# Initialize persistent session state
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

# Sidebar Menu (User requested navigation items in the sidebar)
with st.sidebar:
    st.markdown('<div class="brand-title">📥 Mpesa Analyzer</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Financial statement aggregator</div>', unsafe_allow_html=True)
    
    # Navigation items
    menu = st.radio(
        "Navigation",
        options=[
            "🏠 Home & Upload", 
            "📈 Visual Dashboard", 
            "📁 Merged Statement", 
            "📊 Monthly Pivot", 
            "🖨️ Printable Report"
        ]
    )

# Excel processing cache logic
@st.cache_data(show_spinner="Processing statements...")
def process_statement(file_bytes, required_cols_str, case_insensitive_flag):
    try:
        sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
        required_columns = [c.strip() for c in required_cols_str.split(",") if c.strip()]
        
        if not required_columns:
            return None, "Please specify at least one required column in settings.", {}, sheets, []
            
        merged_dfs = []
        included_sheets = []
        skipped = {}
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
                    if case_insensitive_flag:
                        missing = [r for r in req_lower if r not in cols_lower]
                    else:
                        missing = [r for r in required_columns if r not in cols]
                    skipped[sheet_name] = missing
            except Exception as e:
                skipped[sheet_name] = f"Read error: {e}"
                
        if not merged_dfs:
            return None, f"No sheets contained all required columns: {', '.join(required_columns)}", skipped, sheets, []
            
        final_df = pd.concat(merged_dfs, ignore_index=True)
        return final_df, None, skipped, sheets, included_sheets
    except Exception as exc:
        return None, f"Excel file read error: {exc}", {}, {}, []

# Resolve data logic if file is active
final_df = None
error_msg = None
skipped_sheets = {}
all_sheets = {}
included_sheets = []
paid_col = None
withdrawn_col = None
balance_col = None

if st.session_state["uploaded_file_bytes"] is not None:
    final_df, error_msg, skipped_sheets, all_sheets, included_sheets = process_statement(
        st.session_state["uploaded_file_bytes"],
        st.session_state["required_cols"],
        st.session_state["case_insensitive"]
    )
    
    if final_df is not None:
        paid_col = find_col(final_df, "Paid In")
        withdrawn_col = find_col(final_df, "Withdrawn")
        balance_col = find_col(final_df, "Balance")
        
        if paid_col:
            final_df[paid_col] = pd.to_numeric(final_df[paid_col], errors="coerce")
        if withdrawn_col:
            final_df[withdrawn_col] = pd.to_numeric(final_df[withdrawn_col], errors="coerce")
        if balance_col:
            final_df[balance_col] = pd.to_numeric(final_df[balance_col], errors="coerce")
            
        # Contextual Date selection in sidebar when file is loaded
        candidate_date_cols = [
            c for c in final_df.columns
            if any(k in c.lower() for k in ("date", "time", "completion", "timestamp"))
        ]
        
        with st.sidebar:
            st.markdown("---")
            st.subheader("⚙️ Grouping Options")
            if candidate_date_cols:
                # Resolve date selection state
                saved_date_col = st.session_state["date_col"]
                if saved_date_col not in candidate_date_cols:
                    saved_date_col = candidate_date_cols[0]
                    st.session_state["date_col"] = saved_date_col
                
                selected_date = st.selectbox(
                    "Date Column for Grouping",
                    options=candidate_date_cols,
                    index=candidate_date_cols.index(saved_date_col),
                    help="Column used to group transactions by month."
                )
                st.session_state["date_col"] = selected_date
            else:
                date_input = st.text_input(
                    "Specify Date Column name",
                    value=st.session_state["date_col"],
                    help="Specify the column containing transaction dates."
                )
                st.session_state["date_col"] = date_input

# --- Global Header for Analytical pages ---
if menu != "🏠 Home & Upload":
    if st.session_state["uploaded_file_bytes"] is None:
        st.markdown('<div class="main-title">📥 Mpesa Statement Analyzer</div>', unsafe_allow_html=True)
        st.markdown('<div class="main-subtitle">Select analysis tabs from the sidebar</div>', unsafe_allow_html=True)
        st.warning("⚠️ No workbook uploaded yet. Please navigate to **🏠 Home & Upload** to select your Excel file.")
    elif error_msg:
        st.markdown('<div class="main-title">📥 Mpesa Statement Analyzer</div>', unsafe_allow_html=True)
        st.error(f"❌ {error_msg}")
    elif final_df is not None:
        st.markdown(f'<div class="main-title">📁 {st.session_state["uploaded_file_name"]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="main-subtitle">Parsed transaction metrics and summary insights</div>', unsafe_allow_html=True)
        
        # Calculate global KPIs
        total_transactions = len(final_df)
        total_sheets = len(all_sheets)
        merged_count = len(included_sheets)
        
        sum_paid = final_df[paid_col].sum() if (paid_col and not final_df[paid_col].isna().all()) else 0.0
        sum_withdrawn = final_df[withdrawn_col].sum() if (withdrawn_col and not final_df[withdrawn_col].isna().all()) else 0.0
        net_flow = sum_paid - sum_withdrawn
        
        # Display custom modern Tailwind-style KPI Cards Grid
        sum_paid_val = f"KES {sum_paid:,.2f}"
        sum_withdrawn_val = f"KES {sum_withdrawn:,.2f}"
        net_flow_val = f"KES {net_flow:,.2f}"
        if net_flow < 0:
            net_flow_val = f"- KES {abs(net_flow):,.2f}"
        
        net_theme_bg = "bg-teal-50" if net_flow >= 0 else "bg-rose-50"
        net_theme_badge = "badge-teal" if net_flow >= 0 else "badge-rose"
        net_theme_card = "kpi-card-teal" if net_flow >= 0 else "kpi-card-rose"
        net_trend = "Positive Flow" if net_flow >= 0 else "Negative Flow"

        kpi_html = f"""
        <div class="kpi-grid">
            <div class="kpi-card kpi-card-indigo">
                <div class="kpi-header">
                    <span class="kpi-label">Sheets Combined</span>
                    <div class="kpi-icon-container bg-indigo-50">📁</div>
                </div>
                <div class="kpi-value">{merged_count} / {total_sheets}</div>
                <div class="kpi-badge badge-indigo">Excel Workbook</div>
            </div>
            <div class="kpi-card kpi-card-amber">
                <div class="kpi-header">
                    <span class="kpi-label">Transactions</span>
                    <div class="kpi-icon-container bg-amber-50">🔢</div>
                </div>
                <div class="kpi-value">{total_transactions:,}</div>
                <div class="kpi-badge badge-amber">Total Rows</div>
            </div>
            <div class="kpi-card kpi-card-emerald">
                <div class="kpi-header">
                    <span class="kpi-label">Total Deposits (In)</span>
                    <div class="kpi-icon-container bg-emerald-50">📈</div>
                </div>
                <div class="kpi-value">{sum_paid_val}</div>
                <div class="kpi-badge badge-emerald">Cash Inflow</div>
            </div>
            <div class="kpi-card kpi-card-rose">
                <div class="kpi-header">
                    <span class="kpi-label">Total Withdrawn (Out)</span>
                    <div class="kpi-icon-container bg-rose-50">📉</div>
                </div>
                <div class="kpi-value">{sum_withdrawn_val}</div>
                <div class="kpi-badge badge-rose">Cash Outflow</div>
            </div>
            <div class="kpi-card {net_theme_card}">
                <div class="kpi-header">
                    <span class="kpi-label">Net Cash Flow</span>
                    <div class="kpi-icon-container {net_theme_bg}">💰</div>
                </div>
                <div class="kpi-value">{net_flow_val}</div>
                <div class="kpi-badge {net_theme_badge}">{net_trend}</div>
            </div>
        </div>
        """
        st.markdown(kpi_html, unsafe_allow_html=True)
        st.write("")

# --- Page 1: Home & Upload ---
if menu == "🏠 Home & Upload":
    st.markdown('<div class="main-title">📥 Mpesa Statement Analyzer</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Combine, summarize, and visualize multi-sheet Mpesa data instantly.</div>', unsafe_allow_html=True)
    
    if st.session_state["uploaded_file_bytes"] is None:
        # File uploader is only shown if no active file is loaded
        uploaded_file = st.file_uploader(
            "Upload Excel Workbook (.xlsx, .xls)", 
            type=["xlsx", "xls"],
            help="Upload the Excel file containing one or more client Mpesa sheets."
        )
        if uploaded_file is not None:
            st.session_state["uploaded_file_bytes"] = uploaded_file.getvalue()
            st.session_state["uploaded_file_name"] = uploaded_file.name
            st.rerun()
    else:
        # Active file feedback with an explicit "Clear" action
        st.success(f"✅ Active Statement Loaded: **{st.session_state['uploaded_file_name']}**")
        st.info("👈 Use the sidebar navigation menu to view dashboard charts, records, and pivot tables.")
        
        if st.button("🗑️ Clear & Upload Another File", help="Clear current workbook and upload a new one."):
            st.session_state["uploaded_file_bytes"] = None
            st.session_state["uploaded_file_name"] = None
            st.session_state["date_col"] = ""
            st.rerun()
        
    # Step guide
    st.markdown("""
    <div class="welcome-card">
        <h3 style="margin-top:0; color: #0f172a; font-size: 1.4rem;">Get Started in Seconds</h3>
        <p style="color: #64748b; margin-bottom: 2rem;">Follow these steps to analyze your client statement data:</p>
        <div class="welcome-step">
            <div class="welcome-icon">1</div>
            <div class="welcome-text">
                <strong>Upload Statement:</strong> Drag and drop your client's Excel workbook (.xlsx or .xls) above.
            </div>
        </div>
        <div class="welcome-step">
            <div class="welcome-icon">2</div>
            <div class="welcome-text">
                <strong>Confirm Settings:</strong> Expand the configuration details below to customize statement columns.
            </div>
        </div>
        <div class="welcome-step">
            <div class="welcome-icon">3</div>
            <div class="welcome-text">
                <strong>View Insights:</strong> Use the sidebar navigation on the left to review metrics, visualization charts, monthly pivot summaries, or print reports.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Settings panel on Home page
    with st.expander("⚙️ Configuration Settings", expanded=True):
        req_cols = st.text_input(
            "Required columns (comma-separated)", 
            value=st.session_state["required_cols"],
            help="Only sheets containing all these column headings will be merged."
        )
        if req_cols != st.session_state["required_cols"]:
            st.session_state["required_cols"] = req_cols
            st.rerun()
            
        case_ins = st.checkbox(
            "Case-insensitive column matching", 
            value=st.session_state["case_insensitive"],
            help="Matches columns regardless of casing (e.g. matches 'Paid In' with 'paid in')."
        )
        if case_ins != st.session_state["case_insensitive"]:
            st.session_state["case_insensitive"] = case_ins
            st.rerun()

# --- Page 2: Visual Dashboard ---
elif menu == "📈 Visual Dashboard":
    if final_df is not None:
        st.subheader("Cash Flow Dashboard Insights")
        date_col = st.session_state["date_col"]
        
        if not date_col:
            st.info("ℹ️ Select or specify a date column in the sidebar grouping settings to compute monthly trends.")
        elif paid_col is None or withdrawn_col is None:
            st.warning("⚠️ Visual insights require 'Paid In' and 'Withdrawn' columns to calculate deposits/withdrawals.")
        else:
            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
            valid_dates_count = final_df["_parsed_date"].notna().sum()
            
            if valid_dates_count == 0:
                st.error(f"❌ Could not parse any valid dates from column **'{date_col}'**. Please ensure the column contains dates.")
            else:
                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")
                final_df["_MonthSort"] = final_df["_parsed_date"].dt.strftime("%Y-%m")
                
                # Group monthly values
                dash_pivot = (
                    final_df.groupby(["_MonthSort", "_MonthLabel"])[ [paid_col, withdrawn_col] ]
                    .sum(min_count=1)
                    .reset_index()
                    .rename(columns={"_MonthLabel": "Month", paid_col: "Deposits", withdrawn_col: "Withdrawals"})
                )
                dash_pivot = dash_pivot.sort_values("_MonthSort").reset_index(drop=True)
                
                # Meltdown for Altair plotting
                melt_df = dash_pivot.melt(id_vars=["Month", "_MonthSort"], value_vars=["Deposits", "Withdrawals"], 
                                         var_name="Transaction Type", value_name="Amount")
                
                # Custom styled charts
                c1, c2 = st.columns([2, 1])
                
                with c1:
                    st.markdown("##### Monthly Cash Flow (Deposits vs. Withdrawals)")
                    cash_flow_chart = alt.Chart(melt_df).mark_bar(
                        cornerRadiusTopLeft=6, 
                        cornerRadiusTopRight=6
                    ).encode(
                        x=alt.X('Month:N', sort=alt.SortField(field='_MonthSort', order='ascending'), title='Month'),
                        y=alt.Y('Amount:Q', title='Amount (KES)'),
                        color=alt.Color('Transaction Type:N', scale=alt.Scale(domain=['Deposits', 'Withdrawals'], range=['#10b981', '#f43f5e'])),
                        xOffset='Transaction Type:N',
                        tooltip=[
                            alt.Tooltip('Month:N'),
                            alt.Tooltip('Transaction Type:N'),
                            alt.Tooltip('Amount:Q', format=",.2f")
                        ]
                    ).properties(
                        height=350
                    ).configure_view(
                        strokeOpacity=0
                    )
                    st.altair_chart(cash_flow_chart, use_container_width=True)
                    
                with c2:
                    st.markdown("##### Net Monthly Savings / Cash Balance")
                    dash_pivot["Net Savings"] = dash_pivot["Deposits"].fillna(0) - dash_pivot["Withdrawals"].fillna(0)
                    
                    net_chart = alt.Chart(dash_pivot).mark_area(
                        line={'color': '#0d9488', 'width': 2.5},
                        color=alt.Gradient(
                            gradient='linear',
                            stops=[alt.GradientStop(color='#ccfbf1', offset=0),
                                   alt.GradientStop(color='#ffffff', offset=1)],
                            x1=1, y1=1, x2=1, y2=0
                        )
                    ).encode(
                        x=alt.X('Month:N', sort=alt.SortField(field='_MonthSort', order='ascending'), title='Month'),
                        y=alt.Y('Net Savings:Q', title='Net Amount (KES)'),
                        tooltip=[
                            alt.Tooltip('Month:N'),
                            alt.Tooltip('Net Savings:Q', format=",.2f")
                        ]
                    ).properties(
                        height=350
                    )
                    st.altair_chart(net_chart, use_container_width=True)
                    
                # Distribution charts
                st.markdown("---")
                st.markdown("##### Statement Summary Distribution")
                c_vol, c_net = st.columns(2)
                with c_vol:
                    count_paid = final_df[paid_col].notna().sum()
                    count_withdrawn = final_df[withdrawn_col].notna().sum()
                    counts_df = pd.DataFrame({
                        "Category": ["Deposits (Paid In)", "Withdrawals (Out)"],
                        "Count": [count_paid, count_withdrawn]
                    })
                    volume_chart = alt.Chart(counts_df).mark_arc(innerRadius=60).encode(
                        theta=alt.Theta(field="Count", type="quantitative"),
                        color=alt.Color(field="Category", type="nominal", scale=alt.Scale(range=['#10b981', '#f43f5e'])),
                        tooltip=["Category", "Count"]
                    ).properties(
                        height=250
                    )
                    st.markdown("<p style='text-align: center; color: #64748b;'>Transaction Count Share</p>", unsafe_allow_html=True)
                    st.altair_chart(volume_chart, use_container_width=True)
                    
                with c_net:
                    share_df = pd.DataFrame({
                        "Category": ["Deposits (Paid In)", "Withdrawals (Out)"],
                        "Total Value": [sum_paid, sum_withdrawn]
                    })
                    value_chart = alt.Chart(share_df).mark_arc(innerRadius=60).encode(
                        theta=alt.Theta(field="Total Value", type="quantitative"),
                        color=alt.Color(field="Category", type="nominal", scale=alt.Scale(range=['#10b981', '#f43f5e'])),
                        tooltip=["Category", alt.Tooltip("Total Value", format=",.2f")]
                    ).properties(
                        height=250
                    )
                    st.markdown("<p style='text-align: center; color: #64748b;'>Transaction Volume Value (KES)</p>", unsafe_allow_html=True)
                    st.altair_chart(value_chart, use_container_width=True)

# --- Page 3: Merged Statement ---
elif menu == "📁 Merged Statement":
    if final_df is not None:
        st.subheader("Combined Transaction Records")
        st.write(f"Displaying first 100 rows of merged statement (Total records: {len(final_df)}):")
        st.dataframe(final_df.head(100), use_container_width=True)
        
        st.markdown("##### Download Combined Statement")
        dc1, dc2, _ = st.columns([1, 1, 2])
        
        # Excel download
        try:
            towrite = io.BytesIO()
            final_df.to_excel(towrite, index=False, engine="openpyxl")
            towrite.seek(0)
            dc1.download_button(
                label="📥 Download Merged Excel (.xlsx)",
                data=towrite,
                file_name=f"merged_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception:
            towrite = io.BytesIO()
            final_df.to_excel(towrite, index=False)
            towrite.seek(0)
            dc1.download_button(
                label="📥 Download Merged Excel (.xlsx)",
                data=towrite,
                file_name=f"merged_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
        # CSV download
        csv_bytes = final_df.to_csv(index=False).encode("utf-8")
        dc2.download_button(
            label="📥 Download Merged CSV (.csv)",
            data=csv_bytes,
            file_name=f"merged_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.csv",
            mime="text/csv"
        )
        
        # Details about sheet inclusion/skipping
        st.markdown("---")
        st.subheader("Aggregation Processing Details")
        
        c_inc, c_skp = st.columns(2)
        with c_inc:
            st.markdown(f"🟢 **Merged Sheets ({len(included_sheets)}):**")
            for s in included_sheets:
                st.write(f"- `{s}`")
        with c_skp:
            st.markdown(f"🟡 **Skipped Sheets ({len(skipped_sheets)}):**")
            if skipped_sheets:
                for sheet_name, reason in skipped_sheets.items():
                    if isinstance(reason, list):
                        st.write(f"- `{sheet_name}`: Missing column(s) `{', '.join(reason)}`")
                    else:
                        st.write(f"- `{sheet_name}`: {reason}")
            else:
                st.write("None")

# --- Page 4: Monthly Pivot ---
elif menu == "📊 Monthly Pivot":
    if final_df is not None:
        st.subheader("Monthly Transaction Pivot Summary")
        date_col = st.session_state["date_col"]
        
        if not date_col:
            st.warning("⚠️ Select or specify a date column in the sidebar grouping settings to calculate the Monthly Pivot table.")
        elif paid_col is None or withdrawn_col is None:
            st.warning("⚠️ 'Paid In' and 'Withdrawn' columns are required to construct the pivot table.")
        else:
            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
            
            if final_df["_parsed_date"].isna().all():
                st.error(f"❌ Error: Cannot parse valid datetime values from '{date_col}'. Check the values.")
            else:
                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")
                final_df["_MonthSort"] = final_df["_parsed_date"].dt.strftime("%Y-%m")
                
                # Pivot logic
                pivot_df = (
                    final_df.groupby(["_MonthSort", "_MonthLabel"])[[paid_col, withdrawn_col]]
                    .sum(min_count=1)
                    .reset_index()
                    .rename(columns={"_MonthLabel": "Month", paid_col: "Sum Paid In", withdrawn_col: "Sum Withdrawn"})
                )
                pivot_df = pivot_df.sort_values("_MonthSort").reset_index(drop=True)
                pivot_df["Net Cash Flow"] = pivot_df["Sum Paid In"].fillna(0) - pivot_df["Sum Withdrawn"].fillna(0)
                display_pivot = pivot_df.drop(columns=["_MonthSort"])
                
                st.write("Summary table of deposits, withdrawals, and net savings aggregated chronologically:")
                st.dataframe(display_pivot, use_container_width=True)
                
                st.markdown("##### Download Pivot Reports")
                p_col1, p_col2, _ = st.columns([1, 1, 2])
                
                # Pivot CSV download
                pivot_csv = display_pivot.to_csv(index=False).encode("utf-8")
                p_col1.download_button(
                    "📥 Download Pivot CSV", 
                    data=pivot_csv, 
                    file_name=f"pivot_summary_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.csv", 
                    mime="text/csv"
                )
                
                # Pivot Excel download
                try:
                    piv_bytes = io.BytesIO()
                    display_pivot.to_excel(piv_bytes, index=False, engine="openpyxl")
                    piv_bytes.seek(0)
                    p_col2.download_button(
                        label="📥 Download Pivot Excel (.xlsx)",
                        data=piv_bytes,
                        file_name=f"pivot_summary_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception:
                    piv_bytes = io.BytesIO()
                    display_pivot.to_excel(piv_bytes, index=False)
                    piv_bytes.seek(0)
                    p_col2.download_button(
                        label="📥 Download Pivot Excel (.xlsx)",
                        data=piv_bytes,
                        file_name=f"pivot_summary_{os.path.splitext(st.session_state['uploaded_file_name'])[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

# --- Page 5: Printable Report ---
elif menu == "🖨️ Printable Report":
    if final_df is not None:
        st.subheader("Print Preview & Export")
        st.write("Use this view to print a neat summary statement of client's monthly cash flow activity.")
        date_col = st.session_state["date_col"]
        
        if not date_col or paid_col is None or withdrawn_col is None:
            st.warning("⚠️ Configure settings to compile pivot data before generating a report.")
        else:
            final_df["_parsed_date"] = pd.to_datetime(final_df[date_col], errors="coerce")
            
            if not final_df["_parsed_date"].isna().all():
                final_df["_MonthLabel"] = final_df["_parsed_date"].dt.strftime("%B %Y")
                final_df["_MonthSort"] = final_df["_parsed_date"].dt.strftime("%Y-%m")
                
                pivot_df = (
                    final_df.groupby(["_MonthSort", "_MonthLabel"])[[paid_col, withdrawn_col]]
                    .sum(min_count=1)
                    .reset_index()
                    .rename(columns={"_MonthLabel": "Month", paid_col: "Sum Paid In", withdrawn_col: "Sum Withdrawn"})
                )
                pivot_df = pivot_df.sort_values("_MonthSort").reset_index(drop=True)
                pivot_df["Net Cash Flow"] = pivot_df["Sum Paid In"].fillna(0) - pivot_df["Sum Withdrawn"].fillna(0)
                display_pivot = pivot_df.drop(columns=["_MonthSort"])
                
                # Format numbers for printing
                print_df = display_pivot.copy()
                print_df["Sum Paid In"] = print_df["Sum Paid In"].apply(lambda v: f"{v:,.2f}" if pd.notna(v) else "-")
                print_df["Sum Withdrawn"] = print_df["Sum Withdrawn"].apply(lambda v: f"{v:,.2f}" if pd.notna(v) else "-")
                print_df["Net Cash Flow"] = print_df["Net Cash Flow"].apply(lambda v: f"{v:,.2f}" if pd.notna(v) else "-")
                
                pivot_html_table = print_df.to_html(index=False, classes="pivot-table", border=0)
                
                printable_html = f"""
                <html>
                  <head>
                    <meta charset="utf-8"/>
                    <style>
                      body {{
                        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; 
                        padding: 24px;
                        color: #334155;
                        background-color: #ffffff;
                      }}
                      .header {{
                        border-bottom: 2px solid #1b8a5a;
                        padding-bottom: 12px;
                        margin-bottom: 24px;
                      }}
                      .title {{
                        font-size: 24px;
                        font-weight: bold;
                        color: #0f172a;
                        margin: 0;
                      }}
                      .meta {{
                        font-size: 13px;
                        color: #64748b;
                        margin-top: 4px;
                      }}
                      table.pivot-table {{
                        border-collapse: collapse; 
                        width: 100%;
                        margin-top: 16px;
                      }}
                      table.pivot-table th {{
                        background-color: #f8fafc;
                        border-bottom: 2px solid #cbd5e1;
                        color: #475569;
                        font-weight: 600;
                        padding: 10px 12px;
                        text-align: left;
                        font-size: 13px;
                        text-transform: uppercase;
                      }}
                      table.pivot-table td {{
                        border-bottom: 1px solid #e2e8f0; 
                        padding: 10px 12px; 
                        font-size: 14px;
                        color: #334155;
                      }}
                      table.pivot-table tr:hover {{
                        background-color: #f8fafc;
                      }}
                      .print-btn {{
                        display: inline-block; 
                        margin-bottom: 16px; 
                        padding: 10px 18px; 
                        background: #1b8a5a; 
                        color: white; 
                        border-radius: 6px; 
                        cursor: pointer; 
                        text-decoration: none;
                        font-weight: 500;
                        font-size: 14px;
                        border: none;
                        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                        transition: background 0.2s;
                      }}
                      .print-btn:hover {{
                        background: #0d5c3a;
                      }}
                      @media print {{
                        .print-btn {{ display: none; }}
                        body {{ padding: 0; }}
                      }}
                    </style>
                  </head>
                  <body>
                    <button class="print-btn" onclick="window.print()">🖨️ Click to Print Report</button>
                    <div class="header">
                      <div class="title">M-Pesa Cash Flow Statement Summary</div>
                      <div class="meta">Document: {st.session_state['uploaded_file_name']} | Total Transactions: {total_transactions:,}</div>
                    </div>
                    {pivot_html_table}
                  </body>
                </html>
                """
                # Render printable HTML inside app
                st.components.v1.html(printable_html, height=600, scrolling=True)

# Premium Footer
st.markdown("""
<div class="app-footer">
    Mpesa Statement Analyzer &copy; 2026. Made with &hearts; for Relationship Officers.
</div>
""", unsafe_allow_html=True)

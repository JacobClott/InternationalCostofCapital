"""
Project Cost of Capital Comparator: USD vs. local currency
Compares cost of equity, cost of debt and WACC for a project in USD and in the
currency of a selected country.  Data: Damodaran (betas, BBB spread), OECD
(corporate tax rates), Trading Economics (10Y government yields).
"""
import io
import re
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------
BETAS_URL = "https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls"
RATINGS_URL = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html"
OECD_URL = (
    "https://www.oecd.org/en/publications/corporate-tax-statistics-2024_9c27d6e8-en/"
    "full-report/statutory-corporate-income-tax-rates_440f9286.html"
)
TE_BONDS_URL = "https://tradingeconomics.com/bonds"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; WACC-Comparator/1.0)"}

DEFAULT_INDUSTRY = "Software (System & Application)"
DEFAULT_COUNTRY = "United States"
DEFAULT_ERP = 6.5
DEFAULT_DEBT = 30.0

# ---- Emergency fallbacks (ILLUSTRATIVE PLACEHOLDERS, not live data) --------
# Used only when a live fetch fails.  Refresh by reloading the app or override
# manually in the sidebar.  Tax rates are approximate OECD 2024 combined
# statutory rates; yields are rough placeholders and are NOT current.
SNAPSHOT_AS_OF = "illustrative placeholder, verify before use"
SNAPSHOT_BBB = 0.0111
SNAPSHOT_TAX = {
    "United States": 25.8, "Canada": 26.1, "Mexico": 30.0, "United Kingdom": 25.0,
    "Germany": 29.9, "France": 25.8, "Italy": 27.8, "Spain": 25.0, "Netherlands": 25.8,
    "Belgium": 25.0, "Austria": 23.0, "Portugal": 31.5, "Ireland": 12.5, "Finland": 20.0,
    "Greece": 22.0, "Sweden": 20.6, "Norway": 22.0, "Denmark": 22.0, "Switzerland": 14.7,
    "Poland": 19.0, "Czech Republic": 21.0, "Hungary": 9.0, "Turkey": 25.0, "Israel": 23.0,
    "Japan": 29.7, "South Korea": 26.4, "Australia": 30.0, "New Zealand": 28.0,
    "China": 25.0, "India": 25.2, "Indonesia": 22.0, "Malaysia": 24.0, "Thailand": 20.0,
    "Philippines": 25.0, "Vietnam": 20.0, "Singapore": 17.0, "Hong Kong": 16.5,
    "Brazil": 34.0, "Chile": 27.0, "Colombia": 35.0, "Peru": 29.5, "Argentina": 35.0,
    "South Africa": 27.0, "Egypt": 22.5, "Nigeria": 30.0, "Saudi Arabia": 20.0,
}
SNAPSHOT_YIELD = {
    "United States": 4.2, "Canada": 3.4, "Mexico": 9.5, "United Kingdom": 4.5,
    "Germany": 2.7, "France": 3.4, "Italy": 3.6, "Spain": 3.3, "Netherlands": 2.9,
    "Belgium": 3.3, "Austria": 3.1, "Portugal": 3.2, "Ireland": 3.0, "Finland": 3.2,
    "Greece": 3.4, "Sweden": 2.6, "Norway": 3.8, "Denmark": 2.7, "Switzerland": 0.4,
    "Poland": 5.5, "Czech Republic": 4.3, "Hungary": 6.8, "Turkey": 28.0, "Israel": 4.4,
    "Japan": 1.9, "South Korea": 3.0, "Australia": 4.4, "New Zealand": 4.6,
    "China": 1.8, "India": 6.5, "Indonesia": 6.6, "Malaysia": 3.5, "Thailand": 2.0,
    "Philippines": 6.1, "Vietnam": 3.5, "Singapore": 2.0, "Hong Kong": 3.3,
    "Brazil": 13.5, "Chile": 5.5, "Colombia": 12.0, "Peru": 6.3, "Argentina": 30.0,
    "South Africa": 9.0, "Egypt": 27.0, "Nigeria": 19.0, "Saudi Arabia": 5.0,
}
CURRENCY = {
    "United States": "USD", "Canada": "CAD", "Mexico": "MXN", "United Kingdom": "GBP",
    "Germany": "EUR", "France": "EUR", "Italy": "EUR", "Spain": "EUR", "Netherlands": "EUR",
    "Belgium": "EUR", "Austria": "EUR", "Portugal": "EUR", "Ireland": "EUR", "Finland": "EUR",
    "Greece": "EUR", "Sweden": "SEK", "Norway": "NOK", "Denmark": "DKK", "Switzerland": "CHF",
    "Poland": "PLN", "Czech Republic": "CZK", "Hungary": "HUF", "Turkey": "TRY",
    "Israel": "ILS", "Japan": "JPY", "South Korea": "KRW", "Australia": "AUD",
    "New Zealand": "NZD", "China": "CNY", "India": "INR", "Indonesia": "IDR",
    "Malaysia": "MYR", "Thailand": "THB", "Philippines": "PHP", "Vietnam": "VND",
    "Singapore": "SGD", "Hong Kong": "HKD", "Brazil": "BRL", "Chile": "CLP",
    "Colombia": "COP", "Peru": "PEN", "Argentina": "ARS", "South Africa": "ZAR",
    "Egypt": "EGP", "Nigeria": "NGN", "Saudi Arabia": "SAR",
}
# Name harmonisation between OECD / Trading Economics / this app
ALIASES = {
    "türkiye": "Turkey", "turkiye": "Turkey", "korea": "South Korea",
    "republic of korea": "South Korea", "czechia": "Czech Republic",
    "united states of america": "United States", "usa": "United States",
    "u.s.": "United States", "uk": "United Kingdom", "slovak republic": "Slovakia",
    "hong kong (china)": "Hong Kong", "people's republic of china": "China",
}


def norm_country(name: str) -> str:
    n = re.sub(r"\s+", " ", str(name)).strip()
    return ALIASES.get(n.lower(), n)


# --------------------------------------------------------------------------
# Data fetchers (cached; each returns data + a status string)
# --------------------------------------------------------------------------
def _get(url, timeout=25):
    r = requests.get(url, headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    return r


def _to_frac(x):
    """Convert 0.0111 or 1.11 (percent) or '1.11%' to a fraction."""
    if isinstance(x, str):
        x = float(x.replace("%", "").replace(",", "").strip()) / 100
    x = float(x)
    return x / 100 if x > 1 else x


def parse_betas_bytes(content: bytes):
    """Parse Damodaran betas.xls bytes -> (industry df, BBB spread or None)."""
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    ind_df = None
    for _, raw in sheets.items():
        head = raw.head(40)
        hit = [
            (i, j)
            for i in range(len(head))
            for j in range(head.shape[1])
            if isinstance(head.iat[i, j], str) and head.iat[i, j].strip().lower() == "industry name"
        ]
        if hit:
            r, _ = hit[0]
            df = raw.iloc[r + 1:].copy()
            df.columns = [str(c).strip() for c in raw.iloc[r]]
            ind_df = df
            break
    if ind_df is None:
        raise ValueError("Industry table not found in betas.xls")

    def col(*keys):
        for c in ind_df.columns:
            lc = c.lower()
            if all(k in lc for k in keys):
                return c
        return None

    c_name = col("industry")
    c_firms = col("number")
    c_beta = col("beta")
    c_de = col("d/e")
    c_unl_cash = col("unlevered", "cash")
    c_unl = next((c for c in ind_df.columns if c.lower().startswith("unlevered beta") and "cash" not in c.lower()), None)
    out = pd.DataFrame({
        "Industry": ind_df[c_name],
        "Firms": pd.to_numeric(ind_df[c_firms], errors="coerce") if c_firms else None,
        "Levered beta": pd.to_numeric(ind_df[c_beta], errors="coerce") if c_beta else None,
        "D/E": pd.to_numeric(ind_df[c_de], errors="coerce") if c_de else None,
        "Unlevered beta": pd.to_numeric(ind_df[c_unl], errors="coerce") if c_unl else None,
        "Unlevered beta (cash-corrected)": pd.to_numeric(ind_df[c_unl_cash], errors="coerce") if c_unl_cash else None,
    })
    out = out[out["Industry"].astype(str).str.strip().ne("") & out["Industry"].notna()]
    out = out[~out["Industry"].astype(str).str.lower().str.contains("total market")]
    out = out.dropna(subset=["Unlevered beta", "Unlevered beta (cash-corrected)"], how="all")
    out["Industry"] = out["Industry"].astype(str).str.strip()
    out = out.reset_index(drop=True)

    # Look for a BBB spread anywhere in the workbook
    bbb = None
    for _, raw in sheets.items():
        for i in range(raw.shape[0]):
            row = raw.iloc[i].tolist()
            for j, v in enumerate(row):
                if isinstance(v, str) and ("baa2/bbb" in v.lower() or v.strip().upper() == "BBB"):
                    for w in row[j + 1:]:
                        try:
                            f = _to_frac(w)
                            if 0 < f < 0.15:
                                bbb = f
                                break
                        except (ValueError, TypeError):
                            continue
                if bbb is not None:
                    break
            if bbb is not None:
                break
        if bbb is not None:
            break
    return out, bbb


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_betas():
    try:
        df, bbb = parse_betas_bytes(_get(BETAS_URL, 40).content)
        return df, bbb, "live"
    except Exception as e:  # noqa: BLE001
        return None, None, f"failed: {type(e).__name__}"


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_bbb_from_ratings():
    """Fallback: BBB default spread from Damodaran's ratings page (Baa2/BBB row)."""
    try:
        tables = pd.read_html(io.StringIO(_get(RATINGS_URL).text))
        for t in tables:
            for _, row in t.iterrows():
                vals = [str(v) for v in row.tolist()]
                for k, v in enumerate(vals):
                    if "baa2/bbb" in v.lower():
                        for w in vals[k + 1:]:
                            if "%" in w:
                                return _to_frac(w), "live"
        return None, "not found"
    except Exception as e:  # noqa: BLE001
        return None, f"failed: {type(e).__name__}"


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def fetch_oecd_tax():
    """Best-effort parse of OECD statutory CIT table; falls back to snapshot."""
    try:
        tables = pd.read_html(io.StringIO(_get(OECD_URL, 40).text))
        for t in tables:
            cols = [str(c).lower() for c in t.columns]
            ccol = next((c for c, lc in zip(t.columns, cols) if "jurisdiction" in lc or "country" in lc), None)
            if ccol is None:
                continue
            vcols = [c for c in t.columns if c != ccol and "2024" in str(c)] or \
                    [c for c in t.columns if c != ccol]
            if not vcols:
                continue
            vc = vcols[-1]
            s = pd.to_numeric(t[vc], errors="coerce")
            if s.notna().sum() > 20:
                d = {norm_country(n): float(v) for n, v in zip(t[ccol], s) if pd.notna(v)}
                if "United States" in d:
                    return d, "live"
        return None, "table not found in page HTML"
    except Exception as e:  # noqa: BLE001
        return None, f"failed: {type(e).__name__}"


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_te_table():
    """Trading Economics /bonds table (often JS-rendered, so may be empty)."""
    try:
        res = {}
        for t in pd.read_html(io.StringIO(_get(TE_BONDS_URL).text)):
            cols = [str(c).lower() for c in t.columns]
            if "yield" not in cols:
                continue
            yc = t.columns[cols.index("yield")]
            for n, v in zip(t.iloc[:, 0], t[yc]):
                v = pd.to_numeric(v, errors="coerce")
                if pd.notna(v):
                    res[norm_country(n)] = float(v)
        return res or None
    except Exception:  # noqa: BLE001
        return None


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_te_country(country: str):
    """10Y yield (%) from the country page's static meta description."""
    try:
        slug = country.lower().replace(" ", "-")
        html = _get(f"https://tradingeconomics.com/{slug}/government-bond-yield").text
        m = re.search(r'<meta[^>]+name="description"[^>]+content="([^"]+)"', html, re.I)
        text = m.group(1) if m else html[:5000]
        m2 = re.search(r"(?:to|at|from)\s+(-?\d+(?:\.\d+)?)\s*(?:%|percent)", text)
        if m2:
            return float(m2.group(1))
    except Exception:  # noqa: BLE001
        pass
    return None


def get_yield(country: str):
    """-> (yield %, source label)."""
    tbl = fetch_te_table()
    if tbl and country in tbl:
        return tbl[country], "live (Trading Economics)"
    v = fetch_te_country(country)
    if v is not None:
        return v, "live (Trading Economics)"
    if country in SNAPSHOT_YIELD:
        return SNAPSHOT_YIELD[country], f"PLACEHOLDER ({SNAPSHOT_AS_OF})"
    return None, "unavailable"


# --------------------------------------------------------------------------
# Finance calculations
# --------------------------------------------------------------------------
def compute(beta_u, tax, debt_share, erp, rf_usd, rf_loc, spread_usd, spread_loc):
    """All rates as decimals. Returns dict of results for USD and local currency."""
    d = debt_share
    de = d / (1 - d) if d < 1 else float("inf")
    beta_l = beta_u * (1 + (1 - tax) * de)                      # Hamada relevering
    ke_usd = rf_usd + beta_l * erp                              # CAPM in USD
    ke_loc = (1 + ke_usd) * (1 + rf_loc) / (1 + rf_usd) - 1     # risk-free differential conversion
    kd_usd = rf_usd + spread_usd
    kd_loc = rf_loc + spread_loc
    wacc_usd = (1 - d) * ke_usd + d * kd_usd * (1 - tax)
    wacc_loc = (1 - d) * ke_loc + d * kd_loc * (1 - tax)
    return {
        "D/E": de, "Relevered beta": beta_l,
        "Cost of equity (USD)": ke_usd, "Cost of equity (local)": ke_loc,
        "Pre-tax cost of debt (USD)": kd_usd, "Pre-tax cost of debt (local)": kd_loc,
        "After-tax cost of debt (USD)": kd_usd * (1 - tax),
        "After-tax cost of debt (local)": kd_loc * (1 - tax),
        "WACC (USD)": wacc_usd, "WACC (local)": wacc_loc,
    }


def pct(x, nd=2):
    return f"{x * 100:.{nd}f}%"


# --------------------------------------------------------------------------
# UI helpers
# --------------------------------------------------------------------------
def override_input(label, auto_value, key, step, fmt="%.2f", help_text=None,
                   min_value=None, max_value=None, auto_label="Auto"):
    """Checkbox 'override' + number input. Returns (value, is_manual)."""
    manual = st.checkbox(f"Override: {label}", key=f"ov_{key}", help=help_text)
    if auto_value is None:
        auto_value = 0.0
    if manual:
        val = st.number_input(label, value=float(st.session_state.get(f"val_{key}", auto_value)),
                              step=step, format=fmt, key=f"val_{key}",
                              min_value=min_value, max_value=max_value,
                              label_visibility="collapsed")
        return val, True
    st.number_input(f"{label} ({auto_label})", value=float(auto_value), step=step, format=fmt,
                    disabled=True, key=f"auto_{key}_{round(float(auto_value), 6)}",
                    label_visibility="collapsed")
    return float(auto_value), False


# --------------------------------------------------------------------------
# Page
# --------------------------------------------------------------------------
st.set_page_config(page_title="Cost of Capital: USD vs. Local Currency", page_icon="📊", layout="wide")
st.title("📊 Project Cost of Equity & WACC: USD vs. Local Currency")
st.caption("Same project, same tax rate, debt share and equity risk premium, expressed in two currencies.")

# ---- Load data -----------------------------------------------------------
with st.spinner("Fetching market data…"):
    betas_df, bbb_live, betas_status = fetch_betas()
    oecd_tax, oecd_status = fetch_oecd_tax()
    if bbb_live is None:
        bbb_live, bbb_status = fetch_bbb_from_ratings()
    else:
        bbb_status = "live"

# Manual upload fallback for betas.xls
upload = None
if betas_df is None:
    st.warning(
        "Could not download Damodaran's betas.xls (%s). Download it from the link in the sidebar "
        "and upload it here, or use the manual unlevered-beta override." % betas_status
    )
    upload = st.file_uploader("Upload betas.xls", type=["xls", "xlsx"])
    if upload is not None:
        try:
            betas_df, up_bbb = parse_betas_bytes(upload.getvalue())
            betas_status = "uploaded file"
            if up_bbb is not None:
                bbb_live, bbb_status = up_bbb, "uploaded file"
        except Exception as e:  # noqa: BLE001
            st.error(f"Could not parse uploaded file: {e}")
if betas_df is None:
    betas_df = pd.DataFrame({
        "Industry": [DEFAULT_INDUSTRY], "Firms": [None], "Levered beta": [None], "D/E": [None],
        "Unlevered beta": [1.0], "Unlevered beta (cash-corrected)": [1.0],
    })
    betas_status = "PLACEHOLDER (1.00)"

if bbb_live is None:
    bbb_live, bbb_status = SNAPSHOT_BBB, f"PLACEHOLDER ({SNAPSHOT_AS_OF})"

tax_table = dict(SNAPSHOT_TAX)
tax_source = f"PLACEHOLDER ({SNAPSHOT_AS_OF})"
if oecd_tax:
    tax_table.update(oecd_tax)
    tax_source = "live (OECD)"

countries = sorted(set(tax_table) | set(SNAPSHOT_YIELD) | {DEFAULT_COUNTRY})
industries = sorted(betas_df["Industry"].tolist())

# ---- Sidebar -------------------------------------------------------------
sb = st.sidebar
sb.header("Core inputs")
industry = sb.selectbox(
    "Industry", industries,
    index=industries.index(DEFAULT_INDUSTRY) if DEFAULT_INDUSTRY in industries else 0,
)
country = sb.selectbox("Target country", countries, index=countries.index(DEFAULT_COUNTRY))
erp_pct = sb.number_input("Equity risk premium (%)", value=DEFAULT_ERP, min_value=0.0,
                          max_value=30.0, step=0.1, format="%.2f")
debt_pct = sb.slider("Target debt share, D/(D+E) (%)", 0.0, 95.0, DEFAULT_DEBT, 1.0)
beta_basis = sb.radio(
    "Unlevered beta basis", ["Cash-corrected", "Standard"], horizontal=True,
    help="Damodaran reports unlevered beta and unlevered beta corrected for cash. "
         "Cash-corrected is the operating-business beta.",
)

sb.header("Project tax rate")
if "tax_choice" not in st.session_state:
    st.session_state.tax_choice = "U.S."
is_us = country == DEFAULT_COUNTRY
tcol1, tcol2 = sb.columns(2)
if tcol1.button("U.S. rate", width="stretch",
                type="primary" if st.session_state.tax_choice == "U.S." else "secondary"):
    st.session_state.tax_choice = "U.S."
    st.rerun()
if tcol2.button("Country rate", width="stretch", disabled=is_us,
                type="primary" if st.session_state.tax_choice == "Country" else "secondary"):
    st.session_state.tax_choice = "Country"
    st.rerun()
tax_choice = "U.S." if is_us else st.session_state.tax_choice
sb.caption(f"Active: **{'U.S.' if tax_choice == 'U.S.' else country}** corporate tax rate")

# Auto-fetched values for selected inputs
row = betas_df[betas_df["Industry"] == industry].iloc[0]
beta_col = "Unlevered beta (cash-corrected)" if beta_basis == "Cash-corrected" else "Unlevered beta"
auto_beta = row[beta_col] if pd.notna(row[beta_col]) else row.get("Unlevered beta", 1.0)
auto_tax_us = tax_table.get("United States", 25.8)
auto_tax_c = tax_table.get(country)
auto_y_us, y_us_src = get_yield("United States")
auto_y_c, y_c_src = get_yield(country)

sb.header("Manual overrides")
sb.caption("Tick a box to replace the fetched value.")
with sb:
    beta_u, beta_manual = override_input("Unlevered beta", auto_beta, "beta", 0.01, "%.3f")
    tax_us, tax_us_manual = override_input("U.S. corporate tax rate (%)", auto_tax_us, "tax_us", 0.1, "%.2f", min_value=0.0, max_value=100.0)
    if is_us:
        tax_c, tax_c_manual = tax_us, tax_us_manual
    else:
        tax_c, tax_c_manual = override_input(f"{country} corporate tax rate (%)", auto_tax_c, "tax_c", 0.1, "%.2f", min_value=0.0, max_value=100.0)
    spread_us_pct, spread_us_manual = override_input("USD BBB spread (%)", bbb_live * 100, "bbb_us", 0.05, "%.2f", min_value=0.0)
    y_us_pct, y_us_manual = override_input("U.S. 10Y yield (%)", auto_y_us, "y_us", 0.05, "%.2f")
    if is_us:
        y_c_pct, y_c_manual = y_us_pct, y_us_manual
    else:
        y_c_pct, y_c_manual = override_input(f"{country} 10Y yield (%)", auto_y_c, "y_c", 0.05, "%.2f")
    if is_us:
        spread_c_pct, spread_c_manual = spread_us_pct, False
    else:
        spread_c_pct, spread_c_manual = override_input(
            f"{country} BBB spread (%)", spread_us_pct, "bbb_c", 0.05, "%.2f", min_value=0.0,
            help_text="When off, the foreign BBB spread equals the USD BBB spread.",
            auto_label="auto = USD spread",
        )

sb.header("Data sources")
sb.markdown(
    f"[Damodaran betas]({BETAS_URL}) · [OECD tax rates]({OECD_URL}) · [Trading Economics yields]({TE_BONDS_URL})"
)

if auto_tax_c is None and not tax_c_manual and not is_us:
    st.error(f"No corporate tax rate available for {country}. Use the manual override.")
    st.stop()
if (y_us_pct is None or y_c_pct is None):
    st.error("Missing government yield. Use the manual override.")
    st.stop()

# ---- Calculation ---------------------------------------------------------
tax_active = (tax_us if tax_choice == "U.S." else tax_c) / 100
erp, debt = erp_pct / 100, debt_pct / 100
rf_usd, rf_loc = y_us_pct / 100, y_c_pct / 100
sp_usd, sp_loc = spread_us_pct / 100, spread_c_pct / 100
res = compute(beta_u, tax_active, debt, erp, rf_usd, rf_loc, sp_usd, sp_loc)
cur = CURRENCY.get(country, "Local")
loc_label = f"{country} ({cur})" if not is_us else "United States (USD)"

# ---- Data status ---------------------------------------------------------
with st.expander("Data status", expanded=any("PLACEHOLDER" in s or "failed" in s for s in
                                             [betas_status, bbb_status, tax_source, y_us_src, y_c_src])):
    stat = pd.DataFrame([
        ["Industry unlevered betas", "Manual override" if beta_manual else betas_status],
        ["USD BBB spread", "Manual override" if spread_us_manual else bbb_status],
        ["Corporate tax rates", "Manual override" if (tax_us_manual or tax_c_manual) else tax_source],
        ["U.S. 10Y yield", "Manual override" if y_us_manual else y_us_src],
        [f"{country} 10Y yield", "Manual override" if y_c_manual else y_c_src],
    ], columns=["Item", "Source"])
    st.dataframe(stat, hide_index=True, width="stretch")
    st.caption("Anything marked PLACEHOLDER is an illustrative fallback used only because the live fetch "
               "failed. Do not rely on it; override or refresh before using the results.")

# ---- Results -------------------------------------------------------------
st.subheader(f"Results: {industry}")
st.caption(
    f"Project tax rate: **{tax_active * 100:.2f}%** ({'U.S.' if tax_choice == 'U.S.' else country}) · "
    f"Debt share: **{debt_pct:.0f}%** (D/E {res['D/E']:.2f}) · ERP: **{erp_pct:.2f}%** · "
    f"Unlevered β: **{beta_u:.3f}** → relevered β: **{res['Relevered beta']:.3f}**"
)
c1, c2 = st.columns(2)
for col, title, sfx, rf, sp in [(c1, "USD", "USD", rf_usd, sp_usd), (c2, loc_label, "local", rf_loc, sp_loc)]:
    with col:
        st.markdown(f"#### {title}")
        m1, m2, m3 = st.columns(3)
        m1.metric("Cost of equity", pct(res[f"Cost of equity ({sfx})"]))
        m2.metric("After-tax cost of debt", pct(res[f"After-tax cost of debt ({sfx})"]))
        m3.metric("WACC", pct(res[f"WACC ({sfx})"]))
        st.write(f"Risk-free: **{pct(rf)}** · BBB spread: **{pct(sp)}** · "
                 f"Pre-tax cost of debt: **{pct(res[f'Pre-tax cost of debt ({sfx})'])}**")

summary = pd.DataFrame({
    "Metric": ["Risk-free rate (10Y)", "BBB spread", "Cost of equity", "Pre-tax cost of debt",
               "After-tax cost of debt", "WACC"],
    "USD": [rf_usd, sp_usd, res["Cost of equity (USD)"], res["Pre-tax cost of debt (USD)"],
            res["After-tax cost of debt (USD)"], res["WACC (USD)"]],
    loc_label: [rf_loc, sp_loc, res["Cost of equity (local)"], res["Pre-tax cost of debt (local)"],
                res["After-tax cost of debt (local)"], res["WACC (local)"]],
})
summary["Difference (local − USD)"] = summary[loc_label] - summary["USD"]
disp = summary.copy()
for c in disp.columns[1:]:
    disp[c] = disp[c].map(lambda v: f"{v * 100:.2f}%")
st.dataframe(disp, hide_index=True, width="stretch")

# ---- Assumptions ---------------------------------------------------------
with st.expander("Main assumptions & formulas", expanded=True):
    st.markdown(f"""
- **Beta:** industry unlevered beta ({beta_basis.lower()}, Damodaran, US firms) relevered at the target
  D/E using the project tax rate: βL = βU × [1 + (1 − t) × D/E]. Debt beta is assumed to be zero.
  The same beta is used in both currencies.
- **Tax rate:** one project tax rate (U.S. or {country}) is used for relevering and the debt tax shield
  in *both* currencies. Toggle it in the sidebar.
- **USD cost of equity:** CAPM, Ke = U.S. 10Y yield + βL × ERP.
- **Local-currency cost of equity:** converted through the risk-free differential,
  (1 + Ke<sub>local</sub>) = (1 + Ke<sub>USD</sub>) × (1 + Rf<sub>local</sub>) / (1 + Rf<sub>USD</sub>),
  so the equity risk premium is carried over multiplicatively rather than re-estimated.
- **Cost of debt:** 10Y government yield of the currency + BBB spread. The local BBB spread equals the
  USD BBB spread unless you override it.
- **WACC:** (1 − D/V) × Ke + D/V × Kd × (1 − t), with the same debt share and ERP in both currencies.
- **Government yields** are used as risk-free rates. They include any sovereign default spread, which
  is *not* stripped out. For high-yield countries, interpret the local-currency figures with that in mind.
- **Data:** betas and BBB spread from Damodaran, statutory corporate tax rates from OECD, and 10Y yields
  from Trading Economics. Manual overrides take priority over fetched values.
""", unsafe_allow_html=True)

# ---- All-industry comparison --------------------------------------------
st.subheader("All-industry comparison")
st.caption("Uses the current country, tax rate, ERP, debt share, yields and spreads. Each industry's own "
           "unlevered beta is used (the manual beta override applies only to the selected industry above).")
rows = []
for _, r in betas_df.iterrows():
    b = r[beta_col] if pd.notna(r[beta_col]) else r.get("Unlevered beta")
    if pd.isna(b):
        continue
    o = compute(float(b), tax_active, debt, erp, rf_usd, rf_loc, sp_usd, sp_loc)
    rows.append({
        "Industry": r["Industry"], "Unlevered beta": float(b), "Relevered beta": o["Relevered beta"],
        "Cost of equity (USD)": o["Cost of equity (USD)"], "Cost of equity (local)": o["Cost of equity (local)"],
        "WACC (USD)": o["WACC (USD)"], "WACC (local)": o["WACC (local)"],
    })
allind = pd.DataFrame(rows).sort_values("WACC (USD)", ascending=False).reset_index(drop=True)
fmt_ind = allind.copy()
for c in ["Cost of equity (USD)", "Cost of equity (local)", "WACC (USD)", "WACC (local)"]:
    fmt_ind[c] = fmt_ind[c].map(lambda v: f"{v * 100:.2f}%")
fmt_ind["Unlevered beta"] = fmt_ind["Unlevered beta"].map(lambda v: f"{v:.3f}")
fmt_ind["Relevered beta"] = fmt_ind["Relevered beta"].map(lambda v: f"{v:.3f}")
st.dataframe(fmt_ind, hide_index=True, width="stretch", height=420)
if len(allind) > 1:
    chart = allind.set_index("Industry")[["WACC (USD)", "WACC (local)"]] * 100
    chart.columns = ["WACC USD (%)", f"WACC {cur} (%)" if not is_us else "WACC local (%)"]
    st.bar_chart(chart, height=380)

# ---- Downloads -----------------------------------------------------------
st.subheader("Download results")
inputs = pd.DataFrame({
    "Input": ["Industry", "Target country", "Unlevered beta basis", "Unlevered beta", "Project tax rate used (%)",
              "Tax basis", "Equity risk premium (%)", "Debt share (%)", "U.S. 10Y yield (%)",
              f"{country} 10Y yield (%)", "USD BBB spread (%)", f"{country} BBB spread (%)", "Generated"],
    "Value": [industry, country, beta_basis, round(beta_u, 4), round(tax_active * 100, 3),
              "U.S." if tax_choice == "U.S." else country, erp_pct, debt_pct, y_us_pct, y_c_pct,
              round(spread_us_pct, 4), round(spread_c_pct, 4), datetime.now().strftime("%Y-%m-%d %H:%M")],
})
d1, d2 = st.columns(2)
csv_buf = io.StringIO()
csv_buf.write("# Inputs\n")
inputs.to_csv(csv_buf, index=False)
csv_buf.write("\n# Results (decimal rates)\n")
summary.to_csv(csv_buf, index=False)
csv_buf.write("\n# All industries (decimal rates)\n")
allind.to_csv(csv_buf, index=False)
d1.download_button("⬇️ Download CSV", csv_buf.getvalue(), "cost_of_capital_results.csv", "text/csv",
                   width="stretch")
xbuf = io.BytesIO()
with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
    inputs.to_excel(xw, sheet_name="Inputs", index=False)
    summary.to_excel(xw, sheet_name="Results", index=False)
    allind.to_excel(xw, sheet_name="All industries", index=False)
d2.download_button("⬇️ Download Excel", xbuf.getvalue(), "cost_of_capital_results.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   width="stretch")
st.caption("For analysis and education only; not investment advice. Verify inputs against source data.")

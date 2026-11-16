# Project Cost of Capital: USD vs. Local Currency

A Streamlit app that compares a project's **cost of equity, cost of debt and WACC** in USD and in the
currency of any selected country, using the same project tax rate, debt share and equity risk premium.

## Defaults
Industry: Software (System & Application) · Country: United States · Tax: U.S. rate · ERP 6.5% · Debt share 30%.

## Data sources (fetched live, cached)
| Input | Source |
|---|---|
| Industry unlevered betas | [Damodaran betas.xls](https://pages.stern.nyu.edu/~adamodar/pc/datasets/betas.xls) |
| USD BBB spread | Searched for in betas.xls; if absent, Damodaran's [ratings page](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html) (Baa2/BBB row) |
| Corporate tax rates | [OECD Corporate Tax Statistics 2024](https://www.oecd.org/en/publications/corporate-tax-statistics-2024_9c27d6e8-en/full-report/statutory-corporate-income-tax-rates_440f9286.html) |
| 10Y government yields | [Trading Economics](https://tradingeconomics.com/bonds) (bonds table, then per-country page) |

**Fallbacks.** If a live fetch fails, the app uses an embedded **illustrative placeholder** (tax rates and
yields for ~45 countries, a 1.11% BBB spread) and shows a warning in the *Data status* panel. These
placeholders are not current market data. For betas, you can upload betas.xls manually or use the beta override.
Trading Economics and OECD pages are partly JavaScript-rendered, so live parsing may fail; use the overrides.

## Method
- Relevered beta: `βL = βU × [1 + (1 − t) × D/E]` with `D/E = D% / (1 − D%)`, using the project tax rate.
- USD cost of equity: `Ke = Rf_US + βL × ERP`.
- Local cost of equity: `(1 + Ke_local) = (1 + Ke_USD) × (1 + Rf_local) / (1 + Rf_US)`.
- Cost of debt: `Kd = 10Y yield + BBB spread`. The local BBB spread equals the USD spread unless overridden.
- `WACC = (1 − D%) × Ke + D% × Kd × (1 − t)`.

## Features
Industry/country selection, ERP and debt-share inputs, U.S./country tax toggle, overrides for unlevered beta,
tax rates, USD BBB spread, both 10Y yields and the foreign BBB spread, side-by-side results, assumptions,
all-industry comparison (table + chart), CSV and Excel downloads.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Push `app.py`, `requirements.txt` and `README.md` to a GitHub repo.
2. At [share.streamlit.io](https://share.streamlit.io), choose **New app**, select the repo and branch, and set the main file to `app.py`.
3. Deploy. Data refreshes automatically (betas/BBB every 6h, yields hourly, tax rates daily).

## Notes & limitations
- Government yields are used as risk-free rates and include any sovereign default spread.
- Damodaran betas are US-firm industry averages; the same beta is used in both currencies.
- For education and analysis only; verify inputs before relying on results.

# AppStock Screener (Google Apps Script)

Stock screener for Google Sheets with:
- Universe builder from a Seed sheet
- Parallel real‑time updater (waves + retry/backoff)
- Batched quarterly updater with live progress + resume
- Append‑only Value/Growth/Non‑Profitable screeners
- Manual Add sidebar (ticker + add to Universe)
- README sheet generator and daily schedule installer

## Setup
1) In your Google Sheet, open Extensions → Apps Script and paste `Code.js`.
2) Run “Screener → Set FMP API Key” and save your FinancialModelingPrep key.
3) Create a `Seed` sheet with tickers in column A, names in B, optional price in G.
4) Run “Screener → Reset Universe”, then “Refresh from Seed → Universe”.

## Updating Data
- Real‑Time: “Update Real‑Time Data (parallel)” fetches price, market cap (billions), P/E, PEGs, P/S, P/B, and Industry PE (yesterday snapshot).
- Quarterly: “Update Quarterly Data (batch/resume)” updates D/E, ROE, ROCE, margins, current ratio, EV/Sales, OCF (sum last 4Q), Insider %, growth (true quarter math), and annual CAGRs (3Y/5Y/10Y for Revenue and EPS).

### Growth Math (Quarterly)
Uses the quarterly income statement (not financial‑growth):
- Revenue YoY = (Rev Q0 − Rev Q−4) / |Rev Q−4|
- EPS YoY = (EPS Q0 − EPS Q−4) / |EPS Q−4|
- Revenue QoQ = (Rev Q0 − Rev Q−1) / |Rev Q−1|
- EPS QoQ = (EPS Q0 − EPS Q−1) / |EPS Q−1|
EPS uses `eps`, then `epsdiluted` if missing. Values are written as numbers; format as % in Sheets if preferred.

### CAGR (Annual)
Uses the annual income statement:
- Revenue CAGR 3Y/5Y/10Y = (Rev t / Rev t−N)^(1/N) − 1
- EPS CAGR 3Y/5Y/10Y = (EPS t / EPS t−N)^(1/N) − 1
Only computed when start and end values are positive.

## Screeners (append‑only)
- Value: P/E<20; ROE(5Y)>15%; ROCE>15%; PEG<1; D/E<1; GM>30%; NM>10%; EPS YoY>20%; P/E < Industry PE
- Growth: EPS YoY>20%; PEG<1; P/E<40
- Non‑Profitable: Revenue YoY>30%; P/S<5; EV/Sales<5; D/E<1; Current Ratio>1.5; Market Cap>$500M

## Manual Add
Use “Manual Add (sidebar)” to enter a ticker, choose a target sheet (Value/Growth/Non‑Profitable), and optionally add it to Universe if missing. The row is built live using the same logic as Real‑Time + Quarterly.

## Notes
- Columns: “Institutional Ownership %” has been removed. “Insider Ownership %” is computed as Σ(securitiesOwned last 12 months) ÷ outstandingShares × 100.
- Install “Daily @ 9:05 AM ET” to run the full pipeline automatically.

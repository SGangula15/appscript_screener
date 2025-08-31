# AppStock Screener (Google Apps Script)

Google Sheets stock screener that builds a universe from your Seed tab, updates real‑time and quarterly fundamentals in bulk (with progress + resume), and maintains append‑only screeners (Value/Growth/Non‑Profitable). Includes a manual add flow and an in‑sheet README generator.

## Features
- Universe builder (from `Seed` tab)
- Real‑Time updater (parallel waves + retry/backoff) with live progress sidebar
- Quarterly updater (batched) with live progress + safe resume across runs
- Calculated fields:
  - Operating Cash Flow (TTM)/EBITDA
  - Operating Cash Flow Coverage Ratio (from annual ratios)
  - True quarter‑math growth: Revenue/EPS YoY and QoQ
  - Revenue/EPS CAGRs (3Y/5Y/10Y, annual)
- Append‑only screeners: Value, Growth, Non‑Profitable
- Manual Add (menu prompt) to add a single ticker to a screener and Universe
- README sheet generator and daily trigger installer (9:05 AM ET)

## Setup
1) In your Google Sheet, open Extensions → Apps Script and paste `Code.js`.
2) Set time zone: File → Project properties → Script time zone: America/Detroit (ET).
3) In the Sheet, create a `Seed` sheet with:
   - Column A: Ticker
   - Column B: Name
   - Column G: Price (optional)
4) Menu: Screener → Set FMP API Key, then paste your FinancialModelingPrep API key.
5) Menu: Screener → Reset Universe, then Screener → Refresh from Seed → Universe.

## Menu Reference
- Set FMP API Key: store the FMP key in script properties.
- Reset Universe: recreate the Universe header and reset screener headers.
- Refresh from Seed → Universe: populate Universe from `Seed` (A/B/G).
- Update Real‑Time Data (parallel): fetch price, MC (billions), P/E, Forward P/E, PEGs, P/S, P/B, and Industry PE snapshot (~3 days ago). Live progress sidebar included.
- Update Quarterly Data (batch/resume): update D/E, ROE (TTM), ROE (5Y avg), ROCE, Gross/Net margins, Current Ratio, EV/Sales, Operating Cash Flow (TTM)/EBITDA, Operating Cash Flow Coverage Ratio, Insider %, growth (YoY/QoQ), Revenue/EPS CAGRs. Live progress + resume.
- Reset Quarterly Progress: clear saved progress counters if a run is stuck.
- Update Value / Growth / Non‑Profitable: append new qualifying rows without overwriting existing rows.
- Add Ticker to Value/Growth/Non‑Profitable (Universe copy): copy an existing Universe row.
- Add Ticker Manually: build a row live via APIs and append to a screener; optionally add to Universe.
- Install Daily @ 9:05 AM ET: install the daily automation pipeline.
- Create/Update README: generate the in‑sheet README overview.

## Updating Data
### Real‑Time
- Runs in waves (configurable), with retry/backoff. A live sidebar shows progress.
- If API quota is reached, the run pauses. Run again later to continue. Industry PE snapshot is skipped when paused.

### Quarterly (Batch + Resume)
- Processes tickers in slices and respects a time budget; when nearing the budget it pauses safely and saves progress (offset + last total) so the next run resumes where it left off (unless the Universe size changed).
- Insider ownership is optional during quarterly runs (`INCLUDE_INSIDER_DURING_QUARTERLY`).
- Key computed fields:
  - OCF (TTM)/EBITDA = (sum of last 4 quarters OCF) / income‑statement‑ttm.ebitda
  - Operating Cash Flow Coverage Ratio = annual `ratios.operatingCashFlowCoverageRatio`
  - Growth (YoY/QoQ) computed from quarterly income statements (true quarter math)
  - Revenue/EPS CAGRs from annual income statements

### Growth Math Details
- Revenue YoY = (Rev Q0 − Rev Q−4) / |Rev Q−4|
- EPS YoY = (EPS Q0 − EPS Q−4) / |EPS Q−4|
- Revenue QoQ = (Rev Q0 − Rev Q−1) / |Rev Q−1|
- EPS QoQ = (EPS Q0 − EPS Q−1) / |EPS Q−1|
- EPS field precedence: `eps`, then `epsdiluted`/`epsDiluted`/`epsDilutedGAAP`.

### CAGR (Annual)
- Revenue CAGR N = (Rev t / Rev t−N)^(1/N) − 1 (for N in 3,5,10)
- EPS CAGR N = (EPS t / EPS t−N)^(1/N) − 1 (for N in 3,5,10)
- Only computed when start and end are positive.

## Screeners (append‑only)
- Value: P/E<20; ROE(5Y)>15%; ROCE>15%; PEG<1; D/E<1; GM>30%; NM>10%; EPS YoY>20%; P/E < Industry PE
- Growth: EPS YoY>20%; PEG<1; P/E<40
- Non‑Profitable: Revenue YoY>30%; P/S<5; EV/Sales<5; D/E<1; Current Ratio>1.5; Market Cap>$500M

## Manual Add
Use “Add Ticker Manually” (menu prompt) to enter a ticker, choose a target screener, and optionally add it to Universe. The row is constructed via the same APIs used by the bulk updaters.

## Notes
- “Insider Ownership %” = Σ(securitiesOwned last 12 months) ÷ outstandingShares × 100.
- Install “Daily @ 9:05 AM ET” to run the full pipeline automatically.

## Sharing the GitHub repository (private)
To grant access to specific people:
1) On GitHub, open the repo → Settings → Collaborators and teams (or “Manage access”).
2) Click “Add people”, enter their GitHub username or email, choose a role (e.g., Read, Triage, Write, Maintain, Admin), and invite.
3) They’ll receive an email to accept the invite.

For organizations, you can add users via a team and assign the team to the repo with appropriate permissions.

---
License: see LICENSE.

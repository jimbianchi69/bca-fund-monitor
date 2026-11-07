# BCA quarterly fund-monitoring runbook

Implements the **BCA Investment Monitoring Policy v1.0** (adopted 2026-11-06). Status is assigned only by the
published Toolkit triggers; this runbook moves data, applies the written rules, and drafts documents for a named
person to sign.

## Inputs each quarter
- Fidelity 401(k) Direct Program platform list (FPRS Direct Platform List xlsx), quarter-end data
- Broadridge Fiduciary Focus Toolkit access (Jim signs in)

## Steps
1. `python3 -I tools/build_screen.py <FPRS.xlsx> BCA_Fund_Screening_Fidelity_Q<q>_<yyyy>.xlsx` then recalc.
   Screens (policy §3): open to all investors · on platform · zero recordkeeping offset · lowest-cost zero-offset
   class · ER caps 0.75 / 0.50 / 0.75 / 0.35 / 0.50 (equity, core bond, HY-multisector, cap pres, target-date) ·
   ≥5-yr record · category maps to a slot.
2. If the candidate set changed, re-import the Toolkit list (Research → Lists → Import List, CUSIP column).
3. Investment Screener → Load Existing → `BCA 321 Screened Universe Quarterly Score Run` → View Results.
4. Apply Trigger `BCA 3(21) Acceptable - Score 50 Watch 2of4 Remove 4of4`; record Watch and Remove counts.
   Apply `BCA 3(38) Select - Score 25 and LPL 7 Watch 2of4 Remove 4of4`; filter Historical Trend 12 of 12
   (selection standard; policy §7 permits 10 of 12).
5. Generate Report → all investments → `System: Fi360 Fiduciary Score Summary`. Merge scores into the
   workbook's Candidates tab (cols Q–U).
6. Per client plan: monitoring report with the BCA trigger; flag lineup funds outside the Acceptable Universe.
7. Draft Watch-list notes, Remove recommendations/decision memos, §7 fallback log entries; principal review; sign.
8. File to `_BCA Fund Monitoring`, client folder, Toolkit vault; update Salesforce policy record.

## Known constraints
- Covered Fund Lists override trigger results — never load the universe there.
- Published triggers lock; create a new dated trigger to change criteria.
- Toolkit sessions are per-tab and lock on idle.
- Report downloads land in the browser's Downloads folder; the Report Downloads queue holds generated reports.

## First run (data as of 2026-09-30)
7,316 share classes → 767 candidates + 43 target-date series. 3(21) trigger: 36 Watch / 0 Remove.
3(38) 12-of-12: 116 funds; no passers in Small Cap Growth or Capital Preservation.

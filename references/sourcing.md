# Sourcing: 421-a expiry screen

Use this before underwriting, when the deal lead asks for off-market targets rather than a review of a supplied deal. The screen finds rental buildings whose 421-a exemption is about to end, attaches the owner, and hands the deal lead a list for direct outreach. A building enters the underwriting steps in SKILL.md only after an owner engages and supplies a rent roll.

Scope set by the deal lead on November 18, 2026: start with Williamsburg; 421-a ending within one to three years; 50,000+ building SF as the primary size filter and 40+ residential units as the secondary; deals of roughly $50M to $150M and 50 to 150 units; owners approached directly, not through listings.

## Run

```
python3 scripts/source_421a.py --out <deal-folder>/sourcing-<area> --as-of YYYY-MM-DD
python3 scripts/source_report.py <deal-folder>/sourcing-<area> --name "<Area> 421-a targets.xlsx"
```

Defaults are Brooklyn Community District 1 (`--cd 301`) with primary zips 11211, 11249 and 11206. To move to another area, change `--cd` and `--zips`; nothing else is area-specific. Recalculate the workbook before sending it. The first script writes `universe.csv` and `run-log.json`; keep both with the workbook so the funnel can be audited.

## Sources and what each one decides

| Source (NYC Open Data id) | Decides |
|---|---|
| PLUTO (64uk-42ks) | Universe: building SF, residential units, class, owner of record |
| DOF Property Exemption Detail (muvi-b6kx) | 421-a code, benefit start year, term |
| DOF Exemption Classification Codes (myn9-hwsy) | Code meanings. 5114 is the 25-year and 5116 the 20-year benefit; do not assume the order |
| DOF Property Valuation and Assessment Data (8y4t-faws) | Condo number and owner for each condo unit lot |
| DOF Property Tax Rates by Tax Class (7zb8-7bpk) | Class 2 rate for the step-up estimate |
| HPD Registrations and Contacts (tesw-yqqr, feu5-w2e2) | Registered owner entity, head officer, managing agent, business addresses |
| ACRIS Legals, Master, Parties (8h5j-fqxa, bnx9-e6tj, 636b-3b5g) | Last deed and last mortgage: date, amount, parties |

## Rules

- **End of benefit.** Last exempt fiscal year = benefit start + term - 1, ending June 30. Calibrated on 26 Quincy Street (BBL 3019720010): start 2022, term 36, giving June 30, 2057, the date in the offering memorandum. Read the latest roll year at period 3.
- **Statuses.** In window: ends within the window. Next up: ends within the following three years. Expired: on the prior roll but ended. Later and no 421-a stay on the Universe tab only.
- **Condo-structured buildings.** The exemption sits on unit lots, so match unit lots to the building by condo number, never by block. Treat the building as a rental when the apartments are bundled into few lots or one owner holds 80% or more of the residential units. Individually owned condos stay out of Targets and Watchlist.
- **Tax step-up.** Exempt assessed value on the current roll times the class 2 rate. The benefit phases out in its last years, so this is the remaining increase, and it is a screening figure only.
- **Recent trades.** A deed with consideration in the last five years is noted on the row; the deal lead decides whether a recent buyer is still worth a call.

## Limits to state in the handoff

- End years are derived from DOF exemption records. Spot-check two or three tax bills before outreach.
- No public city dataset carries owner phone numbers. The workbook leaves a column for a paid lookup.
- HPD contacts are self-reported at annual registration and can lag a sale. Compare the registration date with the last deed.
- Multi-lot deeds show the full consideration on every lot they cover.

## Phase 2: back-of-napkin value (not built)

Per the deal lead: revenue from listed rents as a proxy, acknowledged to run high because long-term tenants pay less; operating expenses of $10,000 per unit plus real estate taxes, or $13 per SF; value at a 5.75% cap or higher; vacancy from units advertised on the building's site against the total. Use post-expiry taxes, since that is what a buyer will pay. Before building it, agree the rent source with the deal lead: listing sites restrict automated collection, so use a licensed feed, manual samples, or the owner's own site.

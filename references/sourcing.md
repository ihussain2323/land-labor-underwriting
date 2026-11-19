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

## Phase 2: back-of-napkin value

Purpose: rank the targets by rough value so the deal lead knows which owners to call first. It is not an underwriting and nothing from it goes into a model.

Put `rents.json` in the run folder and rerun `source_report.py`; the workbook gains a Napkin tab and a Rent evidence tab. Without the file the workbook is Phase 1 only. `scripts/source_napkin.py` documents the file format: one object per target with the advertised units (beds, asking rent, SF, source link), the count of units advertised, a confidence grade and notes.

Per the deal lead on November 18, 2026:

| Line | Rule |
|---|---|
| Revenue | Listed asking rents x units. Acknowledged to run high because long-term tenants pay less. A discount input is on the tab, default 0%. |
| Vacancy | Units advertised / total units, with a 3% floor. The floor is an added assumption. |
| Operating expenses | $10,000 per unit. The $13 per SF alternative is shown as a second value. It is read as replacing the $10,000, with RE tax still added; confirm that reading. |
| RE tax | Full post-expiry tax: (assessed total - all exemptions + 421-a exempt value) x class 2 rate. |
| Value | NOI / 5.75%. |

Rules added in the build:

- **Rent source.** Listing sites block automated collection. Use the building's or manager's own site first, then figures visible in search results. Record every unit on the Rent evidence tab with its link, grade the building high, medium or low, and never fill a gap with an invented rent.
- **Unit mix.** A straight average of advertised units is biased by which units happen to be listed. Each listing is divided by the area average for its bedroom count, the building's ratios are averaged into an index, and the index is applied to the area rent at an assumed mix (20% studio, 50% one-bedroom, 30% two-bedroom, editable). A building with no listings gets index 1.00 and is labelled "Area average".
- **Flags.** Note commercial space (taxed but no income counted), large units (per-unit rent understates), and condo structures (assessed value covers every lot).

Limits to state in the handoff:

- Asking rents are not in-place rents. Most of these buildings are rent-stabilized for the life of the 421-a, so in-place rents can sit well below asking.
- Affordable units, unit mix, commercial income and concessions are unknown until an owner sends a rent roll.
- Where a building traded recently, compare the napkin value with the deed price as a sanity check.

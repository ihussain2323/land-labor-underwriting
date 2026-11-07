# Running the skill

An analyst can invoke the skill with a property address, the room or deal email, and the template or its location. The agent performs these steps. Keep working JSON and extracted rows beside the individual deal. A shared skill package should contain no tenant records, credentials or personal computer paths.

**Tools.** Use Python 3.10 or newer. The local Python helpers use the libraries in requirements.txt. If needed, create an isolated environment and install with `python -m pip install -r requirements.txt` from the skill folder. Research uses only the standard library. Use a connected Excel session for live workbook edits. Scripts here read files and produce JSON on standard output. For a saved workbook, Excel must have performed a full recalculation and saved updated caches before capture.

**Deal record.** Copy assets/deal-start.json to the deal folder as deal.json and fill it from the source inventory. The template is intentionally incomplete. The JSON Schema is at schemas/deal.schema.json. The source key points to a record with location, as_of and scope. For a pending field point to a dated inventory of the documents checked. Each normalized rent-roll row has an ID, type, rsf, and separate contract, legal and market monthly rent. A null is permitted for unavailable quantities. Set leased to true, false or null. Record which files supplied dimensions, rent and occupancy in roll.basis. Reconcile unit counts and area against an independent control. For T-12, use the first month's ISO date, expense-only scope, total on the statement, and twelve monthly numbers per nonoverlapping category. For partial history leave t12 null, record the observed period and total in decisions, and flag the incomplete twelve months.

Add optional controls as decisions when supported. Use retail_area in SF, expected_year1_noi in dollars from an independent operating build, history_income_complete as a boolean, and unit_type_order as a list of types in the exact order verified in the copy. Use scalar rates for buyer closing costs, growth, GP share, occupancy, renewal and exit assumptions. For phased rates keep the base scalar and add dated schedule decisions to the manual queue. For the single fixed, full-term interest-only, 30/360 loan drawn at acquisition and retired at sale, loan_structure.value can be `{"kind":"fixed_io_at_close","ltv":0.65,"rate":0.055,"fee_rate":0.0125,"term_months":48}`. An amount can replace ltv. Use `{"kind":"cash"}` for a cash deal or `{"kind":"custom","description":"…"}` and separate evidenced terms for multiple facilities, amortization, floating rates or interim refinancing. Do not apply the single-loan verifier to those structures.

For the T-12 import, use the exact historical account label as the category name or supply a history_accounts decision mapping source labels to target labels. The master map includes rows for personnel, repairs, marketing, utilities, insurance and Other Exp. 1. Verify that each row in the copy is free or belongs to this import before applying it. Consolidate lines deliberately when multiple source accounts belong in one target. Add slots in the deal map for other categories. Keep the tie to the original monthly ledger. For operating forecast candidates, use an operating_lines decision with one entry per included line, e.g. `{"account":"Insurance","selector":"$ / Unit","input":649.25,"growth":0.03}`. Supply the annual dollars per unit for the period actually consumed by that template, or the share of EGR for a percentage line. Trace row growth before converting a T-12 amount to a Year 1 figure. For heterogeneous inputs put each line’s source, period, calculation and exceptions in the decision basis or the individual entry.

**Example of a completed decision** (these figures belong to the example case):

```json
"buyer_cost_rate": {
  "value": 0.03,
  "unit": "fraction of purchase price",
  "status": "precedent",
  "source": "prior-acquisition-model",
  "basis": "Two reviewed acquisition models use 3%. Replace with an itemized buyer budget.",
  "destination": "Transaction Inputs / Transaction Costs / percentage input"
}
```

**Local command sequence.** Run from the unpacked skill folder. Substitute the actual deal folder, Excel file and Zillow CSV paths.

```sh
python scripts/research.py zori.csv zhvi.csv 10017 2026-11-06 2023 > research.json
python scripts/reconcile.py deal.json > reconciliation.json
python scripts/workbook.py inspect "Deal Copy.xlsx" > inspected.json
python scripts/plan.py deal.json inspected.json > plan.json
# Populate and fully recalculate the deal copy in Excel, then save.
python scripts/check_inputs.py "Deal Copy.xlsx" approved-plan.json
python scripts/workbook.py capture "Deal Copy.xlsx" > readback.json
python scripts/verify.py deal.json readback.json > verification.json
```

Save the resolved list of approved scalar writes as approved-plan.json in the same shape as plan.json, including any manual scalar additions, before checking input readbacks. Document intentional linked formulas separately and verify their precedents and outputs in Excel. Apply the plan in the **named copy** after checking the displayed labels, formula status, value units and existing cell contents. ISO dates in the plan are dates to enter as Excel date values. The plan emits candidates for supported scalar inputs, unit mix, a simple fixed interest-only loan, and research. Optionally supply starting_rent_by_type as a type-to-monthly-dollars-per-SF object with the evidence for the chosen rent basis. Supply research_rent and research_value decisions with latest_yoy and mean_yoy from research.py to populate year zero and the forecast benchmarks. The agent completes the manual queue including history import, loan structure, operating schedules, research, taxes, mixed uses, and scenarios. Confirm unused unit slots and unused facilities on refresh. Keep the existing asset mix on a populated copy unless changing it is intended. Re-read actual source and target fields immediately before writing. The plan is an aid for the same identified version, not a standalone import file.

**Connected Excel.** Use the Excel live-control workflow and confirm the open document identity. Inspect labels and formulas in its live input blocks. If a matching saved copy is unavailable, build the same `inspected.json` from live ranges using the map: each key is `Sheet / field` and its value is `{"address":"F12","preserve_formulas":[]}`; include every formula address within each candidate range. For results, prepend `const mapping = <contents of the resolved map JSON>;` to scripts/capture-live.js in an Excel Office.js function-body executor. Run it after full recalc and save its returned object as readback.json. It checks the anchors and scans all used worksheets for formula errors. For a 1904-date-system workbook, use saved-file capture or adapt the serial-date conversion to that workbook. No reader can determine from cached numbers alone whether Excel just recalculated. Record the actual recalculation and source workbook version.

**Result handling.** A malformed record, unresolved layout or failed check returns a nonzero exit code. Read stderr as well as JSON. Reconciliation or verification differences must be explained or corrected. A valid result can still have open items. The verifier returns review_items for these; a successful process exit means the numeric comparisons completed. Inspect every open item, including custom debt checks, unverified assumptions and reported formula errors. Re-run after changing the workbook. For unusual cases compare the applicable schedules and classify any inapplicable baseline check explicitly. Report any checks the available tools could not perform.

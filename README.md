# Multifamily underwriting skill

Prepare a first-pass acquisition model in a deal-specific copy of the underwriting template. The skill gathers deal evidence, applies documented conventions, maps supported inputs to Excel and returns the recalculated workbook with assumptions and open diligence items.

## Architecture

```text
Reusable underwriting skill                 Inputs for one building
---------------------------                 -----------------------
SKILL.md             Workflow               Email and data room
references/          Rules and precedents   Rent roll, leases, T-12, OM
master-map.json      Excel cell map          Meeting instructions
schemas/             Deal record format     ZIP market data and template
            |                                           |
            +---------------------+---------------------+
                                  v
                       Agent reads the evidence
                                  |
                                  v
                              deal.json
                   Value, units, source, basis,
                   effective date and destination
                                  |
                +-----------------+------------------+
                v                 v                  v
           reconcile.py       research.py        cases.py
           Rent roll and      ZIP rent and       Tax changes and
           expense checks     value history      expense scenarios
                |                 |                  |
                +-----------------+------------------+
                                  v
                workbook.py inspects the template
                plan.py maps values to proposed cells
                                  |
                                  v
                 Agent writes the deal copy in Excel
              Transaction Inputs | Asset Inputs | T-12
                                  |
                                  v
                    Excel recalculates the model
                                  |
                                  v
              Capture live results or saved workbook
              Check input values and financial outputs
                                  |
                        +---------+----------+
                        v                    v
                  Discrepancy found       Review package
                  Trace, correct and      Workbook, returns,
                  run checks again        assumptions, open items
```

The agent retrieves documents through authorized apps or supplied exports. It records whether an input was given, derived, instructed, taken from precedent, modeled as a scenario, unknown or inapplicable. For example, the rent roll supports current occupancy. A documented renewal convention can support a forecast. Each decision records which source supports its use.

Python performs repeatable calculations and prepares proposed edits. The agent checks the actual workbook labels and handles edits that need judgment. Excel calculates cash flows and returns. After recalculation, the skill compares the written values with the proposed entries and checks sources and uses, debt service, payoff, equity and dated returns. Missing evidence goes into the handoff with the affected assumption.

## Run a deal

Use a compatible Excel template, access to the deal materials and Python 3.10 or newer. Install the helper dependencies with `python -m pip install -r requirements.txt` from this folder. For personal Codex use, place the repository folder under `~/.codex/skills/land-labor-underwriting`.

> Use $land-labor-underwriting to underwrite [property] from [deal room] in a copy of [template]. Return the saved workbook, the assumptions used and the items that need review.

Each new property starts with a fresh deal record and workbook. Keep deal files in a separate working directory. Record a one-off decision there. When a review establishes a rule for future deals, add the rule, scope and exceptions to the shared conventions.

| Component | Responsibility |
|---|---|
| [Skill instructions](SKILL.md) | Intake, workbook identity, edits, recalculation and handoff |
| [Decision rules](references/decision-rules.md) and [conventions](references/underwriting-conventions.md) | Source selection, assumptions and exceptions |
| [Starter record](assets/deal-start.json) and [schema](schemas/deal.schema.json) | Evidence and decisions for one property |
| [Workbook map](assets/master-map.json) and [mapping guide](references/workbook-map.md) | Candidate cells, surrounding labels and checks for shifted layouts |
| [Scripts](scripts/) and [execution guide](references/run.md) | Research, reconciliation, planning, capture and verification |
| [Conditional cases](references/cases-and-checks.md) | Abatements, expense gaps, debt structures, waterfalls and review checks |

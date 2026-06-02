# Project Balance Memo Updater Spec

## Purpose

Use **Project Balance IC Memo v1.docx** as the base template for an updater that recalculates portfolio exposures when bid sizing changes across multiple LP-led secondary funds.

The core rule is that project-level exposures are weighted by **record date NAV** at the fund level, then blended by the proposed bid amount for each fund.

## Template Map

Project Balance appears to be organized into these memo blocks:

1. `Investment Summary`
2. `Transaction Overview`
3. `Transaction Rationale`
4. `Funds Description`
5. `Relevant Diligence Notes`
6. `Exposure Summary`
7. `Expected Returns`
8. `Model and assumption methodology`
9. Fund-level sections:
   - `BT III`
   - `BT II`
   - `PGIM VI`
10. Supporting appendices / tables:
   - cash flow profile
   - investment thesis
   - investment list
   - current portfolio

The memo package contains 8 chart parts and 13 media files in the Balance file, so the updater should treat visuals as first-class outputs rather than static pasted images.

## Existing Balance Outputs

From the current memo structure, the outputs that need to update cleanly are:

- top summary stats
- fund summary table
- top positions concentration table
- exposure summary charts
- return / IRR / MOIC outputs
- fund-level modeling narrative
- current portfolio / position-level detail tables

## Recommended Data Model

### 1. Fund input

Each fund should have one input record:

```json
{
  "fund_id": "BTIII",
  "fund_name": "Blue Torch III",
  "record_date": "2025-12-31",
  "proposed_bid": 50000000,
  "bid_currency": "USD",
  "source_workbook": "path/to/fund.xlsx",
  "weighting_basis": "record_date_nav"
}
```

### 2. Fund holdings input

Each fund workbook should expose one row per underlying investment:

```json
{
  "fund_id": "BTIII",
  "investment_id": "optional-stable-id",
  "investment_name": "Company A",
  "record_date_nav": 1250000,
  "asset_class_raw": "Corporate Lending",
  "security_type_raw": "First Lien",
  "geography_raw": "North America",
  "sector_raw": "Software",
  "sub_asset_class_raw": "Direct Lending"
}
```

Recommended optional fields:

- `industry`
- `country`
- `region`
- `currency`
- `mark`
- `cost`
- `unfunded`
- `ownership_pct`
- `deal_type`
- `manager_tag`

### 3. Normalization map

The app should maintain a mapping layer from raw fund labels to memo labels.

Example:

- `Corp Lending`, `Corporate Lending`, `Direct Lending` -> `Corporate Lending`
- `Special Sits`, `Situations`, `Structured Equity` -> `Special Situations`
- `U.S.`, `United States`, `North America` -> `North America`

Anything unmapped should be surfaced before export.

### 4. Scenario input

The scenario is the bid mix across funds:

```json
{
  "scenario_name": "Base Case",
  "fund_bids": [
    { "fund_id": "BTIII", "proposed_bid": 50000000 },
    { "fund_id": "BTII", "proposed_bid": 50000000 }
  ]
}
```

## Calculation Rules

1. Compute each fund’s normalized exposure profile from underlying investments using `record_date_nav`.
2. Convert each fund’s exposure profile into percentages by category.
3. Weight each fund’s profile by the proposed bid amount in the scenario.
4. Sum the weighted exposures across all selected funds.
5. Normalize project-level totals to 100%.

Formula:

```text
project_exposure(category) =
  sum over funds [
    fund_weight * fund_exposure(category)
  ]

fund_weight = fund_bid / sum(all fund bids)
```

## Output Model

The app should emit structured outputs for the memo, not just charts:

```json
{
  "project_stats": {},
  "fund_summary": [],
  "top_positions": {},
  "asset_class_exposure": [],
  "security_type_exposure": [],
  "geography_exposure": [],
  "return_metrics": {},
  "narrative_facts": []
}
```

## Chart Families

The Balance package suggests these chart families should be dynamic:

- asset class / sub-asset class exposure
- security type exposure
- geography exposure
- top positions concentration
- return / cash flow profile
- fund-level versus project-level comparison views

## Implementation Shape

The cleanest implementation is:

1. ingest one workbook per fund
2. normalize holdings into a shared schema
3. compute scenario-weighted exposures
4. render charts / tables
5. replace the corresponding memo graphics and values in the Project Balance template

## Working Assumption

Project Balance is the base template for formatting and structure. Mariner, Such, and Yellowstone are style/context examples only and should not override the Balance layout unless specifically requested.

# Memo Helper — Vision & Spec (working draft)

Status: **planning**. This captures where we want the tool to go. Edit freely.

## Goal
Evolve the current exposure-exhibit updater into a **memo helper**: you (1) enter the
deal facts you carry in your head, (2) point it at the deal folder / source files, and
it (3) assembles a populated **draft** IC memo — numbers, tables, charts, and narrative.

## Governing principle (non-negotiable)
**GPT never produces a number.** Every figure — exposures, IRR/MOIC, prices, NAV,
concentration — is computed by the deterministic engine from the models. GPT only drafts
*qualitative prose*, grounded in source documents, **cited**, and watermarked **DRAFT**.
Rationale: an IC memo must be trustworthy; numbers stay auditable, AI text stays clearly
review-required.

## AI layer
- Provider: **Azure OpenAI (GPT)** — GCM has an endpoint/key, so the app can call it
  directly. Wrap it provider-agnostic so the model can be swapped later.
- **Data residency:** only the GCM-sanctioned Azure endpoint may receive deal documents.
  Confirm before sending any confidential source material.
- Output posture: per-section **draft** text + **source citations**; human edits; never
  auto-final.

## Memo anatomy → population method
| Section | Content | Source | Method | Status |
|---|---|---|---|---|
| Investment Summary box | data | models + inputs (dates) | deterministic | partial |
| Recommendation | data + 1 line | inputs (bids/prices) | input-driven | next |
| Transaction Overview | narrative | broker/process notes + inputs | GPT draft | future |
| Transaction Rationale | narrative | analyst view + diligence | GPT draft + edit | future |
| Funds Description | narrative | GP materials | GPT draft | future |
| Diligence Notes | narrative | call notes, GP materials | GPT draft | future |
| Exposure Summary (pies, concentration, asset-type) | data | models/holdings | deterministic | **done** |
| Expected Returns (IRR/MOIC, pricing bridge) | data | models | deterministic | next |
| Modeling methodology | narrative | model assumptions | GPT draft + edit | future |
| Cash-Flow Profile | chart + narrative | model + thesis | deterministic chart / GPT text | next (chart) |
| Investment List (per-fund) | data + narrative | models + diligence | hybrid | future |
| GP Platform & track record | narrative + data | GP docs | GPT draft + data | future |
| Current Portfolio (per-company) | narrative + data | loan tapes, GP one-pagers | hybrid | future |
| Appendix A (overrides) | data | model | deterministic | future |
| ESG | form | manual | input-driven | manual |

## Architecture — four layers on top of today's exporter
1. **Inputs** — form for facts you type (seller, broker/process, bid strategy, dates, recommendation $).
2. **Sources** — point at the deal folder; index **models** (numbers) + **documents** (PDF/Word, narrative grounding).
3. **Population engine** — deterministic for all data; GPT for prose (retrieve relevant source text -> draft section -> return text + citations).
4. **Renderer** — native charts + tables (built) + narrative text into templated sections, marked DRAFT.

## Phased roadmap (each step usable on its own)
1. **Done** — exposure pies + concentration + asset-type table, with formatting.
2. **Deterministic exhibits** — IRR/MOIC + fund-summary tables; cash-flow combo charts; investment-summary box; from the model.
3. **Inputs panel** — type deal facts -> fill Recommendation / Transaction Overview header fields.
4. **Sources + GPT drafting** — point at folder -> draft a narrative section (cited, DRAFT) as a proof of concept; expand section by section.
5. **One-click draft memo** — tie it together.

## Open decisions
- Confirm the Azure OpenAI endpoint is GCM-sanctioned for confidential deal docs.
- v1 "useful draft" scope: data exhibits + header inputs only, or also narrative blocks? (TBD)
- Surgical table writing (edit document.xml directly, like the charts) vs. python-docx
  re-serialization — revisit if byte-fidelity of untouched parts matters.
- Generalize the template beyond Project Balance (chart-map / table-targeting by heading,
  not position) for reuse across Mariner / Such / Yellowstone / future deals.

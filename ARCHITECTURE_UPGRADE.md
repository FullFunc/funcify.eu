# ARCHITECTURE_UPGRADE.md — Funcify Target Architecture
*Generated: 2026-05-07*

---

## OVERVIEW

This document describes the target Funcify architecture and how the current engine.py maps onto it, layer by layer.

---

## THE 6 LAYERS — CURRENT VS TARGET

### LAYER 1 — RAW EXTRACTION
*Extract raw website data exactly as presented. No interpretation.*

**Current state:** Partial.
- `page_text = soup.get_text(" ", strip=True)` extracts visible text
- `product_data["additional_info"] = page_text[:5000]` stores a truncated window
- No raw HTML storage
- No screenshot
- No structured data (JSON-LD) extraction
- No PDF detection

**Target state:**
```
raw_extraction = {
  html: string,           # full page HTML after JS execution
  screenshot_path: string,
  text: string,           # full visible text, no truncation
  structured_data: [],    # JSON-LD, OpenGraph, schema.org
  pdfs_detected: [],      # downloadable lab reports, COA
  raw_ingredient_section: string,  # text between ingredient header and next section
  source_url: string,
  scraped_at: datetime,
  scraper_version: string
}
```

**Gap:** Add Playwright for HTML + screenshot. Add JSON-LD parser. Store full raw text, not [:5000] window.

---

### LAYER 2 — STRUCTURED EXTRACTION
*Convert raw text into structured fields without interpretation.*

**Current state:** Implemented but impure.
- `scrape_product()` extracts fields using regex + CSS selectors
- `extract_with_claude()` overwrites fields using Claude inference (INTERPRETATION — violates Layer 2 purity)
- `anthropic_fallback()` generates fields using Claude inference (INTERPRETATION — violates Layer 2 purity)

**Target state:**
```
structured_extraction = {
  product: {
    name: { value, raw_text, selector_used, confidence },
    brand: { value, raw_text, strategy_used, confidence },
    price: { value, raw_text, regex_match, currency, confidence },
    serving_size: { value_raw, capsule_count, unit, confidence },
    servings_per_container: { value, raw_text, confidence },
    package_size: { value, raw_text, source: "product_name"|"url"|"page_text", confidence },
    usage_instructions: { value, raw_text, confidence },
  },
  ingredients: [{
    raw_text: string,       # exact line as extracted
    name_raw: string,       # before any normalization
    amount: number | null,
    unit: string | null,
    form_raw: string | null,
    parenthetical: string | null,  # content of (...) if present
    is_nested: bool,
    extraction_pass: 1|2|3|"claude",
    line_position: int,
    confidence: "EXACT"|"PARTIAL"|"MISSING"
  }],
  excipients: [{
    raw_text: string,
    name_raw: string,
    function_raw: string | null,
    confidence: "EXACT"|"PARTIAL"|"MISSING"
  }],
  certifications: [{
    value: string,
    raw_text: string,
    source_location: "badge"|"body"|"meta"|"footer"|"pdf",
    confidence: "EXACT"|"PARTIAL"|"MISSING"
  }],
  health_claims: [{
    value: string,
    raw_text: string,
    source_location: string,
    confidence: "EXACT"|"PARTIAL"|"MISSING"
  }],
  warnings: [{
    value: string,
    raw_text: string,
    confidence: "EXACT"|"PARTIAL"|"MISSING"
  }]
}
```

**Key rule for Layer 2:** Claude extraction calls are NOT part of Layer 2. If regex extraction fails, the field is left empty with confidence: "MISSING". Claude is only called in Layer 4+ for validation and in Layer 6 for review.

**Gap:** Separate Claude calls from extraction. Add confidence per field. Add raw_text per field.

---

### LAYER 3 — NORMALIZATION
*Normalize ingredient names, units, forms, and serving logic. No interpretation.*

**Current state:** Partial and destructive.
- `_normalize_name()` maps to `INGREDIENT_SYNONYMS` (hardcoded, 40 entries)
- `simplify_jargon()` rewrites strings in consumer_output (destructive, no raw preservation)
- No unit normalization (IU ↔ mg)
- No form normalization from ingredient_forms sheet

**Target state:**
```
normalized_ingredient = {
  ...raw_extraction fields preserved...,
  name_normalized: string,         # canonical name from ingredient_master
  ingredient_id: string,           # ID from ingredient_master sheet
  parent_ingredient_id: string,    # for "waarvan" nested ingredients
  form_normalized: string,         # from ingredient_forms sheet
  bioavailability_tier: "high"|"medium"|"low"|"reference"|null,
  unit_normalized: "mg"|"mcg"|"IU"|"g"|"CFU",
  amount_normalized: number | null,  # in normalized unit
  ri_percentage: number | null,
  serving_dose: number | null,      # amount × servings_per_day
  daily_max: number | null,
  proprietary_blend: bool,
  normalization_method: string,     # "synonym_match"|"fuzzy_match"|"manual_review"|"unmatched"
  normalization_confidence: "EXACT"|"PARTIAL"|"MISSING"
}
```

**Gap:**
1. Load `Ingredient_synonyms` and `ingredient_forms` sheets from Master Ingredient DB
2. Preserve raw name alongside normalized name (never overwrite)
3. Add unit normalization with known conversion factors
4. Proprietary blend detection

---

### LAYER 4 — VALIDATION
*Validate against Funcify databases.*

**Current state:** None. No validation layer exists.

**Target state:**
```
validation_result = {
  ingredient_match_status: "matched"|"partial_match"|"unmatched",
  matched_ingredient_id: string | null,
  validation_errors: [],
  data_lacunes: [{
    field: string,
    ingredient: string | null,
    reason: string,
    severity: "blocking"|"warning"|"info",
    category_critical: bool
  }],
  engine_readiness: bool,
  engine_readiness_blockers: []
}
```

**Validation checks to implement:**
1. Each ingredient must match ingredient_master or be flagged as unmatched
2. Proprietary blends → DATA_LACUNE for each sub-ingredient without dose
3. Category-critical fields: EPA/DHA for OMEGA3, CFU for PROBIOTIC, etc.
4. Serving size parseable as integer or range
5. Package size ≥ 10 (sanity check)
6. Price parseable as decimal
7. Amount > 0 when unit is present
8. No duplicate ingredient names after normalization

**Gap:** Build entire validation module. Currently nothing exists here.

---

### LAYER 5 — ENGINE PREPARATION
*Prepare engine-ready JSON for the review engine.*

**Current state:** Mixed with scoring. `/score` route does both preparation and scoring.

**Target state:** Clean separation.
```
engine_input = {
  product_identity: {
    name, brand, url, category, type, complexity_tier
  },
  detected_modules: ["OMEGA3", "VITAMIN", ...],
  active_ingredients: [...normalized ingredients with evidence...],
  excipients: [...normalized excipients with evidence...],
  quality_markers: {
    certifications: [...with evidence...],
    third_party_testing: bool,
    coa_available: bool,
    sustainability: bool
  },
  dosage_context: {
    serving_size_raw: string,
    servings_per_container: number,
    servings_per_day: number,
    max_daily_dose: {...per ingredient...}
  },
  interaction_precheck: {
    cofactor_warnings: [],
    antagonist_warnings: [],
    context_flags: []
  },
  data_lacunes: [{...}],
  engine_readiness: bool,
  engine_readiness_blockers: []
}
```

**Gap:** Separate preparation from scoring. `/score` route should accept a fully prepared `engine_input` as input, not raw scraped data.

---

### LAYER 6 — REVIEW ENGINE
*Scoring, evaluation, consumer output. Separate from scraper.*

**Current state:** Implemented in `/score` route.
- `evaluate_criteria_with_claude()` — Claude scoring with DATA_LACUNE semantics
- `calculate_score()` — weighted scoring
- `determine_verdict()` — tier thresholds
- `generate_consumer_output()` — Claude narrative generation

**Assessment:** Layer 6 is the most complete layer. The scoring algorithm is solid. The main gaps are:
- Scoring criteria come from Excel (2 of 6 sheets loaded), but CF rules and ingredient data are hardcoded
- No evidence tracing back to raw extraction in the score output
- Consumer output is generated with current (possibly stale) data lacunes that were generated at Layer 5

**Keep as-is until Layers 1–5 are upgraded.** Then connect Layer 6 to Layer 5 output instead of raw scraped data.

---

## DATA FLOW DIAGRAM

```
URL
 │
 ▼
┌─────────────────────────────┐
│  LAYER 1: RAW EXTRACTION    │  Playwright → HTML, screenshot, text, JSON-LD
└─────────────┬───────────────┘
              │ raw_extraction
              ▼
┌─────────────────────────────┐
│ LAYER 2: STRUCTURED EXTRACT │  regex + CSS → structured fields + confidence
└─────────────┬───────────────┘
              │ structured_extraction
              ▼
┌─────────────────────────────┐
│  LAYER 3: NORMALIZATION     │  Excel DBs → normalized names, forms, units
└─────────────┬───────────────┘
              │ normalized_extraction
              ▼
┌─────────────────────────────┐
│  LAYER 4: VALIDATION        │  ingredient_master match, DATA_LACUNE detection
└─────────────┬───────────────┘
              │ validated_extraction + data_lacunes
              ▼
┌─────────────────────────────┐
│ LAYER 5: ENGINE PREPARATION │  engine_input JSON assembly
└─────────────┬───────────────┘
              │ engine_input
              ▼
┌─────────────────────────────┐
│  LAYER 6: REVIEW ENGINE     │  Claude scoring + consumer output
└─────────────┬───────────────┘
              │ score_output
              ▼
         /score response
```

---

## EXCEL DATABASE MAPPING

### Funcify. Master Ingredient 2.0.xlsx

| Sheet | Current usage | Target usage |
|-------|--------------|--------------|
| ingredient_master | NONE | ingredient ID, canonical name, category → Layer 3 normalization |
| Ingredient_synonyms | NONE | replaces hardcoded INGREDIENT_SYNONYMS dict |
| ingredient_forms | NONE | replaces hardcoded BIOAVAILABILITY_RATIOS dict |
| ingredients_bioavailability_fac | NONE | bioavailability tier per form → Layer 3 enrichment |
| ingredient_stability | NONE | storage-based degradation warnings → Layer 4 validation |
| ingredient_cofactors | NONE | extends evaluate_cofactor_checks() |
| ingredient_antagonist | NONE | new antagonist check module |
| ingredient_transport_competition | NONE | transport conflict warnings |
| ingredient_enzymes | NONE | CYP450 interaction flags |
| ingredient_formulation | NONE | form compatibility rules |
| ingredient_biomarker_effect | NONE | biomarker-to-ingredient mapping |
| biomarkers | NONE | biomarker definitions |
| ingredient_use_case-effect | NONE | use case mapping for consumer output |

### Funcify. Engine Review V4.xlsx

| Sheet | Current usage | Target usage |
|-------|--------------|--------------|
| Engine_review_score | LOADED — core criteria | Keep, extend |
| Category_modules | LOADED — module criteria | Keep, extend |
| Context_Flag_Rules | NOT LOADED | Replace hardcoded CONTEXT_FLAG_RULES |
| Settings | NOT LOADED | Model config, thresholds |
| Engine_weight_explained | NOT LOADED | Documentation only |
| Category_weight_explained | NOT LOADED | Documentation only |

### Funcify. Consumer UI.xlsx

| Sheet | Current usage | Target usage |
|-------|--------------|--------------|
| ui_display_content | NONE | beoordeling_tabel ordering validation |
| ui_recommendation_logic | NONE | verdict-to-recommendation mapping |
| ui_user_filters | NONE | filter logic for frontend |
| ui_output_ranking_rules | NONE | output field ordering |
| ui_problems | NONE | problem definitions |
| ui_problem_biomarker_map | NONE | biomarker-to-problem mapping |
| ui_problems_use_case_map | NONE | use case-to-problem mapping |

---

## CATEGORY MODULE ARCHITECTURE

### Current state
`detect_product_type()` returns one type string. Criteria filtering uses this type for `applies_to` matching. No category-specific extraction logic.

### Target state
Multiple modules can be active simultaneously. A magnesium + vitamin D + K2 product activates MINERAL + VITAMIN modules.

```python
def detect_modules(product_data) -> list[str]:
    # Returns all applicable module types
    modules = []
    text = [product_data["product_name"], ...].join(" ").lower()
    if any(x in text for x in OMEGA3_KEYWORDS):
        modules.append("OMEGA3")
    if any(x in text for x in VITAMIN_KEYWORDS):
        modules.append("VITAMIN")
    # ... etc
    return modules if modules else ["ALL"]
```

Module-specific extractors then run for each detected module.

---

## HALLUCINATION PREVENTION ARCHITECTURE

### Rule: Claude is never called for extraction (Layers 1–3)

The current architecture calls Claude in `extract_with_claude()` and `anthropic_fallback()` during the scraping phase. Under the target architecture:

- **Layer 1–3:** No Claude calls. Pure regex, CSS, and structured data extraction.
- **Layer 4:** Claude MAY be called to resolve ambiguous ingredient name matching (with explicit confidence PARTIAL and raw_text preserved)
- **Layer 5:** Claude is NOT called. Engine preparation is deterministic.
- **Layer 6:** Claude IS called for evaluation and consumer output generation.

**Practical transition (Python PATH A):**
Keep `extract_with_claude()` and `anthropic_fallback()` but:
1. Add system prompt with strict hallucination guardrails
2. Mark all Claude-extracted values with `extraction_method: "claude_extraction"` and `confidence: "PARTIAL"`
3. Never let Claude-extracted values override regex-extracted values for the same field

---

## EVIDENCE OBJECT STANDARD

Every extracted field that flows into Layer 5+ must carry:

```python
{
  "value": any,
  "confidence": "EXACT" | "PARTIAL" | "MISSING",
  "source_url": str,
  "source_type": "visible_text" | "html_meta" | "structured_data" | "image_alt" | "downloadable_file" | "inferred_from_label_text" | "claude_extraction",
  "raw_text": str,
  "extraction_method": str,
  "normalization_method": str | None,
  "matching_logic": str | None,
  "data_lacune": bool,
  "data_lacune_reason": str | None
}
```

### Implementation priority for evidence objects
1. Ingredients (highest impact — drives most scoring)
2. Certifications (high hallucination risk)
3. Serving size / package size (drives price calculations)
4. Health claims (hallucination risk)
5. Brand / product name (lower risk, usually deterministic)

---

## PROPRIETARY BLEND HANDLING

### Current state
Proprietary blends are not detected or flagged. An ingredient line "Eigen Formule 500mg" passes through as a single 500mg active ingredient.

### Target behavior
1. Detect blend: look for "blend", "complex", "matrix", "formule", "proprietary", "eigen mengsel"
2. Flag the blend as `proprietary_blend: true` with total blend amount
3. Extract sub-ingredients within the blend text (they appear as a list, often without individual doses)
4. For each sub-ingredient without individual dose → `data_lacune: true, data_lacune_reason: "dose hidden in proprietary blend"`
5. Set `engine_readiness: false` if a proprietary blend hides critical ingredient doses
6. Report in data_lacunes array with severity: "blocking"

---

## CURRENT CONFLICTS WITH ABSOLUTE SYSTEM RULES

| Rule | Current violation |
|------|------------------|
| Never overwrite raw extracted data with normalized data | `extract_with_claude()` overwrites health_claims, package_size |
| Preserve every original source text | No raw_text stored per field |
| Every normalized field must trace back to raw evidence | No evidence objects |
| Never silently fail | `except Exception: pass` in extract_with_claude |
| Never silently normalize | simplify_jargon() mutates without logging |
| Never silently infer | Claude extraction fills fields without marking as inferred |
| If uncertain → mark as DATA_LACUNE | Missing fields are simply empty strings, not DATA_LACUNE |
| If partially verifiable → mark confidence PARTIAL | No confidence field exists on extracted data |

---

## IMMEDIATE NEXT STEPS

Execute in this order, one step per session to maintain stability:

**Step 1 (TIER 0 — do first):**
- Fix line 699: `f"- {e.get('name', '')}"`
- Remove dead variables lines 648–658
- Add `CLAUDE_MODEL = "claude-sonnet-4-5"` constant
- Syntax check + commit

**Step 2 (TIER 1 — stability):**
- Fix `anthropic_fallback` with system prompt + 4000 chars
- Fix `extract_with_claude` ingredient merge logic
- Tighten `_needs_second_pass` trigger
- Syntax check + commit

**Step 3 (TIER 1 — continued):**
- Package size sanity check (count ≥ 10)
- Strengthen `is_sufficient()`
- Syntax check + commit

**Step 4 (TIER 2 — evidence):**
- Add confidence field to ingredients
- Add data_lacunes array to product_data
- Add engine_readiness flag
- Syntax check + commit

**Step 5 (TIER 3 — database):**
- Load Context_Flag_Rules from Excel (Engine Review V4)
- Load Ingredient_synonyms from Master Ingredient DB
- Load ingredient_forms from Master Ingredient DB
- Test against current scores
- Syntax check + commit

**Step 6 (TIER 3 — continued):**
- Load remaining Master Ingredient sheets
- Wire BIOAVAILABILITY_RATIOS into scraping flow
- Syntax check + commit

**Step 7 (TIER 4 — category extractors):**
- Omega-3 EPA/DHA split extractor
- Probiotic CFU/strain extractor
- Proprietary blend detector
- Syntax check + commit

**Step 8 (TIER 5 — Playwright):**
- Activate Playwright (already installed)
- Cookie banner acceptance
- Accordion expansion
- Replace ScrapingBee as primary
- Test on 10 Dutch supplement sites
- Commit

**Step 9 (TIER 6 — logging):**
- Add structured logging throughout
- Add request timing
- Error classification
- Commit

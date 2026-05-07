# REFACTOR_PLAN.md — Funcify Incremental Upgrade Plan
*Generated: 2026-05-07 | Based on GAP_ANALYSIS.md*

---

## CRITICAL DECISION REQUIRED BEFORE CODING

The target architecture specifies TypeScript, Playwright, Cheerio, Zod, and tsx.
The current codebase is Python 3.12 / Flask.

This is not an incremental refactor. It is a language migration.

**Three paths forward:**

**PATH A — Python-only upgrade**
Keep Python. Add layers, evidence objects, Playwright (already installed), and Excel loading incrementally. No language migration. Fastest path to Funcify compliance. All existing logic preserved.

**PATH B — TypeScript rewrite**
New TypeScript/Node.js backend. All current scoring logic must be ported. Railway deployment config must change. Frontend config.js endpoint must change. Funcify masterprompt V10.docx must be re-integrated. Estimated effort: 3–4 weeks.

**PATH C — Hybrid**
Keep Python Flask API and scoring engine. Add a TypeScript extraction layer (Playwright + Cheerio) as a sidecar service. Python calls TypeScript extractor via HTTP or subprocess. Evidence-compliant JSON passes between layers.

**Recommendation: PATH A for immediate compliance, PATH C as later evolution.**

This plan documents PATH A — all changes in Python, incremental, backward-compatible.

---

## PRIORITY TIERS

### TIER 0 — CRITICAL BUGS (fix before any other work)

| # | File | Line | Issue | Fix |
|---|------|------|-------|-----|
| T0-1 | engine.py | 699 | `f"- {e}"` on dict renders Python repr | `f"- {e.get('name', '')}"` |
| T0-2 | engine.py | 648–658 | Dead pre-assembly variables never used | Remove or wire into user_prompt |
| T0-3 | engine.py | 1029 | `anthropic_fallback` has no system prompt | Add extraction grounding rules |
| T0-4 | engine.py | 1099 | `extract_with_claude` ingredients: block on partial | Merge new ingredients, don't replace |

---

### TIER 1 — STABILITY (no new features, fix broken logic)

**T1-1: Centralize model name**
Add `CLAUDE_MODEL = "claude-sonnet-4-5"` at top of file. Replace all 4 hardcoded strings.

**T1-2: Fix `anthropic_fallback` prompt**
- Add system prompt with hallucination guardrails ("Schrijf NOOIT iets dat niet letterlijk op de pagina staat")
- Increase page_sample from 2000 to 4000 chars
- Add `excipients` field to JSON template
- Add `type` and `form` fields to ingredient schema in prompt

**T1-3: Fix `extract_with_claude` ingredient merge**
Replace "only fill if empty" logic with:
- If `product_data["ingredients"]` is empty → replace with Claude extraction
- If `product_data["ingredients"]` exists → merge: add Claude ingredients not already in `seen_names`, don't overwrite existing

**T1-4: Tighten `_needs_second_pass` trigger**
Current: fires when `certs < 2` (almost always). New logic:
```python
_needs_second_pass = (
    not product_data.get("health_claims")
    or (len(product_data.get("certifications", [])) < 2 and not product_data.get("ingredients"))
    or not product_data.get("usage_instructions")
)
```

**T1-5: Fix package size sanity check**
After extracting from product_name/URL, validate extracted count:
```python
if _m and int(_m.group(1)) >= 10:  # serving counts are < 10, package sizes >= 10
    package_size = _m.group(0)
```

**T1-6: Strengthen `is_sufficient()`**
Require at least 1 ingredient with a non-None amount OR 2 ingredients total:
```python
def is_sufficient(data):
    has_name = bool(data.get("product_name"))
    ingredients = data.get("ingredients", [])
    has_ingredients = (
        any(i.get("amount") is not None for i in ingredients)
        or len(ingredients) >= 2
    )
    return has_name and has_ingredients
```

**T1-7: Strengthen certifications extraction**
Replace full-page substring match with section-aware match:
- Only scan within known certification-relevant sections (div.certifications, footer badges, img alt text)
- Add word-boundary check: `\b{cert}\b` instead of simple `in text_lower`
- Log source of each found certification

---

### TIER 2 — EVIDENCE LAYER (add traceability without changing output)

**T2-1: Add evidence field to ingredients**
Extend ingredient schema:
```python
{
  "name": str,
  "amount": float | None,
  "unit": str,
  "form": str,
  "type": "actief" | "vul-additief",
  "nested": bool,
  # New:
  "confidence": "EXACT" | "PARTIAL" | "MISSING",
  "raw_text": str,
  "extraction_method": "regex_pass1" | "regex_pass2" | "nested_composition" | "claude_extraction",
  "normalization_applied": bool,
  "data_lacune": bool,
  "data_lacune_reason": str | None
}
```
Populate `raw_text` with the original line from which the ingredient was extracted.
Set `confidence: "EXACT"` when amount+unit both present, `"PARTIAL"` when name only, `"MISSING"` when amount=None.

**T2-2: Add product-level data_lacunes array**
After scrape_product() completes, generate:
```python
data_lacunes = []
if not product_data.get("serving_size"):
    data_lacunes.append({"field": "serving_size", "reason": "not found in page text"})
if not product_data.get("package_size"):
    data_lacunes.append({"field": "package_size", "reason": "not found in product name, URL, or page text"})
for ing in product_data.get("ingredients", []):
    if ing.get("amount") is None:
        data_lacunes.append({"field": "ingredient.amount", "ingredient": ing["name"], "reason": "no numeric dose found"})
    if not ing.get("form"):
        data_lacunes.append({"field": "ingredient.form", "ingredient": ing["name"], "reason": "no form specified"})
product_data["data_lacunes"] = data_lacunes
```

**T2-3: Add engine_readiness flag**
```python
blocking_lacunes = [d for d in data_lacunes if d["field"] in ("ingredient.amount", "serving_size")]
product_data["engine_readiness"] = len(blocking_lacunes) == 0
```

**T2-4: Preserve raw_text per field**
For key fields (product_name, brand_name, price, serving_size, package_size), store the original raw text alongside the extracted value:
```python
product_data["_raw"] = {
  "product_name_source": name_texts[0] if name_texts else "",
  "price_match": price_match.group(0) if price_match else "",
  "serving_match": serving_match.group(0) if serving_match else "",
  ...
}
```
This preserves extraction evidence without changing the public API schema.

---

### TIER 3 — DATABASE INTEGRATION

**T3-1: Load `Funcify. Master Ingredient 2.0.xlsx`**
Extend `load_engine_criteria()` pattern to also load:
- `ingredient_master` sheet → ingredient ID, canonical name, category
- `Ingredient_synonyms` sheet → replace hardcoded `INGREDIENT_SYNONYMS` dict
- `ingredient_forms` sheet → replace hardcoded `BIOAVAILABILITY_RATIOS` dict
- `ingredient_cofactors` sheet → extend `evaluate_cofactor_checks()`
- `ingredient_antagonist` sheet → new antagonist warnings
- `Context_Flag_Rules` sheet from Engine Review V4 → replace hardcoded `CONTEXT_FLAG_RULES`

**T3-2: Wire BIOAVAILABILITY_RATIOS into scoring flow**
Currently `get_bioavailability_info()` is never called. Wire it:
- After ingredient extraction, enrich each ingredient with `bioavailability_info`
- Pass enriched data to `evaluate_criteria_with_claude()` so Claude has form quality in context

**T3-3: Load `Funcify. Consumer UI.xlsx`**
Load `ui_display_content`, `ui_recommendation_logic`, `ui_output_ranking_rules` sheets.
Use to validate `beoordeling_tabel` ordering and verdict labeling.

**T3-4: Load `Context_Flag_Rules` from Excel instead of hardcoded list**
After loading, replace `CONTEXT_FLAG_RULES` with Excel-sourced data.
Keep the hardcoded list as fallback if Excel load fails.

---

### TIER 4 — CATEGORY-SPECIFIC EXTRACTORS

**T4-1: Omega-3 extractor**
After `detect_product_type() == "OMEGA3"`, run dedicated extraction:
- Search for EPA/DHA split pattern: `r"EPA[:\s]+(\d+)\s*mg.*DHA[:\s]+(\d+)\s*mg"`
- Search for TOTOX value: `r"TOTOX[:\s]*(\d+)"`
- Search for form: ethyl ester / triglyceride / phospholipid / rtg
- If EPA/DHA not found → add DATA_LACUNE: "EPA/DHA split niet gevonden"

**T4-2: Probiotic extractor**
- Extract CFU per strain pattern: `r"(\d+(?:\.\d+)?\s*(?:miljard|billion|×10)?\s*(?:KVE|CFU))"`
- Extract strain names: Lactobacillus/Bifidobacterium taxonomy
- Check expiry guarantee language
- If CFU per strain not found → DATA_LACUNE

**T4-3: Botanical extractor**
- Extract Latin name from parenthetical after common name
- Extract extract ratio: `r"(\d+:\d+)"`
- Extract standardization: `r"gestandaardiseerd op (\d+)%\s*(\w+)"`
- If extract ratio missing → DATA_LACUNE

**T4-4: Proprietary blend detector**
- Detect blend keywords: "proprietary blend", "eigen mengsel", "complex", "matrix"
- Flag ingredients within blend as `proprietary_blend: true`
- If blend total mg given but individual doses not → DATA_LACUNE per sub-ingredient

---

### TIER 5 — PLAYWRIGHT INTEGRATION

**T5-1: Activate Playwright (already installed)**
Replace or augment ScrapingBee with Playwright Chromium:
```python
from playwright.sync_api import sync_playwright

def fetch_with_playwright(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, wait_until="networkidle", timeout=30000)
        # Accept cookie banners
        for selector in ["#cookie-accept", ".cookie-consent button", "[id*='cookie'] button"]:
            try:
                page.click(selector, timeout=2000)
            except:
                pass
        # Expand accordions
        for selector in ["details", "[aria-expanded='false']", ".accordion-trigger"]:
            try:
                page.click(selector, timeout=1000)
            except:
                pass
        html = page.content()
        browser.close()
        return BeautifulSoup(html, "lxml"), html
```

**T5-2: Save raw HTML per scrape**
Store `raw_html` in product_data["_raw"]["html"] (not in public API output, only for audit).

**T5-3: Screenshot capture (optional)**
```python
page.screenshot(path=f"/tmp/{url_hash}.png", full_page=True)
```

**T5-4: Scraper fallback order**
```
1. Playwright (primary, handles JS, cookie banners, accordions)
2. ScrapingBee (secondary, if Playwright fails)
3. requests + BeautifulSoup (tertiary, for simple static pages)
4. extract_with_claude (fourth pass, for partially extracted data)
5. anthropic_fallback (last resort, full Claude extraction)
```

---

### TIER 6 — STRUCTURED LOGGING AND MONITORING

**T6-1: Add structured logging**
```python
import logging
import json
import time

logging.basicConfig(level=logging.INFO, format='%(message)s')

def log(event, **kwargs):
    logging.info(json.dumps({"event": event, "ts": time.time(), **kwargs}))
```

Use throughout:
```python
log("scrape_start", url=url)
log("scrape_complete", url=url, source=product_data["_source"], ingredient_count=len(ingredients))
log("second_pass_triggered", url=url, reason="missing_health_claims")
log("claude_call", function="evaluate_criteria", model=CLAUDE_MODEL, tokens_out=4000)
log("score_result", url=url, score=score_100, kwalificatie=kwalificatie, critical_fail=critical_fail)
```

**T6-2: Add request timing**
```python
start = time.time()
# ... scraping ...
log("request_complete", url=url, duration_ms=round((time.time()-start)*1000))
```

**T6-3: Error classification**
Replace bare `except Exception` with classified handling:
```python
except anthropic.APIStatusError as e:
    log("claude_error", status=e.status_code, message=str(e))
except requests.Timeout:
    log("scrape_timeout", url=url)
except json.JSONDecodeError as e:
    log("json_parse_error", function="extract_with_claude", raw=raw[:200])
```

---

## EXECUTION ORDER

```
Phase 0 (now):    T0-1 through T0-4   — fix critical bugs
Phase 1 (week 1): T1-1 through T1-7   — stability fixes
Phase 2 (week 2): T2-1 through T2-4   — evidence layer
Phase 3 (week 3): T3-1 through T3-4   — database integration
Phase 4 (week 4): T4-1 through T4-4   — category extractors
Phase 5 (week 5): T5-1 through T5-4   — Playwright activation
Phase 6 (week 6): T6-1 through T6-3   — logging and monitoring
```

---

## BACKWARD COMPATIBILITY RULES

1. `/scrape` and `/score` API signatures must not change
2. `beoordeling_tabel` structure (8 rows, fixed order) must not change
3. `score`, `kwalificatie`, `verdict` fields must not change
4. CORS origins must not change
5. All new fields added to responses must be additive (not replacing existing fields)
6. Evidence objects and data_lacunes should be under `_metadata` key to avoid breaking frontend
7. Frontend `config.js` must not require changes
8. Railway/Procfile deployment must not require changes until Tier 5

---

## DO NOT TOUCH

- `calculate_score()` logic — this is the core engine, working correctly
- `determine_verdict()` tiers — validated and stable
- `JARGON_REPLACEMENTS` — 25 entries, correct
- `generate_consumer_output()` system_prompt — recently stabilized
- `load_engine_criteria()` — Engine_review_score + Category_modules loading works
- Railway deployment configuration — until Playwright is ready
- CORS origins
- `_parse_servings_per_day()` — recently fixed for 1–2 range and usage_instructions fallback

---

## TYPESCRIPT MIGRATION PATH (PATH B/C — FUTURE)

If TypeScript migration is chosen after Phase 6:

1. **Extraction layer only** (PATH C): Rewrite layers 1–3 in TypeScript with Playwright+Cheerio+Zod. Python Flask keeps layers 4–6. TypeScript extractor runs as Railway sidecar, Python scoring engine calls it via HTTP.

2. **Full migration** (PATH B): Rewrite entire stack. Port all Python logic to TypeScript. Estimated: 4–6 weeks. Risk: scoring algorithm drift between Python original and TypeScript port.

3. **Prerequisites before either path:**
   - All Excel databases fully loaded and tested (Tier 3 complete)
   - Evidence layer validated in production (Tier 2 complete)
   - Playwright extraction tested on 20+ Dutch supplement sites (Tier 5 complete)
   - Full test suite with known-good scrape outputs

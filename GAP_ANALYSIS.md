# GAP_ANALYSIS.md — Funcify Codebase Analysis
*Generated: 2026-05-07 | Based on engine.py @ d2f33f1 | 1267 lines*

---

## 1. ARCHITECTURE OVERVIEW

### What exists
Single monolithic `engine.py` (Python 3.12, Flask). No module separation. All logic — scraping, extraction, evaluation, scoring, output generation — lives in one file.

```
engine.py (1267 lines)
├── Global data (30–168)      — CF rules, bioavailability, synonyms, EXCIPIENT_KEYWORDS, JARGON_REPLACEMENTS
├── Helpers (169–392)         — jargon, bioavailability lookups, context flags, cofactor checks, price
├── Excel loader (394–468)    — lazy-load Engine_review_score + Category_modules from .xlsx
├── Scoring engine (471–638)  — detect_product_type, evaluate_criteria_with_claude, calculate_score, determine_verdict
├── Output generator (641–754) — generate_consumer_output (Claude call)
├── Scraper layer (757–1111)  — fetch_html, fetch_with_scrapingbee, scrape_product, extract_with_claude, anthropic_fallback
└── Routes (1113–1267)        — /scrape, /score, /health
```

### Deployment
- Railway (Python 3.12, Gunicorn)
- `Procfile`: `web: python engine.py`
- `nixpacks.toml`: installs Playwright Chromium via `playwright install chromium --with-deps`
- Frontend: static HTML at `funcify.eu` (GitHub Pages via CNAME)
- API: `https://funcifyeu-production.up.railway.app`

### Dependencies (requirements.txt)
```
flask, flask-cors, requests, beautifulsoup4, lxml
anthropic, gunicorn, python-docx, openpyxl, playwright
```

**Critical observation:** `playwright` and `python-docx` are installed but **never imported or used** in `engine.py`. These are dead dependencies. The commit history shows Playwright was once a fallback scraper but has been removed from code while remaining in requirements.

---

## 2. SCRAPER FLOW

### Current flow
```
POST /scrape
  → fetch_with_scrapingbee(url)     [ScrapingBee, render_js=true, wait=2000ms]
      fallback: fetch_html(url)     [plain requests.get with browser UA]
  → scrape_product(soup, url)       [BeautifulSoup extraction]
  → _needs_second_pass check        [(no health_claims) OR (certs < 2)]
      → extract_with_claude()       [Claude, 6000 chars raw text]
  → is_sufficient() check           [product_name AND ≥1 ingredient]
      → anthropic_fallback()        [Claude, 2000 chars, no system prompt]
  → product_data["url"] = url
  → return product_data
```

### What works
- ScrapingBee JS rendering handles most Dutch supplement sites
- 3-pass fallback chain prevents total failure
- Second-pass Claude extraction compensates for missing health_claims/certifications

### What is broken / missing
- **Playwright is installed but never used** — the commit history promised a Playwright fallback but it was removed
- `fetch_html()` uses no retry logic; single timeout=15s failure → exception
- No cookie banner acceptance (JS-blocked content behind consent overlays)
- No accordion/tab expansion before extraction
- No screenshot capture
- No raw HTML preservation
- No PDF detection or download
- No structured data (JSON-LD, schema.org) extraction
- Second-pass trigger fires too eagerly: certifications < 2 is always true for clean products with 0–1 certs, triggering a Claude call even when scrape was complete

---

## 3. EXTRACTION LOGIC

### 3a. Product name
```python
extract_text_blocks(soup, ["h1", ".product-title", ".product-name", '[itemprop="name"]'])
```
- Takes first match, truncates to 200 chars
- **Risk:** Dutch sites often have nav h1 elements before product h1; first-match wins

### 3b. Brand
`_extract_brand()` — 5 strategies:
1. CSS selectors (`[itemprop="brand"]`, `.brand`, etc.)
2. Meta tags (`og:site_name`, `manufacturer`)
3. Title tag last segment after `|–`
4. Second word of title
5. URL hostname

- **Risk:** Fallback to URL hostname returns "bol" for bol.com products — correct brand is vendor, not marketplace
- **Risk:** Title-based fallback (`words[1]`) is completely unreliable

### 3c. Price
```python
re.search(r"[€\$]?\s*(\d+[,\.]\d{2})", page_text)
```
- First `€X,XX` match on entire page text
- **Risk:** Matches cross-sell prices, shipping costs, promo codes before actual product price
- **Risk:** No € symbol required (pattern allows none), so any decimal number matches

### 3d. Serving size / Usage instructions
```python
re.search(r"(?:Gebruik|Dosering|Aanbevolen dagelijkse|Innemen|per dag|dagdosering)[:\s]*([^\n.]{5,200})", page_text)
```
- Captures prose usage instruction into `serving_size`
- **Risk:** "Gebruik dit product als onderdeel van een gevarieerde voeding" captures irrelevant text
- **Risk:** `serving_size` becomes a prose string, not a structured count — downstream `_parse_servings_per_day()` tries to extract a digit from prose

### 3e. Package size
```python
_unit_pat = r"(\d+)\s*(softgels?|capsules?|tabletten?|vegicaps?|stuks?|caps?)"
```
Priority: `product_name` → `url` → page_text line scan (skipping daily-filter lines)
- **Risk:** Product name "Omega 3 2000mg 120 softgels" → correctly returns "120 softgels"
- **Risk:** Product name "3-in-1 Multi 1 capsule" → returns "1 capsule" (serving count, not package)
- **Risk:** URL check has no daily-filter applied — `/omega-3-2-capsules/` would return "2 capsules"

### 3f. Ingredients (3-pass regex)
```python
parse_ingredients_from_text(full_ingredient_text)
```

**Pass 1:** Line-split on `[;|\n]`, then AMOUNT_PAT per line
**Pass 2:** BROAD_INGREDIENT_PAT across full text
**Pass 3:** `_parse_nested_composition()` for "waarvan" sub-ingredients

- **Risk:** Ingredient section boundary detection is fragile:
  ```python
  r"(?:ingrediënten|samenstelling|...) [:\s]*(.{20,4000}?)(?:\n{2,}|©|Bewaar|...)"
  ```
  A page without these exact headers returns `page_text[:4000]` — entire page becomes ingredient source
- **Risk:** BROAD_INGREDIENT_PAT fires on marketing text and table headers
- **Risk:** Duplicate detection uses `seen_names` set but name comparison is exact lowercase — "Vitamin C" and "Vitamine C" are both retained
- **Risk:** `ingredients[:60]` hard truncates — no logging when truncation occurs
- **Risk:** Excipient classification is keyword-based only (`EXCIPIENT_KEYWORDS`) — custom excipients not in list are classified as "actief"
- **Risk:** No molecular form extraction from ingredient names — form is only populated if `\([^)]{3,50}\)` parenthetical exists

### 3g. Certifications
```python
for kw in CERT_KEYWORDS_FIXED:
    if kw.lower() in text_lower and kw not in found:
        found.append(kw)
```
- **Risk:** Simple substring match — "organic" matches "reorganization" in body text
- **Risk:** CERT_KEYWORDS_FIXED is hardcoded (32 terms) — non-standard certifications missed
- **Risk:** No source tracing — can't tell if certification was in a badge, body text, or footer

### 3h. Health claims
CSS selectors first, then regex:
```python
re.finditer(r"(?:ondersteunt|bevordert|helpt|verbetert|zorgt voor)[^.!?\n]{5,120}", page_text)
```
- **Risk:** Fires on all page text — sidebar products, newsletter, footer, cross-sells
- **Risk:** No deduplication against ingredient names
- **Risk:** capped at 8 claims; no priority ordering

---

## 4. NORMALIZATION LOGIC

### What exists
- `_normalize_name()`: lowercase → `INGREDIENT_SYNONYMS` lookup (40 entries, hardcoded)
- `simplify_jargon()`: recursive regex substitution on all string values in consumer_output
- `_split_active_excipients()`: filters by `type == "vul-additief"` only

### What is dangerous
- **Destructive normalization:** `simplify_jargon()` mutates strings without preserving originals. "biobeschikbaarheid" → "opneembaarheid" is applied to all string values recursively, including field names if they happen to match.
- **INGREDIENT_SYNONYMS is a flat dict** — no synonym chains, no fuzzy matching, no partial matching. "magnesium citraat" matches but "magnesium-citraat" does not.
- **No raw value preservation** — after normalization there is no way to trace what the original extracted text was.
- **No normalization logging** — silent substitution, no audit trail.

### What is missing
- Form normalization from BIOAVAILABILITY_RATIOS (these ratios exist but are never used in the main scoring flow)
- Unit normalization (IU ↔ mg conversions for CF threshold comparisons)
- Ingredient matching against Funcify Master Ingredient 2.0.xlsx (file exists, never loaded)
- Evidence objects — no source_url, source_type, raw_text, extraction_method per field

---

## 5. SCHEMA STRUCTURE

### Current schema (untyped Python dict)
```python
{
  "product_name": str,
  "brand_name": str,
  "ingredients": [{
    "name": str,
    "amount": float | None,
    "unit": str,
    "form": str,
    "type": "actief" | "vul-additief",
    "nested": bool
  }],
  "active_ingredients": [ingredient],   # filtered subset
  "excipients": [ingredient],           # filtered subset (still dicts, not names)
  "serving_size": str,                  # prose string, not structured count
  "usage_instructions": str,
  "package_size": str,
  "price": str,
  "health_claims": [str],
  "certifications": [str],
  "warnings": [str],
  "additional_info": str,               # page_text[:5000]
  "url": str,
  "context_flags_triggered": [str],
  "_source": "scraper" | "anthropic_fallback" | "error",
}
```

### What is missing from schema
- No `confidence` per field
- No `raw_text` per field
- No `extraction_method` per field
- No `evidence` object per field
- No `data_lacune` flags per field
- No ingredient `id` or `parent_id`
- No `proprietary_blend` membership flag
- No `ri_percentage` (reference intake)
- No `max_daily_dose`
- No molecular form linked to bioavailability data
- No category-specific fields (EPA/DHA for OMEGA3, CFU/strain for PROBIOTIC, etc.)
- No `engine_readiness` flag
- No `data_lacunes` array at product level

### Schema conflicts
- `excipients` field contains ingredient dicts but the `user_prompt` in `generate_consumer_output` renders them with `f"- {e}"`, producing Python dict repr strings. **This is a live bug.**
- `active_ingredients` and `ingredients` are redundant — `active_ingredients` is a filtered subset of `ingredients`, but both are stored and both are passed to different functions, creating drift risk if one is modified without the other.

---

## 6. OUTPUT STRUCTURE

### Current /score response
```json
{
  "product_name", "brand", "score", "kwalificatie", "verdict",
  "product_type", "product_complexity_tier", "critical_gate",
  "non_verifiable_count", "criteria_evaluated", "confidence",
  "price", "package_size", "price_per_day", "price_per_gram",
  "intake_advice", "context_flags_triggered",
  "wat_doet", "beoordeling_tabel", "highlights",
  "context_flags", "wat_zou_beter", "voor_wie", "consumer_summary",
  "low_confidence_warning"  // optional
}
```

### What is missing from output
- No evidence tracing
- No data_lacunes array
- No extraction coverage report
- No normalization report
- No ingredient-match report
- No per-field confidence
- No raw extracted values
- No source attribution
- No `engine_readiness` flag
- No `validation_errors`
- No raw-page.html, screenshot.png, or file outputs
- No extraction-coverage-report.json

---

## 7. DEPENDENCIES

| Package | Used | Purpose | Status |
|---------|------|---------|--------|
| flask | YES | HTTP API | OK |
| flask-cors | YES | CORS for funcify.eu | OK |
| requests | YES | HTTP fetching | OK |
| beautifulsoup4 | YES | HTML parsing | OK |
| lxml | YES | BS4 parser | OK |
| anthropic | YES | Claude API (4 call sites) | OK |
| gunicorn | YES (deploy) | WSGI server | OK |
| openpyxl | YES | Excel loading | OK (but partial — only 2 of 3 Excel files) |
| python-docx | NO | Docx reading | **DEAD DEPENDENCY** |
| playwright | NO | Browser scraping | **DEAD DEPENDENCY** (installed, never called) |

---

## 8. PLAYWRIGHT USAGE

**Playwright is not used.** It is listed in `requirements.txt` and `nixpacks.toml` installs Chromium, but no import or call exists in `engine.py`. Previous commit history shows it was added and then removed from the scraping chain.

**Gap:** The target Funcify architecture requires Playwright as the primary browser for:
- Cookie banner acceptance
- Accordion/tab expansion
- Network idle waiting
- Screenshot capture
- Raw HTML preservation
- PDF detection

Currently ScrapingBee handles JS rendering as a paid external service. Playwright would make rendering self-hosted and controllable.

---

## 9. INGREDIENT PARSING

### What works
- 3-pass strategy catches most ingredients from well-formatted Dutch supplement pages
- AMOUNT_PAT handles mg, g, mcg, µg, ml, IU, IE, KVE, CFU correctly
- "waarvan" sub-ingredient parsing handles nested compositions
- Excipient classification by EXCIPIENT_KEYWORDS covers common Dutch/English excipients
- `seen_names` deduplication prevents double entries from overlapping passes

### Where parsing fails
1. **Ingredient section boundary not found** → entire `page_text[:4000]` becomes ingredient source → BROAD_INGREDIENT_PAT fires on navigation, marketing text, prices
2. **Proprietary blends** → "Proprietary Blend 500mg" is extracted as a single ingredient with 500mg; sub-ingredients without doses are not extracted or flagged as DATA_LACUNE
3. **"Waarvan" without amount** → `_parse_nested_composition()` requires AMOUNT_PAT match in the sub-text — named sub-components without amounts are silently dropped
4. **Multiple units in one line** → "EPA 180mg / DHA 120mg" — AMOUNT_PAT finds first match (180mg), line is consumed; second match (120mg) is lost unless BROAD_INGREDIENT_PAT catches it
5. **Ingredient name truncation** → `name = line[:match.start()].strip()` — if the ingredient section text is not clean, name captures HTML artifacts
6. **Forms not extracted** → If ingredient is listed as "Magnesium (as Magnesium Bisglycinate) 150mg", the parenthetical "as Magnesium Bisglycinate" is captured as `form` only if it matches `\([^)]{3,50}\)` — the "as" prefix is included in the form string, not normalized
7. **Dutch vs English mixing** → "Vitamin C (ascorbinezuur)" — both name and form would contain Dutch/English mixtures without normalization
8. **No match against Master Ingredient DB** → extracted ingredient names are not validated against `Funcify. Master Ingredient 2.0.xlsx`

### What is missing
- Proprietary blend detection and DATA_LACUNE flagging
- Per-ingredient evidence object (raw_text, line_number, extraction_pass)
- Form normalization against ingredient_forms sheet
- RI% extraction from ingredient lines
- Max daily dose calculation
- Ingredient ID assignment
- Category-specific extraction (EPA/DHA split, CFU/strain, extract ratio)

---

## 10. SUPPLEMENT FACTS HANDLING

### Current state
The engine has **no special handling for Supplement Facts panels.** Dutch pages use "Voedingswaarden", "Samenstelling", or "Ingrediënten" headers. The regex:
```python
r"(?:ingrediënten|samenstelling|inhoudsstoffen|supplement\s*facts|ingredients)[:\s]*(.{20,4000}?)"
```
attempts to capture the section but:
- **Has no table-aware parsing** — `<table>` elements with ingredient rows are parsed as flat text, losing the column structure
- **Has no structured data extraction** — JSON-LD `NutritionInformation` schema.org markup is never checked
- **Has no image-based Supplement Facts extraction** — many Dutch sites use images for the facts panel
- Encoding issues: `ingrediënten` (with ë) is in the regex but page encoding issues may render it as `ingredi?nten`

---

## 11. DATA VALIDATION

### Current state: None

There is **no validation layer.** Data flows directly from extraction to Claude evaluation without:
- Field-type validation
- Required field checking
- Range validation (amount > 0, unit in allowed set)
- Cross-field consistency (serving_size parseable as number)
- Schema validation (no Pydantic, no dataclass, no TypedDict)

**Consequences:**
- `amount: None` reaches `evaluate_criteria_with_claude` and Claude must handle missing data
- `serving_size: "Gebruik 2x daags 1 capsule met water"` reaches `_parse_servings_per_day` which correctly handles it for the 1-2 case but fails on "3-in-1 formula 1 capsule"
- `price: "€29,95"` is passed as string through the system and only parsed at `calculate_price_per_day`

---

## 12. ERROR HANDLING

### What exists
```python
try:
    soup = fetch_with_scrapingbee(url)
    ...
except Exception as e:
    product_data["_error"] = str(e)

try:
    product_data = extract_with_claude(url, page_text, product_data)
except Exception:
    pass  # silent fail

try:
    product_data = anthropic_fallback(url, page_text)
except Exception as e:
    product_data["_fallback_error"] = str(e)
```

### What is wrong
1. **`except Exception: pass`** in second-pass extraction — Claude API failures, JSON parse errors, quota errors all silently swallowed
2. **No error classification** — network timeout, scraping service error, Claude quota, JSON parse error all produce the same `_error` string
3. **No retry logic** — transient network failures cause permanent fallback to lower-quality extraction
4. **`/score` route has one bare `except Exception`** — any error in the 2000-line scoring chain returns generic `{"error": "Engine error: ..."}` with no detail
5. **No structured logging** — no request IDs, no timing, no log levels, no error tracking
6. **`extract_with_claude` silently returns unchanged `product_data`** on exception — caller cannot tell if enrichment occurred

---

## 13. HALLUCINATION RISKS

### H1. `extract_with_claude` — no hallucination guard (HIGH)
Prompt sends 6000 chars raw page text and asks Claude to extract ingredients, certifications, health_claims, package_size, usage_instructions. No rule against inventing data not present in the text. Claude may infer molecular forms, certifications mentioned in marketing ("GMP-gecertificeerd productieproces" ≠ GMP certification), or dosages not explicitly stated.

### H2. `anthropic_fallback` — no system prompt, 2000 chars (CRITICAL)
Only 2000 chars of context. No system prompt for grounding. Highest hallucination risk in the system. Claude is asked to produce a complete product extraction from very limited context, making invented amounts and forms highly likely.

### H3. `generate_consumer_output` — certifications from additional_info (HIGH)
System prompt instructs: "Als informatie in additional_info staat maar niet in gestructureerde velden, gebruik die dan alsnog." Claude receives 3000 chars of raw page text and is told to find certifications and health claims in it. Marketing language ("lab-tested purity") may be elevated to oordeel: "Goed" without being an actual third-party certification.

### H4. `evaluate_criteria_with_claude` — page text in additional_info (MEDIUM)
`additional_info[:1500]` is passed to the scoring evaluator. Claude may use marketing claims in this text as "evidence" for criterion evaluations, elevating scores based on brand claims rather than verifiable data.

### H5. Certifications from free text (MEDIUM)
`_extract_certifications()` matches `CERT_KEYWORDS_FIXED` against full page text. "organic" matches any page mentioning organic products. "GMP" matches any page mentioning "GMP-achtig" or similar. These false positives then get passed to Claude as confirmed certifications.

### H6. `extract_with_claude` certifications replace without validation (MEDIUM)
Claude-returned certifications are merged into `product_data["certifications"]` without validation against `CERT_KEYWORDS_FIXED`. Claude may invent certification names not on any real certification body list.

---

## 14. DATA LOSS RISKS

### D1. `additional_info` capped at 5000 chars
Long ingredient tables and detailed product pages lose content. Supplement facts panels that appear late in the HTML are truncated.

### D2. `extract_with_claude` uses 6000 chars — inconsistent window
Different character window than `additional_info[:5000]`. Certifications found by extract_with_claude are merged into certifications[] but won't appear in the additional_info window used by generate_consumer_output.

### D3. Ingredient list hard-capped at 60 (line 872) — silent truncation
No log entry when truncation occurs. Complex multivitamins with 30+ ingredients lose later ingredients silently.

### D4. `BIOAVAILABILITY_RATIOS` never used in main flow
`get_bioavailability_info()` and `get_better_alternatives()` exist but are called from nowhere in `/scrape` or `/score`. All 28 bioavailability ratios are dead code in the actual scoring path.

### D5. `extract_with_claude` ingredients: only fills when empty
If scraper found 2 partial ingredients (amount=None), Claude's enriched extraction with 10 full ingredients is discarded. Partial bad data blocks good data.

### D6. Health claims capped at 8
Products with extensive claims pages lose coverage.

### D7. Excel criteria silently capped at 50 core + 20 module
If `Funcify__Engine_Review_V4.xlsx` adds more criteria, they're silently dropped.

### D8. `active_ingredients` and `excipients` are filtered copies
After extract_with_claude merges new ingredients into `product_data["ingredients"]`, the `active_ingredients` and `excipients` subsets are NOT re-filtered. They can go stale relative to the master `ingredients` list.

### D9. Fallback score overrides critical_fail cap
Lines 1218–1226: if no Excel criteria loaded, beoordeling_tabel-based fallback score can push score_100 above 50, effectively overriding the critical_fail cap. `determine_verdict` is called again with the new score but consumer_output was generated with the original (capped) score — narrative and score become inconsistent.

---

## 15. OVERWRITE RISKS

### O1. Raw vs normalized data — no separation
`product_data` is mutated in-place throughout the pipeline:
- `scrape_product()` writes initial values
- `extract_with_claude()` overwrites health_claims, certifications, package_size, usage_instructions
- `anthropic_fallback()` replaces the entire dict
- `score()` adds context_flags_triggered, certifications (sustainability)

No original values are preserved. If `extract_with_claude` overwrites a correct serving_size with a wrong one, the original is gone.

### O2. Certification append without ground truth check
Omega-3 sustainability check (lines 1161–1167) appends `"DATA_LACUNE: Geen duurzaamheidscertificering..."` to the certifications array. This DATA_LACUNE string then flows into `generate_consumer_output` as a "certification" and Claude must interpret it correctly. If Claude treats it as a real certification, output is wrong.

### O3. `extract_with_claude` health_claims: full replacement
If scraper found 5 correct claims and Claude returns 3 (possibly different) claims, the 5 originals are discarded. No merge strategy preserves original claims alongside Claude-extracted claims.

### O4. `product_data["url"]` written at route level, not extraction level
URL is added to product_data AFTER scraping. Functions called within scraping that need URL receive it as a parameter, not from product_data. If product_data is later passed without the URL being set, url is missing.

---

## 16. SCALABILITY BOTTLENECKS

### S1. 2–4 synchronous Claude API calls per /score request
- `evaluate_criteria_with_claude()` — 1 call (4000 tokens out)
- `generate_consumer_output()` — 1 call (3000 tokens out)
- `extract_with_claude()` — 1 call (2000 tokens out, on /scrape)
- `anthropic_fallback()` — 1 call (1500 tokens out, if triggered)

All calls are sequential. No batching, no parallelization. A single /score request takes 10–20 seconds minimum.

### S2. Excel file loaded on first request, cached in `_cache` dict
This works for a single process. Under Gunicorn with multiple workers, each worker loads its own copy. Fine for small scale, problematic at scale.

### S3. ScrapingBee dependency
External paid service. Rate limits, pricing, and availability are not controlled. No local fallback beyond plain requests.

### S4. No request queuing
Long-running /score requests block. No async, no queue. High concurrent load will exhaust Gunicorn workers.

---

## 17. MAINTAINABILITY ISSUES

### M1. 1267 lines in one file
No module separation makes it impossible to:
- Test extractors independently
- Swap the scraping strategy
- Add new extraction modules
- Maintain or update CF rules independently

### M2. Four hardcoded model strings
`"claude-sonnet-4-5"` appears at lines 569, 745, 1029, 1077. Upgrading requires 4 changes and a careful audit.

### M3. Global data structures mixed with business logic
`CONTEXT_FLAG_RULES`, `BIOAVAILABILITY_RATIOS`, `INGREDIENT_SYNONYMS` etc. are defined at module level. They should come from Excel (and for CF rules, from `Context_Flag_Rules` sheet which exists in Engine Review V4 but is NOT loaded).

### M4. `Funcify. Master Ingredient 2.0.xlsx` never loaded
This is the authoritative ingredient database. It contains: ingredient_master, Ingredient_synonyms, ingredient_forms, ingredients_bioavailability_fac, ingredient_stability, ingredient_cofactors, ingredient_antagonist, ingredient_transport_competition, ingredient_enzymes, ingredient_formulation, ingredient_biomarker_effect, biomarkers, ingredient_use_case-effect. None of these sheets are read. The hardcoded Python dicts are a poor substitute.

### M5. Dead code after user_prompt refactor
Lines 648–658: `ingredients_text`, `excipients_text`, `usage_instructions_text`, `flags_text`, `certifications_text` — assembled but not used in the current user_prompt.

### M6. CF rules hardcoded instead of loaded from Excel
`Context_Flag_Rules` sheet exists in Engine Review V4 but is never loaded. The hardcoded `CONTEXT_FLAG_RULES` list is a stale copy that diverges from the Excel source of truth over time.

---

## 18. CONFLICTS WITH FUNCIFY REQUIREMENTS

| Requirement | Current State | Gap |
|------------|--------------|-----|
| Layer 1: Raw extraction | Partially — page_text[:5000] stored but not raw HTML | No raw HTML, no screenshot |
| Layer 2: Structured extraction | Partially — scrape_product() extracts structured fields | No evidence objects, no confidence per field |
| Layer 3: Normalization | Partially — INGREDIENT_SYNONYMS, simplify_jargon | Destructive, no audit trail |
| Layer 4: Validation | MISSING | No schema validation |
| Layer 5: Engine preparation | Partially — /score prepares eval | No engine_readiness flag, no data_lacunes array |
| Layer 6: Review engine | Partially — Claude scoring | OK conceptually but mixed into Layer 5 |
| Evidence objects | MISSING | No raw_text, source_type, extraction_method per field |
| DATA_LACUNE per field | MISSING | Only in Claude output, not in extraction |
| Master Ingredient DB | MISSING | File exists, never loaded |
| Consumer UI Excel | MISSING | File exists, never loaded |
| CF rules from Excel | MISSING | Hardcoded instead |
| Playwright | MISSING | Installed, never called |
| Proprietary blend detection | MISSING | No flag, no DATA_LACUNE |
| Category-specific fields | MISSING | No EPA/DHA split, no CFU/strain, no extract ratio |
| Per-ingredient confidence | MISSING | No confidence field on ingredients |
| Output files | MISSING | No raw-page.html, screenshot.png, JSON reports |
| Modular architecture | MISSING | Monolithic single file |
| TypeScript / Zod / Cheerio | MISSING | Python/Flask only |

---

## SUMMARY: WHAT WORKS, WHAT DOESN'T

### Stable and valuable — keep
- `calculate_score()` with DATA_LACUNE semantics (EXACT/AFGELEID/ONVOLLEDIG/NIET_GEVONDEN weighting)
- `determine_verdict()` tier logic
- `evaluate_context_flags()` + `evaluate_cofactor_checks()` (rules themselves valid, but source them from Excel)
- `JARGON_REPLACEMENTS` (25 entries, working)
- `calculate_price_per_day()` and `_parse_servings_per_day()`
- `generate_consumer_output()` system_prompt (recently stabilized)
- `load_engine_criteria()` — Engine_review_score + Category_modules loading
- `_needs_second_pass` logic (concept correct, just too eager)
- CORS configuration
- Railway deployment setup

### Working but needs extension
- `scrape_product()` — works for basic extraction, needs evidence objects and category-specific fields
- `evaluate_criteria_with_claude()` — solid approach, needs serving size fix and model constant
- `_extract_certifications()` — works for known certs, needs tighter source-context matching
- `parse_ingredients_from_text()` — catches most cases, needs proprietary blend detection and evidence objects
- `anthropic_fallback()` — needs system prompt and larger context window

### Needs replacement
- Hardcoded `BIOAVAILABILITY_RATIOS`, `INGREDIENT_SYNONYMS`, `CONTEXT_FLAG_RULES` → load from Excel
- `simplify_jargon()` → run after preserving raw values, not before
- Price regex (first match wins) → smarter price extraction with position heuristic
- Package size extraction (product_name first) → sanity check on extracted count

### Critical bugs to fix immediately
1. **Line 699:** `f"- {e}"` where `e` is a dict → must be `f"- {e.get('name', '')}"`
2. **Dead code lines 648–658** — remove or wire into user_prompt
3. **`anthropic_fallback` no system prompt** — add grounding rules
4. **`extract_with_claude` ingredients: partial block** — merge instead of replace-if-empty

### Dangerous — address before scaling
1. No hallucination guards on extraction Claude calls
2. Destructive normalization without raw value preservation
3. Silent exception swallowing in extract_with_claude
4. Certification false positives from full-page keyword scan
5. Package size from product_name without sanity check

import os
import re
import json
import requests
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)
CORS(app, origins=["https://funcify.eu", "https://fullfunc.github.io"])

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_cache = {}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}

def load_engine_criteria():
    if "criteria" in _cache:
        return _cache["criteria"]
    criteria = []
    if not HAS_OPENPYXL:
        return criteria
    engine_path = os.path.join(BASE_DIR, "Funcify__Engine_Review_V4.xlsx")
    if not os.path.exists(engine_path):
        return criteria
    try:
        wb = openpyxl.load_workbook(engine_path, data_only=True)
        if "Engine_review_score" in wb.sheetnames:
            ws = wb["Engine_review_score"]
            header_row = None
            for row in ws.iter_rows():
                vals = [str(c.value or "").strip() for c in row]
                if "Category" in vals and "Criterion" in vals and "Weight" in vals:
                    header_row = {v: i for i, v in enumerate(vals)}
                    break
            if header_row:
                for row in ws.iter_rows():
                    vals = [c.value for c in row]
                    category = str(vals[header_row.get("Category", 0)] or "")
                    if category in ["Category", "SUMMARY", "INPUT", "HELPERS", "MODEL SETTINGS", ""]:
                        continue
                    criterion = str(vals[header_row.get("Criterion", 1)] or "")
                    weight_raw = vals[header_row.get("Weight", 2)]
                    critical = str(vals[header_row.get("Critical? (Y/N)", 3)] or "N").strip().upper()
                    applies_to = str(vals[header_row.get("AppliesTo", 4)] or "ALL").strip()
                    consumer_impact = str(vals[header_row.get("Consumer_impact", 15)] or "")
                    failure_type = str(vals[header_row.get("Failure_type", 14)] or "")
                    try:
                        weight = float(weight_raw) if weight_raw else 0
                    except (ValueError, TypeError):
                        weight = 0
                    if criterion and weight > 0:
                        criteria.append({
                            "source": "core",
                            "category": category,
                            "criterion": criterion,
                            "weight": weight,
                            "critical": critical == "Y",
                            "applies_to": applies_to,
                            "consumer_impact": consumer_impact,
                            "failure_type": failure_type
                        })
        if "Category_modules" in wb.sheetnames:
            ws = wb["Category_modules"]
            header_row = None
            for row in ws.iter_rows():
                vals = [str(c.value or "").strip() for c in row]
                if "ModuleType" in vals and "Criterion" in vals and "Weight" in vals:
                    header_row = {v: i for i, v in enumerate(vals)}
                    break
            if header_row:
                for row in ws.iter_rows():
                    vals = [c.value for c in row]
                    module_type = str(vals[header_row.get("ModuleType", 0)] or "").strip()
                    if not module_type or module_type in ["ModuleType", "Producttype", ""]:
                        continue
                    criterion = str(vals[header_row.get("Criterion", 2)] or "")
                    weight_raw = vals[header_row.get("Weight", 3)]
                    critical = str(vals[header_row.get("Critical", 4)] or "N").strip().upper()
                    consumer_impact = str(vals[header_row.get("Consumer_impact", 15)] or "")
                    try:
                        weight = float(weight_raw) if weight_raw else 0
                    except (ValueError, TypeError):
                        weight = 0
                    if criterion and weight > 0:
                        criteria.append({
                            "source": "module",
                            "category": f"Module_{module_type}",
                            "criterion": criterion,
                            "weight": weight,
                            "critical": critical == "Y",
                            "applies_to": module_type,
                            "consumer_impact": consumer_impact,
                            "failure_type": "Quality fail"
                        })
        _cache["criteria"] = criteria
    except Exception as e:
        print(f"Error loading criteria: {e}")
    return criteria


def detect_product_type(product_data):
    all_text = " ".join([
        product_data.get("product_name", ""),
        product_data.get("brand_name", ""),
        " ".join([i.get("name", "") for i in product_data.get("ingredients", [])]),
        " ".join(product_data.get("health_claims", []))
    ]).lower()
    if any(x in all_text for x in ["epa", "dha", "omega-3", "omega 3", "visolie", "fish oil", "krill"]):
        return "OMEGA3"
    if any(x in all_text for x in ["probiotic", "lactobacillus", "bifidobacterium", "cfu"]):
        return "PROBIOTIC"
    if any(x in all_text for x in ["vitamine", "vitamin", "d3", "b12", "b6", "folaat", "folate"]):
        return "VITAMIN"
    if any(x in all_text for x in ["magnesium", "zink", "zinc", "ijzer", "iron", "calcium", "selenium"]):
        return "MINERAL"
    if any(x in all_text for x in ["whey", "protein", "eiwit", "caseïne", "casein"]):
        return "PROTEIN"
    if any(x in all_text for x in ["creatine", "bcaa", "leucine", "beta-alanine"]):
        return "SPORT"
    if any(x in all_text for x in ["curcuma", "turmeric", "berberine", "resveratrol", "quercetin"]):
        return "POLYPHENOL"
    if any(x in all_text for x in ["ashwagandha", "rhodiola", "valeriaan", "ginkgo", "extract"]):
        return "BOTANICAL"
    return "ALL"


def evaluate_criteria_with_claude(product_data, criteria, product_type):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    relevant_criteria = []
    for c in criteria:
        applies = c["applies_to"]
        if applies == "ALL" or applies == product_type or (applies == "SPECIFIC" and product_type != "ALL"):
            relevant_criteria.append(c)
    core_criteria = [c for c in relevant_criteria if c["source"] == "core"][:50]
    module_criteria = [c for c in relevant_criteria if c["source"] == "module"][:20]
    all_relevant = core_criteria + module_criteria
    criteria_text = "\n".join([
        f"{i+1}. [{c['category']}] {c['criterion']} (weight={c['weight']}, critical={'JA' if c['critical'] else 'NEE'})"
        for i, c in enumerate(all_relevant)
    ])
    ingredients_text = "\n".join([
        f"- {ing.get('name', 'onbekend')}: {ing.get('amount', '?')} {ing.get('unit', '')} ({ing.get('form', 'vorm onbekend')})"
        for ing in product_data.get("ingredients", [])
    ])
    system_prompt = """Je bent de Funcify beoordelingsengine. Evalueer elk criterium op basis van de productdata.
REGELS:
- Pass=1: informatie aanwezig en voldoet aan criterium
- Fail=0: informatie aanwezig maar voldoet NIET (bewijs van slechte kwaliteit)
- Onbekend=-1: informatie niet beschikbaar (DATA_LACUNE, neutraal, nooit negatief)
- Ontbrekende data is NOOIT negatief. Alleen bewezen slechte kwaliteit telt als Fail.
- VERBODEN: "studies tonen", "literatuur zegt", "klinisch bewezen" - alleen label/website feiten
- Geef alleen valide JSON terug."""
    user_prompt = f"""Product: {product_data.get('product_name', 'Onbekend')}
Merk: {product_data.get('brand_name', 'Onbekend')}
Type: {product_type}
Serving size: {product_data.get('serving_size', 'onbekend')}
Verpakkingsgrootte: {product_data.get('package_size', 'onbekend')}
Prijs: {product_data.get('price', 'onbekend')}

INGREDIENTEN:
{ingredients_text if ingredients_text else 'Geen ingredienten gevonden'}

GEZONDHEIDSCLAIMS:
{chr(10).join(product_data.get('health_claims', [])) or 'Geen'}

CERTIFICERINGEN:
{', '.join(product_data.get('certifications', [])) or 'Geen gevonden'}

AANVULLENDE INFO:
{product_data.get('additional_info', '')[:1500]}

CRITERIA:
{criteria_text}

Geef terug als JSON:
{{
  "evaluations": [
    {{"criterion_index": 1, "pass_value": 1, "evidence": "korte uitleg", "data_quality": "EXACT|AFGELEID|ONVOLLEDIG|NIET_GEVONDEN"}}
  ],
  "critical_gate_triggered": false,
  "key_strengths": ["max 3 sterke punten"],
  "key_weaknesses": ["max 3 zwakke punten"],
  "additives_found": ["vulstoffen lijst"],
  "inferior_forms_found": ["inferieure vormen indien aanwezig"],
  "price_per_serving": "berekend indien mogelijk"
}}"""
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=4000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    return json.loads(raw), all_relevant


def calculate_score(evaluations_data, criteria):
    evaluations = evaluations_data.get("evaluations", [])
    total_raw_score = 0
    total_max_weight = 0
    critical_fail = False
    non_verifiable_count = 0
    for i, criterion in enumerate(criteria):
        eval_item = next((e for e in evaluations if e.get("criterion_index") == i + 1), None)
        pass_value = eval_item.get("pass_value", -1) if eval_item else -1
        weight = criterion["weight"]
        if pass_value == -1:
            non_verifiable_count += 1
            total_max_weight += weight
        elif pass_value == 1:
            total_raw_score += weight
            total_max_weight += weight
        elif pass_value == 0:
            total_max_weight += weight
            if criterion["critical"]:
                critical_fail = True
    if total_max_weight == 0:
        return 0, False, 0
    score_pct = total_raw_score / total_max_weight
    if critical_fail:
        score_pct = min(score_pct, 0.49)
    return score_pct, critical_fail, non_verifiable_count


def determine_verdict(score_pct, critical_fail):
    score_100 = round(score_pct * 100)
    if critical_fail and score_100 < 50:
        return score_100, "Afkeur", "Af te raden"
    elif score_100 >= 85:
        return score_100, "Elite", "Koopwaardig"
    elif score_100 >= 70:
        return score_100, "Degelijk", "Koopwaardig"
    elif score_100 >= 55:
        return score_100, "Matig", "Alleen met context"
    else:
        return score_100, "Afkeur", "Af te raden"


def generate_consumer_output(product_data, evaluations_data, criteria, score_100, kwalificatie, verdict, product_type, critical_fail):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    strengths = evaluations_data.get("key_strengths", [])
    weaknesses = evaluations_data.get("key_weaknesses", [])
    additives = evaluations_data.get("additives_found", [])
    inferior_forms = evaluations_data.get("inferior_forms_found", [])
    price_per_serving = evaluations_data.get("price_per_serving", "")
    ingredients_text = "\n".join([
        f"- {ing.get('name', '')}: {ing.get('amount', '')} {ing.get('unit', '')} ({ing.get('form', '')})"
        for ing in product_data.get("ingredients", [])
    ])
    system_prompt = """Je bent de Funcify consumer output generator. Schrijf heldere eerlijke beoordelingen in Nederlands.
TOON: Direct, eerlijk, consumentvriendelijk. Geen jargon. Geen medisch advies.
VERBODEN: studies tonen, literatuur zegt, klinisch bewezen, therapeutisch.
DATA_LACUNE = neutraal, nooit negatief. Geef alleen valide JSON terug."""
    user_prompt = f"""Product: {product_data.get('product_name', 'Onbekend')} ({product_data.get('brand_name', 'Onbekend')})
Score: {score_100}/100 | Kwalificatie: {kwalificatie} | Verdict: {verdict}
Critical gate: {'JA' if critical_fail else 'NEE'}
Prijs per serving: {price_per_serving or 'onbekend'}

INGREDIENTEN:
{ingredients_text or 'Niet gevonden'}

VULSTOFFEN: {', '.join(additives) or 'Geen'}
INFERIEURE VORMEN: {', '.join(inferior_forms) or 'Geen'}
STERKE PUNTEN: {', '.join(strengths) or 'Geen'}
ZWAKKE PUNTEN: {', '.join(weaknesses) or 'Geen'}

Genereer JSON:
{{
  "wat_doet": "2-3 zinnen wat dit supplement doet in gewone taal",
  "beoordeling_tabel": [
    {{"aspect": "Moleculaire vormen", "bevinding": "wat staat op label over vormen", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Doseringen", "bevinding": "exacte doseringen per dagdosering", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Bioavailabiliteit", "bevinding": "hoe goed worden vormen opgenomen", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Transparantie label", "bevinding": "zijn ingredienten, vormen, doseringen volledig", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Certificeringen", "bevinding": "welke certificeringen aanwezig", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Gezondheidsclaims", "bevinding": "zijn claims reeel en in lijn met ingredienten", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Serving size", "bevinding": "serving size, dagdosering en aantal servings", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}},
    {{"aspect": "Vulstoffen en additieven", "bevinding": "alle vulstoffen en additieven uit ingrediëntenlijst", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}}
  ],
  "highlights": [
    {{"type": "positief", "tekst": "sterk punt"}},
    {{"type": "negatief", "tekst": "zwak punt"}}
  ],
  "context_flags": ["waarschuwingen voor specifieke groepen indien van toepassing"],
  "wat_zou_beter": "wat zou een beter product anders doen",
  "voor_wie": "voor wie nog bruikbaar",
  "consumer_summary": "2-3 zinnen samenvatting"
}}"""
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=3000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    return json.loads(raw)


AMOUNT_PAT = re.compile(r"(\d[\d.,]*)\s*(mg|g|mcg|µg|ug|ml|ie|iu|kve|cfu|%)", re.I)

def fetch_html(url):
    response = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
    response.raise_for_status()
    return BeautifulSoup(response.content, "lxml")

def extract_text_blocks(soup, selectors):
    texts = []
    for sel in selectors:
        for el in soup.select(sel):
            t = el.get_text(" ", strip=True)
            if t:
                texts.append(t)
    return texts

def parse_ingredients_from_text(text):
    ingredients = []
    lines = re.split(r"[,;|\n]", text)
    for line in lines:
        line = line.strip()
        if not line or len(line) < 3:
            continue
        match = AMOUNT_PAT.search(line)
        if match:
            amount_str = match.group(1).replace(",", ".")
            unit = match.group(2).lower()
            name = line[:match.start()].strip(" ()-:")
            if name:
                ingredients.append({"name": name, "amount": float(amount_str) if amount_str else None, "unit": unit, "form": ""})
    return ingredients

def scrape_product(soup):
    selectors_name = ["h1", ".product-title", ".product-name", '[itemprop="name"]']
    selectors_brand = [".brand", ".manufacturer", '[itemprop="brand"]']
    name_texts = extract_text_blocks(soup, selectors_name)
    brand_texts = extract_text_blocks(soup, selectors_brand)
    price_texts = extract_text_blocks(soup, [".price", "[class*='price']", "[itemprop='price']"])
    page_text = soup.get_text(" ", strip=True)
    ing_match = re.search(
        r"(?:ingredi[eë]nten|samenstelling|inhoudsstoffen|ingredients)[:\s]*(.{20,800}?)(?:\n\n|\*|©|Bewaar|Gebruiksaanwijzing|Aanbevolen|$)",
        page_text, re.I | re.S
    )
    full_ingredient_text = ing_match.group(1) if ing_match else page_text[:2000]
    ingredients = parse_ingredients_from_text(full_ingredient_text)
    cert_keywords = ["gmp", "iso", "nzvt", "informed sport", "ifos", "creapure", "vegan", "biologisch", "organic", "kosher", "halal", "glutenvrij", "lactosevrij"]
    certifications = [kw for kw in cert_keywords if kw.lower() in page_text.lower()]
    claims_texts = extract_text_blocks(soup, [".claims", "[class*='claim']", ".usp"])
    health_claims = [c for c in claims_texts if len(c) > 10][:5]
    serving_match = re.search(r"(?:per\s+dagdosering|serving size|dagdosering)[:\s]*([^\n]+)", page_text, re.I)
    serving_size = serving_match.group(1).strip() if serving_match else ""
    package_match = re.search(r"(\d+)\s*(?:capsules?|tabletten?|softgels?|vegicaps?|stuks?)", page_text, re.I)
    package_size = package_match.group(0) if package_match else ""
    return {
        "product_name": (name_texts[0] if name_texts else "")[:200],
        "brand_name": (brand_texts[0] if brand_texts else "")[:100],
        "ingredients": ingredients[:40],
        "serving_size": serving_size[:200],
        "package_size": package_size[:100],
        "price": (price_texts[0] if price_texts else "")[:50],
        "health_claims": health_claims,
        "certifications": certifications,
        "additional_info": page_text[:2000]
    }

def is_sufficient(data):
    return bool(data.get("product_name")) and len(data.get("ingredients", [])) >= 1

def anthropic_fallback(url, page_text=""):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    prompt = f"""Analyseer deze supplementpagina en extraheer productinformatie.
URL: {url}
Pagina tekst: {page_text[:3000] if page_text else 'Niet beschikbaar'}

Geef terug als JSON:
{{"product_name": "naam", "brand_name": "merk", "ingredients": [{{"name": "naam", "amount": 0, "unit": "mg", "form": "vorm"}}], "serving_size": "serving", "package_size": "verpakkingsgrootte", "price": "prijs", "health_claims": ["claim"], "certifications": ["cert"], "additional_info": "info"}}"""
    response = client.messages.create(model="claude-sonnet-4-5", max_tokens=1500, messages=[{"role": "user", "content": prompt}])
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    try:
        return json.loads(raw)
    except:
        return {"product_name": "", "brand_name": "", "ingredients": [], "health_claims": [], "certifications": [], "additional_info": ""}


@app.route("/scrape", methods=["POST"])
def scrape():
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "URL required"}), 400
    product_data = {}
    page_text = ""
    try:
        soup = fetch_html(url)
        page_text = soup.get_text(" ", strip=True)
        product_data = scrape_product(soup)
        product_data["_source"] = "scraper"
    except Exception as e:
        product_data["_source"] = "error"
        product_data["_error"] = str(e)
    if not is_sufficient(product_data):
        try:
            product_data = anthropic_fallback(url, page_text)
            product_data["_source"] = "anthropic_fallback"
        except Exception as e:
            product_data["_fallback_error"] = str(e)
    return jsonify(product_data)


@app.route("/score", methods=["POST"])
def score():
    product_data = request.get_json(silent=True) or {}
    if not product_data.get("product_name") and not product_data.get("ingredients"):
        return jsonify({"error": "No product data provided"}), 400
    try:
        criteria = load_engine_criteria()
        product_type = detect_product_type(product_data)
        evaluations_data, relevant_criteria = evaluate_criteria_with_claude(product_data, criteria, product_type)
        score_pct, critical_fail, non_verifiable_count = calculate_score(evaluations_data, relevant_criteria)
        score_100, kwalificatie, verdict = determine_verdict(score_pct, critical_fail)
        consumer_output = generate_consumer_output(product_data, evaluations_data, relevant_criteria, score_100, kwalificatie, verdict, product_type, critical_fail)
        return jsonify({
            "product_name": product_data.get("product_name", "Onbekend"),
            "brand": product_data.get("brand_name", "Onbekend"),
            "score": score_100,
            "kwalificatie": kwalificatie,
            "verdict": verdict,
            "product_type": product_type,
            "critical_gate": critical_fail,
            "non_verifiable_count": non_verifiable_count,
            "criteria_evaluated": len(relevant_criteria),
            **consumer_output
        })
    except Exception as e:
        return jsonify({"error": f"Engine error: {str(e)}"}), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

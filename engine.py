"""
engine.py — Funcify backend
POST /scrape — haal productdata op van een URL
POST /score  — beoordeel productdata met de scoring engine
GET  /health — readiness check
"""

import os
import re
import json
import requests
import anthropic
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify
from flask_cors import CORS

try:
    import docx as python_docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)
CORS(app, origins=["https://funcify.eu", "https://fullfunc.github.io"])

# ─── Scraper constants ───────────────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Referer": "https://www.google.com/",
}

INGREDIENT_PATTERN = re.compile(
    r"(\d[\d.,]*\s*(?:mg|g|mcg|µg|iu|IU|ml|%|billion\s*CFU|CFU|billion))",
    re.IGNORECASE,
)

# ─── Lazy file loader (loads on first /score request, cached after) ──────────

_cache: dict = {}


def _get_files() -> dict:
    if _cache:
        return _cache

    base = os.path.dirname(os.path.abspath(__file__))

    def read_docx(path: str) -> str:
        if not HAS_DOCX or not os.path.exists(path):
            return ""
        try:
            doc = python_docx.Document(path)
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            app.logger.warning("Could not read %s: %s", path, e)
            return ""

    def read_xlsx(path: str) -> str:
        if not HAS_OPENPYXL or not os.path.exists(path):
            return ""
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            parts = []
            for name in wb.sheetnames:
                ws = wb[name]
                rows = []
                for row in ws.iter_rows(values_only=True):
                    if any(c is not None for c in row):
                        rows.append("\t".join("" if c is None else str(c) for c in row))
                if rows:
                    parts.append(f"=== {name} ===\n" + "\n".join(rows))
            return "\n\n".join(parts)
        except Exception as e:
            app.logger.warning("Could not read %s: %s", path, e)
            return ""

    _cache["masterprompt"]  = read_docx(os.path.join(base, "Funcify masterprompt V10.docx"))
    _cache["ingredient_db"] = read_xlsx(os.path.join(base, "Funcify. Master Ingredient 2.0.xlsx"))
    _cache["engine_review"] = read_xlsx(os.path.join(base, "Funcify. Engine Review V4.xlsx"))
    _cache["consumer_ui"]   = read_xlsx(os.path.join(base, "Funcify. Consumer UI.xlsx"))

    return _cache


# ─── Scraper helpers ─────────────────────────────────────────────────────────

def fetch_html(url: str) -> tuple[str, BeautifulSoup]:
    resp = requests.get(url, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    return resp.text, soup


def extract_text_blocks(soup: BeautifulSoup, *selectors) -> list[str]:
    blocks = []
    for sel in selectors:
        for el in soup.select(sel):
            text = el.get_text(" ", strip=True)
            if text:
                blocks.append(text)
    return blocks


def parse_ingredients_from_text(text: str) -> list[dict]:
    ingredients = []
    lines = [l.strip() for l in re.split(r"[\n;,]", text) if l.strip()]
    for line in lines:
        if len(line) < 3:
            continue
        dosage_match = INGREDIENT_PATTERN.search(line)
        dosage = dosage_match.group(0).strip() if dosage_match else None

        form_match = re.search(r"\bas\s+([A-Za-z][\w\s\-]{2,40})", line)
        form = form_match.group(1).strip() if form_match else None

        name = line
        if dosage:
            name = line[: dosage_match.start()].strip(" -(")
        if form_match:
            name = re.sub(r"\s*\([^)]*\)", "", name).strip()

        if name:
            ingredients.append({"name": name, "dosage": dosage, "molecular_form": form})
    return ingredients


def scrape_product(soup: BeautifulSoup) -> dict:
    result = {
        "product_name": None,
        "brand_name": None,
        "serving_size": None,
        "ingredients": [],
        "health_claims": [],
        "certifications": [],
    }

    for sel in ["h1", '[class*="product-title"]', '[class*="product-name"]', '[itemprop="name"]']:
        el = soup.select_one(sel)
        if el:
            result["product_name"] = el.get_text(strip=True)
            break

    for sel in ['[itemprop="brand"]', '[class*="brand"]', '[class*="vendor"]', "a.brand"]:
        el = soup.select_one(sel)
        if el:
            result["brand_name"] = el.get_text(strip=True)
            break

    full_text = soup.get_text(" ")
    serving_match = re.search(
        r"serving\s+size[:\s]+([^\n.]{3,60})", full_text, re.IGNORECASE
    )
    if serving_match:
        result["serving_size"] = serving_match.group(1).strip()

    ingredient_blocks = extract_text_blocks(
        soup,
        '[class*="ingredient"]',
        '[id*="ingredient"]',
        '[class*="supplement-facts"]',
        '[id*="supplement"]',
        "table",
    )
    raw_ingredient_text = " ".join(ingredient_blocks)

    label_match = re.search(
        r"(?:ingredients?|supplement facts)[:\s]+([\s\S]{10,2000}?)(?:\n\n|\.\s+[A-Z]|$)",
        full_text,
        re.IGNORECASE,
    )
    if label_match:
        raw_ingredient_text += " " + label_match.group(1)

    if raw_ingredient_text.strip():
        result["ingredients"] = parse_ingredients_from_text(raw_ingredient_text)

    for match in re.finditer(
        r"(?:supports?|promotes?|helps?|boosts?|improves?|enhances?)[^.!?\n]{5,120}",
        full_text, re.IGNORECASE
    ):
        claim = match.group(0).strip()
        if claim not in result["health_claims"]:
            result["health_claims"].append(claim)
    result["health_claims"] = result["health_claims"][:10]

    cert_keywords = [
        "certified", "non-gmo", "gmp", "usda organic", "nsf", "informed sport",
        "third.party tested", "gluten.free", "vegan", "kosher", "halal", "organic",
    ]
    for keyword in cert_keywords:
        if re.search(keyword, full_text, re.IGNORECASE):
            label = keyword.replace(".", "-").title()
            if label not in result["certifications"]:
                result["certifications"].append(label)

    return result


def is_sufficient(data: dict) -> bool:
    return bool(data.get("product_name")) and len(data.get("ingredients", [])) > 0


def anthropic_fallback(raw_text: str) -> dict:
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    prompt = f"""You are a supplement product data extractor.

Extract the following fields from the product page text below and return ONLY valid JSON (no markdown, no explanation):

{{
  "product_name": "string or null",
  "brand_name": "string or null",
  "serving_size": "string or null",
  "ingredients": [
    {{"name": "string", "dosage": "string or null", "molecular_form": "string or null"}}
  ],
  "health_claims": ["string"],
  "certifications": ["string"]
}}

Rules:
- molecular_form: the specific chemical form, e.g. "magnesium glycinate", "methylcobalamin"
- dosage: the amount including unit, e.g. "500 mg", "10 mcg"
- health_claims: up to 10 short benefit statements from the page
- certifications: any quality or regulatory badges mentioned

Page text:
{raw_text[:12000]}
"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = message.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


# ─── Engine helpers ──────────────────────────────────────────────────────────

def _ingredient_keywords(product_ingredients: list) -> set:
    words = set()
    for ing in product_ingredients:
        name = (ing.get("name") or "").lower()
        form = (ing.get("molecular_form") or "").lower()
        for part in re.split(r"[\s,\-]+", name + " " + form):
            if len(part) > 3:
                words.add(part)
    return words


def filter_ingredient_db(db_text: str, product_ingredients: list) -> str:
    if not db_text or not product_ingredients:
        return db_text[:10000]

    keywords = _ingredient_keywords(product_ingredients)
    if not keywords:
        return db_text[:10000]

    lines = db_text.splitlines()
    kept = []
    for line in lines:
        if line.startswith("===") or not line.strip():
            kept.append(line)
            continue
        if any(kw in line.lower() for kw in keywords):
            kept.append(line)

    result = "\n".join(kept)
    return result if len(result) > 200 else db_text[:10000]


_OUTPUT_SCHEMA = """\
Geef UITSLUITEND de volgende JSON terug — geen markdown-omhulsel, geen uitleg:

{
  "product_name": "string",
  "brand": "string",
  "score": <integer 0-100>,
  "kwalificatie": "Elite" | "Degelijk" | "Matig" | "Afkeur",
  "verdict": "Koopwaardig" | "Alleen met context" | "Af te raden",
  "wat_doet": "string — wat doet dit supplement in gewone consumententaal",
  "beoordeling_tabel": [
    {
      "criterium": "string",
      "bevinding": "string",
      "oordeel": "Goed" | "Matig" | "Slecht" | "Onbekend"
    }
  ],
  "highlights": {
    "positief": ["string"],
    "negatief": ["string"]
  },
  "context_flags": ["string — specifieke waarschuwingen per doelgroep"],
  "wat_zou_beter": "string — concrete verbeterpunten voor een beter product",
  "voor_wie": "string — voor wie is dit product (nog) bruikbaar ondanks tekortkomingen",
  "consumer_summary": "string — volledige aankoopadviessamenvatting in B1-taalniveau"
}

Scoringsgrenzen:
  Elite    85-100  → Koopwaardig
  Degelijk 65-84   → Koopwaardig
  Matig    45-64   → Alleen met context
  Afkeur   0-44    → Af te raden

Beoordeling_tabel moet minimaal de volgende criteria bevatten (indien van toepassing):
  Bioavailabiliteit, Doseringen, Moleculaire vormen, Certificeringen,
  Transparantie label, Gezondheidsclaims, Serving size, Vulstoffen/additieven.\
"""


def build_messages(product_data: dict) -> tuple[str, str]:
    files = _get_files()

    product_ingredients = product_data.get("ingredients", [])
    relevant_db = filter_ingredient_db(files["ingredient_db"], product_ingredients)

    system_parts = []
    if files["masterprompt"]:
        system_parts.append(files["masterprompt"][:20000])
    system_parts.append(
        "Je bent de Funcify scoring engine. "
        "Je analyseert supplementproducten wetenschappelijk en geeft eerlijk advies aan consumenten. "
        "Je output is altijd geldige JSON."
    )
    system_prompt = "\n\n".join(system_parts)

    sections = []

    if relevant_db.strip():
        sections.append(
            "## INGREDIENT DATABASE (gefilterd op productingredient)\n"
            + relevant_db[:12000]
        )
    if files["engine_review"].strip():
        sections.append(
            "## SCORINGSCRITERIA EN GEWICHTEN (Engine Review V4)\n"
            + files["engine_review"][:6000]
        )
    if files["consumer_ui"].strip():
        sections.append(
            "## KLACHTDEFINITIES EN CONSUMENTENCATEGORIEEN (Consumer UI)\n"
            + files["consumer_ui"][:4000]
        )

    sections.append(
        "## TE BEOORDELEN PRODUCT\n"
        + json.dumps(product_data, ensure_ascii=False, indent=2)
    )
    sections.append("## GEVRAAGDE OUTPUT\n" + _OUTPUT_SCHEMA)

    return system_prompt, "\n\n".join(sections)


def score_product(product_data: dict) -> dict:
    system_prompt, user_message = build_messages(product_data)

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.MULTILINE)
    raw = re.sub(r"\s*```\s*$", "", raw, flags=re.MULTILINE)

    result = json.loads(raw)

    score = int(result.get("score", 0))
    if score >= 85:
        result["kwalificatie"] = "Elite"
        result["verdict"] = "Koopwaardig"
    elif score >= 65:
        result["kwalificatie"] = "Degelijk"
        result["verdict"] = "Koopwaardig"
    elif score >= 45:
        result["kwalificatie"] = "Matig"
        result["verdict"] = "Alleen met context"
    else:
        result["kwalificatie"] = "Afkeur"
        result["verdict"] = "Af te raden"

    return result


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/scrape", methods=["POST"])
def scrape():
    body = request.get_json(force=True, silent=True)
    if not body or not body.get("url"):
        return jsonify({"error": "Missing 'url' in request body"}), 400

    url = body["url"]

    soup = None
    try:
        _, soup = fetch_html(url)
    except (requests.RequestException, requests.exceptions.Timeout) as e:
        app.logger.warning("Fetch failed for %s: %s — falling back to Anthropic", url, e)

    if soup is not None:
        data = scrape_product(soup)
        page_text = soup.get_text(" ", strip=True)
    else:
        data = {}
        page_text = ""

    if not is_sufficient(data):
        try:
            data = anthropic_fallback(page_text)
            data["_source"] = "anthropic_fallback"
        except Exception as e:
            data["_fallback_error"] = str(e)
            data["_source"] = "scraper_partial"
    else:
        data["_source"] = "scraper"

    return jsonify(data)


@app.route("/score", methods=["POST"])
def score():
    data = request.get_json(silent=True) or {}
    url = data.get("url", "") or data.get("product_url", "") or ""

    try:
        result = score_product(data)
        return jsonify(result)
    except json.JSONDecodeError as e:
        return jsonify({"error": f"AI retourneerde ongeldige JSON: {e}"}), 500
    except anthropic.APIError as e:
        return jsonify({"error": f"Anthropic API fout: {e}"}), 502
    except Exception as e:
        app.logger.exception("Onverwachte fout in /score")
        return jsonify({"error": str(e)}), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

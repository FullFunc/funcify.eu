import os
import re
import json
import requests
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic

app = Flask(__name__)
CORS(app, origins=["https://funcify.eu", "https://fullfunc.github.io"])

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

INGREDIENT_PATTERN = re.compile(
    r"(\d[\d.,]*\s*(?:mg|g|mcg|µg|iu|IU|ml|%|billion\s*CFU|CFU|billion))",
    re.IGNORECASE,
)


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

        # Detect molecular form hints (e.g. "as magnesium glycinate")
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

    # --- Product name ---
    for sel in ["h1", '[class*="product-title"]', '[class*="product-name"]', '[itemprop="name"]']:
        el = soup.select_one(sel)
        if el:
            result["product_name"] = el.get_text(strip=True)
            break

    # --- Brand name ---
    for sel in ['[itemprop="brand"]', '[class*="brand"]', '[class*="vendor"]', "a.brand"]:
        el = soup.select_one(sel)
        if el:
            result["brand_name"] = el.get_text(strip=True)
            break

    # --- Serving size ---
    full_text = soup.get_text(" ")
    serving_match = re.search(
        r"serving\s+size[:\s]+([^\n.]{3,60})", full_text, re.IGNORECASE
    )
    if serving_match:
        result["serving_size"] = serving_match.group(1).strip()

    # --- Ingredients ---
    ingredient_blocks = extract_text_blocks(
        soup,
        '[class*="ingredient"]',
        '[id*="ingredient"]',
        '[class*="supplement-facts"]',
        '[id*="supplement"]',
        "table",
    )
    raw_ingredient_text = " ".join(ingredient_blocks)

    # Also try scanning for a "Ingredients:" label in running text
    label_match = re.search(
        r"(?:ingredients?|supplement facts)[:\s]+([\s\S]{10,2000}?)(?:\n\n|\.\s+[A-Z]|$)",
        full_text,
        re.IGNORECASE,
    )
    if label_match:
        raw_ingredient_text += " " + label_match.group(1)

    if raw_ingredient_text.strip():
        result["ingredients"] = parse_ingredients_from_text(raw_ingredient_text)

    # --- Health claims ---
    claim_patterns = [
        r"(?:supports?|promotes?|helps?|boosts?|improves?|enhances?)[^.!?\n]{5,120}",
    ]
    for pattern in claim_patterns:
        for match in re.finditer(pattern, full_text, re.IGNORECASE):
            claim = match.group(0).strip()
            if claim not in result["health_claims"]:
                result["health_claims"].append(claim)
    result["health_claims"] = result["health_claims"][:10]

    # --- Certifications ---
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
    has_name = bool(data.get("product_name"))
    has_ingredients = len(data.get("ingredients", [])) > 0
    return has_name and has_ingredients


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
    # Strip markdown code fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


@app.route("/scrape", methods=["POST"])
def scrape():
    body = request.get_json(force=True, silent=True)
    if not body or not body.get("url"):
        return jsonify({"error": "Missing 'url' in request body"}), 400

    url = body["url"]

    try:
        raw_html, soup = fetch_html(url)
    except requests.RequestException as e:
        return jsonify({"error": f"Failed to fetch URL: {e}"}), 502

    data = scrape_product(soup)

    if not is_sufficient(data):
        try:
            page_text = soup.get_text(" ", strip=True)
            data = anthropic_fallback(page_text)
            data["_source"] = "anthropic_fallback"
        except Exception as e:
            data["_fallback_error"] = str(e)
            data["_source"] = "scraper_partial"
    else:
        data["_source"] = "scraper"

    return jsonify(data)


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

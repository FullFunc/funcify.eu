"""
engine.py — Funcify scoring engine
POST /score  — score a scraped product
GET  /health — readiness check
"""

import os
import re
import json
import functools
import anthropic
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MASTERPROMPT_PATH  = os.path.join(BASE_DIR, "Funcify masterprompt V10.docx")
INGREDIENT_DB_PATH = os.path.join(BASE_DIR, "Funcify. Master Ingredient 2.0.xlsx")
ENGINE_REVIEW_PATH = os.path.join(BASE_DIR, "Funcify. Engine Review V4.xlsx")
CONSUMER_UI_PATH   = os.path.join(BASE_DIR, "Funcify. Consumer UI.xlsx")

# ─── File loaders (process-level cache) ─────────────────────────────────────

@functools.lru_cache(maxsize=None)
def _read_docx(path: str) -> str:
    if not HAS_DOCX or not os.path.exists(path):
        return ""
    try:
        doc = python_docx.Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        app.logger.warning("Could not read %s: %s", path, e)
        return ""


@functools.lru_cache(maxsize=None)
def _read_xlsx(path: str) -> str:
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


def load_masterprompt()  -> str: return _read_docx(MASTERPROMPT_PATH)
def load_ingredient_db() -> str: return _read_xlsx(INGREDIENT_DB_PATH)
def load_engine_review() -> str: return _read_xlsx(ENGINE_REVIEW_PATH)
def load_consumer_ui()   -> str: return _read_xlsx(CONSUMER_UI_PATH)


# ─── Ingredient relevance filter ─────────────────────────────────────────────

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
    """Keep only db rows that mention at least one product ingredient keyword."""
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
    # Fall back to full db if filtering removed everything meaningful
    return result if len(result) > 200 else db_text[:10000]


# ─── Prompt builder ──────────────────────────────────────────────────────────

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
    masterprompt = load_masterprompt()
    ingredient_db_raw = load_ingredient_db()
    engine_review = load_engine_review()
    consumer_ui = load_consumer_ui()

    product_ingredients = product_data.get("ingredients", [])
    relevant_db = filter_ingredient_db(ingredient_db_raw, product_ingredients)

    # System prompt: masterprompt + role instruction
    system_parts = []
    if masterprompt:
        system_parts.append(masterprompt[:20000])
    system_parts.append(
        "Je bent de Funcify scoring engine. "
        "Je analyseert supplementproducten wetenschappelijk en geeft eerlijk advies aan consumenten. "
        "Je output is altijd geldige JSON."
    )
    system_prompt = "\n\n".join(system_parts)

    # User message: databases + product + schema
    sections = []

    if relevant_db.strip():
        sections.append(
            "## INGREDIENT DATABASE (gefilterd op productingredient)\n"
            + relevant_db[:12000]
        )

    if engine_review.strip():
        sections.append(
            "## SCORINGSCRITERIA EN GEWICHTEN (Engine Review V4)\n"
            + engine_review[:6000]
        )

    if consumer_ui.strip():
        sections.append(
            "## KLACHTDEFINITIES EN CONSUMENTENCATEGORIEEN (Consumer UI)\n"
            + consumer_ui[:4000]
        )

    sections.append(
        "## TE BEOORDELEN PRODUCT\n"
        + json.dumps(product_data, ensure_ascii=False, indent=2)
    )

    sections.append("## GEVRAAGDE OUTPUT\n" + _OUTPUT_SCHEMA)

    user_message = "\n\n".join(sections)
    return system_prompt, user_message


# ─── Scoring logic ───────────────────────────────────────────────────────────

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
    # Strip markdown code fences if the model adds them despite the instruction
    raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.MULTILINE)
    raw = re.sub(r"\s*```\s*$", "", raw, flags=re.MULTILINE)

    result = json.loads(raw)

    # Ensure score ↔ kwalificatie ↔ verdict are internally consistent
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

@app.route("/score", methods=["POST"])
def score():
    body = request.get_json(force=True, silent=True)
    if not body:
        return jsonify({"error": "Ongeldige of lege JSON body"}), 400
    if not body.get("product_name") and not body.get("ingredients"):
        return jsonify({"error": "Body moet minimaal 'product_name' of 'ingredients' bevatten"}), 400

    try:
        result = score_product(body)
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
    files = {
        "masterprompt":  os.path.exists(MASTERPROMPT_PATH),
        "ingredient_db": os.path.exists(INGREDIENT_DB_PATH),
        "engine_review": os.path.exists(ENGINE_REVIEW_PATH),
        "consumer_ui":   os.path.exists(CONSUMER_UI_PATH),
    }
    docx_ok   = HAS_DOCX
    xlsx_ok   = HAS_OPENPYXL
    api_key   = bool(os.environ.get("ANTHROPIC_API_KEY"))
    all_ready = all(files.values()) and docx_ok and xlsx_ok and api_key

    return jsonify({
        "status":             "ok" if all_ready else "degraded",
        "files":              files,
        "python_docx":        docx_ok,
        "openpyxl":           xlsx_ok,
        "api_key_configured": api_key,
    }), 200 if all_ready else 503


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 8080)))

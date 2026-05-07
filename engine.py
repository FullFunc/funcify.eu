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

# ─── Context Flag Rules ───────────────────────────────────────────────────────
CONTEXT_FLAG_RULES = [
    {"id": "CF001", "keywords": ["pyridoxine", "vitamine b6", "vitamin b6", "b6"], "operator": "gte", "threshold": 25, "unit": "mg", "severity": "Critical", "message": "Dit product bevat een hoge B6-dosering. Langdurige inname boven 25 mg per dag kan bij sommige mensen perifere neuropathie veroorzaken. Gebruik niet structureel zonder begeleiding."},
    {"id": "CF002", "keywords": ["pyridoxine", "vitamine b6", "b6"], "operator": "gte", "threshold": 50, "unit": "mg", "severity": "Critical", "message": "Dit product bevat een zeer hoge B6-dosering. Alleen bij specifieke therapeutische indicatie onder begeleiding."},
    {"id": "CF003", "keywords": ["retinol", "vitamine a", "vitamin a"], "operator": "gte", "threshold": 3000, "unit": "IU", "severity": "Critical", "message": "Dit product bevat een hoge dosis vitamine A. Accumulatie is mogelijk bij dagelijks gebruik. Niet aanbevolen tijdens zwangerschap."},
    {"id": "CF005", "keywords": ["jodium", "iodine", "kaliumjodide", "natriumjodide", "zeewier", "kelp"], "operator": "gte", "threshold": 150, "unit": "mcg", "severity": "Major", "message": "Dit product bevat jodium. Bij bestaande schildklierproblemen kan extra jodium klachten verergeren."},
    {"id": "CF007", "keywords": ["ashwagandha", "withania somnifera"], "operator": "any", "threshold": 0, "unit": "", "severity": "Major", "message": "Ashwagandha kan de schildklierfunctie beinvloeden. Bij schildklierproblematiek of schildkliermedicatie: gebruik alleen onder begeleiding."},
    {"id": "CF008", "keywords": ["ijzer", "ferrum", "iron", "feso4", "ferrousfumaraat", "ijzerbisglycinaat"], "operator": "any", "threshold": 0, "unit": "", "severity": "Major", "message": "IJzersuppletie is alleen zinvol bij aangetoond tekort. Overmatige ijzerinname is geassocieerd met oxidatieve stress. Niet aanbevolen zonder indicatie."},
    {"id": "CF009", "keywords": ["ijzer", "ferrum", "iron"], "operator": "gte", "threshold": 45, "unit": "mg", "severity": "Critical", "message": "Hoge ijzerdosering. Boven 45 mg per dag verhoogd risico op gastro-intestinale bijwerkingen. Alleen onder medische supervisie."},
    {"id": "CF010", "keywords": ["zink", "zinc"], "operator": "gte", "threshold": 25, "unit": "mg", "severity": "Major", "message": "Hoge zink-inname kan koper verdringen. Langdurig gebruik zonder koper-suppletie kan leiden tot koperdepletie."},
    {"id": "CF011", "keywords": ["zink", "zinc"], "operator": "gte", "threshold": 40, "unit": "mg", "severity": "Critical", "message": "Zinkdosering nadert de veiligheidsgrens van 40 mg per dag. Niet combineren met andere zinkbevattende supplementen."},
    {"id": "CF014", "keywords": ["vitamine k", "vitamin k", "k2", "mk-7", "menaquinon", "phylloquinon"], "operator": "any", "threshold": 0, "unit": "", "severity": "Critical", "message": "Vitamine K interfereert direct met bloedverdunners zoals Acenocoumarol en Warfarine. Bij gebruik van bloedverdunners: overleg altijd met je arts."},
    {"id": "CF015", "keywords": ["epa", "dha", "omega-3", "omega 3", "visolie", "fish oil", "levertraan"], "operator": "gte", "threshold": 3000, "unit": "mg", "severity": "Major", "message": "Hoge omega-3 dosering heeft een mild bloedverdunnend effect. Bij gebruik van antistollingsmedicatie of een geplande operatie: overleg met je arts."},
    {"id": "CF017", "keywords": ["sint-janskruid", "hypericum perforatum", "st john"], "operator": "any", "threshold": 0, "unit": "", "severity": "Critical", "message": "Sint-Janskruid verlaagt de werkzaamheid van veel medicijnen waaronder anticonceptie, antidepressiva en bloedverdunners. Niet combineren met medicijnen."},
    {"id": "CF023", "keywords": ["vitamine d3", "cholecalciferol", "vitamin d3"], "operator": "gte", "threshold": 4000, "unit": "IU", "severity": "Major", "message": "Hoge vitamine D3 dosering. Bij zwangerschap of nierfunctiestoornissen: gebruik niet zonder medisch advies."},
    {"id": "CF024", "keywords": ["vitamine c", "ascorbinezuur", "natriumascorbaat", "ascorbate", "vitamin c"], "operator": "gte", "threshold": 1000, "unit": "mg", "severity": "Major", "message": "Hoge vitamine C dosering verhoogt de oxalaatuitscheiding via de nier. Bij nierstenen of verminderde nierfunctie: gebruik voorzichtig."},
    {"id": "CF025", "keywords": ["magnesium"], "operator": "gte", "threshold": 350, "unit": "mg", "severity": "Major", "message": "Hoge magnesium-inname. Bij nierfunctiestoornissen kan magnesium niet goed worden uitgescheiden wat kan leiden tot ophoping."},
    {"id": "CF026", "keywords": ["creatine"], "operator": "any", "threshold": 0, "unit": "", "severity": "Minor", "message": "Creatine verhoogt creatinine-waarden in het bloed wat bloedtesten kan vertekenen. Bij verminderde nierfunctie: gebruik alleen na medisch overleg."},
    {"id": "CF033", "keywords": ["berberine", "berberis"], "operator": "any", "threshold": 0, "unit": "", "severity": "Critical", "message": "Berberine heeft een significant bloedsuikerverlagend effect. Combinatie met diabetesmedicatie kan een gevaarlijk lage bloedsuiker veroorzaken."},
    {"id": "CF036", "keywords": ["foliumzuur", "folic acid", "pteroylglutaminezuur"], "operator": "any", "threshold": 0, "unit": "", "severity": "Major", "message": "Synthetisch foliumzuur kan bij MTHFR-polymorfisme onvoldoende worden omgezet naar de actieve vorm. De actieve vorm methylfolaat is de betere keuze."},
    {"id": "CF042", "keywords": ["coq10", "coenzym q10", "ubiquinol", "ubiquinone"], "operator": "any", "threshold": 0, "unit": "", "severity": "Major", "message": "CoQ10 is extra relevant als je statines gebruikt. Statines verlagen de aanmaak van CoQ10 in je lichaam."},
    {"id": "CF043", "keywords": ["rode gistrijst", "monascus purpureus", "monacoline"], "operator": "any", "threshold": 0, "unit": "", "severity": "Critical", "message": "Rode gistrijst bevat een natuurlijk statine. Combinatie met statinetherapie geeft een verhoogd risico op spierschade en leverbelasting."},
    {"id": "CF045", "keywords": ["melatonine", "melatonin"], "operator": "gte", "threshold": 0.5, "unit": "mg", "severity": "Major", "message": "Melatonine kan interacteren met antidepressiva en slaap- of kalmeringsmiddelen. Bij gebruik van deze medicijnen: raadpleeg je arts."},
    {"id": "CF046", "keywords": ["5-htp", "5-hydroxytryptofaan"], "operator": "any", "threshold": 0, "unit": "", "severity": "Critical", "message": "5-HTP verhoogt de serotoninesynthese. Niet combineren met antidepressiva of MAO-remmers: risico op serotoninesyndroom."},
    {"id": "CF050", "keywords": ["lactobacillus", "bifidobacterium", "saccharomyces", "probiotica", "probiotic"], "operator": "any", "threshold": 0, "unit": "", "severity": "Critical", "message": "Levende bacterien in probiotica kunnen bij ernstig immuungecompromitteerde personen systemische infecties veroorzaken. Raadpleeg je arts bij chemotherapie of transplantatie."},
    {"id": "CF055", "keywords": ["vitamine e", "tocoferol", "vitamin e"], "operator": "gte", "threshold": 400, "unit": "IU", "severity": "Major", "message": "Hoge vitamine E heeft een antistollend effect. Bij gebruik van bloedverdunners is er een verhoogd risico op bloedingen."},
]

# ─── Bioavailability Ratios ───────────────────────────────────────────────────
BIOAVAILABILITY_RATIOS = {
    "magnesium_oxide": {"ratio": 0.2, "label": "slechte opneembaarheid", "consumer": "wordt grotendeels niet opgenomen door het lichaam"},
    "magnesium_carbonate": {"ratio": 0.2, "label": "slechte opneembaarheid", "consumer": "wordt grotendeels niet opgenomen"},
    "magnesium_citraat": {"ratio": 1.0, "label": "referentievorm", "consumer": "goede opneembaarheid"},
    "magnesium_malaat": {"ratio": 1.1, "label": "goede opneembaarheid", "consumer": "goed opneembaar"},
    "magnesium_bisglycinate": {"ratio": 1.4, "label": "uitstekende opneembaarheid", "consumer": "een van de best opneembare vormen"},
    "magnesium_threonate": {"ratio": 1.6, "label": "uitstekende opneembaarheid", "consumer": "hoogste opneembaarheid, ook in de hersenen"},
    "zinc_oxide": {"ratio": 0.2, "label": "slechte opneembaarheid", "consumer": "wordt nauwelijks opgenomen"},
    "zinc_sulfaat": {"ratio": 0.6, "label": "matige opneembaarheid", "consumer": "matige opneembaarheid"},
    "zinc_gluconaat": {"ratio": 0.9, "label": "goede opneembaarheid", "consumer": "goede opneembaarheid"},
    "zinc_picolinaat": {"ratio": 1.0, "label": "referentievorm", "consumer": "goede opneembaarheid"},
    "zinc_bisglycinate": {"ratio": 1.3, "label": "uitstekende opneembaarheid", "consumer": "uitstekend opneembaar"},
    "iron_sulfaat": {"ratio": 0.5, "label": "matige opneembaarheid maar hoge GI-belasting", "consumer": "matig opneembaar en kan maagklachten geven"},
    "iron_fumaraat": {"ratio": 0.9, "label": "goede opneembaarheid", "consumer": "goede opneembaarheid"},
    "iron_bisglycinate": {"ratio": 1.0, "label": "referentievorm", "consumer": "goed opneembaar zonder maagklachten"},
    "heme_iron": {"ratio": 1.8, "label": "beste opneembaarheid", "consumer": "beste opneembaarheid van alle ijzervormen"},
    "omega3_ethylester": {"ratio": 0.5, "label": "lagere opneembaarheid", "consumer": "goedkopere minder goed opneembare vorm"},
    "omega3_triglyceride": {"ratio": 1.0, "label": "referentievorm", "consumer": "goed opneembare natuurlijke vorm"},
    "omega3_rtg": {"ratio": 1.5, "label": "uitstekende opneembaarheid", "consumer": "beste opneembare visolievorm"},
    "omega3_phospholipid": {"ratio": 1.2, "label": "goede opneembaarheid", "consumer": "goed opneembaar vetmolecuul"},
    "vitamin_d2": {"ratio": 0.7, "label": "minder effectief", "consumer": "minder effectief dan vitamine D3"},
    "vitamin_d3": {"ratio": 1.0, "label": "referentievorm", "consumer": "de beste en meest actieve vitamine D-vorm"},
    "cyanocobalamine": {"ratio": 0.7, "label": "vereist omzetting", "consumer": "vereist extra omzettingsstap in het lichaam"},
    "hydroxocobalamine": {"ratio": 1.0, "label": "referentievorm", "consumer": "goede opneembaarheid"},
    "methylcobalamine": {"ratio": 1.1, "label": "direct actief", "consumer": "direct actieve vorm, geen omzetting nodig"},
    "foliumzuur_synthetisch": {"ratio": 0.6, "label": "vereist omzetting", "consumer": "moet worden omgezet naar de actieve vorm, lukt niet bij iedereen"},
    "methylfolaat_5mthf": {"ratio": 1.0, "label": "actieve referentievorm", "consumer": "direct actieve vorm, wordt door iedereen goed opgenomen"},
    "ubiquinone": {"ratio": 1.0, "label": "referentievorm", "consumer": "standaard CoQ10-vorm"},
    "ubiquinol": {"ratio": 1.7, "label": "uitstekende opneembaarheid", "consumer": "de actieve gereduceerde CoQ10-vorm, significant beter opneembaar"},
}

# ─── Ingredient Synonyms ──────────────────────────────────────────────────────
INGREDIENT_SYNONYMS = {
    "omega 3": "omega_3", "omega-3": "omega_3", "visolie": "omega_3", "fish oil": "omega_3",
    "epa": "omega_3", "dha": "omega_3", "levertraan": "omega_3", "krillolie": "omega_3",
    "magnesium bisglycinaat": "magnesium_bisglycinate", "magnesium bisglycinate": "magnesium_bisglycinate",
    "mg bisglycinate": "magnesium_bisglycinate",
    "magnesium oxide": "magnesium_oxide", "magnesiumoxide": "magnesium_oxide",
    "magnesium citraat": "magnesium_citraat", "magnesiumcitraat": "magnesium_citraat",
    "vitamine d3": "vitamin_d3", "cholecalciferol": "vitamin_d3", "d3": "vitamin_d3",
    "vitamine d2": "vitamin_d2", "ergocalciferol": "vitamin_d2",
    "methylcobalamine": "methylcobalamine", "methylcobalamin": "methylcobalamine",
    "cyanocobalamine": "cyanocobalamine", "cyanocobalamin": "cyanocobalamine",
    "vitamine b12": "vitamin_b12", "cobalamine": "vitamin_b12",
    "foliumzuur": "foliumzuur_synthetisch", "folic acid": "foliumzuur_synthetisch",
    "5-mthf": "methylfolaat_5mthf", "methylfolaat": "methylfolaat_5mthf", "l-methylfolaat": "methylfolaat_5mthf",
    "zink": "zinc", "zinc": "zinc", "zinkbisglycinaat": "zinc_bisglycinate",
    "zinc bisglycinate": "zinc_bisglycinate",
    "zinkoxide": "zinc_oxide", "zinc oxide": "zinc_oxide",
    "ijzer": "iron", "ferrum": "iron", "iron": "iron",
    "ijzerbisglycinaat": "iron_bisglycinate", "iron bisglycinate": "iron_bisglycinate",
    "creatine monohydraat": "creatine", "creatine monohydrate": "creatine",
    "ubiquinol": "ubiquinol", "ubiquinone": "ubiquinone", "coq10": "ubiquinone",
    "co-enzyme q10": "ubiquinone",
    "heemijzer": "heme_iron", "heme iron": "heme_iron",
}

EXCIPIENT_KEYWORDS = [
    "capsulehuls", "hydroxypropylmethylcellulose", "hpmc", "rijstvezel", "antiklontermiddel",
    "bevochtigingsmiddel", "glycerol", "gezuiverd water", "siliciumdioxide", "magnesiumstearaat",
    "talk", "titaniumdioxide", "gelatine", "visgelatine", "zetmeel", "maltodextrine", "silica",
    "cellulose", "stearinezuur", "carrageen", "lecithine", "arabische gom", "dicalciumfosfaat",
    "microkristallijne", "magnesium stearate", "stearic acid", "silicon dioxide",
    "microcrystalline cellulose", "mcc", "hydroxypropyl methylcellulose", "gelatin",
    "dicalcium phosphate", "calcium carbonate", "rice flour", "rice bran", "talc",
    "titanium dioxide", "croscarmellose", "povidone", "polyethylene glycol", "polysorbate",
    "sodium lauryl sulfate", "maltodextrin",
]

CERT_KEYWORDS_FIXED = [
    "IFOS", "ISO 17025", "ISO 22000", "GMP", "HACCP", "USP", "EFSA", "NSF",
    "Informed Sport", "NZVT", "Creapure", "MSC", "Rainforest Alliance", "Fair Trade",
    "Friend of the Sea", "Dolphin Safe", "Green-e", "TRAACS", "Albion", "Golden Omega",
    "WADA", "Keurmerk", "Kosher", "Halal", "Vegan Society", "Soil Association",
    "ISO 9001", "FSSC 22000", "BRC", "IFS", "organic", "biologisch",
]

AMOUNT_PAT = re.compile(r"(\d[\d.,]*)\s*(mg|g|mcg|µg|ug|ml|ie|iu|kve|cfu|%)", re.I)
BROAD_INGREDIENT_PAT = re.compile(
    r"([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-\(\)\/]+?)\s*[:\(]?\s*(\d[\d\.,]*)\s*(mg|mcg|g|IU|ie|kve|cfu|ml)",
    re.I
)

# ─── Jargon Simplification ────────────────────────────────────────────────────
JARGON_REPLACEMENTS = [
    (r"\boxidatiestatus\b", "versheid en stabiliteit van de olie"),
    (r"\bethylester\b", "goedkopere minder goed opneembare vorm"),
    (r"\btriglyceridevorm\b", "goed opneembare natuurlijke vorm"),
    (r"\bbiobeschikbaarheid\b", "opneembaarheid"),
    (r"\btherapeutische dosering\b", "effectieve dagdosering"),
    (r"\btherapeutisch\b", "werkzaam"),
    (r"\bklinisch relevante\b", "werkzame"),
    (r"\bklinisch bewezen\b", ""),
    (r"\bstudies tonen aan\b", ""),
    (r"\bliteratuur wijst uit\b", ""),
    (r"\bfarmaceutisch\b", "van hoge kwaliteit"),
    (r"\bchelaatvorm\b", "goed opneembare gebonden vorm"),
    (r"\bfosfolipide\b", "opneembaar vetmolecuul"),
    (r"\bCOA\b", "onafhankelijk laboratoriumrapport"),
    (r"\bTOTOX\b", "versheidswaarde van de olie"),
    (r"\bproprietary blend\b", "mengsel met verborgen doseringen"),
    (r"\bgesupplementeerd\b", "aangevuld"),
    (r"\bbiologisch actief\b", "direct actief"),
    (r"\bbioactieve\b", "actieve"),
    (r"\bendogene\b", "eigen aanmaak van"),
    (r"\bexogeen\b", "via suppletie"),
    (r"\bfarmacokinetisch\b", "opname-"),
    (r"\bsubklinisch\b", "nog niet meetbaar"),
    (r"\binflammatie\b", "ontsteking"),
    (r"\bsystemisch\b", "via het bloed"),
]


def simplify_jargon(obj):
    if isinstance(obj, str):
        for pattern, replacement in JARGON_REPLACEMENTS:
            obj = re.sub(pattern, replacement, obj, flags=re.I)
        return obj
    elif isinstance(obj, list):
        return [simplify_jargon(item) for item in obj]
    elif isinstance(obj, dict):
        return {k: simplify_jargon(v) for k, v in obj.items()}
    return obj


# ─── Bioavailability helpers ──────────────────────────────────────────────────
def _normalize_name(name):
    key = name.lower().strip()
    return INGREDIENT_SYNONYMS.get(key, key)


def get_bioavailability_info(ingredient_name, form=""):
    key = _normalize_name(ingredient_name)
    if key in BIOAVAILABILITY_RATIOS:
        return BIOAVAILABILITY_RATIOS[key]
    if form:
        form_key = _normalize_name(form)
        if form_key in BIOAVAILABILITY_RATIOS:
            return BIOAVAILABILITY_RATIOS[form_key]
    return None


def get_better_alternatives(ingredient_name, current_ratio):
    if current_ratio >= 0.8:
        return []
    alternatives = []
    base = _normalize_name(ingredient_name).split("_")[0]
    for key, info in BIOAVAILABILITY_RATIOS.items():
        if key.startswith(base) and info["ratio"] > current_ratio:
            alternatives.append({"form": key, **info})
    return sorted(alternatives, key=lambda x: x["ratio"], reverse=True)[:3]


# ─── Context Flag Evaluation ──────────────────────────────────────────────────
def _unit_to_base(value, unit):
    unit = unit.lower()
    if unit in ("mcg", "µg", "ug"):
        return value, "mcg"
    if unit in ("iu", "ie"):
        return value, "IU"
    return value, unit


def evaluate_context_flags(ingredients):
    triggered = []
    triggered_ids = set()
    for rule in CONTEXT_FLAG_RULES:
        rule_id = rule["id"]
        if rule_id in triggered_ids:
            continue
        for ing in ingredients:
            name_lower = ing.get("name", "").lower()
            matched_keyword = any(kw in name_lower for kw in rule["keywords"])
            if not matched_keyword:
                continue
            # CF024 high-dose vitamin C: only trigger for active ingredients, not excipients/antioxidants
            if rule_id == "CF024" and ing.get("type") == "vul-additief":
                continue
            if rule["operator"] == "any":
                triggered.append(rule["message"])
                triggered_ids.add(rule_id)
                break
            elif rule["operator"] == "gte":
                amount = ing.get("amount")
                unit = ing.get("unit", "")
                if amount is None or not unit:
                    continue
                val, norm_unit = _unit_to_base(float(amount), unit)
                _, norm_thresh_unit = _unit_to_base(rule["threshold"], rule["unit"])
                if norm_unit == norm_thresh_unit and val >= rule["threshold"]:
                    triggered.append(rule["message"])
                    triggered_ids.add(rule_id)
                    break
    return triggered


# ─── Cofactor / Antagonist Checks ────────────────────────────────────────────
def evaluate_cofactor_checks(ingredients):
    warnings = []
    names_lower = [i.get("name", "").lower() for i in ingredients]

    def has_ing(keywords):
        return any(any(kw in n for kw in keywords) for n in names_lower)

    def get_amount_iu(keywords):
        for ing in ingredients:
            n = ing.get("name", "").lower()
            if any(kw in n for kw in keywords):
                amt = ing.get("amount")
                u = (ing.get("unit") or "").lower()
                if amt is not None and u in ("iu", "ie"):
                    return float(amt)
        return 0.0

    def get_amount_mg(keywords):
        for ing in ingredients:
            n = ing.get("name", "").lower()
            if any(kw in n for kw in keywords):
                amt = ing.get("amount")
                u = (ing.get("unit") or "").lower()
                if amt is not None and u == "mg":
                    return float(amt)
        return 0.0

    d3_iu = get_amount_iu(["vitamine d3", "cholecalciferol", "vitamin d3", "d3"])
    k2_present = has_ing(["vitamine k", "vitamin k", "k2", "mk-7", "menaquinon"])
    if d3_iu >= 2000 and not k2_present:
        warnings.append("Hoge vitamine D zonder K2 kan calciumdepositie in bloedvaten bevorderen. Overweeg K2 toe te voegen.")

    zinc_mg = get_amount_mg(["zink", "zinc"])
    copper_present = has_ing(["koper", "copper", "cuprum"])
    if zinc_mg >= 25 and not copper_present:
        warnings.append("Hoog zink zonder koper kan bij langdurig gebruik koperdepletie veroorzaken.")

    iron_present = has_ing(["ijzer", "ferrum", "iron"])
    vitc_present = has_ing(["vitamine c", "ascorbinezuur", "natriumascorbaat", "ascorbate", "vitamin c"])
    if iron_present and not vitc_present:
        warnings.append("Vitamine C verbetert ijzeropname significant. Neem bij voorkeur samen in.")

    omega3_present = has_ing(["epa", "dha", "omega-3", "omega 3", "visolie", "fish oil"])
    vite_present = has_ing(["vitamine e", "tocoferol", "tocopherol", "vitamin e"])
    if omega3_present and not vite_present:
        warnings.append("Omega-3 vetzuren zijn gevoelig voor oxidatie. Controleer of het product antioxidantbescherming bevat.")

    return warnings


# ─── Intake Advice ────────────────────────────────────────────────────────────
def get_intake_advice(ingredients):
    names_lower = [i.get("name", "").lower() for i in ingredients]

    def has_ing(keywords):
        return any(any(kw in n for kw in keywords) for n in names_lower)

    advices = []
    if has_ing(["epa", "dha", "omega-3", "omega 3", "visolie", "fish oil", "levertraan"]):
        advices.append("Neem in bij een maaltijd met vet voor optimale opname.")
    if has_ing(["magnesium"]):
        advices.append("Kan nuchter worden ingenomen maar bij gevoelige maag bij een maaltijd innemen.")
    if has_ing(["ijzer", "ferrum", "iron"]):
        advices.append("Neem in op nuchtere maag of met vitamine C. Niet combineren met koffie, thee of zuivel.")
    if has_ing(["vitamine d", "vitamin d", "cholecalciferol", "d3"]):
        advices.append("Neem in bij een vetbevattende maaltijd voor optimale opname.")
    if has_ing(["lactobacillus", "bifidobacterium", "saccharomyces", "probiotica", "probiotic"]):
        advices.append("Neem bij voorkeur op nuchtere maag of 30 minuten voor de maaltijd.")
    if has_ing(["creatine"]):
        advices.append("Tijdstip van inname maakt weinig uit. Consistent dagelijks gebruik is belangrijker.")
    return " ".join(advices)


# ─── Price Calculations ───────────────────────────────────────────────────────
def _parse_price(price_str):
    if not price_str:
        return 0.0
    cleaned = re.sub(r"[^\d,\.]", "", price_str).replace(",", ".")
    m = re.search(r"\d+\.\d+|\d+", cleaned)
    return float(m.group(0)) if m else 0.0


def _parse_units(package_size_str):
    m = re.search(r"(\d+)", package_size_str or "")
    return int(m.group(1)) if m else 0


def _parse_servings_per_day(serving_size_str, usage_instructions_str=""):
    s = serving_size_str or ""
    u = usage_instructions_str or ""
    if re.search(r"1[-–]2|one to two|één tot twee", s, re.I):
        return 2
    m = re.search(r"(\d+)", s)
    spd = max(1, int(m.group(1))) if m else 1
    if spd == 1 and re.search(r"\b2\b.{0,20}(per dag|dagelijks)", u, re.I):
        return 2
    return spd


def calculate_price_per_day(price_str, package_size_str, serving_size_str, usage_instructions_str=""):
    price = _parse_price(price_str)
    units = _parse_units(package_size_str)
    spd = _parse_servings_per_day(serving_size_str, usage_instructions_str)
    if not price or not units:
        return ""
    days = units / spd
    if days <= 0:
        return ""
    ppd = price / days
    return "€" + f"{ppd:.2f}".replace(".", ",") + " per dag"


def calculate_price_per_gram(price_str, ingredients, package_size_str, serving_size_str):
    price = _parse_price(price_str)
    units = _parse_units(package_size_str)
    spd = _parse_servings_per_day(serving_size_str)
    if not price or not units:
        return ""
    total_mg_per_serving = 0.0
    for ing in ingredients:
        if ing.get("type") == "vul-additief":
            continue
        amt = ing.get("amount")
        unit = (ing.get("unit") or "").lower()
        if amt is None:
            continue
        if unit == "mg":
            total_mg_per_serving += float(amt)
        elif unit == "g":
            total_mg_per_serving += float(amt) * 1000
    if total_mg_per_serving <= 0:
        return ""
    total_g = (total_mg_per_serving * units) / spd / 1000
    if total_g <= 0:
        return ""
    ppg = price / total_g
    return "€" + f"{ppg:.2f}".replace(".", ",") + " per gram"


# ─── Excel criteria loading ───────────────────────────────────────────────────
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
                            "source": "core", "category": category, "criterion": criterion,
                            "weight": weight, "critical": critical == "Y", "applies_to": applies_to,
                            "consumer_impact": consumer_impact, "failure_type": failure_type,
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
                            "source": "module", "category": f"Module_{module_type}",
                            "criterion": criterion, "weight": weight, "critical": critical == "Y",
                            "applies_to": module_type, "consumer_impact": consumer_impact,
                            "failure_type": "Quality fail",
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


def evaluate_criteria_with_claude(product_data, criteria, product_type, url=""):
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
- Behandel ontbrekende data NOOIT als bewijs van slechte kwaliteit.
- VERBODEN: "studies tonen", "literatuur zegt", "klinisch bewezen" - alleen label/website feiten
- Geef data_quality terug als EXACT (letterlijk op label met exacte waarde), AFGELEID (logisch af te leiden), ONVOLLEDIG (deels beschikbaar), NIET_GEVONDEN (niet aangetroffen op pagina).

SEVERITY TUNING voor omega-3 producten:
- Ontbrekende IFOS certificering = DATA_LACUNE, niet negatief
- Ontbrekende onafhankelijk laboratoriumrapport = DATA_LACUNE middel
- Ontbrekende EPA/DHA specificatie bij omega-3 = critical fail (dit is kerndata)
- Goedkopere minder goed opneembare vorm = negatief middel
- Mengsel met verborgen doseringen = negatief zwaar
- Ontbrekende batchdata = DATA_LACUNE klein

- Geef alleen valide JSON terug."""
    user_prompt = f"""Product URL: {url}
Product: {product_data.get('product_name', 'Onbekend')}
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
{criteria_text if criteria_text else 'Geen criteria beschikbaar (geen Excel geladen)'}

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
    """
    Score calculation with data_quality awareness.
    - NIET_GEVONDEN → always neutral (pass_value forced to -1)
    - ONVOLLEDIG + fail → half-weight penalty
    - Critical gate only triggers on EXACT or AFGELEID fail evidence
    """
    evaluations = evaluations_data.get("evaluations", [])
    total_raw_score = 0.0
    total_max_weight = 0.0
    critical_fail = False
    non_verifiable_count = 0

    for i, criterion in enumerate(criteria):
        eval_item = next((e for e in evaluations if e.get("criterion_index") == i + 1), None)
        pass_value = eval_item.get("pass_value", -1) if eval_item else -1
        data_quality = (eval_item.get("data_quality", "") if eval_item else "").upper()
        weight = criterion["weight"]

        # NIET_GEVONDEN is always neutral
        if data_quality == "NIET_GEVONDEN":
            pass_value = -1

        if pass_value == -1:
            non_verifiable_count += 1
            total_max_weight += weight
        elif pass_value == 1:
            total_raw_score += weight
            total_max_weight += weight
        elif pass_value == 0:
            # ONVOLLEDIG fail: halve the penalty weight
            effective_weight = weight * 0.5 if data_quality == "ONVOLLEDIG" else weight
            total_max_weight += effective_weight
            # Critical gate only on proven evidence (EXACT or AFGELEID)
            if criterion["critical"] and data_quality in ("EXACT", "AFGELEID", ""):
                critical_fail = True

    if total_max_weight == 0:
        return 0.0, False, 0

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


def generate_consumer_output(product_data, evaluations_data, criteria, score_100, kwalificatie, verdict, product_type, critical_fail, context_flags_triggered=None):
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
    excipients_text = "\n".join([
        f"- {e.get('name', '')}"
        for e in product_data.get("excipients", [])
    ]) or "Geen vulstoffen geidentificeerd."
    usage_instructions_text = product_data.get("usage_instructions", "") or ""
    flags_text = "\n".join(context_flags_triggered or []) or "Geen"
    certifications_text = ", ".join(product_data.get("certifications", [])) or "Geen certificeringen gevonden op de pagina."
    if score_100 > 75:
        _wat_zou_beter_instruction = (
            "ADAPTIEF wat_zou_beter (score >75): beschrijf voor wie dit product het meest geschikt is "
            "en welk type product nóg beter bij hun doelen zou aansluiten. Tweede persoon, maximaal 3 zinnen."
        )
    elif score_100 >= 50:
        _wat_zou_beter_instruction = (
            "ADAPTIEF wat_zou_beter (score 50-75): noem maximaal 2 concrete verbeterpunten die een beter product zou hebben, "
            "en beschrijf kort voor wie dit product toch nog geschikt is. Tweede persoon, maximaal 3 zinnen."
        )
    else:
        _wat_zou_beter_instruction = (
            "ADAPTIEF wat_zou_beter (score <50): noem alleen de concrete verbeterpunten die een kwalitatief beter alternatief "
            "ten opzichte van dit product zou hebben. Geen doelgroep beschrijving. Tweede persoon, maximaal 3 zinnen."
        )
    system_prompt = """KRITIEKE REGEL: Schrijf NOOIT inname-advies dat niet letterlijk uit de usage_instructions of additional_info van dit specifieke product komt. Als usage_instructions leeg is schrijf dan alleen: Zie de verpakking voor innameadvies. Verzin NOOIT timing of combinaties met voedsel op basis van algemene kennis.
TWEEDE KRITIEKE REGEL: Schrijf NOOIT percentages, ratio's of getallen die niet letterlijk op de productpagina of in de verstrekte productdata staan. Geen '50-70% betere opname' of andere berekende claims. Schrijf alleen wat aantoonbaar beschikbaar is. Bij bioavailabiliteit: beschrijf alleen de vorm en waarom die beter is dan alternatieven, zonder getallen te noemen die niet op het label staan.

Schrijf ALTIJD in tweede persoon (je, jij). Maximaal 2-3 zinnen per sectie. Geen Engelse termen tenzij geen Nederlandse equivalent bestaat. Geen wetenschappelijk jargon. Schrijf alsof je het uitlegt aan iemand zonder medische kennis maar die wel serieus is over gezondheid. Begin elke sectie direct met de inhoud, geen inleidende zinnen als 'Dit product' of 'Het supplement'.

Je bent de Funcify consumer output generator. Schrijf heldere eerlijke beoordelingen in Nederlands.
TOON: Direct, eerlijk, consumentvriendelijk.
VERBODEN: studies tonen, literatuur zegt, klinisch bewezen, therapeutisch.
DATA_LACUNE = neutraal, nooit negatief.

De beoordeling_tabel bevat ALTIJD exact deze 8 aspecten in deze volgorde: Moleculaire vormen, Doseringen, Bioavailabiliteit, Transparantie label, Certificeringen, Gezondheidsclaims, Serving size, Vulstoffen en additieven.
Voor de Certificeringen rij: benoem ALLE certificeringen uit de CERTIFICERINGEN GEVONDEN OP PAGINA lijst, niet alleen de bekendste. Als meerdere certificeringen aanwezig zijn, noem ze allemaal.
wat_doet is altijd 2-3 zinnen. consumer_summary is altijd 2-3 zinnen.
highlights bevat altijd minimaal 2 positieve en 2 negatieve punten.
context_flags bevat alleen waarschuwingen die direct getriggerd worden door specifieke ingredienten en doseringen uit dit product, nooit generiek advies.
Als een aspect geen data heeft schrijf dan als bevinding: Geen informatie beschikbaar op de productpagina. — nooit een lege bevinding.

Geef alleen valide JSON terug.""" + "\n\n" + _wat_zou_beter_instruction
    user_prompt = f"""Product: {product_data.get('product_name', 'Onbekend')} ({product_data.get('brand_name', 'Onbekend')})
Score: {score_100}/100 | Kwalificatie: {kwalificatie} | Verdict: {verdict}
Critical gate: {'JA' if critical_fail else 'NEE'}
Prijs per serving: {price_per_serving or 'onbekend'}
Product URL: {product_data.get('url', '')}

INGREDIENTEN:
{ingredients_text or 'Niet gevonden'}

VULSTOFFEN EN HULPSTOFFEN: {excipients_text}
INSTRUCTIE VULSTOFFEN: GEBRUIK ALLEEN DE BOVENSTAANDE VULSTOFFEN. Maak NOOIT aannames over capsulemateriaal of hulpstoffen die niet in de lijst staan. Als visgelatine in de lijst staat schrijf dan visgelatine, niet 'waarschijnlijk gelatine'. Als de lijst leeg is schrijf dan: Geen vulstoffen beschikbaar op de productpagina.
GEBRUIKSINSTRUCTIES VAN LABEL: {usage_instructions_text or 'Niet gevonden op de pagina.'}

VULSTOFFEN: {', '.join(additives) or 'Geen'}
INFERIEURE VORMEN: {', '.join(inferior_forms) or 'Geen'}
STERKE PUNTEN: {', '.join(strengths) or 'Geen'}
ZWAKKE PUNTEN: {', '.join(weaknesses) or 'Geen'}
CONTEXT FLAGS: {flags_text}
CERTIFICERINGEN GEVONDEN OP PAGINA: {certifications_text}

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
    {{"aspect": "Vulstoffen en additieven", "bevinding": "alle vulstoffen en additieven uit ingredientenlijst", "oordeel": "Goed|Matig|Slecht|DATA_LACUNE"}}
  ],
  "highlights": [
    {{"type": "positief", "tekst": "sterk punt 1"}},
    {{"type": "positief", "tekst": "sterk punt 2"}},
    {{"type": "negatief", "tekst": "zwak punt 1"}},
    {{"type": "negatief", "tekst": "zwak punt 2"}}
  ],
  "context_flags": ["alleen product-specifieke waarschuwingen op basis van ingredienten en doseringen"],
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


# ─── Scraper helpers ──────────────────────────────────────────────────────────
def fetch_html(url):
    response = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
    response.raise_for_status()
    return BeautifulSoup(response.content, "lxml")


def fetch_with_scrapingbee(url):
    api_key = os.environ.get("SCRAPINGBEE_KEY")
    if api_key:
        response = requests.get(
            "https://app.scrapingbee.com/api/v1/",
            params={"api_key": api_key, "url": url, "render_js": "true", "wait": "2000"},
            timeout=30,
        )
        response.raise_for_status()
        return BeautifulSoup(response.content, "lxml")
    return fetch_html(url)


def extract_text_blocks(soup, selectors):
    texts = []
    for sel in selectors:
        for el in soup.select(sel):
            t = el.get_text(" ", strip=True)
            if t:
                texts.append(t)
    return texts


def _parse_nested_composition(text):
    nested_results = []
    pattern = re.compile(r"(?:Waarvan|waarvan|w\.o\.|Davon|dont)\s*[:\s]", re.I)
    parts = pattern.split(text)
    if len(parts) <= 1:
        return nested_results
    for part in parts[1:]:
        for m in AMOUNT_PAT.finditer(part[:300]):
            before = part[:m.start()].strip(" ()-:")
            name_candidate = re.split(r"[,;]", before)[-1].strip()
            if name_candidate and len(name_candidate) > 1:
                nested_results.append({
                    "name": name_candidate,
                    "amount": float(m.group(1).replace(",", ".")),
                    "unit": m.group(2).lower(),
                    "form": "", "type": "actief", "nested": True,
                })
    return nested_results


def parse_ingredients_from_text(text):
    ingredients = []
    seen_names = set()

    lines = re.split(r"[;|\n]", text)
    for line in lines:
        line = line.strip()
        if not line or len(line) < 3:
            continue
        is_nested = bool(re.match(r"^(waarvan|w\.o\.|of\s+which)[:\s]", line, re.I))
        if is_nested:
            line = re.sub(r"^(waarvan|w\.o\.|of\s+which)[:\s]*", "", line, flags=re.I).strip()
        form = ""
        form_match = re.search(r"\(([^)]{3,50})\)", line)
        if form_match:
            form = form_match.group(1).strip()
        match = AMOUNT_PAT.search(line)
        if match:
            amount_str = match.group(1).replace(",", ".")
            unit = match.group(2).lower()
            name = line[:match.start()].strip(" ()-:")
            if form:
                name = re.sub(r"\([^)]*\)", "", name).strip(" ()-:")
            if name and name.lower() not in seen_names and len(name) > 1:
                seen_names.add(name.lower())
                excipient = any(kw in name.lower() for kw in EXCIPIENT_KEYWORDS)
                ingredients.append({
                    "name": name,
                    "amount": float(amount_str) if amount_str else None,
                    "unit": unit, "form": form,
                    "type": "vul-additief" if excipient else "actief",
                    "nested": is_nested,
                })
        else:
            name = re.sub(r"\([^)]*\)", "", line).strip(" ()-:")
            if name and len(name) > 2 and len(name) < 60:
                excipient = any(kw in name.lower() for kw in EXCIPIENT_KEYWORDS)
                if excipient and name.lower() not in seen_names:
                    seen_names.add(name.lower())
                    ingredients.append({
                        "name": name, "amount": None, "unit": "",
                        "form": form, "type": "vul-additief", "nested": False,
                    })

    for m in BROAD_INGREDIENT_PAT.finditer(text):
        name = m.group(1).strip(" ()-:")
        if not name or len(name) < 2 or name.lower() in seen_names:
            continue
        try:
            amount = float(m.group(2).replace(",", "."))
        except ValueError:
            continue
        unit = m.group(3).lower()
        seen_names.add(name.lower())
        excipient = any(kw in name.lower() for kw in EXCIPIENT_KEYWORDS)
        ingredients.append({
            "name": name, "amount": amount, "unit": unit,
            "form": "", "type": "vul-additief" if excipient else "actief", "nested": False,
        })

    for nested in _parse_nested_composition(text):
        if nested["name"].lower() not in seen_names:
            seen_names.add(nested["name"].lower())
            ingredients.append(nested)

    return ingredients[:60]


def _split_active_excipients(ingredients):
    active = [i for i in ingredients if i.get("type") != "vul-additief"]
    excipients = [i for i in ingredients if i.get("type") == "vul-additief"]
    return active, excipients


def _extract_brand(soup, page_text, url):
    for sel in ['[itemprop="brand"]', '[itemprop="manufacturer"]', ".brand", ".manufacturer",
                '[class*="brand"]', '[class*="vendor"]', "a.brand"]:
        els = extract_text_blocks(soup, [sel])
        if els and len(els[0]) < 60:
            return els[0]
    for meta in soup.find_all("meta"):
        prop = meta.get("property", "") + meta.get("name", "")
        if "brand" in prop.lower() or "manufacturer" in prop.lower() or "og:site_name" in prop.lower():
            content = meta.get("content", "").strip()
            if content and len(content) < 60:
                return content
    title_tag = soup.find("title")
    if title_tag:
        title = title_tag.get_text(strip=True)
        parts = re.split(r"[|\-–]", title)
        if len(parts) > 1:
            candidate = parts[-1].strip()
            if 2 < len(candidate) < 40:
                return candidate
        words = title.split()
        if len(words) >= 2:
            candidate = words[1]
            if 2 < len(candidate) < 40:
                return candidate
    host_match = re.search(r"https?://(?:www\.)?([^./]+)", url)
    if host_match:
        return host_match.group(1).capitalize()
    return ""


def _extract_certifications(soup, page_text):
    found = []
    text_lower = page_text.lower()
    for kw in CERT_KEYWORDS_FIXED:
        if kw.lower() in text_lower and kw not in found:
            found.append(kw)
    return found


def _extract_warnings(page_text):
    warnings = []
    warn_pattern = re.compile(
        r"(?:Let op|Waarschuwing|Niet gebruiken|Niet aanbevolen|Raadpleeg|Overleg met|Buiten bereik)[^.!?\n]{5,300}",
        re.I
    )
    for m in warn_pattern.finditer(page_text):
        w = m.group(0).strip()
        if w not in warnings:
            warnings.append(w)
    return warnings[:5]


def scrape_product(soup, url=""):
    page_text = soup.get_text(" ", strip=True)

    name_texts = extract_text_blocks(soup, ["h1", ".product-title", ".product-name", '[itemprop="name"]'])
    product_name = (name_texts[0] if name_texts else "")[:200]

    brand_name = _extract_brand(soup, page_text, url)[:100]

    price = ""
    price_match = re.search(r"[€\$]?\s*(\d+[,\.]\d{2})", page_text)
    if price_match:
        price = f"€{price_match.group(1)}"

    serving_size = ""
    usage_instructions = ""
    serving_match = re.search(
        r"(?:Gebruik|Dosering|Aanbevolen dagelijkse|Innemen|per dag|dagdosering)[:\s]*([^\n.]{5,200})",
        page_text, re.I
    )
    if serving_match:
        serving_size = serving_match.group(1).strip()[:200]
        usage_instructions = serving_match.group(0).strip()[:300]

    package_size = ""
    _unit_pat = r"(\d+)\s*(softgels?|capsules?|tabletten?|vegicaps?|stuks?|caps?)"
    _daily_filter = re.compile(r"(per dag|dagelijks|aanbevolen|dosering)", re.I)
    # Priority: search h1 title and URL first
    for _candidate_text in [product_name, url]:
        _m = re.search(_unit_pat, _candidate_text, re.I)
        if _m:
            package_size = _m.group(0)
            break
    # Fallback: scan page_text line by line, skip lines that mention daily dosage context
    if not package_size:
        for _line in page_text.splitlines():
            if _daily_filter.search(_line):
                continue
            _m = re.search(_unit_pat, _line, re.I)
            if _m:
                package_size = _m.group(0)
                break

    ing_match = re.search(
        r"(?:ingredi[eë]nten|samenstelling|inhoudsstoffen|supplement\s*facts|ingredients)[:\s]*(.{20,4000}?)(?:\n{2,}|\*{2,}|©|Bewaar|Gebruiksaanwijzing|Aanbevolen\s+dagelijkse|Disclaimer|$)",
        page_text, re.I | re.S
    )
    full_ingredient_text = ing_match.group(1) if ing_match else page_text[:4000]
    ingredients = parse_ingredients_from_text(full_ingredient_text)
    active_ingredients, excipients = _split_active_excipients(ingredients)

    certifications = _extract_certifications(soup, page_text)

    claims_texts = extract_text_blocks(soup, [".claims", "[class*='claim']", ".usp", "[class*='usp']", "[class*='benefit']"])
    health_claims = [c for c in claims_texts if 10 < len(c) < 300][:8]
    if not health_claims:
        for m in re.finditer(r"(?:ondersteunt|bevordert|helpt|verbetert|zorgt voor)[^.!?\n]{5,120}", page_text, re.I):
            claim = m.group(0).strip()
            if claim not in health_claims:
                health_claims.append(claim)
        health_claims = health_claims[:8]

    warnings = _extract_warnings(page_text)

    return {
        "product_name": product_name,
        "brand_name": brand_name,
        "ingredients": ingredients,
        "active_ingredients": active_ingredients,
        "excipients": excipients,
        "serving_size": serving_size,
        "usage_instructions": usage_instructions,
        "package_size": package_size,
        "price": price,
        "health_claims": health_claims,
        "certifications": certifications,
        "warnings": warnings,
        "additional_info": page_text[:5000],
    }


def is_sufficient(data):
    return bool(data.get("product_name")) and len(data.get("ingredients", [])) >= 1


def anthropic_fallback(url, page_text=""):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    page_sample = page_text[:2000] if page_text else "Niet beschikbaar"
    prompt = f"""Analyseer deze supplementpagina en extraheer productinformatie.
URL: {url}
Pagina tekst: {page_sample}

Identificeer ALLE certificeringen die op de pagina staan vermeld, inclusief certificeringen die niet in standaard lijsten voorkomen maar wel expliciet worden genoemd.

Geef terug als JSON:
{{"product_name": "naam", "brand_name": "merk", "ingredients": [{{"name": "naam", "amount": 0, "unit": "mg", "form": "vorm"}}], "serving_size": "serving", "usage_instructions": "instructies", "package_size": "verpakkingsgrootte", "price": "prijs", "health_claims": ["claim"], "certifications": ["cert"], "warnings": ["waarschuwing"], "additional_info": "info"}}"""
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    try:
        result = json.loads(raw)
        ingredients = result.get("ingredients", [])
        active, excipients = _split_active_excipients(ingredients)
        result["active_ingredients"] = active
        result["excipients"] = excipients
        return result
    except Exception:
        return {
            "product_name": "", "brand_name": "", "ingredients": [],
            "active_ingredients": [], "excipients": [],
            "health_claims": [], "certifications": [], "warnings": [], "additional_info": "",
        }


def extract_with_claude(url, raw_text, product_data):
    """Second-pass Claude extraction when health_claims are missing or certifications < 2."""
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    prompt = f"""Analyseer de onderstaande paginatekst van een supplementproduct en extraheer de volgende informatie zo volledig mogelijk.

URL: {url}
Paginatekst (eerste 6000 tekens):
{raw_text[:6000]}

Geef terug als JSON met ALLEEN deze velden:
{{
  "health_claims": ["volledige gezondheidsclaim 1", "claim 2"],
  "certifications": ["certificering 1", "certificering 2"],
  "ingredients": [{{"name": "naam", "amount": 0, "unit": "mg", "form": "vorm"}}],
  "package_size": "getal + eenheid bv 120 capsules",
  "usage_instructions": "innameadvies letterlijk van de pagina"
}}

Regels:
- health_claims: alle claims die iets beweren over gezondheid, werking of doel van het product
- certifications: alle keurmerken, kwaliteitslogo's en certificeringen die expliciet worden vermeld
- ingredients: volledige ingrediëntenlijst met naam, hoeveelheid en eenheid
- package_size: het totale aantal capsules/tabletten/softgels in de verpakking als getal + eenheid
- usage_instructions: de letterlijke innametekst van de pagina, of leeg als niet gevonden"""
    try:
        response = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"```json\s*", "", raw)
        raw = re.sub(r"```\s*", "", raw)
        extracted = json.loads(raw)
    except Exception:
        return product_data

    # Merge: Claude values take precedence when non-empty
    if extracted.get("health_claims"):
        product_data["health_claims"] = extracted["health_claims"]
    if extracted.get("certifications"):
        existing = set(c.lower() for c in product_data.get("certifications", []))
        for cert in extracted["certifications"]:
            if cert.lower() not in existing:
                product_data.setdefault("certifications", []).append(cert)
                existing.add(cert.lower())
    if extracted.get("ingredients"):
        if not product_data.get("ingredients"):
            ingredients = extracted["ingredients"]
            active, excipients = _split_active_excipients(ingredients)
            product_data["ingredients"] = ingredients
            product_data["active_ingredients"] = active
            product_data["excipients"] = excipients
    if extracted.get("package_size") and not product_data.get("package_size"):
        product_data["package_size"] = extracted["package_size"]
    if extracted.get("usage_instructions") and not product_data.get("usage_instructions"):
        product_data["usage_instructions"] = extracted["usage_instructions"]

    return product_data


# ─── Routes ───────────────────────────────────────────────────────────────────
@app.route("/scrape", methods=["POST"])
def scrape():
    data = request.get_json(silent=True) or {}
    url = data.get("url", "").strip()
    if not url:
        return jsonify({"error": "URL required"}), 400
    product_data = {}
    page_text = ""
    try:
        soup = fetch_with_scrapingbee(url)
        page_text = soup.get_text(" ", strip=True)
        product_data = scrape_product(soup, url)
        product_data["_source"] = "scraper"
        # Second-pass Claude extraction: triggers independently when health_claims are missing OR certifications are sparse
        _needs_second_pass = (
            not product_data.get("health_claims")
            or len(product_data.get("certifications", [])) < 2
        )
        if _needs_second_pass:
            try:
                product_data = extract_with_claude(url, page_text, product_data)
            except Exception:
                pass
    except Exception as e:
        product_data["_source"] = "error"
        product_data["_error"] = str(e)
    if not is_sufficient(product_data):
        try:
            product_data = anthropic_fallback(url, page_text)
            product_data["_source"] = "anthropic_fallback"
        except Exception as e:
            product_data["_fallback_error"] = str(e)
    product_data["url"] = url
    return jsonify(product_data)


@app.route("/score", methods=["POST"])
def score():
    product_data = request.get_json(silent=True) or {}
    try:
        url = product_data.get("url", "")
        product_data["url"] = url
        criteria = load_engine_criteria()
        product_type = detect_product_type(product_data)
        ingredients = product_data.get("ingredients", [])

        # Sustainability check for omega-3
        if product_type == "OMEGA3":
            sus_keywords = ["msc", "friend of the sea", "dolphin safe", "marine stewardship"]
            certs = product_data.get("certifications", [])
            has_sus = any(any(sk in c.lower() for sk in sus_keywords) for c in certs)
            if not has_sus:
                certs = list(certs) + ["DATA_LACUNE: Geen duurzaamheidscertificering gevonden. Herkomst en vangstmethode onbekend."]
                product_data["certifications"] = certs

        # Context flags: standard rules + cofactor checks
        context_flags_triggered = evaluate_context_flags(ingredients)
        context_flags_triggered += evaluate_cofactor_checks(ingredients)
        product_data["context_flags_triggered"] = context_flags_triggered

        # Intake advice
        intake_advice = get_intake_advice(ingredients)

        # Complexity tier
        active_count = len(product_data.get("active_ingredients", ingredients))
        if active_count <= 2:
            product_complexity_tier = "Single"
        elif active_count <= 6:
            product_complexity_tier = "Multi"
        else:
            product_complexity_tier = "Complex"

        # Price calculations
        price_str = product_data.get("price", "")
        package_size_str = product_data.get("package_size", "")
        serving_size_str = product_data.get("serving_size", "")
        price_per_day = calculate_price_per_day(price_str, package_size_str, serving_size_str, product_data.get("usage_instructions", ""))
        price_per_gram = calculate_price_per_gram(price_str, ingredients, package_size_str, serving_size_str)

        # Evaluate with Claude
        evaluations_data, relevant_criteria = evaluate_criteria_with_claude(product_data, criteria, product_type, url)

        # Confidence score
        evals = evaluations_data.get("evaluations", [])
        exact_count = sum(1 for e in evals if e.get("data_quality", "").upper() == "EXACT")
        total_evals = len(evals)
        confidence = exact_count / total_evals if total_evals > 0 else 0.0
        low_confidence_warning = (
            "Beperkte productinformatie beschikbaar. Deze beoordeling is gebaseerd op wat publiek vermeld staat op de productpagina."
            if confidence < 0.4 else None
        )

        # Score calculation
        score_pct, critical_fail, non_verifiable_count = calculate_score(evaluations_data, relevant_criteria)
        score_100, kwalificatie, verdict = determine_verdict(score_pct, critical_fail)

        # Consumer output
        consumer_output = generate_consumer_output(
            product_data, evaluations_data, relevant_criteria,
            score_100, kwalificatie, verdict, product_type, critical_fail,
            context_flags_triggered=context_flags_triggered
        )

        # Fallback score from beoordeling_tabel when no criteria loaded or score is 0
        if (not relevant_criteria or score_100 == 0) and consumer_output.get("beoordeling_tabel"):
            tabel = consumer_output["beoordeling_tabel"]
            goed = sum(1 for r in tabel if r.get("oordeel", "").upper() == "GOED")
            matig = sum(1 for r in tabel if r.get("oordeel", "").upper() == "MATIG")
            fallback = min(100, round((goed * 12 + matig * 6) / 96 * 100))
            if fallback > score_100:
                score_100 = fallback
                score_pct = fallback / 100
                score_100, kwalificatie, verdict = determine_verdict(score_pct, critical_fail)

        # Apply jargon simplification
        consumer_output = simplify_jargon(consumer_output)

        response_body = {
            "product_name": product_data.get("product_name", "Onbekend"),
            "brand": product_data.get("brand_name", "Onbekend"),
            "score": score_100,
            "kwalificatie": kwalificatie,
            "verdict": verdict,
            "product_type": product_type,
            "product_complexity_tier": product_complexity_tier,
            "critical_gate": critical_fail,
            "non_verifiable_count": non_verifiable_count,
            "criteria_evaluated": len(relevant_criteria),
            "confidence": round(confidence, 2),
            "price": product_data.get("price", ""),
            "package_size": product_data.get("package_size", ""),
            "price_per_day": price_per_day,
            "price_per_gram": price_per_gram,
            "intake_advice": intake_advice,
            "context_flags_triggered": context_flags_triggered,
            **consumer_output
        }
        if low_confidence_warning:
            response_body["low_confidence_warning"] = low_confidence_warning

        return jsonify(response_body)
    except Exception as e:
        return jsonify({"error": f"Engine error: {str(e)}"}), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

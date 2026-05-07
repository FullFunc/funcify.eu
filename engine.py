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
CLAUDE_MODEL = "claude-sonnet-4-5"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}

# ─── Certificerings categorieën ──────────────────────────────────────────────
QUALITY_CERT_LIST = [
    "ifos", "gmp", "iso 22000", "iso22000", "informed sport", "nzvt",
    "creapure", "nsf", "haccp", "usp", "iso 17025", "iso17025",
    "fssc 22000", "brc", "ifs", "traacs", "albion",
]

SUSTAINABILITY_CERT_LIST = [
    "msc", "friend of the sea", "dolphin safe", "rainforest alliance",
    "fair trade", "vegan society", "soil association", "organic",
    "biologisch", "green-e", "golden omega",
]

# ─── Product type gewichten (review vs biology) ───────────────────────────────
PRODUCT_TYPE_WEIGHTS = {
    "OMEGA3":       {"review": 0.40, "biology": 0.60},
    "LIPID":        {"review": 0.45, "biology": 0.55},
    "THYROID":      {"review": 0.30, "biology": 0.70},
    "MINERAL":      {"review": 0.40, "biology": 0.60},
    "INFLAMMATION": {"review": 0.35, "biology": 0.65},
    "PROBIOTIC":    {"review": 0.35, "biology": 0.65},
    "STRESS":       {"review": 0.35, "biology": 0.65},
    "VITAMIN":      {"review": 0.40, "biology": 0.60},
    "IRON":         {"review": 0.45, "biology": 0.55},
    "SPORT":        {"review": 0.40, "biology": 0.60},
    "PROTEIN":      {"review": 0.45, "biology": 0.55},
    "BOTANICAL":    {"review": 0.35, "biology": 0.65},
    "DEFAULT":      {"review": 0.40, "biology": 0.60},
}

# ─── Severity volgorde voor sortering ────────────────────────────────────────
SEVERITY_ORDER = {"Critical": 0, "Major": 1, "Minor": 2, "Info": 3}

# ─── Producttype complexiteit labels ─────────────────────────────────────────
COMPLEXITY_LABELS = {
    "Single": "Enkelvoudig",
    "Multi": "Meervoudig",
    "Complex": "Complex",
}

# ─── Context Flag Rules ───────────────────────────────────────────────────────
CONTEXT_FLAG_RULES = [
    {"id":"CF001","keywords":["pyridoxine","vitamine b6","vitamin b6","b6"],"operator":"gte","threshold":25,"unit":"mg","severity":"Critical","message":"Dit product bevat een hoge B6-dosering. Langdurige inname boven 25 mg per dag kan bij sommige mensen perifere neuropathie veroorzaken. Gebruik niet structureel zonder begeleiding."},
    {"id":"CF002","keywords":["pyridoxine","vitamine b6","b6"],"operator":"gte","threshold":50,"unit":"mg","severity":"Critical","message":"Dit product bevat een zeer hoge B6-dosering. Alleen bij specifieke therapeutische indicatie onder begeleiding."},
    {"id":"CF003","keywords":["retinol","vitamine a","vitamin a"],"operator":"gte","threshold":3000,"unit":"IU","severity":"Critical","message":"Dit product bevat een hoge dosis vitamine A. Accumulatie is mogelijk bij dagelijks gebruik. Niet aanbevolen tijdens zwangerschap."},
    {"id":"CF004","keywords":["retinol","vitamine a"],"operator":"gte","threshold":10000,"unit":"IU","severity":"Critical","message":"Hoge retinol-inname is geassocieerd met aangeboren afwijkingen. Dit product is niet geschikt tijdens zwangerschap."},
    {"id":"CF005","keywords":["jodium","iodine","kaliumjodide","natriumjodide","zeewier","kelp"],"operator":"gte","threshold":150,"unit":"mcg","severity":"Major","message":"Dit product bevat jodium. Bij bestaande schildklierproblemen kan extra jodium klachten verergeren."},
    {"id":"CF006","keywords":["jodium","kelp","blaaswier","fucus vesiculosus"],"operator":"gte","threshold":500,"unit":"mcg","severity":"Critical","message":"Zeer hoge jodium-inname. Boven 500 mcg per dag bestaat risico op schildklierproblematiek, ook zonder voorgeschiedenis."},
    {"id":"CF007","keywords":["ashwagandha","withania somnifera"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"Ashwagandha kan de schildklierfunctie beïnvloeden. Bij schildklierproblematiek of schildkliermedicatie: gebruik alleen onder begeleiding."},
    {"id":"CF008","keywords":["ijzer","ferrum","iron","feso4","ferrousfumaraat","ijzerbisglycinaat"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"IJzersuppletie is alleen zinvol bij aangetoond tekort. Overmatige ijzerinname is geassocieerd met oxidatieve stress. Niet aanbevolen zonder indicatie."},
    {"id":"CF009","keywords":["ijzer","ferrum","iron"],"operator":"gte","threshold":45,"unit":"mg","severity":"Critical","message":"Hoge ijzerdosering. Boven 45 mg per dag verhoogd risico op bijwerkingen. Alleen onder medische supervisie."},
    {"id":"CF010","keywords":["zink","zinc"],"operator":"gte","threshold":25,"unit":"mg","severity":"Major","message":"Hoge zink-inname kan koper verdringen. Langdurig gebruik zonder koper-suppletie kan leiden tot koperdepletie."},
    {"id":"CF011","keywords":["zink","zinc"],"operator":"gte","threshold":40,"unit":"mg","severity":"Critical","message":"Zinkdosering nadert de veiligheidsgrens van 40 mg per dag. Niet combineren met andere zinkbevattende supplementen."},
    {"id":"CF012","keywords":["calcium","calciumcarbonaat","calciumcitraat"],"operator":"gte","threshold":1200,"unit":"mg","severity":"Major","message":"Hoge calcium-inname kan de opname van ijzer, zink en magnesium verminderen. Neem niet gelijktijdig in met andere mineralen."},
    {"id":"CF013","keywords":["selenium","natriumseleniet","selenomethionine"],"operator":"gte","threshold":200,"unit":"mcg","severity":"Major","message":"Selenium heeft een smal therapeutisch venster. Chronisch hoge inname is geassocieerd met selenose."},
    {"id":"CF014","keywords":["vitamine k","vitamin k","k2","mk-7","menaquinon","phylloquinon"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Vitamine K interfereert direct met bloedverdunners zoals Acenocoumarol en Warfarine. Bij gebruik van bloedverdunners: overleg altijd met je arts."},
    {"id":"CF015","keywords":["epa","dha","omega-3","omega 3","visolie","fish oil","levertraan"],"operator":"gte","threshold":3000,"unit":"mg","severity":"Major","message":"Hoge omega-3 dosering heeft een mild bloedverdunnend effect. Bij antistollingsmedicatie of geplande operatie: overleg met je arts."},
    {"id":"CF016","keywords":["knoflookextract","allium sativum","ginkgo biloba","ginkgo"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"Dit ingrediënt heeft bloedverdunnende eigenschappen. Niet combineren met antistollingsmedicatie zonder medisch advies."},
    {"id":"CF017","keywords":["sint-janskruid","hypericum perforatum","st john"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Sint-Janskruid verlaagt de werkzaamheid van veel medicijnen waaronder anticonceptie, antidepressiva en bloedverdunners. Niet combineren met medicijnen."},
    {"id":"CF018","keywords":["curcumine","curcuma longa","kurkuma"],"operator":"gte","threshold":500,"unit":"mg","severity":"Major","message":"Hoge curcumine-dosering heeft anticoagulerend effect. Bij galstenen of antistollingsmedicatie: gebruik voorzichtig."},
    {"id":"CF019","keywords":["retinol","vitamine a"],"operator":"gte","threshold":800,"unit":"mcg","severity":"Critical","message":"Vitamine A boven aanbevolen zwangerschapsdosering. Hoog risico bij zwangerschap."},
    {"id":"CF020","keywords":["ashwagandha","withania somnifera","rhodiola","rhodiola rosea"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Adaptogenen zijn onvoldoende onderzocht op veiligheid tijdens zwangerschap en borstvoeding. Gebruik wordt afgeraden."},
    {"id":"CF021","keywords":["vitamine b6","pyridoxine"],"operator":"gte","threshold":10,"unit":"mg","severity":"Major","message":"B6 boven 10 mg per dag tijdens zwangerschap: afstemmen met arts."},
    {"id":"CF022","keywords":["foliumzuur","folaat","5-mthf","methylfolaat"],"operator":"any","threshold":0,"unit":"","severity":"Info","message":"Dit product bevat folaat. Relevant bij zwangerschapswens en zwangerschap. Controleer of de dosering minimaal 400 mcg per dag is."},
    {"id":"CF023","keywords":["vitamine d3","cholecalciferol","vitamin d3"],"operator":"gte","threshold":4000,"unit":"IU","severity":"Major","message":"Hoge vitamine D3 dosering. Bij zwangerschap of nierfunctiestoornissen: gebruik niet zonder medisch advies."},
    {"id":"CF024","keywords":["vitamine c","ascorbinezuur","natriumascorbaat","vitamin c"],"operator":"gte","threshold":1000,"unit":"mg","severity":"Major","message":"Hoge vitamine C dosering verhoogt de oxalaatuitscheiding via de nier. Bij nierstenen of verminderde nierfunctie: gebruik voorzichtig."},
    {"id":"CF025","keywords":["magnesium"],"operator":"gte","threshold":350,"unit":"mg","severity":"Major","message":"Hoge magnesium-inname. Bij nierfunctiestoornissen kan magnesium niet goed worden uitgescheiden."},
    {"id":"CF026","keywords":["creatine"],"operator":"any","threshold":0,"unit":"","severity":"Minor","message":"Creatine verhoogt creatinine-waarden in het bloed wat bloedtesten kan vertekenen. Bij verminderde nierfunctie: gebruik alleen na medisch overleg."},
    {"id":"CF027","keywords":["groene thee","egcg","camellia sinensis"],"operator":"gte","threshold":400,"unit":"mg","severity":"Major","message":"Geconcentreerde groene thee-extracten zijn in zeldzame gevallen geassocieerd met leverschade bij nuchter gebruik."},
    {"id":"CF028","keywords":["niacine","nicotinezuur","vitamine b3"],"operator":"gte","threshold":500,"unit":"mg","severity":"Major","message":"Hoge niacine-dosering kan flush veroorzaken en bij langdurig gebruik lever-enzymstoringen geven."},
    {"id":"CF029","keywords":["informed sport","nzvt","wada","doping"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Dit product heeft geen aantoonbare doping-certificering. Voor competitieve sporters bestaat risico op een onbedoelde positieve dopingtest."},
    {"id":"CF030","keywords":["creatine"],"operator":"any","threshold":0,"unit":"","severity":"Minor","message":"Creatine is een toegestaan supplement voor sporters. Controleer altijd of het specifieke product een doping-keurmerk heeft."},
    {"id":"CF031","keywords":["cafeine","caffeine","cafeïne"],"operator":"gte","threshold":200,"unit":"mg","severity":"Major","message":"Hoge cafeïne-dosering. Bij hart- en vaataandoeningen, slaapproblemen of angstklachten: gebruik voorzichtig."},
    {"id":"CF032","keywords":["beta-alanine"],"operator":"gte","threshold":3200,"unit":"mg","severity":"Minor","message":"Beta-alanine kan tintelingen veroorzaken bij hogere doseringen. Niet gevaarlijk maar kan als oncomfortabel worden ervaren."},
    {"id":"CF033","keywords":["berberine","berberis"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Berberine heeft een significant bloedsuikerverlagend effect. Combinatie met diabetesmedicatie kan een gevaarlijk lage bloedsuiker veroorzaken."},
    {"id":"CF034","keywords":["chroom","chroompicolinaat","chromium"],"operator":"gte","threshold":200,"unit":"mcg","severity":"Major","message":"Chroom kan insulinegevoeligheid verbeteren. Bij gelijktijdig gebruik van insuline of diabetesmedicatie: risico op hypoglykemie."},
    {"id":"CF035","keywords":["alfa-liponzuur","ala","r-liponzuur","alpha lipoic"],"operator":"gte","threshold":300,"unit":"mg","severity":"Major","message":"Alfa-liponzuur verlaagt bloedsuiker. Bij gebruik van diabetesmedicatie of schildkliermedicatie: controleer dosering."},
    {"id":"CF036","keywords":["foliumzuur","folic acid","pteroylglutaminezuur"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"Synthetisch foliumzuur kan bij MTHFR-polymorfisme onvoldoende worden omgezet. De actieve vorm methylfolaat is de betere keuze."},
    {"id":"CF037","keywords":["cyanocobalamine","cyanocobalamin"],"operator":"any","threshold":0,"unit":"","severity":"Minor","message":"Cyanocobalamine vereist een extra omzettingsstap. Bij MTHFR-polymorfisme is methylcobalamine de voorkeursvorm."},
    {"id":"CF038","keywords":["same","s-adenosyl"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"SAMe is een methyldonor en kan serotoninespiegel verhogen. Bij antidepressiva of bipolaire stoornis: niet zonder begeleiding gebruiken."},
    {"id":"CF039","keywords":["echinacea","echinacea purpurea","echinacea angustifolia"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"Echinacea stimuleert het immuunsysteem. Bij auto-immuunziekten of immunosuppressiva: gebruik wordt afgeraden."},
    {"id":"CF040","keywords":["astragalus","astragalus membranaceus","huang qi"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"Astragalus is een immuunstimulant. Bij auto-immuunziekten of immunosuppressiva: vermijd dit supplement."},
    {"id":"CF041","keywords":["vitamine d3","cholecalciferol","vitamin d3"],"operator":"gte","threshold":2000,"unit":"IU","severity":"Info","message":"Vitamine D speelt een rol bij immuunregulatie. Bij auto-immuunziekten kan hogere vitamine D therapeutisch relevant zijn. Bespreek dit met je arts."},
    {"id":"CF042","keywords":["coq10","coenzym q10","ubiquinol","ubiquinone"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"CoQ10 is extra relevant bij statinegebruik. Statines verlagen de aanmaak van CoQ10 in het lichaam."},
    {"id":"CF043","keywords":["rode gistrijst","monascus purpureus","monacoline"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Rode gistrijst bevat een natuurlijk statine. Combinatie met statinetherapie geeft verhoogd risico op spierschade en leverbelasting."},
    {"id":"CF044","keywords":["magnesium"],"operator":"any","threshold":0,"unit":"","severity":"Info","message":"Magnesium heeft een gunstige rol bij bloeddrukregulatie en hartritme. Relevant bij hypertensie of hartritmeproblematiek."},
    {"id":"CF045","keywords":["melatonine","melatonin"],"operator":"gte","threshold":0.5,"unit":"mg","severity":"Major","message":"Melatonine kan interacteren met antidepressiva en slaap- of kalmeringsmiddelen. Bij gebruik van deze medicijnen: raadpleeg je arts."},
    {"id":"CF046","keywords":["5-htp","5-hydroxytryptofaan"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"5-HTP verhoogt de serotoninesynthese. Niet combineren met antidepressiva of MAO-remmers: risico op serotoninesyndroom."},
    {"id":"CF047","keywords":["l-tryptofaan","tryptofaan","tryptophan"],"operator":"gte","threshold":500,"unit":"mg","severity":"Critical","message":"Hoge tryptofaan-dosering verhoogt serotonineproductie. Bij gebruik van SSRI of MAO-remmers: risico op serotoninesyndroom."},
    {"id":"CF048","keywords":["gaba","gamma-aminoboterzuur"],"operator":"any","threshold":0,"unit":"","severity":"Minor","message":"GABA-supplementen kunnen het centrale zenuwstelsel beïnvloeden. Bij benzodiazepinen of gabapentinoïden: gebruik voorzichtig."},
    {"id":"CF049","keywords":["valeriaan","valeriana officinalis","valerian"],"operator":"any","threshold":0,"unit":"","severity":"Minor","message":"Valeriaan heeft milde sedatieve eigenschappen. Bij slaap- of kalmeringsmiddelen: additief effect mogelijk."},
    {"id":"CF050","keywords":["lactobacillus","bifidobacterium","saccharomyces","probiotica","probiotic"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Levende bacteriën in probiotica kunnen bij ernstig immuungecompromitteerde personen systemische infecties veroorzaken. Raadpleeg je arts bij chemotherapie of transplantatie."},
    {"id":"CF051","keywords":["saccharomyces cerevisiae","saccharomyces boulardii"],"operator":"any","threshold":0,"unit":"","severity":"Major","message":"Saccharomyces is een gist. Bij gebruik van antifunginale middelen worden gistprobiotica geïnactiveerd en is suppletie ineffectief."},
    {"id":"CF052","keywords":["kinderen","children","pediatric","kind"],"operator":"any","threshold":0,"unit":"","severity":"Critical","message":"Dit product is niet specifiek geformuleerd of gedoseerd voor kinderen. Doseringen voor volwassenen zijn niet direct toepasbaar voor kinderen."},
    {"id":"CF053","keywords":["vitamine d3","cholecalciferol"],"operator":"gte","threshold":1000,"unit":"IU","severity":"Major","message":"Vitamine D dosering boven 1000 IU per dag voor kinderen: controleer of dit past bij de leeftijdsspecifieke behoefte."},
    {"id":"CF054","keywords":["vitamine k2","mk-7"],"operator":"any","threshold":0,"unit":"","severity":"Info","message":"Vitamine K2 is relevant voor botgezondheid en cardiovasculaire bescherming bij senioren. Controleer interactie met eventuele antistollingsmedicatie."},
    {"id":"CF055","keywords":["vitamine e","tocoferol","vitamin e"],"operator":"gte","threshold":400,"unit":"IU","severity":"Major","message":"Hoge vitamine E heeft een antistollend effect. Bij bloedverdunners is er een verhoogd risico op bloedingen."},
    {"id":"CF056","keywords":["beta-caroteen","beta carotene","vitamine a","retinol"],"operator":"gte","threshold":5000,"unit":"IU","severity":"Critical","message":"Hoge beta-caroteen dosering bij rokers is geassocieerd met verhoogd longkankerrisico. Hoge retinol bij zwangerschap heeft teratogeen risico."},
    {"id":"CF057","keywords":["vitamine d3","cholecalciferol"],"operator":"gte","threshold":4000,"unit":"IU","severity":"Major","message":"Zeer hoge vitamine D3 dosering. Vetoplosbare vitaminen stapelen bij langdurig hoge inname. Controleer serumspiegels bij gebruik boven 4000 IU per dag."},
    {"id":"CF058","keywords":["vitamine d3","cholecalciferol","vitamin d"],"operator":"any","threshold":0,"unit":"","severity":"Info","message":"In de Nederlandse winter is zonlicht onvoldoende voor vitamine D-aanmaak. Een dosering onder 1000 IU per dag kan onvoldoende zijn zonder zonblootstelling."},
    {"id":"CF059","keywords":["vitamine d3","cholecalciferol"],"operator":"gte","threshold":2000,"unit":"IU","severity":"Info","message":"In de zomer met voldoende buitenactiviteit kan endogene vitamine D aanmaak voldoende zijn. Hoge suppletie gecombineerd met zonblootstelling kan bij gevoelige personen tot hypervitaminose leiden."},
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
            if rule_id == "CF024" and ing.get("type") == "vul-additief":
                continue
            if rule["operator"] == "any":
                triggered.append({
                    "id": rule_id,
                    "severity": rule["severity"],
                    "message": rule["message"]
                })
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
                    triggered.append({
                        "id": rule_id,
                        "severity": rule["severity"],
                        "message": rule["message"]
                    })
                    triggered_ids.add(rule_id)
                    break
    triggered.sort(key=lambda f: SEVERITY_ORDER.get(f.get("severity", "Info"), 4))
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
        warnings.append({"id": "COF001", "severity": "Major", "message": "Hoge vitamine D zonder K2 kan calciumdepositie in bloedvaten bevorderen. Overweeg een product met K2 erbij."})

    zinc_mg = get_amount_mg(["zink", "zinc"])
    copper_present = has_ing(["koper", "copper", "cuprum"])
    if zinc_mg >= 25 and not copper_present:
        warnings.append({"id": "COF002", "severity": "Major", "message": "Hoog zink zonder koper kan bij langdurig gebruik koperdepletie veroorzaken."})

    iron_present = has_ing(["ijzer", "ferrum", "iron"])
    vitc_present = has_ing(["vitamine c", "ascorbinezuur", "natriumascorbaat", "vitamin c"])
    if iron_present and not vitc_present:
        warnings.append({"id": "COF003", "severity": "Minor", "message": "Vitamine C verbetert ijzeropname significant. Combineer bij voorkeur met vitamine C of neem apart met vitamine C-rijke voeding."})

    omega3_present = has_ing(["epa", "dha", "omega-3", "omega 3", "visolie", "fish oil"])
    vite_present = has_ing(["vitamine e", "tocoferol", "tocopherol", "vitamin e"])
    if omega3_present and not vite_present:
        warnings.append({"id": "COF004", "severity": "Major", "message": "Omega-3 vetzuren zijn gevoelig voor oxidatie. Controleer of het product antioxidantbescherming bevat zoals vitamine E."})

    warnings.sort(key=lambda f: SEVERITY_ORDER.get(f.get("severity", "Info"), 4))
    return warnings


# ─── Intake Advice ────────────────────────────────────────────────────────────
def get_intake_advice(product_data):
    usage = (product_data.get("usage_instructions") or "").strip()
    if usage and len(usage) > 10:
        return usage
    return ""


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
    engine_path = os.path.join(BASE_DIR, "Funcify. Engine Review V4.xlsx")
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
        model=CLAUDE_MODEL,
        max_tokens=4000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    return json.loads(raw), all_relevant


def calculate_score(evaluations_data, criteria, product_type="DEFAULT"):
    weights = PRODUCT_TYPE_WEIGHTS.get(product_type, PRODUCT_TYPE_WEIGHTS["DEFAULT"])
    review_w = weights["review"]
    biology_w = weights["biology"]

    evaluations = evaluations_data.get("evaluations", [])
    core_score_raw = 0.0
    core_max = 0.0
    module_score_raw = 0.0
    module_max = 0.0
    core_critical_fail = False
    module_critical_fail = False
    non_verifiable_count = 0

    for i, criterion in enumerate(criteria):
        eval_item = next((e for e in evaluations if e.get("criterion_index") == i + 1), None)
        pass_value = eval_item.get("pass_value", -1) if eval_item else -1
        data_quality = (eval_item.get("data_quality", "") if eval_item else "").upper()
        weight = criterion["weight"]
        is_core = criterion.get("source") == "core"

        if data_quality == "NIET_GEVONDEN":
            pass_value = -1

        if pass_value == -1:
            non_verifiable_count += 1
            continue  # Pass=-1 telt NIET mee in noemer
        elif pass_value == 1:
            if is_core:
                core_score_raw += weight
                core_max += weight
            else:
                module_score_raw += weight
                module_max += weight
        elif pass_value == 0:
            if is_core:
                core_max += weight
                if criterion["critical"] and data_quality in ("EXACT", "AFGELEID"):
                    core_critical_fail = True
            else:
                module_max += weight
                if criterion["critical"] and data_quality in ("EXACT", "AFGELEID"):
                    module_critical_fail = True

    core_pct = (core_score_raw / core_max) if core_max > 0 else 0.0
    module_pct = (module_score_raw / module_max) if module_max > 0 else 0.0

    if core_critical_fail:
        core_pct = min(core_pct, 0.49)
    if module_critical_fail:
        module_pct = min(module_pct, 0.49)

    if core_max > 0 and module_max > 0:
        overall_pct = (core_pct * review_w) + (module_pct * biology_w)
    elif core_max > 0:
        overall_pct = core_pct
    else:
        overall_pct = module_pct

    critical_fail = core_critical_fail or module_critical_fail
    if critical_fail:
        overall_pct = min(overall_pct, 0.49)

    return overall_pct, critical_fail, non_verifiable_count


def determine_verdict(score_pct, critical_fail):
    score_100 = round(score_pct * 100)
    if critical_fail and score_100 < 50:
        return score_100, "Afkeur", "Af te raden"
    elif score_100 >= 85:
        return score_100, "Elite", "Koopwaardig"
    elif score_100 >= 70:
        return score_100, "Degelijk", "Koopwaardig"
    elif score_100 >= 50:
        return score_100, "Matig", "Alleen met context"
    else:
        return score_100, "Afkeur", "Af te raden"


def generate_consumer_output(product_data, evaluations_data, criteria, score_100, kwalificatie, verdict, product_type, critical_fail, context_flags_triggered=None):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    # Certificeringen splitsen
    all_certs = product_data.get("certifications", [])
    quality_certs = [c for c in all_certs if any(q in c.lower() for q in QUALITY_CERT_LIST)]
    sustainability_certs = [c for c in all_certs if any(s in c.lower() for s in SUSTAINABILITY_CERT_LIST)]

    # Cert oordeel voor rij 5
    if quality_certs and sustainability_certs:
        cert_oordeel = "GOED"
    elif quality_certs or sustainability_certs:
        cert_oordeel = "MATIG"
    else:
        cert_oordeel = "ONTBREEKT"

    # Portiegrootte beoordeling voor rij 7
    serving_size_str = product_data.get("serving_size", "")
    servings_per_day = _parse_servings_per_day(serving_size_str, product_data.get("usage_instructions", ""))
    portie_waarschuwing = "PORTIEGROOTTE WAARSCHUWING: meer dan 4 eenheden per dag. Beoordeel rij 7 als MATIG." if servings_per_day > 4 else ""

    # Adaptieve sectie 6 instructie
    if score_100 >= 75:
        sectie6_instructie = "VOOR WIE IS DIT PRODUCT GESCHIKT: beschrijf in 2-3 zinnen in tweede persoon voor wie dit product het meest geschikt is. Noem eventueel een alternatief producttype voor een andere doelgroep. Geen verbeterpunten."
        sectie6_kop = "voor_wie_geschikt"
        sectie7_tonen = False
    elif score_100 >= 50:
        sectie6_instructie = "WAT ZOU EEN BETER PRODUCT BEVATTEN: combineer maximaal 2 concrete verbeterpunten met een korte doelgroepomschrijving. Gebruik preferred_format_text als basis maar schrijf in gewone taal. Maximaal 3 zinnen."
        sectie6_kop = "wat_zou_beter"
        sectie7_tonen = True
    else:
        sectie6_instructie = "WAT ZOU EEN BETER PRODUCT BEVATTEN: alleen verbeterpunten. Geen doelgroepomschrijving. Sluit verplicht af met: Vraag bij aankoop altijd om een onafhankelijk testrapport per batch."
        sectie6_kop = "wat_zou_beter"
        sectie7_tonen = True

    # Inname-advies alleen van label
    usage_instructions = (product_data.get("usage_instructions") or "").strip()
    inname_tekst = f"Inname-advies van label: {usage_instructions}" if usage_instructions and len(usage_instructions) > 10 else "Geen inname-advies op label gevonden — schrijf dit dan ook NIET."

    system_prompt = """Je bent de Funcify beoordelingsengine voor Nederlandse consumenten.

UNIVERSELE ABSOLUTE REGELS — NOOIT overtreden:
1. Schrijf ALTIJD in tweede persoon (je, jij). Nooit derde persoon.
2. Maximaal 2-3 zinnen per sectie en per tabelrij in de kolom bevinding.
3. VERBODEN in output: TG, rTG, EE, PL, COA, TOTOX, PV, AV, CFU, UL, EFSA, HPLC, ING-codes, DATA_LACUNE, DB-codes.
4. Oordelen in de tabel zijn uitsluitend: GOED, MATIG, SLECHT of ONTBREEKT. Nooit DATA_LACUNE.
5. ONTBREEKT = informatie niet beschikbaar. SLECHT = informatie aanwezig maar aantoonbaar slecht. Nooit verwarren.
6. Inname-advies in Sectie 2 ALLEEN als het letterlijk op het label staat. Anders niet schrijven.
7. Rij 8 excipients: gebruik ALLEEN de excipients lijst. Maak NOOIT aannames over capsulemateriaal.
8. Minimaal 2 sterktes EN minimaal 2 zwaktes in highlights. Altijd.
9. Cofactor-bevindingen horen als zwakte in highlights: bijv. zink zonder koper, D3 zonder K2.
10. Geen fabricaat. Geen aannames. Alleen wat in de productdata staat.

SECTIE 3 — BEOORDELING TABEL: exact 8 rijen in deze volgorde, geen uitzonderingen:

Rij 1 — Welke vorm krijg je?
Beschrijf de moleculaire vorm van het hoofdingredient in gewone taal. Is het een goed opneembare of minder goed opneembare vorm? Geen afkortingen.

Rij 2 — Krijg je genoeg per dag?
Beschrijf de exacte dagdosering per hoofdingredient. Zeg of het voldoende is voor een merkbaar effect. Bij proprietary blend: vermeld dat exacte doseringen niet beschikbaar zijn.

Rij 3 — Hoe goed wordt het opgenomen?
Beschrijf de opneembaarheid van de gebruikte vormen in gewone taal. Benoem of er cofactoren of de innamevorm de opname ondersteunen. Geen getallen of percentages.

Rij 4 — Is alles duidelijk vermeld?
Beoordeel: (a) moleculaire vormen vermeld, (b) exacte doseringen vermeld, (c) serving size duidelijk, (d) hulpstoffen volledig vermeld. Proprietary blend = automatisch SLECHT.

Rij 5 — Onafhankelijk gecontroleerd?
Benoem ALTIJD twee categorieën apart: kwaliteitscertificering EN duurzaamheidscertificering. Gebruik het voorberekende oordeel dat je ontvangt.

Rij 6 — Kloppen de gezondheidsclaims?
Beschrijf gevonden claims. Sluiten ze logisch aan bij de ingrediënten en doseringen? Als geen claims beschikbaar: ONTBREEKT.

Rij 7 — Hoeveel per verpakking?
Beschrijf serving size, dagdosering en totaal aantal servings. Bereken dagvoorraad in dagen. Prijs per dag als beschikbaar.

Rij 8 — Zijn er onnodige toevoegingen?
Gebruik ALLEEN de excipients lijst. Als leeg: ONTBREEKT. Beoordeel op: onnodige additieven aanwezig (SLECHT) of passend (GOED).

Geef alleen valide JSON terug zonder markdown."""

    critical_gate_str = "JA" if critical_fail else "NEE"
    sectie7_tonen_str = "JA" if sectie7_tonen else "NEE — weglaten"
    voor_wie_placeholder = "voor wie is dit product nog bruikbaar — 1-2 zinnen" if sectie7_tonen else ""

    user_prompt = f"""Product: {product_data.get('product_name', 'Onbekend')} ({product_data.get('brand_name', 'Onbekend')})
Score: {score_100}/100 | Kwalificatie: {kwalificatie} | Verdict: {verdict}
Producttype: {product_type} | Critical gate: {critical_gate_str}

INGREDIENTEN:
{chr(10).join([f"- {ing.get('name','')}: {ing.get('amount','')} {ing.get('unit','')} (vorm: {ing.get('form','niet vermeld')})" for ing in product_data.get('ingredients', [])]) or 'Geen ingredienten gevonden'}

EXCIPIENTS — gebruik ALLEEN dit voor rij 8, maak geen aannames:
{chr(10).join([f"- {e.get('name', e) if isinstance(e, dict) else e}" for e in product_data.get('excipients', [])]) or 'Geen excipients beschikbaar — rij 8 oordeel = ONTBREEKT'}

{inname_tekst}

GEZONDHEIDSCLAIMS:
{chr(10).join([f"- {c}" for c in product_data.get('health_claims', [])]) or 'Geen claims gevonden'}

KWALITEITSCERTIFICERINGEN (voor rij 5 categorie 1):
{', '.join(quality_certs) or 'Geen kwaliteitscertificering gevonden'}

DUURZAAMHEIDSCERTIFICERINGEN (voor rij 5 categorie 2):
{', '.join(sustainability_certs) or 'Geen duurzaamheidscertificering gevonden'}

VOORBEREKEND OORDEEL RIJ 5: {cert_oordeel}

SERVING SIZE: {product_data.get('serving_size', 'Niet vermeld')}
VERPAKKINGSGROOTTE: {product_data.get('package_size', 'Niet vermeld')}
PRIJS PER DAG: {product_data.get('price_per_day', 'Niet beschikbaar')}
{portie_waarschuwing}

STERKE PUNTEN UIT EVALUATIE: {', '.join(evaluations_data.get('key_strengths', []))}
ZWAKKE PUNTEN UIT EVALUATIE: {', '.join(evaluations_data.get('key_weaknesses', []))}

CONTEXT WAARSCHUWINGEN (voor highlights en sectie 5):
{chr(10).join([f.get('message','') for f in (context_flags_triggered or [])]) or 'Geen'}

SECTIE 6 INSTRUCTIE: {sectie6_instructie}
SECTIE 7 TONEN: {sectie7_tonen_str}

AANVULLENDE PAGINATEKST:
{product_data.get('additional_info', '')[:2000]}

Genereer valide JSON met exact deze structuur:
{{
  "wat_doet": "2-3 zinnen wat dit supplement doet. Sluit af met inname-advies ALLEEN als beschikbaar op label. Daarna voor wie relevant.",
  "beoordeling_tabel": [
    {{"aspect": "Welke vorm krijg je?", "bevinding": "beschrijving in gewone taal", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}},
    {{"aspect": "Krijg je genoeg per dag?", "bevinding": "beschrijving in gewone taal", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}},
    {{"aspect": "Hoe goed wordt het opgenomen?", "bevinding": "beschrijving in gewone taal", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}},
    {{"aspect": "Is alles duidelijk vermeld?", "bevinding": "beschrijving in gewone taal", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}},
    {{"aspect": "Onafhankelijk gecontroleerd?", "bevinding": "kwaliteitscert: [lijst] | duurzaamheidscert: [lijst]", "oordeel": "{cert_oordeel}"}},
    {{"aspect": "Kloppen de gezondheidsclaims?", "bevinding": "beschrijving in gewone taal", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}},
    {{"aspect": "Hoeveel per verpakking?", "bevinding": "serving size, dagdosering, dagvoorraad", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}},
    {{"aspect": "Zijn er onnodige toevoegingen?", "bevinding": "ALLEEN excipients lijst gebruiken", "oordeel": "GOED|MATIG|SLECHT|ONTBREEKT"}}
  ],
  "highlights": [
    {{"type": "positief", "tekst": "Goed nieuws — [sterkte 1]"}},
    {{"type": "positief", "tekst": "Goed nieuws — [sterkte 2]"}},
    {{"type": "negatief", "tekst": "Let op — [zwakte 1]"}},
    {{"type": "negatief", "tekst": "Let op — [zwakte 2]"}}
  ],
  "context_flags_output": [],
  "{sectie6_kop}": "tekst conform sectie 6 instructie",
  "voor_wie": "{voor_wie_placeholder}",
  "consumer_summary": "één zin met sterkste punt en zwakste punt, maximaal 30 woorden"
}}"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=3500,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    result = json.loads(raw)

    # Zorg dat wat_zou_beter en voor_wie_geschikt altijd aanwezig zijn
    if "voor_wie_geschikt" not in result:
        result["voor_wie_geschikt"] = ""
    if "wat_zou_beter" not in result:
        result["wat_zou_beter"] = ""

    return result


# ─── Scraper helpers ──────────────────────────────────────────────────────────
def fetch_html(url):
    response = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
    response.raise_for_status()
    return BeautifulSoup(response.content, "lxml")


def fetch_with_scrapingbee(url):
    """
    Haalt pagina op via ScrapingBee als SCRAPINGBEE_KEY beschikbaar is.
    ScrapingBee rendert JavaScript zodat dynamisch geladen content
    (ingrediëntenlijsten, certificeringslogos, prijzen) beschikbaar is.
    Fallback naar gewone requests als key ontbreekt.
    """
    api_key = os.environ.get("SCRAPINGBEE_KEY")
    if not api_key:
        return fetch_html(url)
    try:
        response = requests.get(
            "https://app.scrapingbee.com/api/v1/",
            params={
                "api_key": api_key,
                "url": url,
                "render_js": "true",
                "wait": "3000",
                "wait_for": ".product-description,.ingredients,.composition,.supplement-facts",
                "block_ads": "true",
                "block_resources": "false",
                "premium_proxy": "false",
                "country_code": "nl",
            },
            timeout=45,
        )
        response.raise_for_status()
        return BeautifulSoup(response.content, "lxml")
    except Exception as e:
        print(f"ScrapingBee fout ({e}), fallback naar directe fetch")
        return fetch_html(url)


def get_scraping_source(url):
    """
    Bepaalt de beste scraping strategie op basis van de URL.
    Webshops met veel JavaScript krijgen ScrapingBee.
    Eenvoudige statische paginas krijgen directe fetch.
    """
    js_heavy_domains = [
        "bol.com", "amazon", "vitaminstore", "holland-barrett",
        "vitakruid", "bonusan", "orthica", "solgar", "now-foods",
        "iherb", "myprotein", "bodyenfitshop", "bulk",
    ]
    url_lower = url.lower()
    needs_js = any(domain in url_lower for domain in js_heavy_domains)
    api_key = os.environ.get("SCRAPINGBEE_KEY")
    if needs_js and api_key:
        return "scrapingbee"
    elif api_key:
        return "scrapingbee"
    else:
        return "direct"


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
    page_sample = page_text[:4000] if page_text else "Niet beschikbaar"
    system_prompt = """Je extraheert productinformatie van supplement-pagina's.

ABSOLUTE REGELS:
1. Schrijf ALLEEN informatie die letterlijk op de pagina staat. Verzin NOOIT iets.
2. Als een waarde niet op de pagina staat: geef null of lege string terug.
3. Verzin NOOIT hoeveelheden, vormen of ingrediënten die niet expliciet vermeld worden.
4. Certificeringen: alleen vermelden als expliciet benoemd op de pagina (logo, tekst, badge).
5. Ingredients amount: geef null als de hoeveelheid niet letterlijk op de pagina staat.
6. Geef alleen valide JSON terug zonder markdown."""
    prompt = f"""Analyseer deze supplementpagina en extraheer productinformatie.
URL: {url}
Paginatekst: {page_sample}

Geef terug als JSON:
{{"product_name": "naam", "brand_name": "merk", "ingredients": [{{"name": "naam", "amount": null, "unit": "mg", "form": "vorm of null"}}], "excipients": [{{"name": "naam", "amount": null, "unit": "", "form": "", "type": "vul-additief", "nested": false}}], "serving_size": "serving", "usage_instructions": "instructies", "package_size": "verpakkingsgrootte", "price": "prijs", "health_claims": ["claim"], "certifications": ["cert"], "warnings": ["waarschuwing"], "additional_info": ""}}"""
    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2000,
        system=system_prompt,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = response.content[0].text.strip()
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    try:
        result = json.loads(raw)
        # Combine ingredients + separately-returned excipients, then split
        all_ingredients = result.get("ingredients", []) + result.get("excipients", [])
        active, excipients = _split_active_excipients(all_ingredients)
        result["ingredients"] = all_ingredients
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
    extract_system = """Je extraheert ontbrekende velden van supplement-pagina's. Strikte regels:
1. Schrijf ALLEEN wat letterlijk op de pagina staat. Verzin NOOIT iets.
2. Amounts: geef null als niet letterlijk vermeld. Nooit schatten.
3. Certifications: alleen als expliciet benoemd (logo, badge, tekst). Geen marketing-claims als certificering.
4. Ingredients: gebruik exacte namen en hoeveelheden van de pagina.
5. Geef alleen valide JSON terug zonder markdown."""
    prompt = f"""Extraheer ontbrekende velden van deze supplementpagina.

URL: {url}
Paginatekst (eerste 6000 tekens):
{raw_text[:6000]}

Geef terug als JSON met ALLEEN deze velden:
{{
  "health_claims": ["volledige gezondheidsclaim 1", "claim 2"],
  "certifications": ["certificering 1", "certificering 2"],
  "ingredients": [{{"name": "naam", "amount": null, "unit": "mg", "form": "vorm of null"}}],
  "package_size": "getal + eenheid bv 120 capsules",
  "usage_instructions": "innameadvies letterlijk van de pagina"
}}

Regels:
- health_claims: alleen claims die iets beweren over gezondheid, werking of doel van het product
- certifications: alleen keurmerken die expliciet worden vermeld als certificering (geen marketing)
- ingredients: volledige ingrediëntenlijst met naam, hoeveelheid en eenheid — null als hoeveelheid niet vermeld
- package_size: het totale aantal capsules/tabletten/softgels in de verpakking als getal + eenheid
- usage_instructions: de letterlijke innametekst van de pagina, of leeg als niet gevonden"""
    try:
        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=2000,
            system=extract_system,
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
        existing_ingredients = product_data.get("ingredients", [])
        if not existing_ingredients:
            # No existing ingredients — use Claude extraction directly
            ingredients = extracted["ingredients"]
            active, excipients = _split_active_excipients(ingredients)
            product_data["ingredients"] = ingredients
            product_data["active_ingredients"] = active
            product_data["excipients"] = excipients
        else:
            # Existing ingredients — merge: add Claude ingredients not already seen
            seen = set(i.get("name", "").lower() for i in existing_ingredients)
            added = False
            for ing in extracted["ingredients"]:
                if ing.get("name", "").lower() not in seen:
                    existing_ingredients.append(ing)
                    seen.add(ing.get("name", "").lower())
                    added = True
            if added:
                active, excipients = _split_active_excipients(existing_ingredients)
                product_data["ingredients"] = existing_ingredients
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
        problem_ids = product_data.get("problem_ids", [])
        product_data["problem_ids"] = problem_ids
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
        intake_advice = get_intake_advice(product_data)

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
        score_pct, critical_fail, non_verifiable_count = calculate_score(evaluations_data, relevant_criteria, product_type)
        score_100, kwalificatie, verdict = determine_verdict(score_pct, critical_fail)

        # Consumer output
        consumer_output = generate_consumer_output(
            product_data, evaluations_data, relevant_criteria,
            score_100, kwalificatie, verdict, product_type, critical_fail,
            context_flags_triggered=context_flags_triggered
        )

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
            "product_complexity_label": COMPLEXITY_LABELS.get(product_complexity_tier, product_complexity_tier),
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
            "quality_certs": [c for c in product_data.get("certifications", []) if any(q in c.lower() for q in QUALITY_CERT_LIST)],
            "sustainability_certs": [c for c in product_data.get("certifications", []) if any(s in c.lower() for s in SUSTAINABILITY_CERT_LIST)],
            "problem_ids": product_data.get("problem_ids", []),
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

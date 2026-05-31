
import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="BOM Smart Assistant", page_icon="🛠️", layout="wide")

DATA_FILE = Path(__file__).parent / "BOM_TAB_FIN.xlsx"
SHEET_NAME = "Feuil1"


# =========================================================
# 1. Normalisation et détection des colonnes
# =========================================================

def norm(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def col_like(df, *patterns):
    cols = list(df.columns)
    for p in patterns:
        np = norm(p)
        for c in cols:
            if norm(c) == np:
                return c
        for c in cols:
            if np in norm(c):
                return c
    return None


def row_to_text(row):
    return norm(" ".join([str(v) for v in row.to_list() if str(v).lower() != "nan"]))


def get_value(row, col):
    if col and col in row.index:
        v = row[col]
        if pd.isna(v):
            return ""
        return str(v)
    return ""


# =========================================================
# 2. Moteur de classification maintenance
# =========================================================

def detect_component_category(text):
    t = norm(text)

    if any(k in t for k in ["roulement", "bearing", "antifriction"]):
        return "roulement"

    if any(k in t for k in ["garniture mecanique", "mechanical seal", "seal cartridge", "presse etoupe"]):
        return "garniture"

    if any(k in t for k in ["joint torique", "o ring", "oring", "gasket", "joint", "bague d etancheite", "etancheite"]):
        return "joint"

    if any(k in t for k in ["roue", "impeller", "impulseur"]):
        return "roue"

    if any(k in t for k in ["arbre", "shaft"]):
        return "arbre"

    if any(k in t for k in ["accouplement", "coupling"]):
        return "accouplement"

    if any(k in t for k in ["corps de pompe", "corps de palier", "volute", "casing", "case cover", "corps", "couvercle"]):
        return "corps"

    if any(k in t for k in ["baseplate", "chassis", "support", "pedestal", "plaque de base"]):
        return "support"

    if any(k in t for k in ["bouchon"]):
        return "bouchon"

    if any(k in t for k in ["vis", "screw", "ecrou", "nut", "washer", "rondelle", "goujon", "stud", "bolt"]):
        return "fixation"

    if any(k in t for k in ["viton", "ansi", "astm", "aisi", "dn", "ff", "rf"]):
        return "technique_a_verifier"

    return "inconnu"


def classify_from_category(category):
    if category in ["roulement", "garniture"]:
        return "Élevée", "Stock important"

    if category in ["joint", "roue"]:
        return "Moyenne à élevée", "Stock moyen"

    if category in ["arbre", "accouplement"]:
        return "Moyenne", "Stock moyen"

    if category in ["corps"]:
        return "Faible à moyenne", "Stock faible"

    if category in ["support", "bouchon", "fixation"]:
        return "Faible", "Stock faible"

    return "À vérifier", "À définir"


def justification_from_category(category, row, mapping, row_number):
    desc = get_value(row, mapping.get("desc_courte"))
    desc_long = get_value(row, mapping.get("desc_longue"))
    nom = get_value(row, mapping.get("nom"))
    tag = get_value(row, mapping.get("tag"))

    context = ""
    if tag:
        context = f" pour le TAG {tag}"

    # Plusieurs formulations par catégorie pour éviter les phrases répétées
    variants = {
        "roulement": [
            f"Le roulement assure le guidage de la rotation{context}. Sa dégradation peut provoquer vibrations, échauffement et arrêt de l’équipement ; il doit donc être priorisé en stock.",
            f"Cette pièce est sensible car elle intervient directement dans la rotation de la pompe. Une usure du roulement peut réduire la disponibilité et générer une intervention urgente.",
            f"Le roulement est un composant fonctionnel critique : il influence la stabilité mécanique, le bruit et la température de fonctionnement. Un stock de sécurité est recommandé.",
        ],
        "garniture": [
            f"La garniture mécanique assure l’étanchéité{context}. Sa défaillance peut entraîner une fuite du fluide, une perte de performance ou l’arrêt de la pompe.",
            f"Cette pièce est critique car elle limite les fuites au niveau de l’arbre. En maintenance, son indisponibilité peut prolonger fortement le temps d’intervention.",
            f"La garniture mécanique est prioritaire pour la continuité d’exploitation : elle protège l’installation contre les fuites et les arrêts liés à l’étanchéité.",
        ],
        "joint": [
            f"Le joint participe à l’étanchéité de l’ensemble. Une dégradation peut provoquer des fuites ou des pertes de pression ; un stock moyen est donc conseillé.",
            f"Cette pièce est généralement peu coûteuse mais importante pour éviter les fuites. Elle doit rester identifiable et disponible lors des interventions de maintenance.",
            f"L’élément d’étanchéité doit être suivi car son remplacement est fréquent lors des démontages. Sa criticité dépend de sa position dans la pompe et du fluide véhiculé.",
        ],
        "roue": [
            f"La roue influence directement le débit et la performance hydraulique. Une usure ou détérioration peut réduire le rendement et perturber le fonctionnement.",
            f"Cette pièce est importante pour la fonction de pompage. Son état conditionne la capacité de la pompe à assurer le débit attendu.",
            f"La roue est à surveiller car elle est exposée au fluide et peut subir usure, corrosion ou déséquilibre, ce qui justifie un suivi maintenance spécifique.",
        ],
        "arbre": [
            f"L’arbre transmet le mouvement entre l’entraînement et les parties tournantes. Sa défaillance est moins fréquente, mais l’impact sur l’arrêt de la pompe peut être important.",
            f"Cette pièce a une fonction mécanique centrale. Elle doit être suivie surtout en cas de vibration, désalignement ou usure des portées.",
            f"L’arbre est classé en criticité moyenne car son remplacement est moins courant, mais il reste essentiel pour la transmission du mouvement.",
        ],
        "accouplement": [
            f"L’accouplement participe à la transmission entre moteur et pompe. Un défaut peut générer vibrations, désalignement et usure prématurée.",
            f"Cette pièce doit être surveillée car elle influence la qualité de transmission du mouvement et peut impacter les roulements et l’arbre.",
            f"L’accouplement est important pour la fiabilité mécanique ; sa criticité reste moyenne mais son contrôle est utile lors des interventions préventives.",
        ],
        "corps": [
            f"Le corps ou couvercle est une pièce structurelle. Sa défaillance est moins fréquente, mais elle peut nécessiter un arrêt prolongé si la pièce n’est pas disponible.",
            f"Cette pièce supporte l’ensemble hydraulique ou mécanique. Elle est rarement remplacée, d’où un stock faible mais une identification claire reste nécessaire.",
            f"Le composant est important pour l’intégrité de la pompe, mais son besoin en stock courant reste limité par rapport aux pièces d’usure.",
        ],
        "support": [
            f"Ce composant a principalement une fonction de support ou de fixation. Il présente une faible criticité en stock courant, sauf en cas de dommage mécanique.",
            f"La pièce contribue au maintien de l’ensemble mais n’est pas une pièce d’usure principale. Un stock faible est suffisant dans une logique maintenance.",
            f"Son rôle est plutôt structurel ; elle doit être documentée dans la BOM, mais elle n’est pas prioritaire en stock de sécurité.",
        ],
        "bouchon": [
            f"Le bouchon est une pièce secondaire. Sa disponibilité reste utile, mais son impact direct sur l’arrêt de la pompe est généralement limité.",
            f"Cette pièce est classée faible car elle ne constitue pas un organe principal de fonctionnement, tout en restant nécessaire pour certaines opérations.",
            f"Le bouchon doit être référencé pour faciliter les interventions, mais il ne nécessite pas un niveau de stock élevé.",
        ],
        "fixation": [
            f"Il s’agit d’un élément de fixation. Sa criticité unitaire est faible, mais sa disponibilité facilite le remontage et évite les retards d’intervention.",
            f"Les fixations sont nécessaires aux opérations de maintenance, mais elles ne représentent pas en général une pièce critique de fonctionnement.",
            f"Cette pièce est classée faible car elle a un rôle d’assemblage. Elle doit rester disponible en quantité raisonnable pour les travaux de démontage/remontage.",
        ],
        "technique_a_verifier": [
            f"La désignation contient surtout des informations techniques ou matière. La fonction exacte de la pièce doit être confirmée à partir de la BOM ou du datasheet.",
            f"La description ne permet pas d’identifier clairement le rôle maintenance de la pièce. Une validation technique est nécessaire avant de fixer la criticité.",
            f"La pièce présente des informations dimensionnelles ou de matériau, mais son usage exact reste à confirmer pour éviter une classification erronée.",
        ],
        "inconnu": [
            f"La désignation disponible n’est pas suffisante pour attribuer une criticité fiable. Une vérification avec le document BOM ou le retour d’expérience est recommandée.",
            f"La pièce ne contient pas de mot-clé technique exploitable. Elle doit être contrôlée manuellement afin d’éviter une mauvaise priorité de stock.",
            f"La classification automatique reste prudente : le rôle exact du composant doit être validé par un responsable technique ou par la documentation fournisseur.",
        ],
    }

    options = variants.get(category, variants["inconnu"])
    return options[row_number % len(options)]


@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME, dtype=str).dropna(how="all")

    mapping = {
        "tag": col_like(df, "TAG D'EQUIPEMENT", "TAG EQUIPEMENT"),
        "nom": col_like(df, "NOM D'EQUIPEMENT", "NOM EQUIPEMENT"),
        "desc_courte": col_like(df, "DESCRIPTION (40 CARACTERES)", "DESCRIPTION COURTE"),
        "desc_longue": col_like(df, "CARACTERISTIQUES TECHNIQUES", "DESCRIPTION LONGUE"),
    }

    criticites = []
    stocks = []
    justifications = []

    for idx, row in df.iterrows():
        text = " ".join([
            get_value(row, mapping.get("desc_courte")),
            get_value(row, mapping.get("desc_longue")),
            get_value(row, mapping.get("nom")),
        ])
        category = detect_component_category(text)
        crit, stock = classify_from_category(category)
        justif = justification_from_category(category, row, mapping, int(idx))
        criticites.append(crit)
        stocks.append(stock)
        justifications.append(justif)

    df["Criticité proposée"] = criticites
    df["Niveau de stock proposé"] = stocks
    df["Justification maintenance"] = justifications

    return df, mapping


# =========================================================
# 3. Affichage et export
# =========================================================

def visible_cols(df, mapping):
    cols = [
        mapping.get("tag"),
        mapping.get("nom"),
        mapping.get("desc_courte"),
        mapping.get("desc_longue"),
        "Criticité proposée",
        "Niveau de stock proposé",
        "Justification maintenance",
    ]
    return [c for c in cols if c and c in df.columns]


def export_xlsx(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BOM_ANALYSE")
        ws = writer.book["BOM_ANALYSE"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
            cell.alignment = cell.alignment.copy(horizontal="center", vertical="center", wrap_text=True)
    out.seek(0)
    return out


# =========================================================
# 4. Assistant questions-réponses amélioré
# =========================================================

def find_tag(question, tags):
    q = norm(question)
    for tag in tags:
        if tag != "Tous" and norm(tag) in q:
            return tag
    return None


def filter_by_question(question, df, mapping, tags):
    q = norm(question)
    data = df.copy()

    tag = find_tag(question, tags)
    tag_col = mapping.get("tag")

    if tag and tag_col:
        data = data[data[tag_col].fillna("").astype(str) == tag]

    row_text = data.apply(row_to_text, axis=1)

    intent = "general"
    title = "Résultat de la recherche dans le tableau BOM"

    if any(w in q for w in ["resume", "résumé", "resumer", "résumer", "synthese", "synthèse"]):
        intent = "resume"
        title = "Résumé maintenance"

    elif any(w in q for w in ["critique", "critical", "elevee", "élevée", "prioritaire"]):
        intent = "critique"
        data = data[data["Criticité proposée"] == "Élevée"]
        title = "Pièces à criticité élevée"

    elif any(w in q for w in ["stock", "disponible", "rechange"]):
        intent = "stock"
        data = data[data["Niveau de stock proposé"].isin(["Stock important", "Stock moyen"])]
        title = "Pièces à suivre en priorité pour le stock"

    elif any(w in q for w in ["verifier", "vérifier", "a verifier", "à vérifier", "validation"]):
        intent = "verifier"
        data = data[data["Criticité proposée"] == "À vérifier"]
        title = "Pièces nécessitant une vérification"

    elif any(w in q for w in ["roulement", "bearing"]):
        intent = "roulement"
        data = data[row_text.apply(lambda x: "roulement" in x or "bearing" in x)]
        title = "Pièces liées aux roulements"

    elif any(w in q for w in ["joint", "gasket", "oring", "o ring", "etancheite", "étanchéité"]):
        intent = "joint"
        data = data[row_text.apply(lambda x: any(k in x for k in ["joint", "gasket", "oring", "o ring", "etancheite"]))]
        title = "Pièces liées aux joints et à l’étanchéité"

    elif any(w in q for w in ["pourquoi", "justification", "explique", "raison"]):
        intent = "justification"
        title = "Explication maintenance"

    if tag:
        title += f" du TAG {tag}"

    return title, intent, tag, data


def counts_text(data):
    total = len(data)
    high = int((data["Criticité proposée"] == "Élevée").sum())
    med_high = int((data["Criticité proposée"] == "Moyenne à élevée").sum())
    medium = int((data["Criticité proposée"] == "Moyenne").sum())
    check = int((data["Criticité proposée"] == "À vérifier").sum())
    stock_imp = int((data["Niveau de stock proposé"] == "Stock important").sum())

    return total, high, med_high, medium, check, stock_imp


def answer_question(question, df, mapping, tags):
    title, intent, tag, data = filter_by_question(question, df, mapping, tags)
    total, high, med_high, medium, check, stock_imp = counts_text(data)

    if intent == "resume":
        answer = f"""
**{title}**

L’équipement sélectionné regroupe **{total} lignes BOM**. L’analyse fait ressortir **{high} pièces à criticité élevée**, **{med_high} pièces à criticité moyenne à élevée** et **{check} éléments à vérifier**.

Les pièces à criticité élevée doivent être traitées en priorité, car elles peuvent avoir un impact direct sur la disponibilité de la pompe, notamment en cas de fuite, vibration, échauffement ou arrêt. Les éléments classés **à vérifier** nécessitent une validation complémentaire à partir de la documentation fournisseur ou du retour d’expérience maintenance.

**Priorité recommandée :**
1. contrôler les pièces à criticité élevée ;
2. vérifier les éléments non clairement identifiés ;
3. confirmer le niveau de stock selon la fréquence d’intervention et la politique de maintenance.
"""

    elif intent == "critique":
        answer = f"""
**{title}**

J’ai identifié **{total} pièce(s)** classée(s) en **criticité élevée**. Ces éléments doivent être suivis en priorité, car leur indisponibilité peut prolonger le temps d’arrêt ou compliquer une intervention corrective.

Cette liste est utile pour préparer le stock de sécurité, organiser les interventions et repérer les composants qui peuvent avoir un impact direct sur la continuité de fonctionnement.
"""

    elif intent == "stock":
        answer = f"""
**{title}**

Le tableau fait apparaître **{total} pièce(s)** nécessitant un suivi stock prioritaire, dont **{stock_imp}** avec un **stock important** proposé. Ces pièces sont à considérer en priorité pour éviter les retards lors des interventions.

Le niveau de stock proposé reste indicatif : il doit être ajusté selon la fréquence de remplacement, la criticité réelle de l’équipement et les délais d’approvisionnement.
"""

    elif intent == "verifier":
        answer = f"""
**{title}**

J’ai relevé **{total} ligne(s)** dont la classification automatique reste **à vérifier**. Cela signifie que la désignation disponible n’est pas assez claire pour attribuer une criticité fiable.

Ces lignes doivent être contrôlées à partir du datasheet, du plan constructeur ou de la BOM d’origine afin d’éviter une erreur de codification ou une mauvaise priorité de stock.
"""

    elif intent == "roulement":
        answer = f"""
**{title}**

J’ai trouvé **{total} ligne(s)** liées aux roulements. Ces pièces sont généralement importantes pour la fiabilité mécanique de la pompe, car elles influencent la rotation, les vibrations et l’échauffement.

En maintenance, les roulements doivent être facilement identifiables et disponibles, surtout lorsque la pompe est critique pour la continuité de service.
"""

    elif intent == "joint":
        answer = f"""
**{title}**

J’ai identifié **{total} ligne(s)** liées aux joints ou à l’étanchéité. Ces éléments sont importants lors des démontages et remontages, car leur dégradation peut entraîner des fuites ou une perte de performance.

Il est recommandé de les suivre avec attention, surtout pour les pompes manipulant des fluides sensibles ou corrosifs.
"""

    elif intent == "justification":
        answer = f"""
**{title}**

La justification maintenance est établie selon le rôle probable de la pièce : rotation, étanchéité, transmission, fonction hydraulique, support ou fixation. Plus la pièce influence directement l’arrêt, la fuite, la vibration ou la performance de la pompe, plus sa criticité proposée est élevée.

Cette logique reste volontairement prudente : lorsqu’une désignation est ambiguë, l’outil classe la pièce **à vérifier** au lieu d’inventer une criticité non confirmée.
"""

    else:
        answer = f"""
**{title}**

La recherche a retourné **{total} ligne(s)**. Parmi elles, on trouve **{high} pièce(s) à criticité élevée**, **{med_high} pièce(s) à criticité moyenne à élevée** et **{check} élément(s) à vérifier**.

Tu peux préciser la question avec un TAG, un type de pièce ou une criticité pour obtenir une réponse plus ciblée.
"""

    answer += "\n\n*Remarque : cette analyse est indicative et doit être validée par l’historique des pannes, le retour d’expérience maintenance et la politique de stock de l’entreprise.*"
    return answer, data


# =========================================================
# 5. Interface
# =========================================================

st.title("🛠️ BOM Smart Assistant")
st.caption("Assistant intelligent pour exploiter le tableau final BOM des pompes et aider la maintenance.")

df, mapping = load_data()
tag_col = mapping.get("tag")
tags = ["Tous"] + sorted(df[tag_col].dropna().astype(str).unique().tolist()) if tag_col else ["Tous"]

st.sidebar.success("Tableau final BOM chargé automatiquement")
tag_value = st.sidebar.selectbox("Filtrer par TAG", tags)
keyword = st.sidebar.text_input("Recherche par mot-clé", placeholder="Ex. roulement, joint, corps, SULZER...")
criticity = st.sidebar.selectbox("Criticité", ["Toutes"] + sorted(df["Criticité proposée"].dropna().unique().tolist()))

filtered = df.copy()
if tag_value != "Tous" and tag_col:
    filtered = filtered[filtered[tag_col].fillna("").astype(str) == tag_value]

if keyword:
    kw = keyword.lower()
    filtered = filtered[filtered.apply(lambda r: kw in " ".join([str(v) for v in r.to_list()]).lower(), axis=1)]

if criticity != "Toutes":
    filtered = filtered[filtered["Criticité proposée"] == criticity]

cols = visible_cols(filtered, mapping)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Lignes analysées", len(df))
c2.metric("Lignes affichées", len(filtered))
c3.metric("Pièces critiques", int((df["Criticité proposée"] == "Élevée").sum()))
c4.metric("À vérifier", int((df["Criticité proposée"] == "À vérifier").sum()))

st.divider()
st.subheader("💡 Exemples de questions possibles")
st.markdown("""
- Quelles sont les pièces critiques du TAG 120AP01 ?
- Résumé du TAG 120AP01.
- Donne-moi les joints du TAG 120AP01.
- Quels sont les roulements disponibles dans le tableau ?
- Quelles pièces nécessitent un stock important ?
- Quelles pièces sont à vérifier ?
- Pourquoi cette pièce est classée critique ?
""")

st.subheader("🤖 Assistant questions-réponses")
question = st.text_input("Votre question", placeholder="Ex. Quelles sont les pièces critiques du TAG 120AP01 ?")

if question:
    answer, res_df = answer_question(question, df, mapping, tags)
    st.markdown(answer)
    res_cols = visible_cols(res_df, mapping)
    st.dataframe(res_df[res_cols].head(50), use_container_width=True, height=350)

st.divider()
st.subheader("Résumé maintenance")

if tag_value != "Tous" and tag_col:
    tdf = df[df[tag_col].astype(str) == tag_value]
    st.write(f"**TAG sélectionné :** {tag_value}")
    st.write(f"Nombre de composants associés : **{len(tdf)}**")
    st.write(f"Pièces à criticité élevée : **{int((tdf['Criticité proposée'] == 'Élevée').sum())}**")
    st.write(f"Pièces à vérifier : **{int((tdf['Criticité proposée'] == 'À vérifier').sum())}**")
else:
    st.write("Sélectionne un TAG dans le menu à gauche pour obtenir un résumé spécifique d’une pompe.")

st.caption("Remarque : la criticité proposée est qualitative et indicative. Elle doit être confirmée par l’historique des pannes, le retour d’expérience maintenance et la politique de stock de l’entreprise.")

st.divider()
st.subheader("Tableau BOM analysé")
st.dataframe(filtered[cols], use_container_width=True, height=520)

st.download_button(
    "📥 Télécharger le tableau analysé en Excel",
    data=export_xlsx(filtered[cols]),
    file_name="BOM_analyse_criticite_stock.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

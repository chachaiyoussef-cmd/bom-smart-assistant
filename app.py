
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
# 1. Outils de base
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


def safe_str(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)


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


def get_value(row, col):
    if col and col in row.index:
        return safe_str(row[col])
    return ""


def row_to_text(row):
    return norm(" ".join([safe_str(v) for v in row.to_list()]))


def contains_any(text, keywords):
    t = norm(text)
    return any(k in t for k in keywords)


# =========================================================
# 2. Classification maintenance
# =========================================================

def detect_component_family(text):
    t = norm(text)

    if any(k in t for k in ["roulement", "bearing", "antifriction"]):
        return "Rotation", "Roulement"

    if any(k in t for k in ["garniture mecanique", "mechanical seal", "seal cartridge", "presse etoupe"]):
        return "Étanchéité", "Garniture mécanique"

    if any(k in t for k in ["joint torique", "o ring", "oring", "gasket", "joint", "bague d etancheite", "etancheite"]):
        return "Étanchéité", "Joint / élément d’étanchéité"

    if any(k in t for k in ["roue", "impeller", "impulseur"]):
        return "Hydraulique", "Roue / impeller"

    if any(k in t for k in ["arbre", "shaft"]):
        return "Transmission", "Arbre"

    if any(k in t for k in ["accouplement", "coupling"]):
        return "Transmission", "Accouplement"

    if any(k in t for k in ["corps de pompe", "corps de palier", "volute", "casing", "case cover", "corps", "couvercle"]):
        return "Structure", "Corps / couvercle"

    if any(k in t for k in ["baseplate", "chassis", "support", "pedestal", "plaque de base"]):
        return "Support", "Support / châssis"

    if any(k in t for k in ["bouchon"]):
        return "Accessoire", "Bouchon"

    if any(k in t for k in ["vis", "screw", "ecrou", "nut", "washer", "rondelle", "goujon", "stud", "bolt"]):
        return "Fixation", "Élément de fixation"

    if any(k in t for k in ["viton", "ansi", "astm", "aisi", "dn", "ff", "rf", "grade"]):
        return "À vérifier", "Donnée technique à confirmer"

    return "À vérifier", "Non identifié automatiquement"


def classify_from_family(family, component_type):
    if component_type in ["Roulement", "Garniture mécanique"]:
        return "Élevée", "Stock important", "P1 – Prioritaire"

    if component_type in ["Joint / élément d’étanchéité", "Roue / impeller"]:
        return "Moyenne à élevée", "Stock moyen", "P2 – Important"

    if component_type in ["Arbre", "Accouplement"]:
        return "Moyenne", "Stock moyen", "P2 – Important"

    if component_type in ["Corps / couvercle"]:
        return "Faible à moyenne", "Stock faible", "P3 – Standard"

    if component_type in ["Support / châssis", "Bouchon", "Élément de fixation"]:
        return "Faible", "Stock faible", "P3 – Standard"

    return "À vérifier", "À définir", "À vérifier"


def justification_from_family(family, component_type, row, mapping, index):
    tag = get_value(row, mapping.get("tag"))
    tag_txt = f" pour le TAG {tag}" if tag else ""

    variants = {
        "Roulement": [
            f"Le roulement assure le guidage de la rotation{tag_txt}. Sa défaillance peut générer vibrations, échauffement et arrêt de l’équipement ; un stock de sécurité est donc recommandé.",
            "Composant sensible de la chaîne de rotation : son usure peut impacter directement la disponibilité de la pompe et augmenter le risque d’intervention urgente.",
            "Le roulement influence la stabilité mécanique, le bruit et la température de fonctionnement. Il est classé prioritaire en raison de son impact sur la fiabilité.",
        ],
        "Garniture mécanique": [
            f"La garniture mécanique assure l’étanchéité{tag_txt}. Une défaillance peut provoquer une fuite, une perte de performance ou l’arrêt de la pompe.",
            "Pièce critique pour la maîtrise des fuites au niveau de l’arbre. Son indisponibilité peut prolonger fortement la durée d’intervention.",
            "Élément prioritaire pour la continuité d’exploitation, car il protège l’installation contre les fuites et les arrêts liés à l’étanchéité.",
        ],
        "Joint / élément d’étanchéité": [
            "Le joint participe à l’étanchéité de l’ensemble. Sa dégradation peut entraîner des fuites ou des pertes de pression ; un stock moyen est conseillé.",
            "Pièce généralement consommable lors des démontages. Sa disponibilité facilite le remontage et limite les risques de fuite après intervention.",
            "Élément important pour maintenir l’étanchéité ; sa criticité dépend de sa position dans la pompe et du fluide véhiculé.",
        ],
        "Roue / impeller": [
            "La roue influence directement le débit et la performance hydraulique. Une usure ou détérioration peut réduire le rendement de la pompe.",
            "Composant essentiel pour la fonction de pompage : son état conditionne la capacité à assurer le débit attendu.",
            "Pièce exposée au fluide, pouvant subir usure, corrosion ou déséquilibre ; elle nécessite un suivi maintenance spécifique.",
        ],
        "Arbre": [
            "L’arbre transmet le mouvement vers les parties tournantes. Sa défaillance est moins fréquente, mais son impact sur l’arrêt peut être important.",
            "Pièce mécanique centrale à surveiller en cas de vibration, désalignement ou usure des portées.",
            "Classé en criticité moyenne : remplacement moins courant, mais rôle essentiel dans la transmission du mouvement.",
        ],
        "Accouplement": [
            "L’accouplement assure la transmission entre moteur et pompe. Un défaut peut provoquer vibrations, désalignement et usure prématurée.",
            "Pièce importante pour la qualité de transmission du mouvement ; elle peut impacter indirectement l’arbre et les roulements.",
            "À surveiller lors des contrôles préventifs, surtout en présence de vibrations ou d’écarts d’alignement.",
        ],
        "Corps / couvercle": [
            "Pièce structurelle importante. Sa défaillance est moins fréquente, mais peut entraîner un arrêt prolongé si la pièce n’est pas disponible.",
            "Le composant supporte l’ensemble hydraulique ou mécanique ; il est rarement remplacé, d’où un stock faible mais une identification claire nécessaire.",
            "Son besoin en stock courant reste limité par rapport aux pièces d’usure, mais il doit rester correctement référencé dans la BOM.",
        ],
        "Support / châssis": [
            "Composant à fonction de support. Sa criticité est faible en stock courant, sauf en cas de dommage mécanique.",
            "Pièce utile au maintien de l’ensemble, mais non considérée comme pièce d’usure principale.",
            "Rôle principalement structurel ; la disponibilité en stock peut rester limitée.",
        ],
        "Bouchon": [
            "Pièce secondaire utile pour certaines opérations de maintenance, mais avec impact direct généralement limité sur l’arrêt de la pompe.",
            "Élément à référencer pour faciliter les interventions, sans nécessiter un niveau de stock élevé.",
            "Criticité faible, mais disponibilité utile pour éviter les petites indisponibilités lors du remontage.",
        ],
        "Élément de fixation": [
            "Élément de fixation nécessaire au montage/remontage. Criticité unitaire faible, mais disponibilité utile pour éviter les retards d’intervention.",
            "Les fixations ne sont pas des organes fonctionnels principaux, mais elles facilitent les opérations de maintenance.",
            "Pièce d’assemblage classée faible ; elle doit rester disponible en quantité raisonnable pour les interventions.",
        ],
        "Donnée technique à confirmer": [
            "La désignation contient surtout des informations de matière, de norme ou de dimension. La fonction exacte doit être confirmée avant classification finale.",
            "Les informations disponibles ne permettent pas d’identifier clairement le rôle maintenance ; une validation technique est nécessaire.",
            "Classification prudente : le composant doit être contrôlé avec la BOM d’origine ou le datasheet fournisseur.",
        ],
        "Non identifié automatiquement": [
            "La désignation disponible n’est pas suffisante pour attribuer une criticité fiable. Une vérification documentaire est recommandée.",
            "Aucun mot-clé technique exploitable n’a été détecté ; la pièce doit être validée manuellement pour éviter une mauvaise priorité de stock.",
            "L’outil reste prudent et classe cette ligne à vérifier afin d’éviter une classification non justifiée.",
        ],
    }

    choices = variants.get(component_type, variants["Non identifié automatiquement"])
    return choices[int(index) % len(choices)]


@st.cache_data
def load_data():
    if not DATA_FILE.exists():
        st.error("Le fichier BOM_TAB_FIN.xlsx est introuvable dans le repository.")
        st.stop()

    df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME, dtype=str).dropna(how="all")

    mapping = {
        "tag": col_like(df, "TAG D'EQUIPEMENT", "TAG EQUIPEMENT"),
        "nom": col_like(df, "NOM D'EQUIPEMENT", "NOM EQUIPEMENT"),
        "desc_courte": col_like(df, "DESCRIPTION (40 CARACTERES)", "DESCRIPTION COURTE"),
        "desc_longue": col_like(df, "CARACTERISTIQUES TECHNIQUES", "DESCRIPTION LONGUE"),
    }

    families, component_types, criticites, stocks, priorities, justifs = [], [], [], [], [], []

    for idx, row in df.iterrows():
        text = " ".join([
            get_value(row, mapping.get("desc_courte")),
            get_value(row, mapping.get("desc_longue")),
            get_value(row, mapping.get("nom")),
        ])
        family, comp_type = detect_component_family(text)
        crit, stock, priority = classify_from_family(family, comp_type)
        justif = justification_from_family(family, comp_type, row, mapping, idx)

        families.append(family)
        component_types.append(comp_type)
        criticites.append(crit)
        stocks.append(stock)
        priorities.append(priority)
        justifs.append(justif)

    df["Famille maintenance"] = families
    df["Type de composant"] = component_types
    df["Criticité proposée"] = criticites
    df["Niveau de stock proposé"] = stocks
    df["Priorité maintenance"] = priorities
    df["Justification maintenance"] = justifs

    return df, mapping


# =========================================================
# Affichage
# =========================================================

def visible_cols(df, mapping):
    cols = [
        mapping.get("tag"),
        mapping.get("nom"),
        mapping.get("desc_courte"),
        mapping.get("desc_longue"),
        "Famille maintenance",
        "Type de composant",
        "Criticité proposée",
        "Niveau de stock proposé",
        "Priorité maintenance",
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
# Chatbot amélioré
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
    title = "Résultat de la recherche"

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

    elif any(w in q for w in ["roulement", "bearing", "rotation"]):
        intent = "roulement"
        data = data[(data["Famille maintenance"] == "Rotation") | row_text.apply(lambda x: "roulement" in x or "bearing" in x)]
        title = "Pièces liées à la rotation"

    elif any(w in q for w in ["joint", "gasket", "oring", "o ring", "etancheite", "étanchéité", "fuite"]):
        intent = "joint"
        data = data[(data["Famille maintenance"] == "Étanchéité") | row_text.apply(lambda x: any(k in x for k in ["joint", "gasket", "oring", "o ring", "etancheite"]))]
        title = "Pièces liées à l’étanchéité"

    elif any(w in q for w in ["hydraulique", "roue", "impeller", "debit", "débit"]):
        intent = "hydraulique"
        data = data[data["Famille maintenance"] == "Hydraulique"]
        title = "Pièces liées à la fonction hydraulique"

    elif any(w in q for w in ["famille", "type", "categorie", "catégorie"]):
        intent = "famille"
        title = "Répartition par famille maintenance"

    elif any(w in q for w in ["recommandation", "recommander", "conseil", "actions"]):
        intent = "recommandation"
        title = "Recommandations maintenance"

    elif any(w in q for w in ["top", "plus", "maximum"]):
        intent = "top"
        title = "TAGs les plus sensibles"

    if tag:
        title += f" du TAG {tag}"

    return title, intent, tag, data


def counts(data):
    return {
        "total": len(data),
        "high": int((data["Criticité proposée"] == "Élevée").sum()),
        "med_high": int((data["Criticité proposée"] == "Moyenne à élevée").sum()),
        "medium": int((data["Criticité proposée"] == "Moyenne").sum()),
        "check": int((data["Criticité proposée"] == "À vérifier").sum()),
        "stock_imp": int((data["Niveau de stock proposé"] == "Stock important").sum()),
        "stock_moy": int((data["Niveau de stock proposé"] == "Stock moyen").sum()),
    }


def format_answer(title, intent, data, mapping):
    c = counts(data)

    if intent == "resume":
        return f"""
**{title}**

L’analyse porte sur **{c['total']} ligne(s) BOM**. Elle met en évidence **{c['high']} pièce(s) à criticité élevée**, **{c['med_high']} pièce(s) à criticité moyenne à élevée** et **{c['check']} élément(s) à vérifier**.

Les priorités concernent principalement les composants qui peuvent influencer directement l’arrêt, la fuite, la vibration ou la performance de la pompe. Les lignes classées **à vérifier** doivent être contrôlées à partir de la documentation fournisseur ou de la BOM d’origine.

**Lecture maintenance :**
1. sécuriser les pièces à criticité élevée ;
2. vérifier les désignations ambiguës ;
3. ajuster le stock selon le retour d’expérience et les délais d’approvisionnement.
"""

    if intent == "critique":
        return f"""
**{title}**

J’ai identifié **{c['total']} pièce(s)** classée(s) en **criticité élevée**. Ces pièces sont prioritaires car leur défaillance peut entraîner un arrêt de la pompe, une fuite, une vibration importante ou une perte de performance.

Elles doivent être suivies en priorité dans la politique de stock et dans la préparation des interventions de maintenance.
"""

    if intent == "stock":
        return f"""
**{title}**

Le résultat contient **{c['total']} pièce(s)** nécessitant un suivi de stock, dont **{c['stock_imp']}** avec **stock important** et **{c['stock_moy']}** avec **stock moyen**.

L’objectif est d’éviter les retards d’intervention sur les composants sensibles, surtout lorsque les délais d’approvisionnement sont longs ou que l’équipement est important pour la continuité de service.
"""

    if intent == "verifier":
        return f"""
**{title}**

J’ai trouvé **{c['total']} ligne(s)** à vérifier. Cela signifie que la désignation disponible n’est pas assez explicite pour définir une criticité fiable.

Ces lignes nécessitent une validation technique : datasheet, plan constructeur, BOM d’origine ou retour d’expérience maintenance.
"""

    if intent == "roulement":
        return f"""
**{title}**

Le tableau retourne **{c['total']} ligne(s)** liées à la rotation. Ces pièces sont importantes car elles influencent la stabilité mécanique, les vibrations et l’échauffement.

En maintenance, les roulements et éléments associés doivent être facilement identifiables et disponibles lorsque la pompe est critique.
"""

    if intent == "joint":
        return f"""
**{title}**

J’ai identifié **{c['total']} ligne(s)** liées à l’étanchéité. Ces éléments sont importants pour limiter les fuites et garantir le bon fonctionnement après démontage/remontage.

Un suivi stock est recommandé, surtout pour les joints et garnitures utilisés lors des interventions préventives ou correctives.
"""

    if intent == "hydraulique":
        return f"""
**{title}**

Le résultat contient **{c['total']} composant(s)** liés à la fonction hydraulique. Ces pièces peuvent influencer le débit, le rendement et la performance de la pompe.

Elles doivent être suivies lorsque la pompe est soumise à usure, corrosion ou conditions de fonctionnement sévères.
"""

    if intent == "famille":
        family_count = data["Famille maintenance"].value_counts().head(6)
        txt = "\n".join([f"- **{fam}** : {nb} ligne(s)" for fam, nb in family_count.items()])
        return f"""
**{title}**

Voici la répartition principale par famille maintenance :

{txt}

Cette classification permet de mieux distinguer les pièces de rotation, d’étanchéité, hydrauliques, structurelles et les éléments à vérifier.
"""

    if intent == "recommandation":
        return f"""
**{title}**

À partir du tableau analysé, les recommandations prioritaires sont :

1. **Sécuriser les pièces P1**, notamment les roulements et garnitures mécaniques.
2. **Valider les lignes à vérifier** avec les documents fournisseur afin d’éviter les erreurs de criticité.
3. **Structurer le stock** selon la criticité, la fréquence de remplacement et les délais d’achat.
4. **Mettre à jour les descriptions ambiguës** pour améliorer la qualité de la BOM.
5. **Exploiter les résultats dans SAP/GMAO** pour faciliter les futures interventions.
"""

    if intent == "top":
        tag_col = mapping.get("tag")
        if tag_col and tag_col in data.columns:
            temp = data[data["Criticité proposée"] == "Élevée"]
            top = temp[tag_col].value_counts().head(5)
            if len(top) > 0:
                txt = "\n".join([f"- **{tag}** : {nb} pièce(s) critiques" for tag, nb in top.items()])
            else:
                txt = "- Aucune pièce critique identifiée dans le périmètre sélectionné."
        else:
            txt = "- Colonne TAG non détectée."
        return f"""
**{title}**

Les TAGs les plus sensibles selon le nombre de pièces critiques sont :

{txt}

Cette lecture permet d’orienter les contrôles et la priorisation du stock.
"""

    return f"""
**{title}**

La recherche retourne **{c['total']} ligne(s)**. On y trouve **{c['high']} pièce(s) à criticité élevée**, **{c['med_high']} pièce(s) à criticité moyenne à élevée** et **{c['check']} élément(s) à vérifier**.

Pour obtenir une réponse plus précise, indique un TAG, une famille de pièce ou une priorité maintenance.
"""


def answer_question(question, df, mapping, tags):
    try:
        title, intent, tag, data = filter_by_question(question, df, mapping, tags)
        answer = format_answer(title, intent, data, mapping)
        answer += "\n\n*Remarque : cette analyse est indicative et doit être validée par l’historique des pannes, le retour d’expérience maintenance et la politique de stock de l’entreprise.*"
        return answer, data
    except Exception as e:
        fallback = """
**Je n’ai pas pu traiter la question exactement comme formulée.**

Essaie de reformuler avec un TAG, une famille de pièce ou une priorité.  
Exemples :  
- Quelles sont les pièces critiques du TAG 120AP01 ?  
- Résumé du TAG 120AP01.  
- Quelles pièces nécessitent un stock important ?  
- Quelles pièces sont à vérifier ?
"""
        return fallback, df.head(0)


# =========================================================
# Interface Streamlit
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
family_filter = st.sidebar.selectbox("Famille maintenance", ["Toutes"] + sorted(df["Famille maintenance"].dropna().unique().tolist()))

filtered = df.copy()

if tag_value != "Tous" and tag_col:
    filtered = filtered[filtered[tag_col].fillna("").astype(str) == tag_value]

if keyword:
    kw = keyword.lower()
    filtered = filtered[filtered.apply(lambda r: kw in " ".join([safe_str(v) for v in r.to_list()]).lower(), axis=1)]

if criticity != "Toutes":
    filtered = filtered[filtered["Criticité proposée"] == criticity]

if family_filter != "Toutes":
    filtered = filtered[filtered["Famille maintenance"] == family_filter]

cols = visible_cols(filtered, mapping)

# Dashboard
c1, c2, c3, c4 = st.columns(4)
c1.metric("Lignes analysées", len(df))
c2.metric("Lignes affichées", len(filtered))
c3.metric("Pièces critiques", int((df["Criticité proposée"] == "Élevée").sum()))
c4.metric("À vérifier", int((df["Criticité proposée"] == "À vérifier").sum()))

st.divider()
st.subheader("📊 Dashboard maintenance")

g1, g2 = st.columns(2)

with g1:
    st.write("**Répartition par criticité**")
    crit_chart = df["Criticité proposée"].value_counts().rename_axis("Criticité").reset_index(name="Nombre")
    st.bar_chart(crit_chart, x="Criticité", y="Nombre", use_container_width=True)

with g2:
    st.write("**Répartition par famille maintenance**")
    fam_chart = df["Famille maintenance"].value_counts().rename_axis("Famille").reset_index(name="Nombre")
    st.bar_chart(fam_chart, x="Famille", y="Nombre", use_container_width=True)

st.divider()
st.subheader("💡 Exemples de questions possibles")
st.markdown("""
- Quelles sont les pièces critiques du TAG 120AP01 ?
- Résumé du TAG 120AP01.
- Donne-moi les joints du TAG 120AP01.
- Quels sont les roulements disponibles dans le tableau ?
- Quelles pièces nécessitent un stock important ?
- Quelles pièces sont à vérifier ?
- Quelles sont les pièces liées à l’étanchéité ?
- Donne-moi la répartition par famille maintenance.
- Quels sont les TAGs les plus sensibles ?
- Donne-moi les recommandations maintenance.
""")

st.subheader("🤖 Assistant questions-réponses")
question = st.text_input("Votre question", placeholder="Ex. Quelles sont les pièces critiques du TAG 120AP01 ?")

if question:
    answer, res_df = answer_question(question, df, mapping, tags)
    st.markdown(answer)
    res_cols = visible_cols(res_df, mapping)
    if len(res_df) > 0 and res_cols:
        st.dataframe(res_df[res_cols].head(50), use_container_width=True, height=350)
    else:
        st.info("Aucun tableau à afficher pour cette question.")

st.divider()
st.subheader("Résumé automatique par TAG")

if tag_value != "Tous" and tag_col:
    tdf = df[df[tag_col].astype(str) == tag_value]
    total = len(tdf)
    high = int((tdf["Criticité proposée"] == "Élevée").sum())
    check = int((tdf["Criticité proposée"] == "À vérifier").sum())
    main_family = tdf["Famille maintenance"].value_counts().idxmax() if total > 0 else "Non disponible"

    st.markdown(f"""
Le TAG **{tag_value}** regroupe **{total} composant(s)** dans la BOM.  
L’analyse indique **{high} pièce(s) à criticité élevée** et **{check} élément(s) à vérifier**.  
La famille maintenance la plus représentée est **{main_family}**.

Cette lecture permet d’identifier rapidement les pièces prioritaires, les éléments nécessitant validation et les besoins potentiels en stock de rechange.
""")
else:
    st.write("Sélectionne un TAG dans le menu à gauche pour obtenir un résumé automatique spécifique.")

st.divider()
st.subheader("✅ Recommandations finales")
st.markdown("""
- Prioriser la validation des lignes classées **À vérifier**.
- Sécuriser en stock les pièces **P1 – Prioritaire**, notamment les roulements et garnitures mécaniques.
- Exploiter la **famille maintenance** pour mieux préparer les interventions.
- Mettre à jour les descriptions ambiguës afin d’améliorer la qualité de la BOM.
- Utiliser les résultats comme support d’aide à la décision, puis les valider par l’expérience maintenance.
""")

st.divider()
st.subheader("Tableau BOM analysé")
st.dataframe(filtered[cols], use_container_width=True, height=520)

st.download_button(
    "📥 Télécharger le tableau analysé en Excel",
    data=export_xlsx(filtered[cols]),
    file_name="BOM_analyse_criticite_stock.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

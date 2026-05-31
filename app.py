
import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="BOM Smart Assistant", page_icon="🛠️", layout="wide")

DATA_FILE = Path(__file__).parent / "BOM_TAB_FIN.xlsx"
SHEET_NAME = "Feuil1"


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


def classify(text):
    t = norm(text)
    rules = [
        (["roulement", "bearing", "antifriction"], "Élevée", "Stock important",
         "Pièce liée à la rotation ; risque de vibration, échauffement ou arrêt."),
        (["garniture mecanique", "mechanical seal", "seal cartridge", "presse etoupe"], "Élevée", "Stock important",
         "Pièce d’étanchéité sensible ; risque de fuite et arrêt possible."),
        (["joint torique", "o ring", "oring", "gasket", "joint", "bague d etancheite", "etancheite"], "Moyenne à élevée", "Stock moyen",
         "Élément d’étanchéité ; risque de fuite ou perte de performance."),
        (["roue", "impeller", "impulseur"], "Moyenne à élevée", "Stock moyen",
         "Pièce hydraulique importante ; impact direct sur le débit et la performance."),
        (["arbre", "shaft"], "Moyenne", "Stock limité",
         "Pièce de transmission ; critique mais remplacement moins fréquent."),
        (["accouplement", "coupling"], "Moyenne", "Stock moyen",
         "Élément de transmission moteur-pompe ; risque de vibration ou désalignement."),
        (["corps de pompe", "corps de palier", "volute", "casing", "case cover", "corps", "couvercle"], "Faible à moyenne", "Stock faible",
         "Pièce structurelle importante mais rarement remplacée en stock courant."),
        (["baseplate", "chassis", "support", "pedestal", "plaque de base"], "Faible", "Stock faible",
         "Élément support ; remplacement généralement peu fréquent."),
        (["bouchon"], "Faible", "Stock faible", "Pièce secondaire ; faible criticité unitaire."),
        (["vis", "screw", "ecrou", "nut", "washer", "rondelle", "goujon", "stud", "bolt"], "Faible", "Stock faible",
         "Élément de fixation ; faible criticité unitaire."),
    ]
    for keys, crit, stock, justif in rules:
        if any(k in t for k in keys):
            return crit, stock, justif
    return "À vérifier", "À définir", "Classification non automatique ; validation technique nécessaire."


@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME, dtype=str).dropna(how="all")

    mapping = {
        "tag": col_like(df, "TAG D'EQUIPEMENT", "TAG EQUIPEMENT"),
        "nom": col_like(df, "NOM D'EQUIPEMENT", "NOM EQUIPEMENT"),
        "desc_courte": col_like(df, "DESCRIPTION (40 CARACTERES)", "DESCRIPTION COURTE"),
        "desc_longue": col_like(df, "CARACTERISTIQUES TECHNIQUES", "DESCRIPTION LONGUE"),
    }

    text = pd.Series([""] * len(df), index=df.index)
    for key in ["desc_courte", "desc_longue"]:
        if mapping[key]:
            text = text + " " + df[mapping[key]].fillna("").astype(str)

    result = text.apply(classify)
    df["Criticité proposée"] = result.apply(lambda x: x[0])
    df["Niveau de stock proposé"] = result.apply(lambda x: x[1])
    df["Justification maintenance"] = result.apply(lambda x: x[2])

    return df, mapping


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


def find_tag(question, tags):
    q = norm(question)
    for tag in tags:
        if tag != "Tous" and norm(tag) in q:
            return tag
    return None


def safe_row_text(data):
    # Correctif V7 : conversion robuste de chaque valeur en texte
    # pour éviter l'erreur de chatbot sur Streamlit Cloud.
    return data.apply(lambda r: norm(" ".join([str(v) for v in r.to_list()])), axis=1)


def answer_question(question, df, mapping, tags):
    data = df.copy()
    tag = find_tag(question, tags)
    tag_col = mapping.get("tag")

    if tag and tag_col:
        data = data[data[tag_col].fillna("").astype(str) == tag]

    q = norm(question)
    row_text = safe_row_text(data)

    intro = "Résultat de la recherche dans le tableau BOM"

    if "critique" in q or "elevee" in q or "élevée" in q:
        data = data[data["Criticité proposée"] == "Élevée"]
        intro = "Voici les pièces à criticité élevée"

    elif "stock" in q:
        data = data[data["Niveau de stock proposé"].isin(["Stock important", "Stock moyen"])]
        intro = "Voici les pièces nécessitant un suivi stock prioritaire"

    elif "verifier" in q or "vérifier" in q:
        data = data[data["Criticité proposée"] == "À vérifier"]
        intro = "Voici les pièces à vérifier"

    elif "roulement" in q or "bearing" in q:
        data = data[row_text.apply(lambda x: "roulement" in x or "bearing" in x)]
        intro = "Voici les lignes liées aux roulements"

    elif "joint" in q or "gasket" in q or "etancheite" in q or "étanchéité" in q:
        data = data[row_text.apply(lambda x: any(k in x for k in ["joint", "gasket", "etancheite", "oring", "o ring"]))]
        intro = "Voici les lignes liées aux joints et éléments d’étanchéité"

    elif "resume" in q or "résumé" in q:
        intro = "Résumé maintenance"

    if tag:
        intro += f" pour le TAG {tag}"

    msg = f"""{intro}.

- Nombre de lignes trouvées : {len(data)}
- Pièces à criticité élevée : {int((data['Criticité proposée'] == 'Élevée').sum())}
- Pièces à criticité moyenne à élevée : {int((data['Criticité proposée'] == 'Moyenne à élevée').sum())}
- Pièces à vérifier : {int((data['Criticité proposée'] == 'À vérifier').sum())}

Remarque : les résultats restent indicatifs et doivent être validés par l’historique des pannes, le retour d’expérience maintenance et la politique de stock de l’entreprise.
"""
    return msg, data


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
    filtered = filtered[filtered.astype(str).apply(lambda r: kw in " ".join([str(v) for v in r.to_list()]).lower(), axis=1)]
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
- Donne-moi les joints du TAG 120AP01.
- Quels sont les roulements disponibles dans le tableau ?
- Quelles pièces nécessitent un stock important ?
- Résumé du TAG 120AP01.
- Quelles pièces sont à vérifier ?
""")

st.subheader("🤖 Assistant questions-réponses")
question = st.text_input("Votre question", placeholder="Ex. Quelles sont les pièces critiques du TAG 120AP01 ?")
if question:
    msg, res_df = answer_question(question, df, mapping, tags)
    st.markdown(msg)
    st.dataframe(res_df[visible_cols(res_df, mapping)].head(50), use_container_width=True, height=350)

st.divider()
st.subheader("Résumé maintenance")
if tag_value != "Tous" and tag_col:
    tdf = df[df[tag_col].astype(str) == tag_value]
    st.write(f"**TAG sélectionné :** {tag_value}")
    st.write(f"Nombre de composants associés : **{len(tdf)}**")
    st.write(f"Pièces à criticité élevée : **{int((tdf['Criticité proposée'] == 'Élevée').sum())}**")
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

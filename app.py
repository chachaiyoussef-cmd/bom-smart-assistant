
import re
import unicodedata
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="BOM Smart Assistant", page_icon="🛠️", layout="wide")

DATA_FILE = Path(__file__).parent / "BOM_TAB_FIN.xlsx"
SHEET_NAME = "Feuil1"

def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def tokens(value):
    return [t for t in normalize_text(value).split() if len(t) > 1]

def column_score(col_name, patterns):
    ncol = normalize_text(col_name)
    col_tokens = set(tokens(col_name))
    best = 0
    for pattern in patterns:
        npat = normalize_text(pattern)
        pat_tokens = set(tokens(pattern))
        if ncol == npat:
            best = max(best, 100)
        elif npat and npat in ncol:
            best = max(best, 80)
        elif pat_tokens and pat_tokens.issubset(col_tokens):
            best = max(best, 70)
        elif pat_tokens:
            overlap = len(pat_tokens.intersection(col_tokens))
            best = max(best, int(40 * overlap / len(pat_tokens)))
    return best

def map_columns(df):
    candidates = {
        "numero": ["n", "numero"],
        "famille": ["famille d equipement", "famille equipement"],
        "tag_equipement": ["tag d equipement", "tag equipement"],
        "tag_sap": ["tag sur sap", "tag sap"],
        "nom_equipement": ["nom d equipement", "nom equipement"],
        "sous_ensemble": ["sous ensemble de l equipement", "sous ensemble"],
        "code_sap": ["code sap", "ca sap", "code article"],
        "description_courte": ["description 40 caracteres", "description courte"],
        "description_longue": ["caracteristiques techniques ou reference description longue", "description longue", "caracteristiques techniques"],
        "unite": ["unite", "unite de mesure"],
        "reference": ["reference"],
        "quantite": ["quantite installe", "quantite", "qte"],
        "marque": ["marque", "fabricant"],
        "documentation": ["documentation extrait du datasheet plan manuel", "documentation"],
        "observations": ["observations", "observation", "remarque"],
    }
    mapping = {}
    used_cols = set()
    for std, patterns in candidates.items():
        scored = []
        for col in df.columns:
            if col in used_cols:
                continue
            scored.append((column_score(col, patterns), col))
        scored.sort(reverse=True)
        if scored and scored[0][0] >= 45:
            mapping[std] = scored[0][1]
            used_cols.add(scored[0][1])
        else:
            mapping[std] = None
    return mapping

def get_col(df, mapping, std_name):
    col = mapping.get(std_name)
    if col and col in df.columns:
        return df[col].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index)

def classify_piece(text):
    t = normalize_text(text)
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
        (["bouchon"], "Faible", "Stock faible",
         "Pièce secondaire ; faible criticité unitaire."),
        (["vis", "screw", "ecrou", "nut", "washer", "rondelle", "goujon", "stud", "bolt"], "Faible", "Stock faible",
         "Élément de fixation ; faible criticité unitaire."),
    ]
    for keywords, criticite, stock, justification in rules:
        if any(k in t for k in keywords):
            return criticite, stock, justification
    return "À vérifier", "À définir", "Classification non automatique ; validation technique nécessaire."

@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME, header=0, dtype=str)
    df = df.dropna(how="all")
    mapping = map_columns(df)
    combined = (
        get_col(df, mapping, "description_courte")
        + " "
        + get_col(df, mapping, "description_longue")
        + " "
        + get_col(df, mapping, "reference")
    )
    res = combined.apply(classify_piece)
    df["Criticité proposée"] = res.apply(lambda x: x[0])
    df["Niveau de stock proposé"] = res.apply(lambda x: x[1])
    df["Justification maintenance"] = res.apply(lambda x: x[2])
    return df, mapping

def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BOM_ANALYSE")
        ws = writer.book["BOM_ANALYSE"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
            cell.alignment = cell.alignment.copy(horizontal="center", vertical="center", wrap_text=True)
        for col in ws.columns:
            letter = col[0].column_letter
            max_len = max(len("" if c.value is None else str(c.value)) for c in col[:120])
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)
    output.seek(0)
    return output

def find_tag_in_question(question, tags):
    q = normalize_text(question)
    for tag in tags:
        if normalize_text(tag) in q:
            return tag
    return None

def select_display_columns(df, mapping):
    preferred = [
        mapping.get("tag_equipement"),
        mapping.get("description_courte"),
        mapping.get("description_longue"),
        mapping.get("reference"),
        mapping.get("quantite"),
        mapping.get("marque"),
        "Criticité proposée",
        "Niveau de stock proposé",
        "Justification maintenance",
    ]
    return [c for c in preferred if c and c in df.columns]

def chatbot_answer(question, analyzed, mapping, tags):
    q_norm = normalize_text(question)
    tag = find_tag_in_question(question, tags)
    data = analyzed.copy()
    tag_col = mapping.get("tag_equipement")
    if tag and tag_col in data.columns:
        data = data[data[tag_col].fillna("").astype(str) == tag]

    row_text = data.astype(str).apply(lambda row: normalize_text(" ".join(row.values)), axis=1)

    if any(w in q_norm for w in ["critique", "critical", "elevee", "élevée"]):
        data = data[data["Criticité proposée"] == "Élevée"]
        intro = "Voici les pièces à criticité élevée"
    elif any(w in q_norm for w in ["verifier", "vérifier", "a verifier", "à vérifier"]):
        data = data[data["Criticité proposée"] == "À vérifier"]
        intro = "Voici les pièces à vérifier"
    elif "stock" in q_norm:
        data = data[data["Niveau de stock proposé"].isin(["Stock important", "Stock moyen"])]
        intro = "Voici les pièces nécessitant un suivi stock prioritaire"
    elif any(w in q_norm for w in ["roulement", "bearing"]):
        mask = row_text.apply(lambda txt: "roulement" in txt or "bearing" in txt)
        data = data[mask]
        intro = "Voici les lignes liées aux roulements"
    elif any(w in q_norm for w in ["joint", "gasket", "oring", "o ring"]):
        mask = row_text.apply(lambda txt: any(k in txt for k in ["joint", "gasket", "oring", "o ring"]))
        data = data[mask]
        intro = "Voici les lignes liées aux joints et éléments d’étanchéité"
    elif any(w in q_norm for w in ["resume", "résumé", "resumer", "résumer"]):
        intro = "Résumé maintenance"
    else:
        intro = "Résultat de la recherche dans le tableau BOM"

    if tag:
        intro += f" pour le TAG {tag}"

    answer = (
        f"{intro}.\n\n"
        f"- Nombre de lignes trouvées : {len(data)}\n"
        f"- Pièces à criticité élevée : {int((data['Criticité proposée'] == 'Élevée').sum())}\n"
        f"- Pièces à criticité moyenne à élevée : {int((data['Criticité proposée'] == 'Moyenne à élevée').sum())}\n"
        f"- Pièces à vérifier : {int((data['Criticité proposée'] == 'À vérifier').sum())}\n\n"
        "Remarque : les résultats restent indicatifs et doivent être validés par l’historique des pannes, "
        "le retour d’expérience maintenance et la politique de stock de l’entreprise."
    )
    return answer, data

st.title("🛠️ BOM Smart Assistant")
st.caption("Assistant intelligent pour exploiter le tableau final BOM des pompes et aider la maintenance.")

analyzed, mapping = load_data()
tag_col = mapping.get("tag_equipement")
tags = ["Tous"] + sorted([x for x in analyzed[tag_col].dropna().astype(str).unique() if x.strip()]) if tag_col else ["Tous"]

st.sidebar.success("Tableau final BOM chargé automatiquement")
tag_value = st.sidebar.selectbox("Filtrer par TAG", tags)
keyword = st.sidebar.text_input("Recherche par mot-clé", placeholder="Ex. roulement, joint, corps, SULZER...")
criticity_values = ["Toutes"] + sorted(analyzed["Criticité proposée"].dropna().unique().tolist())
criticity = st.sidebar.selectbox("Criticité", criticity_values)

filtered = analyzed.copy()
if tag_value != "Tous" and tag_col:
    filtered = filtered[filtered[tag_col].fillna("").astype(str) == tag_value]
if keyword:
    kw = keyword.strip().lower()
    filtered = filtered[filtered.astype(str).apply(lambda row: kw in " ".join(row.values).lower(), axis=1)]
if criticity != "Toutes":
    filtered = filtered[filtered["Criticité proposée"] == criticity]

col1, col2, col3, col4 = st.columns(4)
col1.metric("Lignes analysées", len(analyzed))
col2.metric("Lignes affichées", len(filtered))
col3.metric("Pièces critiques", int((analyzed["Criticité proposée"] == "Élevée").sum()))
col4.metric("À vérifier", int((analyzed["Criticité proposée"] == "À vérifier").sum()))

st.divider()
st.subheader("🤖 Assistant questions-réponses")
st.write("Pose une question sur le tableau BOM.")
question = st.text_input("Votre question", placeholder="Ex. Quelles sont les pièces critiques du TAG 120AP01 ?")

if question:
    answer, chat_df = chatbot_answer(question, analyzed, mapping, tags)
    st.markdown(answer)
    cols = select_display_columns(chat_df, mapping)
    st.dataframe(chat_df[cols].head(50), use_container_width=True, height=350)

st.divider()
st.subheader("Résumé maintenance")
if tag_value != "Tous" and tag_col:
    tag_df = analyzed[analyzed[tag_col].astype(str) == tag_value]
    st.write(f"**TAG sélectionné :** {tag_value}")
    st.write(f"Nombre de composants associés : **{len(tag_df)}**")
    st.write(f"Pièces à criticité élevée : **{int((tag_df['Criticité proposée'] == 'Élevée').sum())}**")
else:
    st.write("Sélectionne un TAG dans le menu à gauche pour obtenir un résumé spécifique d’une pompe.")

st.caption(
    "Remarque : la criticité proposée est qualitative et indicative. "
    "Elle doit être confirmée par l’historique des pannes, le retour d’expérience maintenance "
    "et la politique de stock de l’entreprise."
)

st.divider()
st.subheader("Tableau BOM analysé")
st.dataframe(filtered, use_container_width=True, height=520)

st.download_button(
    "📥 Télécharger le tableau analysé en Excel",
    data=export_excel(filtered),
    file_name="BOM_analyse_criticite_stock.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.divider()
st.subheader("Colonnes détectées")
detected = pd.DataFrame([{"Champ standard": k, "Colonne détectée": v or "Non détectée"} for k, v in mapping.items()])
st.dataframe(detected, use_container_width=True)

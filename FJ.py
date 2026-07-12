from urllib.request import urlopen

SOURCE_URL = "https://raw.githubusercontent.com/BBQfish23/Feuille-Journaliere/2b94417ec1333946d196cedb1d1936adcb098729/FJ.py"

source = urlopen(SOURCE_URL, timeout=20).read().decode("utf-8")

# Retirer l'ancien correctif imghdr, inutile sous Python 3.11.
imports_marker = "import streamlit as st"
source = source[source.index(imports_marker):]

# Mettre à jour le texte d'introduction.
source = source.replace(
    "st.markdown(\"Générez la feuille journalière à partir d'un export **Deputy (CSV)** ou **Emprez (Excel hebdomadaire)**.\")",
    "st.markdown(\"Générez la feuille journalière à partir d'un export **Emprez (Excel hebdomadaire)**.\")",
)

# Configurer l'impression sur une page Lettre US, centrée horizontalement et verticalement.
print_anchor = """    # --- RETOUR DES DONNÉES ---
    output = io.BytesIO()
"""
print_settings = """    # --- PARAMÈTRES D'IMPRESSION ---
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.print_area = f\"A1:E{max_row}\"

    # --- RETOUR DES DONNÉES ---
    output = io.BytesIO()
"""
if print_anchor not in source:
    raise RuntimeError("Impossible de trouver la section de sauvegarde Excel.")
source = source.replace(print_anchor, print_settings, 1)

# Remplacer l'interface Deputy/Emprez par une interface Emprez seulement.
start_marker = "# --- INTERFACE PRINCIPALE : CHOIX DE LA SOURCE ---"
end_marker = "# --- TÉLÉCHARGEMENT ---"
start = source.index(start_marker)
end = source.index(end_marker)

emprez_interface = '''# --- INTERFACE PRINCIPALE : IMPORT EMPREZ ---
df = None
date_obj = None
date_formatee = None

st.markdown(
    "Déposez l'export **Emprez** (Excel d'une semaine), puis choisissez la **journée** "
    "à générer (un seul jour est produit à la fois)."
)
uploaded_file = st.file_uploader("Choisir le fichier Excel", type=["xls", "xlsx"])
if uploaded_file is not None:
    try:
        xls = lire_excel(uploaded_file)
        sheet_name = xls.sheet_names[0]
        df_raw = xls.parse(sheet_name, header=None)
    except Exception:
        st.error("Impossible de lire le fichier Excel Emprez.")
        st.stop()

    m = re.search(r'(\\d{4})-(\\d{2})-(\\d{2})', str(sheet_name))
    if m:
        lundi = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    else:
        today = datetime.today()
        lundi = today - timedelta(days=today.weekday())

    jours_dispo = [lundi + timedelta(days=i) for i in range(7)]
    labels = [formater_date_fr(d) for d in jours_dispo]
    choix = st.selectbox("📅 Choisissez la journée à générer :", labels)
    idx_jour = labels.index(choix)
    date_obj = jours_dispo[idx_jour]
    date_formatee = labels[idx_jour]
    jour_col = idx_jour + 1

    df = charger_emprez(df_raw, jour_col, date_obj.strftime('%Y-%m-%d'))
    if df.empty:
        st.warning("Aucun quart trouvé pour cette journée.")
        df = None
    else:
        st.success(f"{len(df)} employé(s) trouvé(s) pour le : {date_formatee}")

'''

source = source[:start] + emprez_interface + source[end:]
exec(compile(source, SOURCE_URL, "exec"), globals(), globals())

import sys
# Hack de compatibilité imghdr
try:
    import imghdr
except ImportError:
    import types
    m = types.ModuleType("imghdr")
    m.what = lambda x, h=None: None
    sys.modules["imghdr"] = m

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import random
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURATION STREAMLIT ---
st.set_page_config(page_title="Générateur d'Horaire Spa", layout="centered")
st.title("📅 Générateur d'Horaire Réception")

def extraire_prenom(nom_complet):
    if pd.isna(nom_complet): return "N/A"
    return str(nom_complet).split(' ')[0]

def attribuer_taches_recep(nb_staff):
    taches = ["Bye Bye Clés"]
    pool = ["Objet perdus", "Aide entre dept.", "Nettoyage boutique"]
    random.shuffle(pool)
    while len(taches) < nb_staff: taches.extend(pool)
    return taches[:nb_staff]

uploaded_file = st.file_uploader("Glissez le fichier Deputy CSV ici", type="csv")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    df['End Time'] = df['End Time'].fillna('')
    df['Note'] = df['Note'].fillna('')
    
    date_brute = df['Start Date'].iloc[0]
    date_obj = datetime.strptime(date_brute, '%Y-%m-%d')
    jours_fr = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    mois_fr = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    date_txt = f"{jours_fr[date_obj.weekday()]} {date_obj.day} {mois_fr[date_obj.month - 1]} {date_obj.year}"

    st.success(f"✅ Prêt pour le {date_txt}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Horaire"

    # --- DÉFINITION DES STYLES ---
    dark_gray = PatternFill(start_color="7F7F7F", end_color="7F7F7F", fill_type="solid")
    light_gray = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    border_thick_bottom = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thick'))
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. EN-TÊTE PRINCIPALE
    ws.merge_cells('A1:E1')
    ws['A1'] = f"Réception | {date_txt}"
    ws['A1'].font = Font(color="FFFFFF", bold=True, size=18)
    ws['A1'].fill = dark_gray
    ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 40

    curr_row = 2
    pauses_attribuees = {}

    def calculer_pause(dept, h_debut, total_h):
        try:
            if float(total_h) < 5.5: return "15 min"
            base = datetime.strptime(h_debut, "%H:%M") + timedelta(hours=4)
            if dept not in pauses_attribuees: pauses_attribuees[dept] = []
            while any(abs((base - p).total_seconds()) < 1800 for p in pauses_attribuees[dept]):
                base += timedelta(minutes=30)
            pauses_attribuees[dept].append(base)
            return base.strftime("%H:%M")
        except: return "-"

    # 2. SECTION SUPERVISEURS
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
    cell_sup = ws.cell(row=curr_row, column=1, value="SUPERVISEURS ET ADJOINTS")
    cell_sup.fill = light_gray
    cell_sup.font = Font(bold=True, size=14)
    cell_sup.alignment = align_center
    ws.row_dimensions[curr_row].height = 25
    curr_row += 1

    ordres_sup = [("Réception", "Réception - Supervision"), ("Bistro", "Bistro - Supervision"), 
                  ("Opérations", "Site extérieur - Supervision"), ("Entretien", "ENTRETIEN MÉNAGER"),
                  ("Soins", "MASSO"), ("Maintenance", "Maintenance- spa")]

    for label, search in ordres_sup:
        found = df[df['Area'].str.contains(search, case=False, na=False)]
        if label == "Maintenance":
            found = found[df['Note'].str.contains('Sur Appel', case=False, na=False)]
            if found.empty: found = pd.DataFrame([{'Employee': 'Adam', 'Start Time': 'Sur Appel', 'End Time': '', 'Total Time': 0}])
        else:
            found = found[(df['Area'].str.contains('Supervision|Responsable|Chef', case=False) | 
                           df['Note'].str.contains('Responsable|Supervision', case=False, na=False))]
        
        if not found.empty:
            start_m = curr_row
            for i, (_, r) in enumerate(found.iterrows()):
                ws.cell(row=curr_row, column=1, value=label.upper()).font = Font(bold=True)
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=5)
                h_end = str(r['End Time']).strip()
                h_info = f" ({r['Start Time']} - {h_end})" if h_end and h_end != 'nan' else f" ({r['Start Time']})"
                ws.cell(row=curr_row, column=2, value=f"{extraire_prenom(r['Employee'])}{h_info}").alignment = Alignment(horizontal="left", vertical="center")
                for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thin
                curr_row += 1
            if curr_row - start_m > 1:
                ws.merge_cells(start_row=start_m, start_column=1, end_row=curr_row-1, end_column=1)
                ws.cell(row=start_m, column=1).alignment = align_center

    # 3. SECTION RÉCEPTION
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
    cell_rec = ws.cell(row=curr_row, column=1, value="RÉCEPTION")
    cell_rec.fill = light_gray
    cell_rec.font = Font(bold=True, size=14)
    cell_rec.alignment = align_center
    ws.row_dimensions[curr_row].height = 25
    curr_row += 1

    headers = ["NOM", "QUART / POSTE", "Caisse", "Lunch", "TÂCHE / RESP."]
    for i, h in enumerate(headers, 1):
        ws.cell(row=curr_row, column=i, value=h).font = Font(bold=True)
        ws.cell(row=curr_row, column=i).border = border_thin
        ws.cell(row=curr_row, column=i).alignment = align_center
    curr_row += 1

    # Logique Postes intelligente
    rm = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False) & (df['Start Time'] < '12:00')].sort_values('Start Time')
    qbe = df[df['Area'].str.contains('QUALITÉ ET BIEN ÊTRE', case=False)].sort_values('Start Time')
    rs = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False) & (df['Start Time'] >= '12:00')].sort_values('Start Time')
    
    reception_staff = pd.concat([rm, rs]).sort_values(['Start Time', 'End Time'])
    postes_occupes, assign_postes = {}, {}
    ordre_pref = [4, 1, 3, 2, 5]

    for idx, emp in reception_staff.iterrows():
        h_deb = emp['Start Time']
        for p in [p for p, fin in postes_occupes.items() if fin <= h_deb]: del postes_occupes[p]
        p_trouve = next((p for p in ordre_pref if p not in postes_occupes), None)
        if p_trouve is None: p_trouve = min(postes_occupes, key=postes_occupes.get)
        assign_postes[idx] = f"Poste {p_trouve}"
        postes_occupes[p_trouve] = emp['End Time']

    caisse_n = 1
    for group, is_qbe in [(rm, False), (qbe, True), (rs, False)]:
        t_p = attribuer_taches_recep(len(group))
        for i, (idx, r) in enumerate(group.iterrows()):
            ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
            pos = "QBE" if is_qbe else assign_postes.get(idx, "")
            ws.cell(row=curr_row, column=2, value=f"{r['Start Time']}-{r['End Time']} / {pos}").font = Font(bold=True)
            ws.cell(row=curr_row, column=3, value=("VAM" if i == 0 else "VPM") if is_qbe else caisse_n)
            if not is_qbe: caisse_n += 1
            ws.cell(row=curr_row, column=4, value=calculer_pause("REC", r['Start Time'], r['Total Time']))
            ws.cell(row=curr_row, column=5, value="QBE" if is_qbe else t_p[i])
            for c in range(1, 6):
                ws.cell(row=curr_row, column=c).border = border_thin
                ws.cell(row=curr_row, column=c).alignment = align_center
                if is_qbe:
                    ws.cell(row=curr_row, column=c).fill = dark_gray
                    ws.cell(row=curr_row, column=c).font = Font(color="FFFFFF", bold=True)
            curr_row += 1

    # 4. SECTION LOUNGE (si existe)
    lg_data = df[df['Area'].str.contains('Lounge', case=False)].sort_values('Start Time')
    if not lg_data.empty:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        cell_lg = ws.cell(row=curr_row, column=1, value="LOUNGE")
        cell_lg.fill = light_gray
        cell_lg.font = Font(bold=True, size=14)
        cell_lg.alignment = align_center
        curr_row += 1
        for i, (_, r) in enumerate(lg_data.iterrows()):
            ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
            ws.cell(row=curr_row, column=2, value=f"{r['Start Time']}-{r['End Time']}").font = Font(bold=True)
            t_lg = "Ouverture" if i == 0 else "Fermeture" if i == len(lg_data)-1 else "Accueil"
            ws.cell(row=curr_row, column=5, value=t_lg)
            if t_lg == "Accueil": ws.cell(row=curr_row, column=3).fill = black_fill
            else: ws.cell(row=curr_row, column=3, value=caisse_n); caisse_n += 1
            ws.cell(row=curr_row, column=4, value=calculer_pause("LG", r['Start Time'], r['Total Time']))
            for c in range(1, 6):
                ws.cell(row=curr_row, column=c).border = border_thin
                ws.cell(row=curr_row, column=c).alignment = align_center
            curr_row += 1

    # 5. CERCLES FINAUX
    for t in ["VENTES FIDÉLITÉS", "SOINS"]:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        ws.cell(row=curr_row, column=1, value=t).fill = light_gray
        ws.cell(row=curr_row, column=1).font = Font(bold=True)
        ws.cell(row=curr_row, column=1).alignment = align_center
        curr_row += 1
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        ws.cell(row=curr_row, column=1, value="○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○ ○").font = Font(size=20)
        ws.cell(row=curr_row, column=1).alignment = align_center
        ws.row_dimensions[curr_row].height = 40
        for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thin
        curr_row += 1

    # --- AJUSTEMENT LARGEUR COLONNES ---
    ws.column_dimensions['A'].width = 18 # Nom
    ws.column_dimensions['B'].width = 28 # Quart / Poste
    ws.column_dimensions['C'].width = 12 # Caisse
    ws.column_dimensions['D'].width = 12 # Lunch
    ws.column_dimensions['E'].width = 25 # Tâche

    # --- TÉLÉCHARGEMENT ---
    output = io.BytesIO()
    wb.save(output)
    st.download_button(label="📥 Télécharger l'Excel formaté", data=output.getvalue(), 
                       file_name=f"Horaire_{date_obj.strftime('%Y-%m-%d')}.xlsx", 
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

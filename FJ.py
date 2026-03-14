import sys
# Hack de compatibilité pour Python 3.11+ sur Streamlit Cloud
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

# --- CONFIGURATION INTERFACE ---
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

# --- CHARGEMENT DU FICHIER ---
uploaded_file = st.file_uploader("Glissez le fichier Deputy CSV ici", type="csv")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    df['End Time'] = df['End Time'].fillna('')
    df['Note'] = df['Note'].fillna('')
    
    # Traitement de la date
    date_brute = df['Start Date'].iloc[0]
    date_obj = datetime.strptime(date_brute, '%Y-%m-%d')
    jours_fr = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    mois_fr = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    date_txt = f"{jours_fr[date_obj.weekday()]} {date_obj.day} {mois_fr[date_obj.month - 1]} {date_obj.year}"

    st.success(f"✅ Fichier chargé : {date_txt}")

    # --- CRÉATION DU WORKBOOK ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Horaire"

    # Styles
    font_titre_principal = Font(color="FFFFFF", bold=True, size=18)
    font_titre_secondaire = Font(color="000000", bold=True, size=16)
    font_normal = Font(color="000000", size=11)
    font_normal_bold = Font(color="000000", bold=True, size=11)
    font_cercles = Font(size=24) 
    font_qbe = Font(color="FFFFFF", bold=True, size=11)

    dark_gray_fill = PatternFill(start_color="7F7F7F", end_color="7F7F7F", fill_type="solid")
    light_gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

    thin = Side(style='thin'); thick = Side(style='thick')
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_thick_bottom = Border(left=thin, right=thin, top=thin, bottom=thick)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    pauses_attribuees = {} 

    def calculer_pause_unique(dept, heure_debut_str, duree_heures):
        try:
            val = float(duree_heures)
            if val < 5.5: return "15 min"
            fmt = "%H:%M"
            base_pause = datetime.strptime(heure_debut_str, fmt) + timedelta(hours=4)
            if dept not in pauses_attribuees: pauses_attribuees[dept] = []
            while any(abs((base_pause - p).total_seconds()) < 1800 for p in pauses_attribuees[dept]):
                base_pause += timedelta(minutes=30) 
            pauses_attribuees[dept].append(base_pause)
            return base_pause.strftime(fmt)
        except: return "-"

    # 1. Titre
    ws.merge_cells('A1:E1')
    ws['A1'] = f"Réception | {date_txt}"
    ws['A1'].font = font_titre_principal
    ws['A1'].fill = dark_gray_fill
    ws['A1'].alignment = center_align
    ws.row_dimensions[1].height = 45

    curr_row = 2

    # 2. Section Superviseurs
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
    ws.cell(row=curr_row, column=1, value="SUPERVISEURS ET ADJOINTS").fill = light_gray_fill
    ws.cell(row=curr_row, column=1).font = font_titre_secondaire
    ws.cell(row=curr_row, column=1).alignment = center_align
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
                ws.cell(row=curr_row, column=2, value=f"{extraire_prenom(r['Employee'])}{h_info}")
                curr_row += 1
            if curr_row - start_m > 1: ws.merge_cells(start_row=start_m, start_column=1, end_row=curr_row-1, end_column=1)

    # 3. Sections Employés (Réception & Lounge)
    caisse_num = 1
    lg_data = df[df['Area'].str.contains('Lounge', case=False)].sort_values('Start Time')
    sections_a_afficher = ["RÉCEPTION"]
    if not lg_data.empty: sections_a_afficher.append("LOUNGE")

    for section in sections_a_afficher:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        ws.cell(row=curr_row, column=1, value=section).fill = light_gray_fill
        ws.cell(row=curr_row, column=1).font = font_titre_secondaire
        curr_row += 1
        headers = ["NOM", "QUART / POSTE", "Caisse", "Lunch", "TÂCHE / RESP."]
        for i, h in enumerate(headers, 1):
            ws.cell(row=curr_row, column=i, value=h).font = font_normal_bold
            ws.cell(row=curr_row, column=i).border = border_thin
        curr_row += 1

        if section == "RÉCEPTION":
            rm = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False) & (df['Start Time'] < '12:00')].sort_values('Start Time')
            qbe = df[df['Area'].str.contains('QUALITÉ ET BIEN ÊTRE', case=False)].sort_values('Start Time')
            rs = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False) & (df['Start Time'] >= '12:00')].sort_values('Start Time')
            
            reception_staff = pd.concat([rm, rs]).sort_values(['Start Time', 'End Time'])
            postes_occupes, assignations_finales = {}, {}
            ordre_prefere = [4, 1, 3, 2, 5]

            for idx, emp in reception_staff.iterrows():
                h_deb = emp['Start Time']
                liberes = [p for p, fin in postes_occupes.items() if fin <= h_deb]
                for p in liberes: del postes_occupes[p]
                
                poste_trouve = next((p for p in ordre_prefere if p not in postes_occupes), None)
                if poste_trouve is None: poste_trouve = min(postes_occupes, key=postes_occupes.get)
                
                assignations_finales[idx] = f"Poste {poste_trouve}"
                postes_occupes[poste_trouve] = emp['End Time']

            for group, is_qbe in [(rm, False), (qbe, True), (rs, False)]:
                t_p = attribuer_taches_recep(len(group))
                for i, (idx, r) in enumerate(group.iterrows()):
                    ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
                    pos = "QBE" if is_qbe else assignations_finales.get(idx, "")
                    h_end = str(r['End Time']).strip() if str(r['End Time']).strip() != 'nan' else ""
                    ws.cell(row=curr_row, column=2, value=f"{r['Start Time']}-{h_end} / {pos}").font = font_normal_bold
                    ws.cell(row=curr_row, column=3, value=("VAM" if i == 0 else "VPM") if is_qbe else caisse_num)
                    if not is_qbe: caisse_num += 1
                    ws.cell(row=curr_row, column=4, value=calculer_pause_unique("REC", r['Start Time'], r['Total Time']))
                    ws.cell(row=curr_row, column=5, value="QBE" if is_qbe else t_p[i])
                    for c in range(1, 6):
                        cell = ws.cell(row=curr_row, column=c)
                        cell.border = border_thin; cell.font = font_normal
                        if is_qbe: cell.fill = dark_gray_fill; cell.font = font_qbe
                    curr_row += 1
            for c in range(1, 6): ws.cell(row=curr_row-1, column=c).border = border_thick_bottom

        elif section == "LOUNGE":
            lg_list = list(lg_data.iterrows())
            for i, (_, r) in enumerate(lg_list):
                ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
                h_e = str(r['End Time']).strip() if str(r['End Time']).strip() != 'nan' else ""
                ws.cell(row=curr_row, column=2, value=f"{r['Start Time']}-{h_e}").font = font_normal_bold
                t = "Ouverture" if i == 0 else "Fermeture" if i == len(lg_list)-1 else "Accueil" if i == 1 else ""
                ws.cell(row=curr_row, column=5, value=t)
                if t == "Accueil": ws.cell(row=curr_row, column=3).fill = black_fill
                else: ws.cell(row=curr_row, column=3, value=caisse_num); caisse_num += 1
                ws.cell(row=curr_row, column=4, value=calculer_pause_unique("LG", r['Start Time'], r['Total Time']))
                for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thin; ws.cell(row=curr_row, column=c).font = font_normal
                curr_row += 1
            for c in range(1, 6): ws.cell(row=curr_row-1, column=c).border = border_thick_bottom

    # 4. Section Finale (Cercles)
    for titre in ["VENTES FIDÉLITÉS", "SOINS", "BONNE JOURNÉE !"]:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        ws.cell(row=curr_row, column=1, value=titre).fill = light_gray_fill
        ws.cell(row=curr_row, column=1).font = font_titre_secondaire
        curr_row += 1
        if titre != "BONNE JOURNÉE !":
            ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
            ws.cell(row=curr_row, column=1, value="○" * 18).font = font_cercles
            ws.cell(row=curr_row, column=1).alignment = Alignment(horizontal="center")
            ws.row_dimensions[curr_row].height = 45
            for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thin
            curr_row += 1

    # Auto-size colonnes
    for i in range(1, 6):
        ws.column_dimensions[get_column_letter(i)].width = 25

    # --- TÉLÉCHARGEMENT ---
    output = io.BytesIO()
    wb.save(output)
    st.download_button(label="📥 Télécharger l'horaire Excel", data=output.getvalue(), 
                       file_name=f"FJ_{date_obj.strftime('%Y%m%d')}.xlsx", 
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

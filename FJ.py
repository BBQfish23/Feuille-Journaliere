import sys
# Hack de compatibilité pour Python 3.13+ (Streamlit Cloud)
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

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur d'Horaire Spa", layout="centered")

st.title("📅 Générateur d'Horaire Réception")
st.markdown("Déposez votre export **Deputy CSV** pour générer la feuille journalière formatée.")

# --- SÉLECTION DU THÈME ---
liste_themes = [
    "Standard", 
    "Printemps", "Été", "Automne", "Hiver",
    "Nouvel An", "Saint-Valentin", "Pâques", "Fête des Mères", "Fête des Pères", 
    "Saint-Jean-Baptiste", "Fête du Canada", "Action de grâce", "Halloween", "Temps des Fêtes",
    "Semaine de Relâche", "Jour de Pluie (Cocooning)", "Canicule", "Grosse Journée", "Événement VIP"
]

theme_choisi = st.selectbox("🎭 Choisissez un thème pour l'horaire :", liste_themes)

# Dictionnaire complet des thèmes
themes_config = {
    "Standard": {"emoji": "", "msg_fin": "BONNE JOURNÉE !", "font": "Calibri"},
    "Printemps": {"emoji": "🌱", "msg_fin": "BONNE JOURNÉE ! 🌷", "font": "Comic Sans MS"},
    "Été": {"emoji": "☀️", "msg_fin": "BONNE JOURNÉE SOUS LE SOLEIL ! 🕶️", "font": "Trebuchet MS"},
    "Automne": {"emoji": "🍂", "msg_fin": "BEL AUTOMNE ET BONNE JOURNÉE ! 🍁", "font": "Georgia"},
    "Hiver": {"emoji": "❄️", "msg_fin": "BONNE JOURNÉE ! ⛄", "font": "Century Gothic"},
    "Nouvel An": {"emoji": "🎉", "msg_fin": "BONNE ANNÉE ! 🥂", "font": "Trebuchet MS"},
    "Saint-Valentin": {"emoji": "🤍", "msg_fin": "JOYEUSE SAINT-VALENTIN ! 🕊️", "font": "Georgia"},
    "Pâques": {"emoji": "🐇", "msg_fin": "JOYEUSES PÂQUES ! 🥚", "font": "Comic Sans MS"},
    "Fête des Mères": {"emoji": "💐", "msg_fin": "BONNE FÊTE DES MÈRES ! 🌸", "font": "Georgia"},
    "Fête des Pères": {"emoji": "☕", "msg_fin": "BONNE FÊTE DES PÈRES ! 👔", "font": "Times New Roman"},
    "Saint-Jean-Baptiste": {"emoji": "⚜️", "msg_fin": "BONNE FÊTE NATIONALE ! ⚜️", "font": "Impact"},
    "Fête du Canada": {"emoji": "🍁", "msg_fin": "BONNE FÊTE DU CANADA ! 🎆", "font": "Impact"},
    "Action de grâce": {"emoji": "🦃", "msg_fin": "JOYEUSE ACTION DE GRÂCE ! 🍂", "font": "Georgia"},
    "Halloween": {"emoji": "👻", "msg_fin": "JOYEUSE HALLOWEEN ! 🎃", "font": "Impact"},
    "Temps des Fêtes": {"emoji": "🎄", "msg_fin": "JOYEUSES FÊTES ! ✨", "font": "Courier New"},
    "Semaine de Relâche": {"emoji": "⛷️", "msg_fin": "BONNE RELÂCHE ! ☕", "font": "Comic Sans MS"},
    "Jour de Pluie (Cocooning)": {"emoji": "🌧️", "msg_fin": "BONNE JOURNÉE COCOONING ! 🍵", "font": "Courier New"},
    "Canicule": {"emoji": "🌡️", "msg_fin": "RESTEZ AU FRAIS ! 🧊", "font": "Arial Black"},
    "Grosse Journée": {"emoji": "💪", "msg_fin": "EXCELLENTE JOURNÉE, ON LÂCHE PAS ! 🔥", "font": "Arial Black"},
    "Événement VIP": {"emoji": "⭐", "msg_fin": "EXCELLENTE JOURNÉE VIP ! 🥂", "font": "Times New Roman"}
}

theme_actuel = themes_config[theme_choisi]

# --- FONCTIONS UTILITAIRES ---
def extraire_prenom(nom_complet):
    if pd.isna(nom_complet): return "N/A"
    return str(nom_complet).split(' ')[0]

def str_to_minutes(t_str):
    try:
        h, m = map(int, str(t_str).strip().split(':'))
        return h * 60 + m
    except:
        return 1440 

# --- INTERFACE DE CHARGEMENT ---
uploaded_file = st.file_uploader("Choisir le fichier CSV", type="csv")

if uploaded_file is not None:
    # 1. BLINDAGE DU CHARGEMENT DES DONNÉES
    df = pd.read_csv(uploaded_file)
    df.columns = df.columns.str.strip() # Enlève les espaces invisibles
    
    # Remplissage sécuritaire des valeurs vides (NaN)
    for col in ['Start Time', 'End Time', 'Note', 'Area', 'Employee']:
        if col in df.columns:
            df[col] = df[col].fillna('')

    try:
        date_brute = df['Start Date'].iloc[0]
        date_obj = datetime.strptime(str(date_brute).strip(), '%Y-%m-%d')
        jours_fr = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
        mois_fr = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
        date_formatee = f"{jours_fr[date_obj.weekday()]} {date_obj.day} {mois_fr[date_obj.month - 1]} {date_obj.year}"
        st.success(f"Fichier détecté pour le : {date_formatee}")
    except Exception as e:
        st.error("Erreur avec la date. Vérifiez le format du fichier CSV.")
        st.stop()

    # --- CRÉATION DU EXCEL ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Horaire Spa"

    # Styles
    font_titre_secondaire = Font(color="000000", bold=True, size=16)
    font_superviseur_dept = Font(color="000000", bold=True, size=12)
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
    center_no_wrap = Alignment(horizontal="center", vertical="center", wrap_text=False)
    border_thick_all = Border(left=thick, right=thick, top=thick, bottom=thick)

    # --- LOGIQUE DE PAUSES GLOBALES ---
    pauses_globales = set() 

    def calculer_pause_unique(is_qbe, heure_debut_str, duree_heures):
        try:
            val = float(duree_heures)
            if val < 5.5: return "15 min"
            fmt = "%H:%M"
            
            # Base 3 heures après l'arrivée
            base = datetime.strptime(str(heure_debut_str).strip(), fmt) + timedelta(hours=3)
            # Pas de pause avant 11h30
            earliest = datetime.strptime("11:30", fmt)
            current_pause = max(base, earliest)
            
            # Arrondi à 30 min supérieures
            minute = current_pause.minute
            if 0 < minute <= 30:
                current_pause = current_pause.replace(minute=30)
            elif minute > 30:
                current_pause = current_pause + timedelta(hours=1)
                current_pause = current_pause.replace(minute=0)
                
            interdits_globaux = ["18:00", "18:30"]
            interdits_qbe = ["10:00", "11:30", "13:00", "14:30", "16:00", "17:30", "19:00", "20:30"]
            
            while True:
                str_pause = current_pause.strftime(fmt)
                if str_pause in interdits_globaux:
                    pass 
                elif is_qbe and str_pause in interdits_qbe:
                    pass 
                elif str_pause in pauses_globales:
                    pass 
                else:
                    pauses_globales.add(str_pause)
                    return str_pause
                
                current_pause += timedelta(minutes=30)
                if current_pause.hour < 6: return "-" # Sécurité anti-boucle infinie
        except: 
            return "-"

    # 1. Titre Principal
    ws.merge_cells('A1:E1')
    ws.row_dimensions[1].height = 45
    
    titre_texte = f"Département Réception | {date_formatee}"
    if theme_actuel["emoji"]:
        titre_texte = f"{theme_actuel['emoji']} {titre_texte} {theme_actuel['emoji']}"
        
    ws['A1'] = titre_texte
    ws['A1'].fill = dark_gray_fill
    ws['A1'].font = Font(name=theme_actuel["font"], color="FFFFFF", bold=True, size=18)
    ws['A1'].alignment = center_align
    for c in range(1, 6): ws.cell(row=1, column=c).border = border_thick_all

    curr_row = 2

    # 2. Section SUPERVISEURS
    ws.row_dimensions[curr_row].height = 30
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
    ws.cell(row=curr_row, column=1, value="SUPERVISEURS ET ADJOINTS").fill = light_gray_fill
    ws.cell(row=curr_row, column=1).font = font_titre_secondaire
    ws.cell(row=curr_row, column=1).alignment = center_align
    for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thick_all
    curr_row += 1

    ordres_sup = [
        ("Réception", "Réception - Supervision"), ("Bistro", "Bistro - Supervision"), 
        ("Opérations", "Site extérieur - Supervision"), ("Entretien", "ENTRETIEN MÉNAGER"),
        ("Soins", "MASSO"), ("Maintenance", "Maintenance- spa")
    ]

    for label, search in ordres_sup:
        found = df[df['Area'].str.contains(search, case=False, na=False)]
        
        if label == "Maintenance":
            found = found[found['Note'].str.contains('Sur Appel', case=False, na=False)]
            if found.empty:
                found = pd.DataFrame([{'Employee': 'Adam', 'Start Time': 'Sur Appel', 'End Time': '', 'Total Time': 0}])
        else:
            found = found[(found['Area'].str.contains('Supervision|Responsable|Chef', case=False, na=False) | 
                           found['Note'].str.contains('Responsable|Supervision', case=False, na=False))]

        if not found.empty:
            start_m = curr_row
            items = found.sort_values('Start Time') if 'Start Time' in found.columns else found
            total_items = len(items) 
            
            for i, (_, r) in enumerate(items.iterrows()):
                is_first = (i == 0)
                is_last = (i == total_items - 1)
                
                cell_label = ws.cell(row=curr_row, column=1, value=label.upper())
                cell_label.font = font_superviseur_dept
                cell_label.border = Border(
                    left=thick, right=thick, 
                    top=thick if is_first else thin, bottom=thick if is_last else thin
                )
                
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=5)
                h_end = str(r['End Time']).strip()
                h_info = f" ({r['Start Time']} - {h_end})" if h_end and h_end != 'nan' and h_end != '' else f" ({r['Start Time']})"
                txt = f"{extraire_prenom(r['Employee'])}{h_info}"
                
                for c in range(2, 6):
                    cell = ws.cell(row=curr_row, column=c)
                    cell.font = font_normal
                    if c == 2: cell.value = txt
                    
                    cell.border = Border(
                        left=thick if c == 2 else thin, right=thick if c == 5 else thin, 
                        top=thick if is_first else thin, bottom=thick if is_last else thin
                    )
                curr_row += 1
                
            if total_items > 1:
                ws.merge_cells(start_row=start_m, start_column=1, end_row=curr_row-1, end_column=1)

    # --- PRÉPARATION GLOBALE DES CAISSES ---
    rm = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False, na=False) & (df['Start Time'].astype(str) < '12:00')].sort_values('Start Time')
    qbe = df[df['Area'].str.contains('QUALITÉ ET BIEN ÊTRE', case=False, na=False)].sort_values('Start Time')
    rs = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False, na=False) & (df['Start Time'].astype(str) >= '12:00')].sort_values('Start Time')
    
    lg_data = df[df['Area'].str.contains('Lounge', case=False, na=False)].sort_values('Start Time')
    lg_list = list(lg_data.iterrows())
    total_lg = len(lg_list)

    candidats_caisses = []
    
    for _, r in pd.concat([rm, rs]).iterrows():
        candidats_caisses.append((r['Employee'], str(r['Start Time'])))
        
    for _, r in lg_list:
        note_str = str(r['Note']).lower()
        if 'accueil' not in note_str and 'acceuil' not in note_str:
            candidats_caisses.append((r['Employee'], str(r['Start Time'])))

    candidats_caisses.sort(key=lambda x: str_to_minutes(x[1]))

    map_caisses = {}
    for idx, candidat in enumerate(candidats_caisses):
        map_caisses[candidat] = idx + 1

    # 3. SECTIONS EMPLOYÉS
    sections_a_afficher = ["RÉCEPTION"]
    if not lg_data.empty: sections_a_afficher.append("LOUNGE")

    for section in sections_a_afficher:
        ws.row_dimensions[curr_row].height = 30
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        ws.cell(row=curr_row, column=1, value=section).fill = light_gray_fill
        ws.cell(row=curr_row, column=1).font = font_titre_secondaire
        for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thick_all
        curr_row += 1
        
        headers = {1: "NOM", 2: "QUART / POSTE", 3: "Caisse", 4: "Lunch", 5: "TÂCHE / RESP."}
        for c in range(1, 6):
            cell = ws.cell(row=curr_row, column=c)
            cell.font = font_normal_bold
            cell.border = border_thin
            if c in headers: cell.value = headers[c]
        curr_row += 1

        if section == "RÉCEPTION":
            # Pool de tâches agrandi (* 50) 
            pool_taches = ["Nettoyage boutique"] + ["Objet perdus", "Aide entre dept."] * 50
            random.shuffle(pool_taches) 
            
            ordre_postes = ["Poste 4", "Poste 1", "Poste 3", "Poste 2", "Poste 5"]
            fin_postes = {p: "00:00" for p in ordre_postes} 

            for group, is_qbe in [(rm, False), (qbe, True), (rs, False)]:
                for i, (_, r) in enumerate(group.iterrows()):
                    ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
                    
                    h_end = str(r['End Time']).strip()
                    h_end = h_end if h_end != 'nan' and h_end != '' else ""
                    
                    if not is_qbe:
                        start_min = str_to_minutes(r['Start Time'])
                        postes_libres = [p for p in ordre_postes if str_to_minutes(fin_postes[p]) <= start_min]
                        
                        if postes_libres:
                            poste_attribue = postes_libres[0]
                        else:
                            poste_attribue = min(ordre_postes, key=lambda p: (str_to_minutes(fin_postes[p]), ordre_postes.index(p)))
                        
                        fin_postes[poste_attribue] = h_end if h_end else "23:59"
                        quart_txt = f"{str(r['Start Time']).strip()}-{h_end} / {poste_attribue}"
                        
                        if poste_attribue == "Poste 1":
                            tache_finale = "Bye Bye Clés"
                        else:
                            tache_finale = pool_taches.pop(0)
                    else:
                        quart_txt = f"{str(r['Start Time']).strip()}-{h_end}"
                        tache_finale = "QBE"
                        
                    ws.cell(row=curr_row, column=2, value=quart_txt).font = font_normal_bold
                    
                    if is_qbe:
                        c_val = "VAM" if i == 0 else "VPM"
                    else:
                        c_val = map_caisses.get((r['Employee'], str(r['Start Time'])), "")
                        
                    ws.cell(row=curr_row, column=3, value=c_val)
                    ws.cell(row=curr_row, column=4, value=calculer_pause_unique(is_qbe, str(r['Start Time']), r['Total Time']))
                    ws.cell(row=curr_row, column=5, value=tache_finale)
                    
                    for c in range(1, 6):
                        cell = ws.cell(row=curr_row, column=c)
                        cell.border = border_thin
                        cell.font = font_normal
                        if is_qbe: 
                            cell.fill = dark_gray_fill
                            cell.font = font_qbe
                    curr_row += 1
            
            for c in range(1, 6): 
                ws.cell(row=curr_row-1, column=c).border = border_thick_bottom
                
        elif section == "LOUNGE":
            for i, (_, r) in enumerate(lg_list):
                ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
                h_end_lg = str(r['End Time']).strip() if str(r['End Time']).strip() != 'nan' else ""
                ws.cell(row=curr_row, column=2, value=f"{str(r['Start Time']).strip()}-{h_end_lg}").font = font_normal_bold
                
                note_str = str(r['Note']).lower()
                is_accueil = 'accueil' in note_str or 'acceuil' in note_str
                
                if is_accueil:
                    t = "Accueil"
                elif i == 0:
                    t = "Ouverture"
                elif i == total_lg - 1:
                    t = "Fermeture"
                else:
                    t = ""
                    
                ws.cell(row=curr_row, column=5, value=t)
                
                cell_c = ws.cell(row=curr_row, column=3)
                if is_accueil: 
                    cell_c.fill = black_fill
                else: 
                    cell_c.value = map_caisses.get((r['Employee'], str(r['Start Time'])), "")
                    
                ws.cell(row=curr_row, column=4, value=calculer_pause_unique(False, str(r['Start Time']), r['Total Time']))
                
                for c in range(1, 6): 
                    ws.cell(row=curr_row, column=c).border = border_thin
                    ws.cell(row=curr_row, column=c).font = font_normal
                curr_row += 1
                
            for c in range(1, 6): 
                ws.cell(row=curr_row-1, column=c).border = border_thick_bottom

    # 4. SECTION FINALE
    cercles_18_colles = "○" * 18
    
    police_fin_theme = Font(name=theme_actuel["font"], color="000000", bold=True, size=16)
    
    sections_finales = [
        ("VENTES FIDÉLITÉS", cercles_18_colles, font_titre_secondaire), 
        ("SOINS", cercles_18_colles, font_titre_secondaire), 
        (theme_actuel["msg_fin"], None, police_fin_theme)
    ]
    
    for titre, contenu, police_a_utiliser in sections_finales:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        cell_t = ws.cell(row=curr_row, column=1, value=titre)
        cell_t.fill = light_gray_fill
        cell_t.font = police_a_utiliser
        ws.row_dimensions[curr_row].height = 30
        
        for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thick_all
        curr_row += 1
        
        if contenu:
            ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
            cell_c = ws.cell(row=curr_row, column=1, value=contenu)
            cell_c.font = font_cercles
            cell_c.alignment = center_no_wrap
            ws.row_dimensions[curr_row].height = 45
            for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thin
            curr_row += 1

   # --- AUTO-SIZE ET BORDURES FINALES ---
    max_row = curr_row - 1
    for r in range(1, max_row + 1):
        for c in range(1, 6):
            cell = ws.cell(row=r, column=c)
            if not cell.alignment.horizontal: cell.alignment = center_align
            cb = cell.border
            cell.border = Border(left=(thick if c == 1 else cb.left), right=(thick if c == 5 else cb.right), 
                                top=(thick if r == 1 else cb.top), bottom=(thick if r == max_row else cb.bottom))

    # Ajustement des largeurs de colonnes
    for i in range(1, 6):
        col_letter = get_column_letter(i)
        
        if i == 1:
            ws.column_dimensions[col_letter].width = 25.5
        else:
            max_length = 0
            for row in ws.iter_rows(min_col=i, max_col=i):
                for cell in row:
                    if cell.coordinate not in ws.merged_cells or cell.column == 1:
                        try:
                            if cell.value:
                                length = len(str(cell.value))
                                if cell.font and cell.font.bold: length += 2
                                if length > max_length: max_length = length
                        except: pass
            ws.column_dimensions[col_letter].width = max_length + 3

    # --- TÉLÉCHARGEMENT ---
    output = io.BytesIO()
    wb.save(output)
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Télécharger l'horaire Excel",
        data=processed_data,
        file_name=f"Horaire_Spa_{date_obj.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

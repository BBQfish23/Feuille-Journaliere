import pandas as pd
from datetime import datetime, timedelta
import random
import glob
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- 1. CHARGEMENT DYNAMIQUE ET DATE ---
fichiers_csv = glob.glob('DeputyRosterDaily_*.csv')
if not fichiers_csv:
    print("Aucun fichier CSV trouvé.")
    exit()

fichier_recent = max(fichiers_csv, key=os.path.getctime)
df = pd.read_csv(fichier_recent)
df['End Time'] = df['End Time'].fillna('')
df['Note'] = df['Note'].fillna('')

date_brute = df['Start Date'].iloc[0]
date_obj = datetime.strptime(date_brute, '%Y-%m-%d')

jours_fr = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
mois_fr = ["janvier", "février", "mars", "avril", "mai", "juin", 
           "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
date_formatee = f"{jours_fr[date_obj.weekday()]} {date_obj.day} {mois_fr[date_obj.month - 1]} {date_obj.year}"

wb = Workbook()
ws = wb.active
ws.title = "Horaire Spa"

# --- STYLES ET TAILLES DE POLICE ---
font_titre_principal = Font(color="FFFFFF", bold=True, size=18)
font_titre_secondaire = Font(color="000000", bold=True, size=16)
font_superviseur_dept = Font(color="000000", bold=True, size=12)
font_normal = Font(color="000000", size=11)
font_normal_bold = Font(color="000000", bold=True, size=11)
font_cercles = Font(size=24) 
font_qbe = Font(color="FFFFFF", bold=True, size=11)

dark_gray_fill = PatternFill(start_color="7F7F7F", end_color="7F7F7F", fill_type="solid")
light_gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

thin = Side(style='thin')
thick = Side(style='thick')
border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
border_thick_bottom = Border(left=thin, right=thin, top=thin, bottom=thick)
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
center_no_wrap = Alignment(horizontal="center", vertical="center", wrap_text=False)

# --- SYSTÈME DE GESTION DES PAUSES ---
pauses_attribuees = {} 

def calculer_pause_unique(dept, heure_debut_str, duree_heures):
    try:
        val = float(duree_heures)
        if val < 5.5: return "15 min"
        fmt = "%H:%M"
        base_pause = datetime.strptime(heure_debut_str, fmt) + timedelta(hours=4)
        if dept not in pauses_attribuees:
            pauses_attribuees[dept] = []
        while any(abs((base_pause - p).total_seconds()) < 1800 for p in pauses_attribuees[dept]):
            base_pause += timedelta(minutes=30) 
        pauses_attribuees[dept].append(base_pause)
        return base_pause.strftime(fmt)
    except: return "-"

def extraire_prenom(nom_complet):
    if pd.isna(nom_complet): return "N/A"
    return str(nom_complet).split(' ')[0]

def attribuer_taches_recep(nb_staff):
    taches = ["Bye Bye Clés"]
    pool = ["Objet perdus", "Aide entre dept.", "Nettoyage boutique"]
    random.shuffle(pool)
    while len(taches) < nb_staff: taches.extend(pool)
    return taches[:nb_staff]

# --- CONSTRUCTION ---

# 1. Titre Principal
ws.merge_cells('A1:E1')
ws.row_dimensions[1].height = 45
ws['A1'] = f"Réception | {date_formatee}"
ws['A1'].fill = dark_gray_fill
ws['A1'].font = font_titre_principal
ws['A1'].alignment = center_align

curr_row = 2

# 2. Section SUPERVISEURS
ws.row_dimensions[curr_row].height = 30
ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
ws.cell(row=curr_row, column=1, value="SUPERVISEURS ET ADJOINTS").fill = light_gray_fill
ws.cell(row=curr_row, column=1).font = font_titre_secondaire
ws.cell(row=curr_row, column=1).alignment = center_align
curr_row += 1

ordres_sup = [
    ("Réception", "Réception - Supervision"), ("Bistro", "Bistro - Supervision"), 
    ("Opérations", "Site extérieur - Supervision"), ("Entretien", "ENTRETIEN MÉNAGER"),
    ("Soins", "MASSO"), ("Maintenance", "Maintenance- spa")
]

for label, search in ordres_sup:
    found = df[df['Area'].str.contains(search, case=False, na=False)]
    if label == "Maintenance":
        found = found[df['Note'].str.contains('Sur Appel', case=False, na=False)]
        if found.empty:
            found = pd.DataFrame([{'Employee': 'Adam', 'Start Time': 'Sur Appel', 'End Time': '', 'Total Time': 0}])
    else:
        found = found[(df['Area'].str.contains('Supervision|Responsable|Chef', case=False) | 
                       df['Note'].str.contains('Responsable|Supervision', case=False, na=False))]

    if not found.empty:
        start_m = curr_row
        items = found.sort_values('Start Time') if 'Start Time' in found.columns else found
        for i, (_, r) in enumerate(items.iterrows()):
            current_border = border_thick_bottom if i == len(items) - 1 else border_thin
            cell_label = ws.cell(row=curr_row, column=1, value=label.upper())
            cell_label.font = font_superviseur_dept
            cell_label.border = current_border
            
            ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=5)
            h_end = str(r['End Time']).strip()
            h_info = f" ({r['Start Time']} - {h_end})" if h_end and h_end != 'nan' and h_end != '' else f" ({r['Start Time']})"
            txt = f"{extraire_prenom(r['Employee'])}{h_info}"
            
            for c in range(2, 6):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = current_border
                cell.font = font_normal
                if c == 2: cell.value = txt
            curr_row += 1
        if curr_row - start_m > 1:
            ws.merge_cells(start_row=start_m, start_column=1, end_row=curr_row-1, end_column=1)

# 3. SECTIONS EMPLOYÉS
caisse_num = 1
lg_data = df[df['Area'].str.contains('Lounge', case=False)].sort_values('Start Time')
sections_a_afficher = ["RÉCEPTION"]
if not lg_data.empty: sections_a_afficher.append("LOUNGE")

for section in sections_a_afficher:
    ws.row_dimensions[curr_row].height = 30
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
    cell_sec = ws.cell(row=curr_row, column=1, value=section)
    cell_sec.fill = light_gray_fill
    cell_sec.font = font_titre_secondaire
    curr_row += 1
    
    headers = {1: "NOM", 2: "QUART / POSTE", 3: "Caisse", 4: "Lunch", 5: "TÂCHE / RESP."}
    for c in range(1, 6):
        cell = ws.cell(row=curr_row, column=c)
        cell.font = font_normal_bold
        cell.border = border_thin
        if c in headers: cell.value = headers[c]
    curr_row += 1

    if section == "RÉCEPTION":
        rm = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False) & (df['Start Time'] < '12:00')].sort_values('Start Time')
        qbe = df[df['Area'].str.contains('QUALITÉ ET BIEN ÊTRE', case=False)].sort_values('Start Time')
        rs = df[df['Area'].str.contains('RÉCEPTION- Responsable', case=False) & (df['Start Time'] >= '12:00')].sort_values('Start Time')
        cycle_p = ["Poste 4", "Poste 1", "Poste 3", "Poste 2", "Poste 5"]

        for group, is_qbe in [(rm, False), (qbe, True), (rs, False)]:
            t_p = attribuer_taches_recep(len(group))
            for i, (_, r) in enumerate(group.iterrows()):
                ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
                h_end = str(r['End Time']).strip() if str(r['End Time']).strip() != 'nan' else ""
                quart_txt = f"{r['Start Time']}-{h_end}" + ("" if is_qbe else f" / {cycle_p[i%5]}")
                ws.cell(row=curr_row, column=2, value=quart_txt).font = font_normal_bold
                c_val = ("VAM" if i == 0 else "VPM") if is_qbe else caisse_num
                ws.cell(row=curr_row, column=3, value=c_val)
                if not is_qbe: caisse_num += 1
                ws.cell(row=curr_row, column=4, value=calculer_pause_unique("RECEPTION", r['Start Time'], r['Total Time']))
                ws.cell(row=curr_row, column=5, value="QBE" if is_qbe else t_p[i])
                for c in range(1, 6):
                    cell = ws.cell(row=curr_row, column=c)
                    cell.border = border_thin
                    cell.font = font_normal
                    if is_qbe:
                        cell.fill = dark_gray_fill
                        cell.font = font_qbe
                curr_row += 1
        for c in range(1, 6): ws.cell(row=curr_row-1, column=c).border = border_thick_bottom

    elif section == "LOUNGE":
        lg_list = list(lg_data.iterrows())
        for i, (_, r) in enumerate(lg_list):
            ws.cell(row=curr_row, column=1, value=extraire_prenom(r['Employee']))
            h_end_lg = str(r['End Time']).strip() if str(r['End Time']).strip() != 'nan' else ""
            ws.cell(row=curr_row, column=2, value=f"{r['Start Time']}-{h_end_lg}").font = font_normal_bold
            t = "Ouverture" if i == 0 else "Fermeture" if i == len(lg_list)-1 else "Accueil" if i == 1 else ""
            ws.cell(row=curr_row, column=5, value=t)
            cell_c = ws.cell(row=curr_row, column=3)
            if t == "Accueil": cell_c.fill = black_fill
            else: cell_c.value = caisse_num; caisse_num += 1
            ws.cell(row=curr_row, column=4, value=calculer_pause_unique("LOUNGE", r['Start Time'], r['Total Time']))
            for c in range(1, 6):
                ws.cell(row=curr_row, column=c).border = border_thin
                ws.cell(row=curr_row, column=c).font = font_normal
            curr_row += 1
        for c in range(1, 6): ws.cell(row=curr_row-1, column=c).border = border_thick_bottom

# 4. SECTION FINALE
cercles_18_colles = "○" * 18
sections_finales = [("VENTES FIDÉLITÉS", cercles_18_colles), ("SOINS", cercles_18_colles), ("BONNE JOURNÉE !", None)]

for titre, contenu in sections_finales:
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
    cell_t = ws.cell(row=curr_row, column=1, value=titre)
    cell_t.fill = light_gray_fill
    cell_t.font = font_titre_secondaire 
    ws.row_dimensions[curr_row].height = 30
    curr_row += 1
    if contenu:
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=5)
        cell_c = ws.cell(row=curr_row, column=1, value=contenu)
        cell_c.font = font_cercles
        cell_c.alignment = center_no_wrap
        ws.row_dimensions[curr_row].height = 45
        for c in range(1, 6): ws.cell(row=curr_row, column=c).border = border_thin
        curr_row += 1

# --- POST-TRAITEMENT : BORDURES ET ALIGNEMENT ---
max_row = curr_row - 1
for r in range(1, max_row + 1):
    for c in range(1, 6):
        cell = ws.cell(row=r, column=c)
        if not cell.alignment.horizontal:
            cell.alignment = center_align
        cb = cell.border
        l = thick if c == 1 else cb.left
        ri = thick if c == 5 else cb.right
        t = thick if r == 1 else cb.top
        b = thick if r == max_row else cb.bottom
        cell.border = Border(left=l, right=ri, top=t, bottom=b)

# --- 5. AUTO-SIZE FINAL (S'assure que tout rentre sur une ligne) ---
for i in range(1, 6):
    max_length = 0
    column_letter = get_column_letter(i)
    for row in ws.iter_rows(min_col=i, max_col=i):
        for cell in row:
            # On ne calcule l'auto-size que sur les cellules non-fusionnées
            # pour éviter que le titre principal (A1:E1) ne fausse tout
            if cell.coordinate not in ws.merged_cells or cell.column == 1:
                try:
                    if cell.value:
                        length = len(str(cell.value))
                        # Petit ajustement pour les polices bold (colonne A et headers)
                        if cell.font.bold: length += 2
                        if length > max_length: max_length = length
                except: pass
    
    # Largeur finale avec un buffer de sécurité minimal
    ws.column_dimensions[column_letter].width = max_length + 3

wb.save(f"Horaire_Spa_Final_{date_obj.strftime('%Y%m%d')}.xlsx")
print(f"Généré : Horaire_Spa_Final_{date_obj.strftime('%Y%m%d')}.xlsx")

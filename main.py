import os
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF, Align
from datetime import datetime

# Chemin vers le fichier Excel

new_DOSSIER = input("entrer le chemin de votre dossier : ").strip('"').strip("'")
new_DOSSIER = new_DOSSIER.replace("\\", "\\\\")

DOSSIER = new_DOSSIER  

print(f"votre chemin : {DOSSIER}") 

if os.path.exists(DOSSIER):
    print("chemin valide")
else:
    print("ATTENTION, chemin introuvable")

retard_path = os.path.join(DOSSIER, "retard.xlsx")
etat_path = os.path.join(DOSSIER, "etat.xlsx")
activite_path = os.path.join(DOSSIER, "activite.xlsx")
charge_path = os.path.join(DOSSIER, "charge.xlsx")
immobilisation_path = os.path.join(DOSSIER, "immobilisation.xlsx")
retardpresta_path=os.path.join(DOSSIER, "retardpresta.xlsx")
    
# Lire exactement les lignes utiles
df_retard = pd.read_excel(retard_path)
df_etat = pd.read_excel(etat_path)
df_activite = pd.read_excel(activite_path)
df_charge=pd.read_excel(charge_path)
df_immobilisation=pd.read_excel(immobilisation_path)
def_retardpresta=pd.read_excel(retardpresta_path)

##################################################################################################
# Générer le camembert de retard
# On sélectionne les colonnes appropriées
labels_retard = df_retard.iloc[0:2, 0]  # Colonne A, lignes 2 et 3 (index 0-based → 0:2)
values_retard = df_retard.iloc[0:2, 1]  # Colonne B


plt.figure(figsize=(4,4))
plt.pie(values_retard, labels=labels_retard, autopct='%1.1f%%', colors=['#19AA6E', '#DC3223'])
plt.title("Taux de Retard")
plt.savefig("retard_pie.png")
plt.close()

##################################################################################################
#diagramme baton pour Etat

plt.figure(figsize=(20, 8))
plt.bar (df_etat.iloc[0:27,0], df_etat.iloc[0:27,1], color="#19AA6E")
plt.xlabel("Etat")
plt.ylabel("Nombre d'appareils")
plt.title("répartition des états des appareils")
plt.xticks(rotation=45)
plt.savefig("etat.png")
plt.close()

##################################################################################################
#graphique pour l'activite

plt.figure(figsize=(10, 6))

# Courbe action interne
plt.plot(df_activite.iloc[0:12, 0], df_activite.iloc[0:12, 2], marker='o', label='Interne', color=("#19AA6E"))

# Courbe action externe
plt.plot(df_activite.iloc[0:12, 0], df_activite.iloc[0:12, 1], marker='o', label='Externe', color=("#1E3246"))

plt.xlabel('Mois')
plt.ylabel('Nombre d\'activités')
plt.title('Activités en internes et externes')
plt.xticks(rotation=45)  # Pour mieux lire les mois
plt.legend()             # Affiche la légende
plt.grid(True)           # Ajoute une grille pour la lisibilité
plt.tight_layout()
plt.savefig("activite.png")
plt.close()

##################################################################################################
# Diagramme des charges 

plt.figure(figsize=(20, 8))
plt.bar (df_charge.iloc[0:16,0], df_charge.iloc[0:16,1], color="#1E3246")
plt.ylabel("Nombre d'appareils")
plt.title("Charge prévisionelle")
plt.xticks(rotation=90)
plt.savefig("charge.png")
plt.close()

##################################################################################################
# graph temps immobilisation

# Convertir les colonnes date en datetime (avec heure si besoin)
df_immobilisation['date_debut'] = pd.to_datetime(df_immobilisation['date_debut'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
df_immobilisation['date_fin'] = pd.to_datetime(df_immobilisation['date_fin'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

# Si certaines n'ont pas l'heure 
df_immobilisation['date_debut'] = df_immobilisation['date_debut'].fillna(pd.to_datetime(df_immobilisation['date_debut'], format='%d/%m/%Y'))
df_immobilisation['date_fin'] = df_immobilisation['date_fin'].fillna(pd.to_datetime(df_immobilisation['date_fin'], format='%d/%m/%Y'))

# Garde que la date (en tronquant l'heure)
df_immobilisation['date_debut'] = df_immobilisation['date_debut'].dt.floor('D')
df_immobilisation['date_fin'] = df_immobilisation['date_fin'].dt.floor('D')

# Calculer la durée en jours
df_immobilisation['Durée'] = (df_immobilisation['date_fin'] - df_immobilisation['date_debut']).dt.days

# Créer une colonne Mois
df_immobilisation['Mois'] = df_immobilisation['date_debut'].dt.to_period('M')

# Grouper par prestataire et mois et calculer la moyenne
df_grouped = df_immobilisation.groupby(['prestataire', 'Mois'])['Durée'].mean().reset_index()

#print(df_grouped)

couleurs=["#1E3246", '#DC3223',"#19AA6E",'#788291',"#72D7AA","#EB827D"]
# Tracer le graphique
plt.figure(figsize=(12, 6))

# boucle pour différentier les prestataires
for i, prestataire in enumerate(df_grouped['prestataire'].unique()):
    data = df_grouped[df_grouped['prestataire'] == prestataire]
    couleur = couleurs[i % len(couleurs)] 
    plt.plot(data['Mois'].astype(str), data['Durée'], marker='o', label=prestataire, color=couleur)

plt.xlabel('Mois')
plt.ylabel("Durée moyenne d'immobilisation (jours)")
plt.title('Durée moyenne de verification/étalonnage par prestataire et par mois')
plt.xticks(rotation=45)
plt.legend(title="Prestataire")
plt.grid(True)
plt.tight_layout()
plt.savefig("immobilisation.png")
plt.close()

##################################################################################################
# graph retard prestataire

def_grouped_retard = def_retardpresta.groupby('prestataire').agg(total_lignes = ('duree_depassee', 'count'), total_depasse = ('duree_depassee', 'sum'), total_respecte = ('duree_respectee', 'sum')) 

def_grouped_retard['%_depasse'] = def_grouped_retard['total_depasse'] / def_grouped_retard['total_lignes']*100
def_grouped_retard['%_respecte'] = def_grouped_retard['total_respecte'] / def_grouped_retard['total_lignes']*100

#print(def_grouped_retard)

plt.figure(figsize=(8,6))
plt.bar(def_grouped_retard.index, def_grouped_retard['%_depasse'], label='Dépassé', color = '#DC3223' )
plt.ylabel('pourcentage %')
plt.title('retard des prestataires en % tous les mois')
plt.savefig("presta_depassee.png")
plt.close()

plt.figure(figsize=(8,6))
plt.bar(def_grouped_retard.index, def_grouped_retard['%_respecte'], label='Respecté', color = "#19AA6E" )
plt.ylabel('pourcentage %')
plt.title('durée respectée par les prestataires en % tous les mois')
plt.savefig("presta_respectee.png")
plt.close()


##################################################################################################
# Créer le PDF
today = datetime.today().strftime("%d/%m/%y")

pdf = FPDF()
pdf.add_page()

pdf.set_font("Helvetica", size=16)
pdf.cell(200, 10, txt="Rapport Indicateurs Métrologie", ln=True, align='C')
pdf.set_font("Helvetica", size=8)
pdf.cell(0, -10, f"{today}", ln=True, align = 'R')

logo_path = os.path.join(DOSSIER, "Alstom.png")
if os.path.exists(logo_path):
    pdf.image(logo_path, x=Align.L, y=5, w=30, h=20)
else:
    print("image introuvable")

pdf.image("retard_pie.png", x=130, y=35, w=60, h=60)

pdf.image("etat.png", x=0.5, y=30, w=120, h=60)

pdf.image("activite.png", x=10, y=110, w=100, h=50)

pdf.image("charge.png", x=110, y=110, w=100, h=50)

pdf.image("immobilisation.png", x=Align.C, y=180, w=100, h=50)

pdf.image("presta_depassee.png", x=Align.L, y=240, w=100, h=50)

pdf.image("presta_respectee.png", x=Align.R, y=240, w=100, h=50)

pdf.set_line_width(0.5)

pdf.set_fill_color("#1E3246")
pdf.rect(x=5, y=25, w=200, h=5, style='F')
pdf.set_xy(x=91, y=22.5)
pdf.set_text_color("#FFFFFF")
pdf.multi_cell(100, 10, "Métrologie de Villeurbanne")

pdf.rect(x=5, y=100, w=200, h=5, style='F')
pdf.set_xy(x=95, y=97.5)
pdf.set_text_color("#FFFFFF")
pdf.multi_cell(100, 10, "Activité sur site")

pdf.rect(x=5, y=165, w=200, h=5, style='F')
pdf.set_xy(x=93, y=162.5)
pdf.set_text_color("#FFFFFF")
pdf.multi_cell(100, 10, "Activité des prestataires")

pdf.creation_date

chemin_pdf = os.path.join(DOSSIER, "Indicateur.pdf")
pdf.output(chemin_pdf)

print("✅ Rapport PDF généré avec succès !")

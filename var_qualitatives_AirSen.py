import pandas as pd

# Charger le fichier Excel
df = pd.read_excel("airsen.xlsx", sheet_name="Vols")

# Liste des variables qualitatives à analyser
variables_qualitatives = [
    "Compagnie",
    "Statut_Vol",
    "Type_Avion",
    "Sexe",
    "Météo"
]

# Fonction pour créer tableau de fréquence
def tableau_frequence(data, variable):
    freq = data[variable].value_counts(dropna=False).reset_index()
    freq.columns = ["Libellé", "Effectif"]
    
    freq["Pourcentage"] = round(
        (freq["Effectif"] / freq["Effectif"].sum()) * 100, 2
    )
    
    return freq

# Générer et exporter les tableaux
with pd.ExcelWriter("tableaux_frequences.xlsx") as writer:
    for var in variables_qualitatives:
        table = tableau_frequence(df, var)
        print(f"\nTableau de fréquence : {var}")
        print(table)
        
        # Export dans une feuille Excel
        table.to_excel(writer, sheet_name=var[:31], index=False)

print("Tous les tableaux ont été exportés dans tableaux_frequences.xlsx")
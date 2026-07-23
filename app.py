import pandas as pd
import openpyxl

def generer_rapport_vente_mensuel(fichier_entree, fichier_sortie="vente_mensuel_rapport.xlsx"):
    """
    Lit la feuille 'data', permet une sélection multiple par succursale,
    et génère un rapport mensuel avec Ventes ($) et Qté Livrée par produit,
    incluant une colonne Total à la fin.
    """
    # 1. Charger les feuilles du fichier Excel source
    try:
        df_data = pd.read_excel(fichier_entree, sheet_name="data")
        df_rapport = pd.read_excel(fichier_entree, sheet_name="rapport mensuel")
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier : {e}")
        return
    
    # 2. Sélection multiple interactive de la Succursale
    if "Succursale" in df_data.columns:
        succursales_dispo = df_data["Succursale"].dropna().unique()
        print(f"\nSuccursales disponibles : {list(succursales_dispo)}")
        
        choix = input("Entrez la ou les succursales souhaitées séparées par des virgules (ex: 1, 2, 3)\nOu appuyez sur Entrée pour TOUT inclure : ")
        
        if choix.strip():
            if pd.api.types.is_numeric_dtype(df_data["Succursale"]):
                valeurs_choisies = [float(s.strip()) for s in choix.split(",") if s.strip().replace('.', '', 1).isdigit()]
            else:
                valeurs_choisies = [s.strip() for s in choix.split(",")]
            
            df_data = df_data[df_data["Succursale"].isin(valeurs_choisies)]
            print(f"-> Filtre appliqué : {len(df_data)} lignes conservées.")

    # 3. Traitement des dates et extraction des mois (calendrier fiscal d'avril à mars)
    df_data["Date_Facture"] = pd.to_datetime(df_data["Date_Facture"])
    df_data["Mois_Num"] = df_data["Date_Facture"].dt.month
    
    mois_map = {
        4: "avril", 5: "mai", 6: "juin", 7: "juillet", 
        8: "août", 9: "septembre", 10: "octobre", 11: "novembre", 
        12: "décembre", 1: "janvier", 2: "février", 3: "mars"
    }
    mois_ordre = ["avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", "janvier", "février", "mars"]
    df_data["Mois_Nom"] = df_data["Mois_Num"].map(mois_map)
    
    # 4. Identification des colonnes descriptives du produit
    index_cols = [c for c in df_rapport.columns if c not in mois_ordre and c != "Date_Facture"]
    merge_keys = [c for c in index_cols if c in df_data.columns]
    
    # 5. Création du pivot pour les Ventes ($$$)
    df_pivot_vente = df_data.pivot_table(
        index=merge_keys,
        columns="Mois_Nom",
        values="Vente$$$",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    df_pivot_vente["Indicateur"] = "Ventes ($)"
    
    # 6. Création du pivot pour la Quantité Livrée (Qté_Livrée)
    df_pivot_qte = df_data.pivot_table(
        index=merge_keys,
        columns="Mois_Nom",
        values="Qté_Livrée",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    df_pivot_qte["Indicateur"] = "Qté Livrée"
    
    # 7. Regroupement des tableaux
    df_combined = pd.concat([df_pivot_vente, df_pivot_qte], ignore_index=True)
    
    # S'assurer de la présence de tous les mois et remplir les valeurs manquantes par 0
    for m in mois_ordre:
        if m not in df_combined.columns:
            df_combined[m] = 0.0
        else:
            df_combined[m] = df_combined[m].fillna(0.0)
            
    # 8. Calcul de la colonne Total (somme des 12 mois)
    df_combined["Total"] = df_combined[mois_ordre].sum(axis=1)
    
    # Ordonnancement final des colonnes : attributs produit + Indicateur + mois + Total
    final_cols = index_cols + ["Indicateur"] + mois_ordre + ["Total"]
    df_final = df_combined[final_cols]
    
    # 9. Export Excel final
    with pd.ExcelWriter(fichier_sortie, engine="openpyxl") as writer:
        df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").to_excel(writer, sheet_name="data", index=False)
        df_final.to_excel(writer, sheet_name="rapport mensuel", index=False)
        
    print(f"\nRapport de vente mensuel généré avec succès dans '{fichier_sortie}' (avec colonne Total) !")

if __name__ == "__main__":
    generer_rapport_vente_mensuel("data a convertir.xlsx")

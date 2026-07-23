import streamlit as st
import pandas as pd
import io

st.title("📊 Générateur de Rapport des Ventes Mensuelles")
st.write("Importez votre fichier Excel source pour générer le rapport croisé par produit (d'avril à mars) filtré par **Département**.")

# 1. Zone de téléchargement du fichier Excel
uploaded_file = st.file_uploader("Choisissez votre fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Charger les feuilles
        df_data = pd.read_excel(uploaded_file, sheet_name="data")
        df_rapport = pd.read_excel(uploaded_file, sheet_name="rapport mensuel")
        
        st.success("Fichier chargé avec succès !")
        
        # 2. Sélection multiple des Départements
        if "Département" in df_data.columns:
            departements_dispo = sorted(df_data["Département"].dropna().unique())
            selected_departements = st.multiselect(
                "Sélectionnez le ou les départements (laisser vide pour tout inclure) :",
                options=departements_dispo
            )
            
            if selected_departements:
                df_data = df_data[df_data["Département"].isin(selected_departements)]
                st.info(f"Filtre appliqué : {len(df_data)} lignes conservées.")

        # 3. Traitement des dates et extraction des mois
        df_data["Date_Facture"] = pd.to_datetime(df_data["Date_Facture"])
        df_data["Mois_Num"] = df_data["Date_Facture"].dt.month
        
        mois_map = {
            4: "avril", 5: "mai", 6: "juin", 7: "juillet", 
            8: "août", 9: "septembre", 10: "octobre", 11: "novembre", 
            12: "décembre", 1: "janvier", 2: "février", 3: "mars"
        }
        mois_ordre = ["avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", "janvier", "février", "mars"]
        df_data["Mois_Nom"] = df_data["Mois_Num"].map(mois_map)
        
        index_cols = [c for c in df_rapport.columns if c not in mois_ordre and c != "Date_Facture"]
        merge_keys = [c for c in index_cols if c in df_data.columns]
        
        # Pivot Ventes
        df_pivot_vente = df_data.pivot_table(
            index=merge_keys,
            columns="Mois_Nom",
            values="Vente$$$",
            aggfunc="sum",
            fill_value=0
        ).reset_index()
        df_pivot_vente["Indicateur"] = "Ventes ($)"
        
        # Pivot Quantité Livrée
        df_pivot_qte = df_data.pivot_table(
            index=merge_keys,
            columns="Mois_Nom",
            values="Qté_Livrée",
            aggfunc="sum",
            fill_value=0
        ).reset_index()
        df_pivot_qte["Indicateur"] = "Qté Livrée"
        
        # Combinaison
        df_combined = pd.concat([df_pivot_vente, df_pivot_qte], ignore_index=True)
        
        for m in mois_ordre:
            if m not in df_combined.columns:
                df_combined[m] = 0.0
            else:
                df_combined[m] = df_combined[m].fillna(0.0)
                
        # Calcul du Total
        df_combined["Total"] = df_combined[mois_ordre].sum(axis=1)
        
        final_cols = index_cols + ["Indicateur"] + mois_ordre + ["Total"]
        df_final = df_combined[final_cols]
        
        # 4. Génération du fichier Excel en mémoire pour téléchargement
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").to_excel(writer, sheet_name="data", index=False)
            df_final.to_excel(writer, sheet_name="rapport mensuel", index=False)
        processed_data = output.getvalue()
        
        st.subheader("Aperçu du rapport généré :")
        st.dataframe(df_final.head(10))
        
        # Bouton de téléchargement
        st.download_button(
            label="📥 Télécharger le rapport Excel final",
            data=processed_data,
            file_name="rapport_ventes_mensuelles.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Une erreur est survenue lors du traitement : {e}")

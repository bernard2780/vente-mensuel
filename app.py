import streamlit as st
import pandas as pd
import io

st.title("📊 Générateur de Rapport des Ventes Mensuelles")
st.write("Importez votre fichier Excel source pour générer les onglets séparés (**rapport ventes** et **rapport qté**).")

uploaded_file = st.file_uploader("Choisissez votre fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_data = pd.read_excel(uploaded_file, sheet_name="data")
        st.success("Fichier chargé avec succès !")
        
        # Sécurité sur les codes produits
        if "No_Produit" in df_data.columns:
            df_data["No_Produit"] = df_data["No_Produit"].astype(str).str.replace(r'\.0$', '', regex=True)
            df_data.loc[df_data["No_Produit"] == "nan", "No_Produit"] = ""

        exclude_cols = ["Date_Facture", "Vente$$$", "Qté_Livrée", "Mois_Num", "Mois_Nom"]
        index_cols = [c for c in df_data.columns if c not in exclude_cols]
        
        col_vente_source = "Vente$$$"
        if col_vente_source not in df_data.columns:
            alternatives = [c for c in df_data.columns if "vente" in c.lower()]
            if alternatives:
                col_vente_source = alternatives[0]
            else:
                st.error("Erreur : La colonne 'Vente$$$' est introuvable.")
                st.stop()

        # Gestion des départements
        if "Département" in df_data.columns:
            df_data["Département"] = df_data["Département"].fillna("Inconnu")
            departements_dispo = sorted(df_data["Département"].unique(), key=str)
            
            selected_departements = st.multiselect(
                "Sélectionnez le ou les départements :",
                options=departements_dispo,
                default=departements_dispo
            )
            
            if selected_departements:
                df_data = df_data[df_data["Département"].isin(selected_departements)]
                st.info(f"Filtre appliqué : {len(df_data)} lignes conservées.")
            else:
                st.warning("Veuillez sélectionner au moins un département.")
                st.stop()

        # Traitement des dates et mois
        df_data["Date_Facture"] = pd.to_datetime(df_data["Date_Facture"])
        df_data["Mois_Num"] = df_data["Date_Facture"].dt.month
        
        mois_ordre = ["avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", "janvier", "février", "mars"]
        mois_map = {
            4: "avril", 5: "mai", 6: "juin", 7: "juillet", 
            8: "août", 9: "septembre", 10: "octobre", 11: "novembre", 
            12: "décembre", 1: "janvier", 2: "février", 3: "mars"
        }
        df_data["Mois_Nom"] = df_data["Mois_Num"].map(mois_map)
        
        merge_keys = [c for c in index_cols if c in df_data.columns and c != "Département"]

        # Fonction de pivot simple et robuste (sans sous-totaux manuels)
        def simple_pivot(data, keys, months, val_col):
            valid_keys = [k for k in keys if k in data.columns]
            pivoted = data.pivot_table(
                index=valid_keys,
                columns="Mois_Nom",
                values=val_col,
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            
            for m in months:
                if m not in pivoted.columns:
                    pivoted[m] = 0.0
            return pivoted

        df_vente_final = simple_pivot(df_data, merge_keys, mois_ordre, col_vente_source)
        df_qte_final = simple_pivot(df_data, merge_keys, mois_ordre, "Qté_Livrée")
        
        for df_res in [df_vente_final, df_qte_final]:
            for m in mois_ordre:
                if m not in df_res.columns:
                    df_res[m] = 0.0
            df_res["Total"] = df_res[mois_ordre].sum(axis=1)
            
        final_cols = merge_keys + mois_ordre + ["Total"]
        for col in final_cols:
            if col not in df_vente_final.columns:
                df_vente_final[col] = ""
            if col not in df_qte_final.columns:
                df_qte_final[col] = ""
                
        df_vente_final = df_vente_final[final_cols]
        df_qte_final = df_qte_final[final_cols]

        # Sécurité Excel
        data_clean = df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").copy()
        for df_tocheck in [data_clean, df_vente_final, df_qte_final]:
            for col in df_tocheck.select_dtypes(include=['object', 'string']).columns:
                df_tocheck[col] = df_tocheck[col].apply(
                    lambda x: str(x).lstrip('=') if isinstance(x, str) and x.startswith('=') else x
                )

        # Génération du fichier Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data_clean.to_excel(writer, sheet_name="data", index=False)
            df_vente_final.to_excel(writer, sheet_name="rapport ventes", index=False)
            df_qte_final.to_excel(writer, sheet_name="rapport qté", index=False)
            
            for sheetname in writer.sheets:
                worksheet = writer.sheets[sheetname]
                for col in worksheet.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    worksheet.column_dimensions[col_letter].width = max(max_length + 3, 12)
                    
        processed_data = output.getvalue()
        
        st.subheader("Aperçu de l'onglet 'rapport ventes' :")
        st.dataframe(df_vente_final.head(10))
        
        st.download_button(
            label="📥 Télécharger le rapport Excel final",
            data=processed_data,
            file_name="rapport_ventes_mensuelles.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Une erreur est survenue lors du traitement : {e}")

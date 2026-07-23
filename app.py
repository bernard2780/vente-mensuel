import streamlit as st
import pandas as pd
import io

st.title("📊 Générateur de Rapport des Ventes Mensuelles")
st.write("Importez votre fichier Excel source pour générer les onglets séparés (**rapport ventes** et **rapport qté**) avec sous-totaux par **Division**, **Regroupement** et **Catégorie**.")

# 1. Zone de téléchargement du fichier Excel
uploaded_file = st.file_uploader("Choisissez votre fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Charger les feuilles
        df_data = pd.read_excel(uploaded_file, sheet_name="data")
        df_rapport = pd.read_excel(uploaded_file, sheet_name="rapport mensuel")
        
        st.success("Fichier chargé avec succès !")
        
        # Vérification et sélection de la colonne des ventes ($)
        col_vente_source = "Vente$$$"
        if col_vente_source not in df_data.columns:
            alternatives = [c for c in df_data.columns if "vente" in c.lower()]
            if alternatives:
                col_vente_source = alternatives[0]
            else:
                st.error("Erreur : La colonne 'Vente$$$' est introuvable dans votre feuille 'data'.")
                st.stop()

        # 2. Gestion et sélection multiple des Départements (Tous sélectionnés par défaut)
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

        # 3. Normalisation de la description du produit par No_Produit (Priorité au Département 01)
        if "No_Produit" in df_data.columns and "Description_Produit" in df_data.columns and "Département" in df_data.columns:
            desc_mapping = {}
            for prod, group in df_data.groupby('No_Produit'):
                dept1_rows = group[group['Département'].isin([1, 1.0, '1', '1.0'])]
                if not dept1_rows.empty and not dept1_rows['Description_Produit'].dropna().empty:
                    desc_mapping[prod] = dept1_rows['Description_Produit'].dropna().iloc[0]
                else:
                    valid_descs = group['Description_Produit'].dropna()
                    desc_mapping[prod] = valid_descs.iloc[0] if not valid_descs.empty else ""
            
            df_data['Description_Produit'] = df_data['No_Produit'].map(desc_mapping).fillna("")

        # 4. Traitement des dates et extraction des mois
        df_data["Date_Facture"] = pd.to_datetime(df_data["Date_Facture"])
        df_data["Mois_Num"] = df_data["Date_Facture"].dt.month
        
        mois_map = {
            4: "avril", 5: "mai", 6: "juin", 7: "juillet", 
            8: "août", 9: "septembre", 10: "octobre", 11: "novembre", 
            12: "décembre", 1: "janvier", 2: "février", 3: "mars"
        }
        mois_ordre = ["avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", "janvier", "février", "mars"]
        df_data["Mois_Nom"] = df_data["Mois_Num"].map(mois_map)
        
        index_cols = [c for c in df_rapport.columns if c not in mois_ordre and c != "Date_Facture" and c != "Indicateur"]

        # 5. Fonction de génération des rapports avec sous-totaux hiérarchiques
        def generate_report_with_subtotals(data, keys, months, val_col):
            valid_keys = [k for k in keys if k in data.columns]
            
            # Pivot initial pour regrouper par clé et par mois
            base_pivoted = data.pivot_table(
                index=valid_keys,
                columns="Mois_Nom",
                values=val_col,
                aggfunc="sum",
                fill_value=0,
                dropna=False
            ).reset_index()
            
            # S'assurer que tous les mois de l'ordre existent dans le dataframe
            for m in months:
                if m not in base_pivoted.columns:
                    base_pivoted[m] = 0.0
            
            div_col = next((k for k in valid_keys if k.lower() == "division"), None)
            reg_col = next((k for k in valid_keys if k.lower() in ["regroupement", "regroup"]), None)
            cat_col = next((k for k in valid_keys if k.lower() in ["catégorie", "categorie"]), None)
            
            rows = []
            
            if div_col and reg_col and cat_col:
                divisions = sorted(base_pivoted[div_col].dropna().unique(), key=str)
                for div in divisions:
                    df_div = base_pivoted[base_pivoted[div_col] == div]
                    regroupements = sorted(df_div[reg_col].dropna().unique(), key=str)
                    
                    for reg in regroupements:
                        df_reg = df_div[df_div[reg_col] == reg]
                        categories = sorted(df_reg[cat_col].dropna().unique(), key=str)
                        
                        for cat in categories:
                            df_cat = df_reg[df_reg[cat_col] == cat]
                            
                            # Lignes de détail
                            for _, row in df_cat.iterrows():
                                rows.append(row.to_dict())
                                
                            # Sous-total Catégorie
                            sub_cat = df_cat[months].sum().to_dict()
                            for k in valid_keys:
                                sub_cat[k] = df_cat[k].iloc[0] if not df_cat.empty else ""
                            sub_cat[cat_col] = f"Total Catégorie: {cat}"
                            if "No_Produit" in valid_keys:
                                sub_cat["No_Produit"] = ""
                            if "Description_Produit" in valid_keys:
                                sub_cat["Description_Produit"] = ""
                            rows.append(sub_cat)
                            
                        # Sous-total Regroupement
                        sub_reg = df_reg[months].sum().to_dict()
                        for k in valid_keys:
                            sub_reg[k] = df_reg[k].iloc[0] if not df_reg.empty else ""
                        sub_reg[reg_col] = f"Total Regroupement: {reg}"
                        if cat_col in valid_keys:
                            sub_reg[cat_col] = ""
                        if "No_Produit" in valid_keys:
                            sub_reg["No_Produit"] = ""
                        if "Description_Produit" in valid_keys:
                            sub_reg["Description_Produit"] = ""
                        rows.append(sub_reg)
                        
                    # Sous-total Division
                    sub_div = df_div[months].sum().to_dict()
                    for k in valid_keys:
                        sub_div[k] = df_div[k].iloc[0] if not df_div.empty else ""
                    sub_div[div_col] = f"Total Division: {div}"
                    if reg_col in valid_keys:
                        sub_div[reg_col] = ""
                    if cat_col in valid_keys:
                        sub_div[cat_col] = ""
                    if "No_Produit" in valid_keys:
                        sub_div["No_Produit"] = ""
                    if "Description_Produit" in valid_keys:
                        sub_div["Description_Produit"] = ""
                    rows.append(sub_div)
            else:
                for _, row in base_pivoted.iterrows():
                    rows.append(row.to_dict())
                    
            return pd.DataFrame(rows)

        # Générer les deux rapports distincts
        df_vente_final = generate_report_with_subtotals(df_data, index_cols, mois_ordre, col_vente_source)
        df_qte_final = generate_report_with_subtotals(df_data, index_cols, mois_ordre, "Qté_Livrée")
        
        # S'assurer que toutes les colonnes de mois et totaux sont présentes
        for df_res in [df_vente_final, df_qte_final]:
            for m in mois_ordre:
                if m not in df_res.columns:
                    df_res[m] = 0.0
                else:
                    df_res[m] = df_res[m].fillna(0.0)
            df_res["Total"] = df_res[mois_ordre].sum(axis=1)
            
        final_cols = index_cols + mois_ordre + ["Total"]
        for col in final_cols:
            if col not in df_vente_final.columns:
                df_vente_final[col] = ""
            if col not in df_qte_final.columns:
                df_qte_final[col] = ""
                
        df_vente_final = df_vente_final[final_cols]
        df_qte_final = df_qte_final[final_cols]
        
        # 6. CONTRÔLE DE BALANCEMENT AUTOMATIQUE
        total_source = df_data[col_vente_source].sum()
        # Exclure les lignes de sous-totaux pour vérifier le balancement exact des détails
        is_detail_row = ~df_vente_final.astype(str).apply(lambda x: x.str.contains("Total")).any(axis=1)
        total_rapport_detail = df_vente_final[is_detail_row]["Total"].sum()
        
        st.subheader("🔍 Validation du balancement des ventes :")
        col1, col2 = st.columns(2)
        col1.metric("Total Ventes (Source Data)", f"{total_source:,.2f} $")
        col2.metric("Total Ventes (Détails Rapport)", f"{total_rapport_detail:,.2f} $")
        
        if abs(total_source - total_rapport_detail) < 0.01:
            st.success("✅ Les montants balancent parfaitement entre la source et le rapport !")
        else:
            st.error(f"⚠️ Écart détecté : {abs(total_source - total_rapport_detail):,.2f} $")

        # 7. SÉCURITÉ EXCEL : Neutraliser les textes commençant par '='
        data_clean = df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").copy()
        for df_tocheck in [data_clean, df_vente_final, df_qte_final]:
            for col in df_tocheck.select_dtypes(include=['object', 'string']).columns:
                df_tocheck[col] = df_tocheck[col].apply(
                    lambda x: str(x).lstrip('=') if isinstance(x, str) and x.startswith('=') else x
                )

        # 8. Génération du fichier Excel en mémoire avec les 3 onglets distincts
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data_clean.to_excel(writer, sheet_name="data", index=False)
            df_vente_final.to_excel(writer, sheet_name="rapport ventes", index=False)
            df_qte_final.to_excel(writer, sheet_name="rapport qté", index=False)
        processed_data = output.getvalue()
        
        st.subheader("Aperçu de l'onglet 'rapport ventes' :")
        st.dataframe(df_vente_final.head(10))
        
        # Bouton de téléchargement
        st.download_button(
            label="📥 Télécharger le rapport Excel final (avec onglets séparés)",
            data=processed_data,
            file_name="rapport_ventes_mensuelles.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Une erreur est survenue lors du traitement : {e}")

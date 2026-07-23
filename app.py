import streamlit as st
import pandas as pd
import io

st.title("📊 Générateur de Rapport des Ventes Mensuelles")
st.write("Importez votre fichier Excel source pour générer les onglets séparés (**rapport ventes** et **rapport qté**) avec sous-totaux hiérarchiques.")

# 1. Zone de téléchargement du fichier Excel
uploaded_file = st.file_uploader("Choisissez votre fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Charger uniquement la feuille data de manière sécurisée
        df_data = pd.read_excel(uploaded_file, sheet_name="data")
        st.success("Fichier chargé avec succès !")
        
        # Définition automatique des colonnes d'index (tout sauf les dates et valeurs numériques de base)
        exclude_cols = ["Date_Facture", "Vente$$$", "Qté_Livrée", "Mois_Num", "Mois_Nom"]
        index_cols = [c for c in df_data.columns if c not in exclude_cols]
        
        # Vérification de la colonne des ventes ($)
        col_vente_source = "Vente$$$"
        if col_vente_source not in df_data.columns:
            alternatives = [c for c in df_data.columns if "vente" in c.lower()]
            if alternatives:
                col_vente_source = alternatives[0]
            else:
                st.error("Erreur : La colonne 'Vente$$$' est introuvable dans votre feuille 'data'.")
                st.stop()

        # 2. Gestion et sélection multiple des Départements
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
        
        mois_ordre = ["avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre", "janvier", "février", "mars"]
        mois_map = {
            4: "avril", 5: "mai", 6: "juin", 7: "juillet", 
            8: "août", 9: "septembre", 10: "octobre", 11: "novembre", 
            12: "décembre", 1: "janvier", 2: "février", 3: "mars"
        }
        df_data["Mois_Nom"] = df_data["Mois_Num"].map(mois_map)
        
        merge_keys = [c for c in index_cols if c in df_data.columns and c != "Département"]

        # 5. Fonction de génération des rapports optimisée (sans surcharge mémoire)
        def generate_report_with_subtotals(data, keys, months, val_col):
            valid_keys = [k for k in keys if k in data.columns]
            
            # Utilisation standard sans dropna=False pour éviter l'explosion mémoire
            base_pivoted = data.pivot_table(
                index=valid_keys,
                columns="Mois_Nom",
                values=val_col,
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            
            for m in months:
                if m not in base_pivoted.columns:
                    base_pivoted[m] = 0.0
            
            div_col = next((k for k in valid_keys if "division" in k.lower()), None)
            reg_col = next((k for k in valid_keys if "regroup" in k.lower()), None)
            cat_col = next((k for k in valid_keys if "catég" in k.lower() or "categ" in k.lower()), None)
            
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
                            
                            for _, row in df_cat.iterrows():
                                rows.append(row.to_dict())
                                
                            sub_cat = df_cat[months].sum().to_dict()
                            for k in valid_keys:
                                sub_cat[k] = df_cat[k].iloc[0] if not df_cat.empty else ""
                            sub_cat[cat_col] = f"Total Catégorie: {cat}"
                            for col_to_blank in ["No_Produit", "Description_Produit", "Sous_Catégorie"]:
                                if col_to_blank in valid_keys:
                                    sub_cat[col_to_blank] = ""
                            rows.append(sub_cat)
                            
                        sub_reg = df_reg[months].sum().to_dict()
                        for k in valid_keys:
                            sub_reg[k] = df_reg[k].iloc[0] if not df_reg.empty else ""
                        sub_reg[reg_col] = f"Total Regroupement: {reg}"
                        for col_to_blank in ["Catégorie", "No_Produit", "Description_Produit", "Sous_Catégorie"]:
                            if col_to_blank in valid_keys:
                                sub_reg[col_to_blank] = ""
                        rows.append(sub_reg)
                        
                    sub_div = df_div[months].sum().to_dict()
                    for k in valid_keys:
                        sub_div[k] = df_div[k].iloc[0] if not df_div.empty else ""
                    sub_div[div_col] = f"Total Division: {div}"
                    for col_to_blank in ["Regroupement", "Regroupe", "Catégorie", "No_Produit", "Description_Produit", "Sous_Catégorie"]:
                        if col_to_blank in valid_keys:
                            sub_div[col_to_blank] = ""
                    rows.append(sub_div)
            else:
                for _, row in base_pivoted.iterrows():
                    rows.append(row.to_dict())
                    
            return pd.DataFrame(rows)

        # Générer les deux rapports distincts
        df_vente_final = generate_report_with_subtotals(df_data, merge_keys, mois_ordre, col_vente_source)
        df_qte_final = generate_report_with_subtotals(df_data, merge_keys, mois_ordre, "Qté_Livrée")
        
        for df_res in [df_vente_final, df_qte_final]:
            for m in mois_ordre:
                if m not in df_res.columns:
                    df_res[m] = 0.0
                else:
                    df_res[m] = df_res[m].fillna(0.0)
            df_res["Total"] = df_res[mois_ordre].sum(axis=1)
            
        final_cols = merge_keys + mois_ordre + ["Total"]
        for col in final_cols:
            if col not in df_vente_final.columns:
                df_vente_final[col] = ""
            if col not in df_qte_final.columns:
                df_qte_final[col] = ""
                
        df_vente_final = df_vente_final[final_cols]
        df_qte_final = df_qte_final[final_cols]
        
        # 6. CONTRÔLE DE BALANCEMENT AUTOMATIQUE
        total_source = df_data[col_vente_source].sum()
        is_detail_row = ~df_vente_final.astype(str).apply(lambda x: x.str.contains("Total")).any(axis=1)
        total_rapport_detail = df_vente_final[is_detail_row]["Total"].sum()
        
        st.subheader("🔍 Validation du balancement des ventes :")
        col1, col2 = st.columns(2)
        col1.metric("Total Ventes (Source Data)", f"{total_source:,.2f} $")
        col2.metric("Total Ventes (Détails Rapport)", f"{total_rapport_detail:,.2f} $")
        
        if abs(total_source - total_rapport_detail) < 0.01:
            st.success("✅ Les montants balancent parfaitement entre la source et le rapport !")
        else:
            st.warning(f"⚠️ Écart détecté : {abs(total_source - total_rapport_detail):,.2f} $")

        # 7. SÉCURITÉ EXCEL
        data_clean = df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").copy()
        for df_tocheck in [data_clean, df_vente_final, df_qte_final]:
            for col in df_tocheck.select_dtypes(include=['object', 'string']).columns:
                df_tocheck[col] = df_tocheck[col].apply(
                    lambda x: str(x).lstrip('=') if isinstance(x, str) and x.startswith('=') else x
                )

        # 8. Génération du fichier Excel avec onglets multiples et mise en page auto
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

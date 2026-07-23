import streamlit as st
import pandas as pd
import io

st.title("📊 Générateur de Rapport des Ventes Mensuelles")
st.write("Importez votre fichier Excel source pour générer le rapport croisé par produit (d'avril à mars) filtré par **Département** avec balancement rigoureux.")

# 1. Zone de téléchargement du fichier Excel
uploaded_file = st.file_uploader("Choisissez votre fichier Excel source (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Charger les feuilles
        df_data = pd.read_excel(uploaded_file, sheet_name="data")
        df_rapport = pd.read_excel(uploaded_file, sheet_name="rapport mensuel")
        
        st.success("Fichier chargé avec succès !")
        
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
        
        index_cols = [c for c in df_rapport.columns if c not in mois_ordre and c != "Date_Facture"]
        merge_keys = [c for c in index_cols if c in df_data.columns]
        
        # 5. Pivot Ventes (avec dropna=False pour ne perdre aucune ligne)
        df_pivot_vente = df_data.pivot_table(
            index=merge_keys,
            columns="Mois_Nom",
            values="Vente$$$",
            aggfunc="sum",
            fill_value=0,
            dropna=False
        ).reset_index()
        df_pivot_vente["Indicateur"] = "Ventes ($)"
        
        # 6. Pivot Quantité Livrée (avec dropna=False)
        df_pivot_qte = df_data.pivot_table(
            index=merge_keys,
            columns="Mois_Nom",
            values="Qté_Livrée",
            aggfunc="sum",
            fill_value=0,
            dropna=False
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
        
        # 7. CONTRÔLE DE BALANCEMENT AUTOMATIQUE
        total_source = df_data["Vente$$$"].sum()
        total_rapport = df_final[df_final["Indicateur"] == "Ventes ($)"]["Total"].sum()
        
        st.subheader("🔍 Validation du balancement des ventes :")
        col1, col2 = st.columns(2)
        col1.metric("Total Ventes (Source Data)", f"{total_source:,.2f} $")
        col2.metric("Total Ventes (Rapport)", f"{total_rapport:,.2f} $")
        
        if abs(total_source - total_rapport) < 0.01:
            st.success("✅ Les montants balancent parfaitement entre la source et le rapport !")
        else:
            st.error(f"⚠️ Écart détecté : {abs(total_source - total_rapport):,.2f} $")

        # 8. SÉCURITÉ EXCEL : Neutraliser les textes commençant par '='
        data_clean = df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").copy()
        for df_tocheck in [data_clean, df_final]:
            for col in df_tocheck.select_dtypes(include=['object', 'string']).columns:
                df_tocheck[col] = df_tocheck[col].apply(
                    lambda x: str(x).lstrip('=') if isinstance(x, str) and x.startswith('=') else x
                )

        # 9. Génération du fichier Excel en mémoire pour téléchargement
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            data_clean.to_excel(writer, sheet_name="data", index=False)
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

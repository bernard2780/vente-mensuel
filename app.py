import io
import pandas as pd
import streamlit as st

st.title("📊 Générateur de Rapport des Ventes Mensuelles")
st.write(
    "Importez votre fichier Excel source pour générer le rapport croisé par"
    " produit (d'avril à mars) filtré par **Département**."
)

uploaded_file = st.file_uploader(
    "Choisissez votre fichier Excel source (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
  try:
    # 1. Charger TOUT en texte brut pour bloquer les grands entiers
    df_data = pd.read_excel(uploaded_file, sheet_name="data", dtype=str)
    st.success("Fichier chargé avec succès !")

    # 2. Nettoyer et convertir les colonnes de chiffres
    for col in ["Vente$$$", "Qté_Livrée"]:
      if col in df_data.columns:
        df_data[col] = (
            df_data[col]
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df_data[col] = pd.to_numeric(df_data[col], errors="coerce").fillna(0)

    # 3. Forcer TOUTES les autres colonnes en texte pur (Évite l'erreur C long)
    for col in df_data.columns:
      if col not in ["Vente$$$", "Qté_Livrée"]:
        df_data[col] = (
            df_data[col]
            .fillna("")
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
        )
        df_data[col] = df_data[col].replace("nan", "")

    # 4. Gestion des départements
    if "Département" in df_data.columns:
      departements_dispo = sorted(
          [d for d in df_data["Département"].unique() if d != ""]
      )
      selected_departements = st.multiselect(
          "Sélectionnez le ou les départements :",
          options=departements_dispo,
          default=departements_dispo,
      )

      if selected_departements:
        df_data = df_data[df_data["Département"].isin(selected_departements)]
        st.info(f"Filtre appliqué : {len(df_data)} lignes conservées.")
      else:
        st.warning("Veuillez sélectionner au moins un département.")
        st.stop()

    # 5. Dates et Mois
    df_data["Date_Facture"] = pd.to_datetime(
        df_data["Date_Facture"], errors="coerce"
    )
    df_data["Mois_Num"] = df_data["Date_Facture"].dt.month

    mois_map = {
        4: "avril",
        5: "mai",
        6: "juin",
        7: "juillet",
        8: "août",
        9: "septembre",
        10: "octobre",
        11: "novembre",
        12: "décembre",
        1: "janvier",
        2: "février",
        3: "mars",
    }
    mois_ordre = [
        "avril",
        "mai",
        "juin",
        "juillet",
        "août",
        "septembre",
        "octobre",
        "novembre",
        "décembre",
        "janvier",
        "février",
        "mars",
    ]
    df_data["Mois_Nom"] = df_data["Mois_Num"].map(mois_map).fillna("inconnu")

    exclude_cols = [
        "Date_Facture",
        "Vente$$$",
        "Qté_Livrée",
        "Mois_Num",
        "Mois_Nom",
        "Département",
    ]
    index_cols = [c for c in df_data.columns if c not in exclude_cols]

    # 6. Regroupement sécurisé sans pivot_table pour éliminer l'erreur C long
    def safe_group(data, val_col, ind_name):
      piv = (
          data.groupby(index_cols + ["Mois_Nom"])[val_col]
          .sum()
          .unstack(fill_value=0)
          .reset_index()
      )
      for m in mois_ordre:
        if m not in piv.columns:
          piv[m] = 0.0
      piv["Indicateur"] = ind_name
      return piv

    df_piv_v = safe_group(df_data, "Vente$$$", "Ventes ($)")
    df_piv_q = safe_group(df_data, "Qté_Livrée", "Qté Livrée")

    df_combined = pd.concat([df_piv_v, df_piv_q], ignore_index=True)
    df_combined["Total"] = df_combined[mois_ordre].sum(axis=1)

    final_cols = index_cols + ["Indicateur"] + mois_ordre + ["Total"]
    df_final = df_combined[final_cols]

    # 7. Contrôle de balancement
    total_source = df_data["Vente$$$"].sum()
    total_rapport = df_final[df_final["Indicateur"] == "Ventes ($)"][
        "Total"
    ].sum()

    st.subheader("🔍 Validation du balancement des ventes :")
    col1, col2 = st.columns(2)
    col1.metric("Total Ventes (Source Data)", f"{total_source:,.2f} $")
    col2.metric("Total Ventes (Rapport)", f"{total_rapport:,.2f} $")

    # 8. Export Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_data.drop(columns=["Mois_Num", "Mois_Nom"], errors="ignore").to_excel(
          writer, sheet_name="data", index=False
      )
      df_final.to_excel(writer, sheet_name="rapport mensuel", index=False)
    processed_data = output.getvalue()

    st.subheader("Aperçu du rapport généré :")
    st.dataframe(df_final.head(10))

    st.download_button(
        label="📥 Télécharger le rapport Excel final",
        data=processed_data,
        file_name="rapport_ventes_mensuelles.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")

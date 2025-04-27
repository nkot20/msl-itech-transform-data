import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

mapping_accounts = {
    700100: "x_studio_loyer_actuel_index",
    700200: "x_studio_loyer_actuel_index",
    700500: "x_studio_intervention_obligatoire",
    704000: "x_studio_forfait",
    701000: "x_studio_provision_pour_charge",
    600100: "x_studio_loyer_actuel_index",
    600200: "x_studio_loyer_actuel_index",
    601900: "x_studio_provision_pour_charge"
}


# Fonction pour extraire les valeurs spécifiques de `comment-int`
def extract_analytical_code(comment):
    """ Extrait la dernière valeur après '/' """
    parts = comment.split("/") if isinstance(comment, str) else []
    return parts[-1] if len(parts) >= 1 else ""

def extract_address(comment):
    """ Extrait la valeur après l'avant-dernier '/' """
    parts = comment.split("/") if isinstance(comment, str) else []
    return parts[-2] if len(parts) >= 2 else ""

# Fonction de transformation
def prepare_data_for_journal(df, journal_name):
    df_filtered = df[df['journal'] == journal_name].copy()

    # Génération du champ 'name' en fopnction du journal
    if journal_name in ["AC2", "GESTIO"]:
        df_filtered.loc[:, 'name'] = "2500-" + df_filtered['docnumber'].astype(str).str.zfill(4)
    elif journal_name == "ODGEST":
        df_filtered['datedoc'] = pd.to_datetime(df_filtered['datedoc'])
        df_filtered['name'] = (
                df_filtered['journal'] + "/" +
                df_filtered['datedoc'].dt.year.astype(str) + "/" +
                df_filtered['datedoc'].dt.month.astype(str).str.zfill(2) + "/" +
                df_filtered['docnumber'].astype(str).str.zfill(4)
        )
    else:
        df_filtered.loc[:, 'name'] = df_filtered['bookyear'].astype(str) + '-' + df_filtered['docnumber'].astype(
            str).str.zfill(4)

    if journal_name == "GESTIO":
        df_filtered['journal'] = "GESTI"

    # Nettoyage et conversion de 'montant-gen' en nombre
    df_filtered['montant-gen'] = df_filtered['montant-gen'].replace(',', '.', regex=True).replace('[^\d.]', '',
                                                                                                  regex=True)
    df_filtered['montant-gen'] = pd.to_numeric(df_filtered['montant-gen'], errors='coerce').fillna(0)

    # Conversion des dates en format sans heure
    df_filtered['datedoc'] = pd.to_datetime(df_filtered['datedoc']).dt.strftime('%Y.%m.%d')
    df_filtered['duedate'] = pd.to_datetime(df_filtered['duedate']).dt.strftime('%Y.%m.%d')

    # **Ajout de la colonne 'Référence' basée sur le comment-int du compte spécifique**
    if journal_name in ["GESTIO", "AC2", "VEN"]:
        reference_account = 400000 if journal_name in ["VEN", "GESTIO"] else 440100

        # Récupérer `comment-int` pour chaque groupe (docnumber + account-id)
        df_filtered['Référence'] = df_filtered.groupby(['docnumber', 'account-id'])['comment-int'].transform(
            lambda x: x[df_filtered['accountgl'] == reference_account].iloc[0]
            if (df_filtered['accountgl'] == reference_account).any() else x.iloc[0]
        )
    else:
        df_filtered['Référence'] = df_filtered['comment-int']

    # Suppression des lignes en fonction du journal
    if journal_name in ["VEN", "GESTIO"]:
        df_filtered = df_filtered[df_filtered['accountgl'] != 400000]
    if journal_name == "AC2":
        df_filtered = df_filtered[df_filtered['accountgl'] != 440100]

    # Cas spécifique pour les journaux VEN, AC2 et GESTIO
    if journal_name in ["VEN", "GESTIO"]:
        price_unit = np.where(df_filtered['D-C'] == 'D', -df_filtered['montant-gen'], df_filtered['montant-gen'])
    elif journal_name == "AC2":
        price_unit = np.where(df_filtered['D-C'] == 'D', df_filtered['montant-gen'], -df_filtered['montant-gen'])
    else:
        price_unit = np.zeros(len(df_filtered))  # Valeur par défaut

    # Gestion spécifique pour le journal ODGES
    if journal_name == "ODGEST":
        df_filtered['journal'] = "ODGES"
        df_destination = pd.DataFrame({
            'Numéro': df_filtered['name'],
            'Écritures comptables/Partenaire': df_filtered['account-id'],
            'Date': df_filtered['datedoc'],
            'Journal': df_filtered['journal'],
            'Écritures comptables/Crédit': np.where(df_filtered['D-C'] == 'C', df_filtered['montant-gen'], 0),
            'Écritures comptables/Débit': np.where(df_filtered['D-C'] == 'D', df_filtered['montant-gen'], 0),
            'Écritures comptables/Libellé': df_filtered['comment-int'],
            'Écritures comptables/Compte/Code': df_filtered['accountgl'],  # Dernière colonne
        })
    else:
        # DataFrame standard pour les autres journaux
        df_destination = pd.DataFrame({
            'name': df_filtered['name'],
            'partner_id': df_filtered['account-id'],
            'invoice_date': df_filtered['datedoc'],
            'invoice_date_due': df_filtered['duedate'],
            'journal_code': df_filtered['journal'],
            'account_id': df_filtered['accountgl'],
            'invoice_line_ids/price_unit': price_unit,  # Colonne ajoutée avant Référence
            'Référence': df_filtered['Référence'],
        })

    # Suppression des doublons pour éviter la répétition des valeurs
    cols_to_check = ['name', 'partner_id', 'invoice_date', 'invoice_date_due', 'journal_code', 'Référence']
    if journal_name == "ODGEST":
        cols_to_check = ['Numéro', 'Date', 'Journal']

    df_destination.loc[df_destination.duplicated(subset=cols_to_check, keep='first'), cols_to_check] = ''

    return df_destination


# ======= FONCTION 2 : Extraction des commentaires =======
def extract_comments(df):
    df_filtered = df[df['journal'].isin(["AC2", "VEN"])].copy()
    df_filtered = df_filtered[df_filtered['accountgl'].isin([400000, 440100])]

    df_filtered['comment-int'] = df_filtered['comment-int'].apply(lambda x: x.split("/")[-1] if isinstance(x, str) else x)

    df_result = df_filtered[['journal', 'accountgl', 'account-id', 'comment-int']]

    return df_result


# ======= FONCTION 3 : Extraction des valeurs après l'avant-dernier slash =======
def extract_second_last_comment(df):
    df_filtered = df[df['journal'].isin(["AC2", "VEN"])].copy()
    df_filtered = df_filtered[~df_filtered['accountgl'].isin([400000, 440100, 499200])]

    def get_second_last_part(comment):
        if isinstance(comment, str) and comment.count("/") >= 2:
            return comment.split("/")[-2]  # Récupérer l'avant-dernier élément
        return comment  # Retourner inchangé si moins de 2 "/"

    df_filtered['comment-int'] = df_filtered['comment-int'].apply(get_second_last_part)

    df_result = df_filtered[['journal', 'accountgl', 'account-id', 'comment-int', 'montant-gen']]

    return df_result


def transform_hms_to_odoo(df_hms, df_destination_template):
    df_filtered = df_hms[df_hms["journal"].isin(["VEN", "AC2"])].copy()
    df_filtered["montant-gen"] = df_filtered["montant-gen"].replace(",", ".", regex=True)
    df_filtered["montant-gen"] = pd.to_numeric(df_filtered["montant-gen"], errors="coerce").fillna(0)
    df_filtered.sort_values(by=["account-id", "docnumber"], inplace=True)

    grouped_data = df_filtered.groupby(["account-id", "docnumber"])
    df_unmatched = pd.DataFrame(columns=df_destination_template.columns)

    for (account_id, doc_number), group in grouped_data:
        if account_id in df_destination_template["x_studio_rf_wb"].values:
            dest_df = df_destination_template
        else:
            if account_id not in df_unmatched["x_studio_rf_wb"].values:
                new_row = pd.Series("", index=df_unmatched.columns)
                new_row["x_studio_rf_wb"] = account_id
                df_unmatched = pd.concat([df_unmatched, pd.DataFrame([new_row])], ignore_index=True)
            dest_df = df_unmatched

        dest_index = dest_df[dest_df["x_studio_rf_wb"] == account_id].index[0]

        # Adresse actuelle à écrire
        current_analytical = str(extract_analytical_code(group.iloc[0]["comment-int"]))
        current_address = str(extract_address(group.iloc[0]["comment-int"]))

        suffix = ""
        found_existing_block = False
        for i in range(20):
            suffix_try = f"_{i}" if i > 0 else ""
            analytical_col = f"x_studio_code_analytique{suffix_try}"
            address_col = f"x_studio_adresse{suffix_try}"

            current_block_analytical = dest_df.at[dest_index, analytical_col] if analytical_col in dest_df.columns else ""
            current_block_address = dest_df.at[dest_index, address_col] if address_col in dest_df.columns else ""

            if (current_block_analytical == current_analytical and current_block_address == current_address):
                suffix = suffix_try
                found_existing_block = True
                break
            elif (pd.isna(current_block_analytical) or current_block_analytical == "") and (pd.isna(current_block_address) or current_block_address == ""):
                suffix = suffix_try
                if analytical_col in dest_df.columns:
                    dest_df.at[dest_index, analytical_col] = current_analytical
                if address_col in dest_df.columns:
                    dest_df.at[dest_index, address_col] = current_address
                found_existing_block = True
                break

        if not found_existing_block:
            continue  # Par sécurité, éviter d'écrire dans un bloc non trouvé

        # Ajout montant principal
        main_rent_account = None
        if "VEN" in group["journal"].values:
            main_rent_account = 700100 if 700100 in group["accountgl"].values else 700200
        elif "AC2" in group["journal"].values:
            main_rent_account = 600100 if 600100 in group["accountgl"].values else 600200

        if main_rent_account is not None:
            column_name = mapping_accounts[main_rent_account] + suffix
            montant_value = group[group["accountgl"] == main_rent_account]["montant-gen"].sum()
            if column_name in dest_df.columns:
                dest_df.at[dest_index, column_name] = float(montant_value)

        for _, row in group.iterrows():
            account_gl = row["accountgl"]
            montant_gen = row["montant-gen"]
            if pd.notna(montant_gen) and montant_gen != 0 and account_gl in mapping_accounts:
                column_name = mapping_accounts[account_gl] + suffix
                if column_name in dest_df.columns:
                    dest_df.at[dest_index, column_name] = float(montant_gen)

    return df_destination_template, df_unmatched

def generate_excel_with_two_sheets(df1, df2):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Données transformées", index=False)
        if not df2.empty:
            df2.to_excel(writer, sheet_name="Non présents dans modèle", index=False)
    output.seek(0)
    return output

def extract_missing_partner_ids(df_update, transformed_data_dict):
    """
    Compare les anciens partner_id du fichier de mise à jour avec ceux présents
    dans les feuilles transformées. Retourne les lignes absentes.
    """
    # 1. Extraire tous les partner_id déjà présents dans les résultats transformés
    all_present_ids = set()
    for journal, df in transformed_data_dict.items():
        if journal == "ODGEST" and "Écritures comptables/Partenaire" in df.columns:
            all_present_ids.update(df["Écritures comptables/Partenaire"].dropna().astype(str).unique())
        elif "partner_id" in df.columns:
            all_present_ids.update(df["partner_id"].dropna().astype(str).unique())

    # 2. S'assurer que df_update a les bonnes colonnes
    df_update.columns = ["ancien", "nouveau"]
    df_update = df_update.astype(str)

    # 3. Filtrer ceux qui ne sont pas présents
    missing_rows = df_update[~df_update["ancien"].isin(all_present_ids)]

    return missing_rows

def extract_ids_missing_from_update(df_update, transformed_data_dict):
    """
    Compare les partner_id présents dans les feuilles transformées avec ceux du fichier de mise à jour.
    Retourne les partner_id absents dans le fichier de mise à jour avec le nom de la feuille d'origine.
    """
    df_update.columns = ["Réf WB", "Nom"]
    update_ids = set(df_update["Nom"].astype(str))

    missing_records = []

    for journal, df in transformed_data_dict.items():
        if journal == "ODGEST" and "Écritures comptables/Partenaire" in df.columns:
            present_ids = df["Écritures comptables/Partenaire"].dropna().astype(str).unique()
        elif journal in ["VEN", "AC2", "GESTIO"] and "partner_id" in df.columns:
            present_ids = df["partner_id"].dropna().astype(str).unique()
        else:
            continue

        for pid in present_ids:
            if pid not in update_ids:
                missing_records.append({"partner_id": pid, "feuille": journal})

    df_missing_ids = pd.DataFrame(missing_records)
    return df_missing_ids

# ======= INTERFACE UTILISATEUR STREAMLIT =======
st.title("📂 MSL-ITECH - Transformation de fichier Excel HMS")

# 🌟 Personnalisation du style CSS
st.markdown("""
    <style>
        .stDownloadButton>button {
            background-color: #4CAF50;
            color: white;
            font-size: 16px;
            padding: 10px;
            border-radius: 5px;
            border: none;
        }
        .stDownloadButton>button:hover {
            background-color: #45a049;
        }
        .stFileUploader {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 10px;
            border: 1px solid #ddd;
        }
        .stTabs {
            font-size: 18px;
        }
    </style>
""", unsafe_allow_html=True)


# 🌟 Création des onglets
tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Transformation du fichier HMS",
    "🔄 Extraction des commentaires",
    "📌 Extraction avancée",
    "📂 Transformation vers le format Odoo"
])

# 🟢 Onglet 1 : Transformation du fichier HMS vers ODOO
with tab1:
    st.header("🚀 Transformation du fichier HMS vers ODOO")

    # 📂 Téléchargement du fichier principal
    uploaded_file = st.file_uploader("📥 **Téléchargez le fichier source HMS (Excel)**", type=['xlsx'], key="file1")

    # 📂 Téléchargement du fichier de mise à jour des `partner_id`
    uploaded_update_file = st.file_uploader("🔄 **Téléchargez le fichier de mise à jour des Partner ID**", type=['xlsx'],
                                            key="update_file")

    if uploaded_file is not None:
        st.success("✅ **Fichier principal chargé avec succès !**")
        df_source = pd.read_excel(uploaded_file)

        journals = df_source['journal'].unique()
        output_buffer = BytesIO()
        transformed_data_dict = {}  # Dictionnaire pour stocker les DataFrames par feuille

        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            for journal in journals:
                df_journal = prepare_data_for_journal(df_source, journal)  # ✅ **L'algorithme d'origine est conservé**
                if not df_journal.empty:
                    df_journal.to_excel(writer, sheet_name=journal, index=False)
                    transformed_data_dict[journal] = df_journal  # Stocker chaque feuille

        output_buffer.seek(0)

        # 📥 **Téléchargement du fichier transformé (sans mise à jour)**
        st.download_button(
            label="📥 **Télécharger le fichier transformé**",
            data=output_buffer,
            file_name="HMS_RESULT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 📊 **Aperçu des premières lignes**
        if transformed_data_dict:
            df_preview = pd.concat(transformed_data_dict.values()).head(20)
            st.write("🔍 **Aperçu des données transformées :**")
            st.dataframe(df_preview)

        # 🛠 **Mise à jour des Partner ID si un fichier est fourni**
        if uploaded_update_file is not None:
            st.success("✅ **Fichier de mise à jour des Partner ID chargé avec succès !**")

            # Charger le fichier de mise à jour
            df_update = pd.read_excel(uploaded_update_file)

            if df_update.shape[1] != 2:
                st.error(
                    "⚠️ **Le fichier de mise à jour doit contenir 2 colonnes : Ancien partner_id et Nouveau partner_id.**")
            else:
                update_dict = df_update.set_index(df_update.columns[0])[df_update.columns[1]].to_dict()

                # Mise à jour du `partner_id` dans **toutes** les feuilles du fichier transformé
                for journal, df in transformed_data_dict.items():
                    if journal == "ODGEST" and "Écritures comptables/Partenaire" in df.columns:
                        df["Écritures comptables/Partenaire"] = df["Écritures comptables/Partenaire"].map(
                            update_dict).fillna(df["Écritures comptables/Partenaire"])
                    elif "partner_id" in df.columns:
                        df["partner_id"] = df["partner_id"].map(update_dict).fillna(df["partner_id"])

                    transformed_data_dict[journal] = df  # Mise à jour du dictionnaire

                output_buffer_updated = BytesIO()
                with pd.ExcelWriter(output_buffer_updated, engine='openpyxl') as writer:
                    for journal, df in transformed_data_dict.items():
                        df.to_excel(writer, sheet_name=journal, index=False)
                output_buffer_updated.seek(0)

                # 📥 **Télécharger le fichier transformé mis à jour**
                st.download_button(
                    label="📥 **Télécharger le fichier transformé mis à jour**",
                    data=output_buffer_updated,
                    file_name="HMS_RESULT_UPDATED.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.success(
                    "✅ **Mise à jour des Partner ID effectuée avec succès sur toutes les feuilles, y compris ODGEST !**")

                # 🔍 Extraction des partner_id absents après mise à jour
                df_missing_partners = extract_ids_missing_from_update(df_update, transformed_data_dict)

                if not df_missing_partners.empty:
                    st.warning(
                        "⚠️ Certains `partner_id` du fichier de mise à jour sont absents dans le fichier transformé.")

                    # 📄 Réécriture du fichier avec la feuille MISSING_IDS
                    output_buffer_with_missing = BytesIO()
                    with pd.ExcelWriter(output_buffer_with_missing, engine='openpyxl') as writer:
                        # Réécriture de toutes les feuilles transformées
                        for journal, df in transformed_data_dict.items():
                            df.to_excel(writer, sheet_name=journal, index=False)
                        # Ajout des ids manquants
                        df_missing_partners.to_excel(writer, sheet_name="MISSING_IDS", index=False)
                    output_buffer_with_missing.seek(0)

                    # 📥 Bouton de téléchargement avec la feuille MISSING_IDS
                    st.download_button(
                        label="📥 Télécharger le fichier final avec les partner_id manquants",
                        data=output_buffer_with_missing,
                        file_name="HMS_RESULT_UPDATED_WITH_MISSING.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # 👁️ Affichage des partner_id manquants
                    st.subheader("📋 Partner ID absents dans les feuilles transformées :")
                    st.dataframe(df_missing_partners)
                else:
                    st.success(
                        "✅ Tous les `partner_id` du fichier de mise à jour sont présents dans le fichier transformé.")

# 🟠 Onglet 2 : Extraction des commentaires
with tab2:
    st.header("🔄 Extraction des commentaires")

    uploaded_file_2 = st.file_uploader("📥 **Téléchargez le fichier source HMS (Excel)**", type=['xlsx'], key="file2")

    if uploaded_file_2 is not None:
        st.success("✅ **Fichier chargé avec succès !**")
        df_source_2 = pd.read_excel(uploaded_file_2)

        df_extracted = extract_comments(df_source_2)  # 💡 L'algorithme d'origine est conservé

        output = BytesIO()
        df_extracted.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button("📥 **Télécharger les commentaires extraits**", data=output, file_name="Commentaires.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.write("🔍 **Aperçu des commentaires extraits :**")
        st.dataframe(df_extracted)


# 🔵 Onglet 3 : Extraction avancée
with tab3:
    st.header("📌 Extraction avancée")

    uploaded_file_3 = st.file_uploader("📥 **Téléchargez le fichier source HMS (Excel)**", type=['xlsx'], key="file3")

    if uploaded_file_3 is not None:
        st.success("✅ **Fichier chargé avec succès !**")
        df_source_3 = pd.read_excel(uploaded_file_3)

        df_advanced = extract_second_last_comment(df_source_3)  # 💡 L'algorithme d'origine est conservé

        output = BytesIO()
        df_advanced.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button("📥 **Télécharger les données extraites**", data=output, file_name="Extraction_Avancee.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.write("🔍 **Aperçu des données extraites :**")
        st.dataframe(df_advanced)


with tab4:
    st.header("📂 Extraction vers Odoo")

    uploaded_hms = st.file_uploader("📂 Téléchargez le fichier HMS (Excel)", type=["xlsx"], key="hms_file")
    uploaded_destination = st.file_uploader("📂 Téléchargez le fichier modèle de destination (Excel)", type=["xlsx"], key="destination_file")

    if uploaded_hms and uploaded_destination:
        st.success("✅ Fichiers chargés avec succès !")

        df_hms = pd.read_excel(uploaded_hms)
        df_destination = pd.read_excel(uploaded_destination)

        # Appel de la fonction de transformation
        df_transformed, df_unmatched = transform_hms_to_odoo(df_hms, df_destination)

        # Génération du fichier Excel avec deux feuilles
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_transformed.to_excel(writer, sheet_name="Données transformées", index=False)
            if not df_unmatched.empty:
                df_unmatched.to_excel(writer, sheet_name="Non présents dans modèle", index=False)
        output.seek(0)

        st.download_button(
            label="📥 Télécharger le fichier transformé",
            data=output,
            file_name="HMS_to_ODOO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Aperçu des deux dataframes
        st.write("🔍 **Aperçu : Données transformées (feuille 1)**")
        st.dataframe(df_transformed.head(30))

        if not df_unmatched.empty:
            st.write("⚠️ **Aperçu : Nouveaux account-id non présents dans le modèle (feuille 2)**")
            st.dataframe(df_unmatched.head(30))

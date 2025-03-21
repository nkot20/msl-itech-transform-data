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


def transform_hms_to_odoo(df_hms, df_destination):
    df_filtered = df_hms[df_hms["journal"].isin(["VEN", "AC2"])].copy()

    # Assurer que 'montant-gen' est bien un float
    df_filtered["montant-gen"] = df_filtered["montant-gen"].replace(",", ".", regex=True)
    df_filtered["montant-gen"] = pd.to_numeric(df_filtered["montant-gen"], errors="coerce")

    # Trier les données par account-id et document number pour garantir l'ordre
    df_filtered.sort_values(by=["account-id", "docnumber"], inplace=True)

    # Création d'un dictionnaire pour stocker les groupes
    grouped_data = df_filtered.groupby(["account-id", "docnumber"])

    for (account_id, doc_number), group in grouped_data:
        if account_id not in df_destination["x_studio_rf_wb"].values:
            continue  # Si le account-id n'existe pas dans le fichier destination, on passe

        # Trouver l'index de la ligne correspondante dans le fichier destination
        dest_index = df_destination[df_destination["x_studio_rf_wb"] == account_id].index[0]

        # Vérifier combien de groupes existent déjà pour cet `account-id`
        existing_groups = [col for col in df_destination.columns if col.startswith("x_studio_code_analytique")]
        group_count = df_destination.loc[dest_index, existing_groups].notna().sum()

        # Déterminer le suffixe en fonction du nombre de groupes existants
        suffix = f"_{group_count}" if group_count > 0 else ""

        # Vérifier que les colonnes existent dans df_destination avant de les modifier
        analytical_col = f"x_studio_code_analytique{suffix}"
        address_col = f"x_studio_adresse{suffix}"

        if analytical_col in df_destination.columns:
            df_destination.at[dest_index, analytical_col] = str(extract_analytical_code(group.iloc[0]["comment-int"]))

        if address_col in df_destination.columns:
            df_destination.at[dest_index, address_col] = str(extract_address(group.iloc[0]["comment-int"]))

        # Parcourir les lignes du groupe et remplir les valeurs correspondantes
        for _, row in group.iterrows():
            account_gl = row["accountgl"]
            montant_gen = row["montant-gen"]

            if account_gl in mapping_accounts:
                column_name = mapping_accounts[account_gl] + suffix
                if column_name in df_destination.columns:  # Vérifier que la colonne existe avant de modifier
                    df_destination.at[dest_index, column_name] = float(montant_gen)  # 🔹 Conversion explicite en float

    return df_destination

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
                st.write(f"🛠️ **Traitement du journal :** `{journal}`")
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

    # Téléchargement du fichier destination (template)
    uploaded_destination = st.file_uploader("📂 Téléchargez le fichier modèle de destination (Excel)", type=["xlsx"],
                                            key="destination_file")

    if uploaded_hms and uploaded_destination:
        st.success("✅ Fichiers chargés avec succès !")

        df_hms = pd.read_excel(uploaded_hms)
        df_destination = pd.read_excel(uploaded_destination)

        df_transformed = transform_hms_to_odoo(df_hms, df_destination)

        # Générer le fichier transformé
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_transformed.to_excel(writer, index=False)
        output.seek(0)

        # Bouton de téléchargement
        st.download_button(
            label="📥 Télécharger le fichier transformé",
            data=output,
            file_name="HMS_to_ODOO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Afficher un aperçu des données transformées
        st.write("🔍 **Aperçu des données transformées :**")
        st.dataframe(df_transformed.head(50))
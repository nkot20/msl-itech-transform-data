import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


# ======= FONCTION 1 : Transformation du fichier HMS (déjà existante) =======
def prepare_data_for_journal(df, journal_name):
    df_filtered = df[df['journal'] == journal_name].copy()

    # Suppression des lignes en fonction du journal
    if journal_name in ["VEN", "GESTIO"]:
        df_filtered = df_filtered[df_filtered['accountgl'] != 400000]
    if journal_name == "AC2":
        df_filtered = df_filtered[df_filtered['accountgl'] != 440100]

    # Génération du champ 'name' en fonction du journal
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

    # Conversion des dates en format yyyy.mm.dd
    df_filtered['datedoc'] = pd.to_datetime(df_filtered['datedoc']).dt.strftime('%Y.%m.%d')
    df_filtered['duedate'] = pd.to_datetime(df_filtered['duedate']).dt.strftime('%Y.%m.%d')

    # **Ajout de la colonne 'Référence'**
    if journal_name in ["VEN", "GESTIO", "AC2"]:
        reference_account = 400000 if journal_name in ["VEN", "GESTIO"] else 440100

        # Récupérer `comment-int` pour chaque groupe (docnumber + account-id)
        df_filtered['Référence'] = df_filtered.groupby(['docnumber', 'account-id'])['comment-int'].transform(
            lambda x: x[df_filtered['accountgl'] == reference_account].iloc[0]
            if (df_filtered['accountgl'] == reference_account).any() else x.iloc[0]
        )
    else:
        df_filtered['Référence'] = df_filtered['comment-int']

    # Cas spécifique pour les journaux VEN, AC2 et GESTIO
    if journal_name in ["VEN", "GESTIO"]:
        price_unit = np.where(df_filtered['D-C'] == 'D', -df_filtered['montant-gen'], df_filtered['montant-gen'])
    elif journal_name == "AC2":
        price_unit = np.where(df_filtered['D-C'] == 'D', df_filtered['montant-gen'], -df_filtered['montant-gen'])
    else:
        price_unit = np.zeros(len(df_filtered))  # Valeur par défaut

    # Création du fichier destination
    df_destination = pd.DataFrame({
        'name': df_filtered['name'],
        'partner_id': df_filtered['account-id'],
        'invoice_date': df_filtered['datedoc'],
        'invoice_date_due': df_filtered['duedate'],
        'journal_code': df_filtered['journal'],
        'account_id': df_filtered['accountgl'],
        'invoice_line_ids/price_unit': price_unit,
        'Référence': df_filtered['Référence'],
    })

    return df_destination


# ======= FONCTION 2 : Extraction des commentaires =======
def extract_comments(df):
    # Filtrer uniquement les journaux AC2 et VEN
    df_filtered = df[df['journal'].isin(["AC2", "VEN"])].copy()

    # Filtrer uniquement les comptes 400000 et 440100
    df_filtered = df_filtered[df_filtered['accountgl'].isin([400000, 440100])]

    # Extraire uniquement la dernière partie du champ comment-int après le dernier "/"
    df_filtered['comment-int'] = df_filtered['comment-int'].apply(lambda x: x.split("/")[-1] if isinstance(x, str) else x)

    # Conserver uniquement les colonnes demandées
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


# ======= INTERFACE UTILISATEUR STREAMLIT =======
st.title("📂 MSL-ITECH - Transformation de fichier Excel HMS")

# Création des onglets
tab1, tab2, tab3 = st.tabs(["🚀 Transformation du fichier (1)", "📝 Extraction des commentaires (2)", "📌 Extraction avancée (3)"])

with tab1:
    st.header("🚀 Transformation de fichier HMS vers ODOO")

    uploaded_file = st.file_uploader("📂 Téléchargez le fichier source HMS (Excel)", type=['xlsx'])

    if uploaded_file is not None:
        st.success("✅ Fichier téléchargé avec succès !")
        df_source = pd.read_excel(uploaded_file)

        journals = df_source['journal'].unique()
        output_buffer = BytesIO()

        all_transformed_data = []
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            for journal in journals:
                st.write(f"🛠️ Traitement du journal : {journal}")
                df_journal = prepare_data_for_journal(df_source, journal)
                if not df_journal.empty:
                    df_journal.to_excel(writer, sheet_name=journal, index=False)
                    all_transformed_data.append(df_journal)

        output_buffer.seek(0)

        # Affichage des données transformées
        if all_transformed_data:
            df_preview = pd.concat(all_transformed_data).head(50)
            st.write("🔍 **Aperçu des premières lignes des données transformées :**")
            st.dataframe(df_preview)


        st.download_button(
            label="📥 Télécharger le fichier transformé",
            data=output_buffer,
            file_name="HMS_RESULT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab2:
    st.header("📝 Extraction des commentaires spécifiques")

    uploaded_file_2 = st.file_uploader("📂 Téléchargez le fichier source HMS (Excel)", type=['xlsx'], key="file2")

    if uploaded_file_2 is not None:
        st.success("✅ Fichier téléchargé avec succès !")
        df_source_2 = pd.read_excel(uploaded_file_2)

        # Appliquer l'extraction des commentaires
        df_extracted = extract_comments(df_source_2)

        # Affichage du tableau dans l'application
        st.write("🔍 Aperçu des données extraites :")
        st.dataframe(df_extracted)

        # Création du fichier téléchargeable
        output_buffer_2 = BytesIO()
        with pd.ExcelWriter(output_buffer_2, engine='openpyxl') as writer:
            df_extracted.to_excel(writer, sheet_name="Commentaires_Extraits", index=False)

        output_buffer_2.seek(0)
        st.download_button(
            label="📥 Télécharger les commentaires extraits",
            data=output_buffer_2,
            file_name="Commentaires_Extraits.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 🔵 Onglet 3 : Extraction avancée
with tab3:
    st.header("📌 Extraction avancée des commentaires")

    uploaded_file_3 = st.file_uploader("📂 Téléchargez le fichier source HMS (Excel)", type=['xlsx'], key="file3")

    if uploaded_file_3 is not None:
        st.success("✅ Fichier téléchargé avec succès !")
        df_source_3 = pd.read_excel(uploaded_file_3)

        df_advanced = extract_second_last_comment(df_source_3)

        st.write("🔍 Aperçu des données extraites :")
        st.dataframe(df_advanced)
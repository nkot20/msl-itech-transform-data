import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


# Fonction de transformation
def prepare_data_for_journal(df, journal_name):
    df_filtered = df[df['journal'] == journal_name].copy()

    # Suppression des lignes en fonction du journal
    if journal_name in ["VEN", "GESTIO"]:
        df_filtered = df_filtered[df_filtered['accountgl'] != 400000]
    if journal_name == "AC2":
        df_filtered = df_filtered[df_filtered['accountgl'] != 440100]

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
    if journal_name in ["GESTIO", "AC2"]:
        reference_account = 400000 if journal_name in ["VEN", "GESTIO"] else 440100

        # Sélectionner la première valeur de 'comment-int' par 'account-id' pour éviter les doublons
        reference_dict = df[df['accountgl'] == reference_account].groupby('account-id')['comment-int'].first().to_dict()

        # Appliquer la référence si elle existe, sinon utiliser 'comment-int' de la ligne courante
        df_filtered['Référence'] = df_filtered['account-id'].map(reference_dict).fillna(df_filtered['comment-int'])
    else:
        df_filtered['Référence'] = df_filtered['comment-int']

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


# Interface utilisateur avec Streamlit
st.title("MSL-ITECH Transformation de fichier Excel HMS")

uploaded_file = st.file_uploader("Téléchargez le fichier source HMS (Excel)", type=['xlsx'])

if uploaded_file is not None:
    st.success("Fichier téléchargé avec succès !")
    df_source = pd.read_excel(uploaded_file)

    journals = df_source['journal'].unique()
    output_buffer = BytesIO()

    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        for journal in journals:
            st.write(f"Traitement du journal : {journal}")
            df_journal = prepare_data_for_journal(df_source, journal)
            if not df_journal.empty:
                df_journal.to_excel(writer, sheet_name=journal, index=False)

    output_buffer.seek(0)
    st.download_button(
        label="Télécharger le fichier transformé pour intégrer dans ODOO",
        data=output_buffer,
        file_name="fichier_transformé.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

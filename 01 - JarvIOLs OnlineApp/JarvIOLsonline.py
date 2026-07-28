import os
import re
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURATION DE LA PAGE STREAMLIT
# ---------------------------------------------------------
st.set_page_config(page_title="Analyse Optique MTF", layout="wide")

# ---------------------------------------------------------
# 1. GESTION DES CHEMINS ET CHARGEMENT DES FICHIERS EXCEL
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Chemins des fichiers Excel dans le même dossier
PATH_EXCEL_PYTHONZ = os.path.join(BASE_DIR, "Excel_PythonZ.xlsx")
PATH_EXCEL_FS = os.path.join(BASE_DIR, "Excel_FS.xlsx")


@st.cache_data
def load_excel_data(file_path):
    """Charge et nettoie les données d'un fichier Excel."""
    if not os.path.exists(file_path):
        return None
    try:
        df = pd.read_excel(file_path)
        # Nettoyage des noms de colonnes
        df.columns = df.columns.str.strip()

        # Formatage de l'ouverture pupillaire pour correspondre aux chaînes (ex: "3.5", "3.5mm")
        if "Ouverture pupillaire" in df.columns:
            df["Ouverture pupillaire"] = df["Ouverture pupillaire"].astype(str)
        if "Ouverture" in df.columns:
            df["Ouverture"] = df["Ouverture"].astype(str)

        return df
    except Exception as e:
        st.error(f"Erreur lors du chargement de {os.path.basename(file_path)} : {e}")
        return None


# Chargement des deux dataframes
df_pythonz = load_excel_data(PATH_EXCEL_PYTHONZ)
df_fs = load_excel_data(PATH_EXCEL_FS)


# Helper function pour renommer les colonnes de manière uniforme si besoin
def normalize_df_columns(df):
    if df is None:
        return None
    rename_mapping = {
        "Nom de l'IOL": "Nom_IOL",
        "Nom de l'I": "Nom_IOL",
        "Ouverture pupillaire": "Ouverture",
        "Condition de lumière": "Condition",
        "Type de cornée": "Type_cornee",
        "Type de cc": "Type_cornee",
    }
    return df.rename(columns=rename_mapping)


df_pythonz = normalize_df_columns(df_pythonz)
df_fs = normalize_df_columns(df_fs)


# ---------------------------------------------------------
# 2. BARRE LATÉRALE (SIDEBAR) - FILTRES ET PARAMÈTRES
# ---------------------------------------------------------
st.sidebar.title("Configuration & Filtres")

# Extraction des options de filtrage disponibles (en se basant sur le DF principal ou FS)
reference_df = df_fs if df_fs is not None else df_pythonz

if reference_df is not None:
    iol_list = reference_df["Nom_IOL"].unique().tolist()
    aperture_list = reference_df["Ouverture"].unique().tolist()
    light_list = reference_df["Condition"].unique().tolist()
    cornea_list = reference_df["Type_cornee"].unique().tolist()

    selected_iol = st.sidebar.selectbox("Modèle IOL", iol_list)
    selected_aperture = st.sidebar.selectbox("Ouverture pupillaire", aperture_list)
    selected_light = st.sidebar.selectbox("Condition de lumière", light_list)
    selected_cornea = st.sidebar.selectbox("Type de cornée", cornea_list)
else:
    st.sidebar.warning("Veuillez vérifier la présence des fichiers Excel dans le dossier.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("Graphes - Échelles Y")

# Limites Y globales/pour les graphes MTF / VA / MTFa
y_min_global = st.sidebar.number_input("Y min (MTF/VA/MTFa)", value=0.0, step=0.1)
y_max_global = st.sidebar.number_input("Y max (MTF/VA/MTFa)", value=1.0, step=0.1)

st.sidebar.markdown("---")
st.sidebar.subheader("Paramètres Frequency Sweep")

# 2) Menu déroulant pour le graphe de Frequency Sweep
fs_column_choice = st.sidebar.selectbox(
    "Condition FS à afficher",
    options=["FS at distance", "FS at 1.50D", "FS at 2.5D"],
    index=0,
)

# 2) Champ pour définir la valeur Frequency max (entre 0 et 600)
fs_max_freq = st.sidebar.number_input(
    "Frequency max",
    min_value=0,
    max_value=600,
    value=100,
    step=10,
    help="Valeur maximale de la fréquence affichée sur l'axe X (0 à 600)",
)

# 6) Gestion des limites Y spécifique au graphe Frequency Sweep
fs_y_min = st.sidebar.number_input("Y min (Frequency Sweep)", value=0.0, step=0.1)
fs_y_max = st.sidebar.number_input("Y max (Frequency Sweep)", value=1.0, step=0.1)


# ---------------------------------------------------------
# 3. AFFICHAGE DES GRAPHIQUES
# ---------------------------------------------------------
st.title("Analyse Optique IOL")

# --- GRAPHE THROUGH FOCUS MTF ---
st.header("Through Focus MTF")

if df_pythonz is not None:
    # Filtrage du DataFrame PythonZ
    df_tf_filtered = df_pythonz[
        (df_pythonz["Nom_IOL"] == selected_iol)
        & (df_pythonz["Ouverture"] == str(selected_aperture))
        & (df_pythonz["Condition"] == selected_light)
        & (df_pythonz["Type_cornee"] == selected_cornea)
    ]

    if not df_tf_filtered.empty and "Defocus" in df_tf_filtered.columns and "MTF" in df_tf_filtered.columns:
        fig_tf, ax_tf = plt.subplots(figsize=(9, 4.5))
        ax_tf.plot(
            df_tf_filtered["Defocus"],
            df_tf_filtered["MTF"],
            marker="o",
            linestyle="-",
            color="#2ca02c",
            label="Through Focus MTF",
        )
        ax_tf.set_title(f"Through Focus MTF - {selected_iol} ({selected_aperture}, {selected_light}, {selected_cornea})")
        ax_tf.set_xlabel("Defocus (D)")
        ax_tf.set_ylabel("MTF")
        ax_tf.set_ylim(y_min_global, y_max_global)
        ax_tf.grid(True, linestyle="--", alpha=0.5)
        ax_tf.legend()
        st.pyplot(fig_tf)
    else:
        st.info("Aucune donnée Through Focus MTF pour cette sélection.")
else:
    st.warning("Fichier Excel_PythonZ non trouvé ou vide.")

st.markdown("---")

# --- 1) GRAPHE FREQUENCY SWEEP (En dessous de Through Focus MTF) ---
st.header("Frequency Sweep")

if df_fs is not None:
    # 4) Filtrage selon l'IOL, l'ouverture, la lumière et la cornée
    df_fs_filtered = df_fs[
        (df_fs["Nom_IOL"] == selected_iol)
        & (df_fs["Ouverture"] == str(selected_aperture))
        & (df_fs["Condition"] == selected_light)
        & (df_fs["Type_cornee"] == selected_cornea)
    ]

    # 4) Imposer la valeur max de Frequency (0 à Frequency max)
    df_fs_filtered = df_fs_filtered[
        (df_fs_filtered["Frequency"] >= 0)
        & (df_fs_filtered["Frequency"] <= fs_max_freq)
    ]

    if not df_fs_filtered.empty:
        # Vérification si la colonne demandée ("FS at distance", "FS at 1.50D", "FS at 2.5D") existe
        if fs_column_choice in df_fs_filtered.columns:
            fig_fs, ax_fs = plt.subplots(figsize=(9, 4.5))

            # 4) X = Frequency, 5) Y = Colonne choisie dans le menu déroulant
            ax_fs.plot(
                df_fs_filtered["Frequency"],
                df_fs_filtered[fs_column_choice],
                marker="",
                linestyle="-",
                linewidth=2,
                color="#1f77b4",
                label=f"{fs_column_choice}",
            )

            ax_fs.set_title(f"Frequency Sweep - {fs_column_choice}")
            ax_fs.set_xlabel("Frequency (lines/mm)")
            ax_fs.set_ylabel("MTF / Value")

            # 6) Gestion des limites de l'axe Y
            ax_fs.set_ylim(fs_y_min, fs_y_max)
            ax_fs.grid(True, linestyle="--", alpha=0.5)
            ax_fs.legend()

            st.pyplot(fig_fs)
        else:
            st.error(f"La colonne '{fs_column_choice}' n'a pas été trouvée dans le fichier Excel_FS.")
    else:
        st.warning("Aucune donnée Frequency Sweep trouvée pour ces critères.")
else:
    st.error("Le fichier 'Excel_FS.xlsx' est introuvable au même emplacement que l'application.")

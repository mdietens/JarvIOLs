import os
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURATION DE LA PAGE STREAMLIT
# ---------------------------------------------------------
st.set_page_config(page_title="Analyse Optique MTF & FS", layout="wide")

# ---------------------------------------------------------
# 1. CHARGEMENT ET NETTOYAGE DES FICHIERS EXCEL
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PATH_EXCEL_PYTHONZ = os.path.join(BASE_DIR, "Excel_PythonZ.xlsx")
PATH_EXCEL_FS = os.path.join(BASE_DIR, "Excel_FS.xlsx")


def find_column(df, candidates):
    """Recherche une colonne indépendamment des espaces ou de la casse."""
    for cand in candidates:
        for col in df.columns:
            if str(col).strip().lower() == cand.lower():
                return col
    return None


@st.cache_data
def load_excel_fs(file_path):
    """Charge et nettoie spécifiquement le fichier Excel_FS."""
    if not os.path.exists(file_path):
        return None
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()

        # Mapping avec les en-têtes exacts indiqués
        col_iol = find_column(df, ["IOL_Model", "Nom de l'IOL", "Nom_IOL"])
        col_pupil = find_column(df, ["Pupil_Size", "Ouverture pupillaire", "Ouverture"])
        col_light = find_column(df, ["Light_Condition", "Condition de lumière", "Condition"])
        col_cornea = find_column(df, ["Cornea_Type", "Type de cornée", "Type_cornee"])

        rename_map = {}
        if col_iol: rename_map[col_iol] = "IOL_Model"
        if col_pupil: rename_map[col_pupil] = "Pupil_Size"
        if col_light: rename_map[col_light] = "Light_Condition"
        if col_cornea: rename_map[col_cornea] = "Cornea_Type"

        df = df.rename(columns=rename_map)

        # Uniformisation en chaîne de caractères pour les filtres textuels
        for col in ["IOL_Model", "Pupil_Size", "Light_Condition", "Cornea_Type"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.replace(",", ".")

        # Nettoyage des valeurs numériques (remplacement des virgules par des points)
        num_cols = ["Frequency", "FS at distance", "FS at 1.50D", "FS at 2.5D"]
        for col in num_cols:
            if col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype(str).str.replace(",", ".").astype(float)
                else:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

        return df
    except Exception as e:
        st.error(f"Erreur de lecture du fichier {os.path.basename(file_path)} : {e}")
        return None


@st.cache_data
def load_excel_pythonz(file_path):
    """Charge le fichier Excel_PythonZ."""
    if not os.path.exists(file_path):
        return None
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()

        col_iol = find_column(df, ["IOL_Model", "Nom de l'IOL", "Nom_IOL"])
        col_pupil = find_column(df, ["Pupil_Size", "Ouverture pupillaire", "Ouverture"])
        col_light = find_column(df, ["Light_Condition", "Condition de lumière", "Condition"])
        col_cornea = find_column(df, ["Cornea_Type", "Type de cornée", "Type_cornee"])

        rename_map = {}
        if col_iol: rename_map[col_iol] = "IOL_Model"
        if col_pupil: rename_map[col_pupil] = "Pupil_Size"
        if col_light: rename_map[col_light] = "Light_Condition"
        if col_cornea: rename_map[col_cornea] = "Cornea_Type"

        df = df.rename(columns=rename_map)

        for col in ["IOL_Model", "Pupil_Size", "Light_Condition", "Cornea_Type"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.replace(",", ".")

        return df
    except Exception as e:
        st.error(f"Erreur de lecture de {os.path.basename(file_path)} : {e}")
        return None


df_fs = load_excel_fs(PATH_EXCEL_FS)
df_pythonz = load_excel_pythonz(PATH_EXCEL_PYTHONZ)

# ---------------------------------------------------------
# 2. PANNEAU LATÉRAL (SIDEBAR)
# ---------------------------------------------------------
st.sidebar.title("Filtres & Configuration")

# Sélection de la source pour peupler les listes déroulantes
ref_df = df_fs if df_fs is not None else df_pythonz

if ref_df is not None:
    iol_options = sorted(ref_df["IOL_Model"].unique().tolist())
    pupil_options = sorted(ref_df["Pupil_Size"].unique().tolist())
    light_options = sorted(ref_df["Light_Condition"].unique().tolist())
    cornea_options = sorted(ref_df["Cornea_Type"].unique().tolist())

    selected_iol = st.sidebar.selectbox("Sélectionner l'IOL Model", iol_options)
    selected_pupil = st.sidebar.selectbox("Sélectionner Pupil Size", pupil_options)
    selected_light = st.sidebar.selectbox("Sélectionner Light Condition", light_options)
    selected_cornea = st.sidebar.selectbox("Sélectionner Cornea Type", cornea_options)
else:
    st.sidebar.error("Aucun fichier Excel trouvé à l'emplacement indiqué.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.subheader("Configuration Frequency Sweep")

# 2) Menu déroulant pour Frequency Sweep
fs_choice = st.sidebar.selectbox(
    "Condition FS",
    options=["FS at distance", "FS at 1.50D", "FS at 2.5D"],
    index=0,
)

# 2) Champ vide / nombre pour Frequency max compris entre 0 et 600
fs_max_freq = st.sidebar.number_input(
    "Frequency max",
    min_value=0,
    max_value=600,
    value=100,
    step=10,
    help="Valeur maximale de la fréquence affichée (0 à 600)",
)

# 6) Gestion des limites d'axe Y pour Frequency Sweep
st.sidebar.markdown("**Échelle Y - Frequency Sweep**")
fs_y_min = st.sidebar.number_input("Y min (FS)", value=0.0, step=0.1)
fs_y_max = st.sidebar.number_input("Y max (FS)", value=1.0, step=0.1)

# Échelle Y globale
st.sidebar.markdown("---")
st.sidebar.markdown("**Échelle Y - Autres Graphes (MTF/VA)**")
global_y_min = st.sidebar.number_input("Y min (Autres)", value=0.0, step=0.1)
global_y_max = st.sidebar.number_input("Y max (Autres)", value=1.0, step=0.1)

# ---------------------------------------------------------
# 3. AFFICHAGE ET TRACÉ DES GRAPHES
# ---------------------------------------------------------
st.title("Analyse Optique")

# --- THROUGH FOCUS MTF ---
st.header("Through Focus MTF")

if df_pythonz is not None:
    df_tf_filtered = df_pythonz[
        (df_pythonz["IOL_Model"] == selected_iol)
        & (df_pythonz["Pupil_Size"] == selected_pupil)
        & (df_pythonz["Light_Condition"] == selected_light)
        & (df_pythonz["Cornea_Type"] == selected_cornea)
    ]

    col_defocus = find_column(df_pythonz, ["Defocus", "Defocus (D)"])
    col_mtf = find_column(df_pythonz, ["MTF", "Through focus MTF"])

    if not df_tf_filtered.empty and col_defocus and col_mtf:
        fig_tf, ax_tf = plt.subplots(figsize=(9, 4.5))
        ax_tf.plot(
            df_tf_filtered[col_defocus],
            df_tf_filtered[col_mtf],
            marker="o",
            linestyle="-",
            color="#2ca02c",
            label="Through Focus MTF",
        )
        ax_tf.set_title(f"Through focus MTF - {selected_iol} ({selected_pupil}, {selected_light}, {selected_cornea})")
        ax_tf.set_xlabel("Defocus (D)")
        ax_tf.set_ylabel("MTF")
        ax_tf.set_ylim(global_y_min, global_y_max)
        ax_tf.grid(True, linestyle="--", alpha=0.5)
        ax_tf.legend()
        st.pyplot(fig_tf)
    else:
        st.info("Aucune donnée disponible pour le Through focus MTF avec ces critères.")
else:
    st.warning("Fichier Excel_PythonZ.xlsx introuvable.")

st.markdown("---")

# --- 1) GRAPHE FREQUENCY SWEEP (EN DESSOUS DE THROUGH FOCUS MTF) ---
st.header("Frequency Sweep")

if df_fs is not None:
    # 4) Filtrage dynamique selon IOL, Ouverture, Lumière, Cornée
    df_fs_filtered = df_fs[
        (df_fs["IOL_Model"] == selected_iol)
        & (df_fs["Pupil_Size"] == selected_pupil)
        & (df_fs["Light_Condition"] == selected_light)
        & (df_fs["Cornea_Type"] == selected_cornea)
    ]

    # 4) Filtrage de la fréquence max entre 0 et la valeur fs_max_freq
    df_fs_filtered = df_fs_filtered[
        (df_fs_filtered["Frequency"] >= 0)
        & (df_fs_filtered["Frequency"] <= fs_max_freq)
    ]

    if not df_fs_filtered.empty:
        if fs_choice in df_fs_filtered.columns:
            fig_fs, ax_fs = plt.subplots(figsize=(9, 4.5))

            # 4) X = Frequency, 5) Y = Choix du menu déroulant (at distance, 1.50D, 2.5D)
            ax_fs.plot(
                df_fs_filtered["Frequency"],
                df_fs_filtered[fs_choice],
                marker="",
                linestyle="-",
                linewidth=2,
                color="#1f77b4",
                label=fs_choice,
            )

            ax_fs.set_title(f"Frequency Sweep - {fs_choice}")
            ax_fs.set_xlabel("Frequency (lines/mm)")
            ax_fs.set_ylabel("MTF")
            
            # 6) Gestion des bornes Min-Max de l'axe Y
            ax_fs.set_ylim(fs_y_min, fs_y_max)
            ax_fs.grid(True, linestyle="--", alpha=0.5)
            ax_fs.legend()

            st.pyplot(fig_fs)
        else:
            st.error(f"La colonne '{fs_choice}' est absente du fichier Excel_FS.")
    else:
        st.warning("Aucune donnée Frequency Sweep trouvée pour cette sélection.")
else:
    st.error("Le fichier Excel_FS.xlsx est introuvable au même emplacement que l'application.")

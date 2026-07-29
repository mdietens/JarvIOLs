import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as px
import plotly.express as px_express
from PIL import Image
import os

# --------------------------------------------------------------------
# 1. CONFIGURATION DE LA PAGE
# --------------------------------------------------------------------
st.set_page_config(
    page_title="IOL Comparative Analysis Tool",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --------------------------------------------------------------------
# 2. AUTHENTIFICATION SIMPLIFIÉE
# --------------------------------------------------------------------
def check_password():
    """Vérification d'authentification pour déploiement Cloud."""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔒 Connexion requise")
    password = st.text_input("Veuillez entrer le mot de passe :", type="password")
    if st.button("Se connecter"):
        if "PASSWORD" in st.secrets and password == st.secrets["PASSWORD"]:
            st.session_state["password_correct"] = True
            st.rerun()
        elif password == "iol2026":  # Mot de passe par défaut/fallback
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect")
    return False

if not check_password():
    st.stop()

# --------------------------------------------------------------------
# 3. CHARGEMENT ET PRÉPARATION DES DONNÉES
# --------------------------------------------------------------------
@st.cache_data
def load_data(file_main, file_fs=None):
    df_main = pd.read_excel(file_main) if file_main else pd.DataFrame()
    df_fs = pd.read_excel(file_fs) if file_fs else pd.DataFrame()
    return df_main, df_fs

st.sidebar.title("📁 Upload & Filtres")

file_1 = st.sidebar.file_uploader("Fichier principal (Through Focus & VA)", type=["xlsx", "xls"])
file_2 = st.sidebar.file_uploader("Fichier secondaire (Frequency Sweep)", type=["xlsx", "xls"])

if file_1 is not None:
    df, df_fs = load_data(file_1, file_2)
else:
    st.info("Veuillez charger au moins le fichier principal dans la barre latérale pour démarrer.")
    st.stop()

# Extraction des options disponibles
available_models = sorted(df['IOL_Model'].dropna().astype(str).unique()) if 'IOL_Model' in df.columns else []
available_pupils = sorted(df['Pupil_Size'].dropna().unique()) if 'Pupil_Size' in df.columns else []
available_lights = sorted(df['Light_Condition'].dropna().astype(str).unique()) if 'Light_Condition' in df.columns else []
available_corneas = sorted(df['Cornea_Type'].dropna().astype(str).unique()) if 'Cornea_Type' in df.columns else []

st.sidebar.markdown("---")
st.sidebar.subheader("🎛️ Paramètres Globaux")

selected_pupil = st.sidebar.selectbox("Pupil Size (mm)", available_pupils if available_pupils else [3.5])
selected_light = st.sidebar.selectbox("Light Condition", available_lights if available_lights else ["Photopic"])
selected_cornea = st.sidebar.selectbox("Cornea Type", available_corneas if available_corneas else ["Monofocal"])

num_configs = st.sidebar.slider("Nombre de configurations à comparer", min_value=1, max_value=6, value=2)

# --------------------------------------------------------------------
# 4. INTERFACE PRINCIPALE & FILTRAGE DES DONNÉES
# --------------------------------------------------------------------
st.title("📊 IOL Comparative Analysis Tool")

main_cols = st.columns(num_configs)
selected_iols = []
filtered_dfs = []
filtered_fs_dfs = []

for i in range(num_configs):
    with main_cols[i]:
        st.header(f"Config {chr(65+i)}")
        default_idx = min(i, len(available_models) - 1) if available_models else 0
        iol_selected = st.selectbox(
            f"Choose IOL ({chr(65+i)})", 
            available_models, 
            index=default_idx, 
            key=f"select_iol_{i}"
        )
        selected_iols.append(iol_selected)
        
        # Filtrage fichier 1 (Through Focus & VA)
        df_f = df[
            (df['IOL_Model'] == iol_selected) & 
            (df['Pupil_Size'] == selected_pupil) & 
            (df['Light_Condition'] == selected_light) & 
            (df['Cornea_Type'] == selected_cornea)
        ].sort_values(by='Defocus') if not df.empty else pd.DataFrame()
        filtered_dfs.append(df_f)

        # ----------------------------------------------------------------
        # Filtrage fichier 2 (Frequency Sweep) - CODE BLINDÉ
        # ----------------------------------------------------------------
        if not df_fs.empty:
            # 1. Normalisation forcée des colonnes de filtrage pour éviter tout conflit de type ou d'espace
            df_fs_clean = df_fs.copy()
            
            # Nettoyage des chaînes de caractères
            for col in ['IOL_Model', 'Light_Condition', 'Cornea_Type']:
                if col in df_fs_clean.columns:
                    df_fs_clean[col] = df_fs_clean[col].astype(str).str.strip()

            # Normalisation de Pupil_Size en flottant (ex: '3.5' -> 3.5)
            if 'Pupil_Size' in df_fs_clean.columns:
                df_fs_clean['Pupil_Size'] = (
                    df_fs_clean['Pupil_Size']
                    .astype(str)
                    .str.replace(',', '.', regex=False)
                    .astype(float)
                )

            # Convertir aussi les variables sélectionnées en Sidebar pour correspondre parfaitement
            target_model = str(iol_selected).strip()
            target_pupil = float(selected_pupil)
            target_light = str(selected_light).strip()
            target_cornea = str(selected_cornea).strip()

            # 2. Filtrage multicritère
            df_fs_f = df_fs_clean[
                (df_fs_clean['IOL_Model'] == target_model) & 
                (np.isclose(df_fs_clean['Pupil_Size'], target_pupil)) &  # Comparaison float sécurisée
                (df_fs_clean['Light_Condition'] == target_light) & 
                (df_fs_clean['Cornea_Type'] == target_cornea)
            ].sort_values(by='Frequency')

            # --- LIGNE DE DÉBOGAGE TEMPORAIRE ---
            st.caption(f"🔍 Résultat du filtre pour {target_model}: {len(df_fs_f)} lignes trouvées.")

            filtered_fs_dfs.append(df_fs_f)
        else:
            filtered_fs_dfs.append(pd.DataFrame())

# --------------------------------------------------------------------
# 5. VISUALISATION DES RÉSULTATS
# --------------------------------------------------------------------
st.markdown("---")
tab1, tab2, tab3 = st.tabs(["📉 Through-Focus MTF / VA", "🌀 Frequency Sweep (MTF)", "🖼️ USAF Target Images"])

with tab1:
    st.subheader("Courbes Through-Focus MTF / Visual Acuity")
    fig_tf = px.Figure()
    for i, df_curr in enumerate(filtered_dfs):
        if not df_curr.empty and 'Defocus' in df_curr.columns:
            y_col = 'MTF' if 'MTF' in df_curr.columns else df_curr.columns[-1]
            fig_tf.add_trace(px.Scatter(
                x=df_curr['Defocus'], 
                y=df_curr[y_col], 
                mode='lines+markers', 
                name=f"{selected_iols[i]} (Config {chr(65+i)})"
            ))
    fig_tf.update_layout(
        title="Through Focus MTF",
        xaxis_title="Defocus (D)",
        yaxis_title="MTF",
        template="plotly_white"
    )
    st.plotly_chart(fig_tf, use_container_width=True)

with tab2:
    st.subheader("Frequency Sweep MTF")
    if file_2 is None:
        st.info("Veuillez uploader le fichier secondaire (Frequency Sweep) dans la sidebar pour afficher ce graphique.")
    else:
        fig_fs = px.Figure()
        for i, df_fs_curr in enumerate(filtered_fs_dfs):
            if not df_fs_curr.empty and 'Frequency' in df_fs_curr.columns:
                y_col = 'MTF' if 'MTF' in df_fs_curr.columns else df_fs_curr.columns[-1]
                fig_fs.add_trace(px.Scatter(
                    x=df_fs_curr['Frequency'], 
                    y=df_fs_curr[y_col], 
                    mode='lines+markers', 
                    name=f"{selected_iols[i]} (Config {chr(65+i)})"
                ))
        fig_fs.update_layout(
            title="MTF vs Spatial Frequency",
            xaxis_title="Spatial Frequency (lp/mm)",
            yaxis_title="MTF",
            template="plotly_white"
        )
        st.plotly_chart(fig_fs, use_container_width=True)

with tab3:
    st.subheader("Simulation Cibles USAF")
    img_cols = st.columns(num_configs)
    for i in range(num_configs):
        with img_cols[i]:
            st.markdown(f"**Config {chr(65+i)} : {selected_iols[i]}**")
            img_path = f"usaf_images/{selected_iols[i]}_{selected_pupil}mm.png"
            if os.path.exists(img_path):
                image = Image.open(img_path)
                st.image(image, caption=f"USAF Target - {selected_iols[i]}")
            else:
                st.caption("Aperçu image indisponible (fichier non trouvé).")

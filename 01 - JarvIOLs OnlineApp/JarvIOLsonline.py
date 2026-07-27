import streamlit as st
from pathlib import Path
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from PIL import Image
import os

# --------------------------------------------------------------------
# CONFIGURATION DE LA PAGE
# --------------------------------------------------------------------
st.set_page_config(layout="wide", page_title="IOL Comparator - BVI", page_icon="🔬")

# --------------------------------------------------------------------
# 1. AUTHENTIFICATION PAR MOT DE PASSE
# --------------------------------------------------------------------
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.title("🔒 Accès Sécurisé - IOL Comparator BVI")
    st.markdown("Veuillez saisir le mot de passe d'accès pour accéder à l'outil en ligne.")
    
    pwd_input = st.text_input("Mot de passe", type="password")
    if st.button("Se connecter", type="primary"):
        target_pwd = st.secrets.get("PASSWORD", "JarvIOLs-BVI")
        if pwd_input == target_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Mot de passe incorrect.")
    return False

if not check_password():
    st.stop()

# --------------------------------------------------------------------
# 2. PATHS & CACHING (OPTIMISATIONS)
# --------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Excel_PythonZ.xlsx"
IMAGES_BASE_DIR = BASE_DIR / "Images"

@st.cache_data(show_spinner="Chargement des données...")
def load_iol_data(path):
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_image_cached(img_path):
    """Mise en cache de l'ouverture des images pour fluidifier l'affichage."""
    if os.path.exists(img_path):
        return Image.open(img_path)
    return None

df = load_iol_data(EXCEL_PATH)

if df.empty:
    st.error(f"Fichier Excel introuvable ou vide à l'emplacement : {EXCEL_PATH}")
    st.stop()

# Extraction des paramètres
available_models = sorted(df['IOL_Model'].unique())
available_pupils = sorted(df['Pupil_Size'].unique())
available_lights = sorted(df['Light_Condition'].unique())
available_corneas = sorted(df['Cornea_Type'].unique())

mtf_columns = [col for col in df.columns if col.startswith('MTF_')]
available_frequencies = sorted([int(col.split('_')[1]) for col in mtf_columns]) if mtf_columns else [25]

COLOR_PALETTE = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
COLOR_PALETTE_CYAN = ["#2BD9FE", "#e377c2", "#bcbd22", "#17becf", "#f7b6d2", "#c49c94"]

# --------------------------------------------------------------------
# 3. BARRE LATÉRALE
# --------------------------------------------------------------------
st.sidebar.header("🔬 Optical Parameters")
selected_pupil = st.sidebar.selectbox("Pupil aperture (mm)", available_pupils)
selected_light = st.sidebar.selectbox("Light Condition", available_lights)
selected_cornea = st.sidebar.selectbox("Cornea model", available_corneas)

selected_freq = st.sidebar.selectbox("Spatial frequencies for Through-Focus (LP/mm)", available_frequencies)
target_mtf_column = f"MTF_{selected_freq}"

st.sidebar.markdown("---")
st.sidebar.header("⚡ Light Energy Integration Range")
col_min, col_max = st.sidebar.columns(2)
default_defocus_min = float(df['Defocus'].min())
default_defocus_max = float(df['Defocus'].max())

with col_min:
    defocus_min_energy = st.number_input("Defocus Min (D)", value=default_defocus_min, step=0.25, format="%.2f")
with col_max:
    defocus_max_energy = st.number_input("Defocus Max (D)", value=default_defocus_max, step=0.25, format="%.2f")

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Comparison Mode")
comparison_mode = st.sidebar.radio("Choose visualization:", ["1) Multi-Columns (Side-by-Side)", "2) Superimposed (Same graph)"])

st.sidebar.markdown("---")
st.sidebar.header("🎯 Selection of Defocus (-1 to 5 D)")
defocus_inputs = []
cols_defocus = st.sidebar.columns(5)
default_defocus = [-0.5, 0.0, 1.0, 2.0, 3.0]

for i in range(5):
    with cols_defocus[i]:
        val = st.number_input(f"D{i+1}", min_value=-1.0, max_value=5.0, value=float(default_defocus[i]), step=0.1, format="%.2f")
        defocus_inputs.append(round(val, 2))

sorted_defocus_inputs = sorted(defocus_inputs)

st.sidebar.markdown("---")
st.sidebar.header("🔢 Number of Configurations")
num_configs = st.sidebar.number_input("Select number of IOLs to compare (2 - 6):", min_value=2, max_value=6, value=2, step=1)

if st.sidebar.button("🚪 Déconnexion"):
    st.session_state["authenticated"] = False
    st.rerun()

# --------------------------------------------------------------------
# 4. INTERFACE PRINCIPALE & FILTRAGE
# --------------------------------------------------------------------
st.title("📊 IOL Comparative Analysis Tool")

main_cols = st.columns(num_configs)
selected_iols = []
filtered_dfs = []

for i in range(num_configs):
    with main_cols[i]:
        st.header(f"Config {chr(65+i)}")
        default_idx = min(i, len(available_models) - 1)
        iol_selected = st.selectbox(f"Choose IOL ({chr(65+i)})", available_models, index=default_idx, key=f"select_iol_{i}")
        selected_iols.append(iol_selected)
        
        df_f = df[
            (df['IOL_Model'] == iol_selected) & 
            (df['Pupil_Size'] == selected_pupil) & 
            (df['Light_Condition'] == selected_light) & 
            (df['Cornea_Type'] == selected_cornea)
        ].sort_values(by='Defocus')
        filtered_dfs.append(df_f)

# Plages Y globales
all_va_vals = [val for sub in filtered_dfs if not sub.empty for val in sub['VA'].dropna().tolist()]
all_tf_vals = [val for sub in filtered_dfs if not sub.empty and target_mtf_column in sub.columns for val in sub[target_mtf_column].dropna().tolist()]

global_va_range = [min(all_va_vals) - 0.05, max(all_va_vals) + 0.05] if all_va_vals else None
global_tf_range = [min(all_tf_vals) - 0.05, max(all_tf_vals) + 0.05] if all_tf_vals else None

# --------------------------------------------------------------------
# 5. FONCTIONS DE GRAPHIQUES
# --------------------------------------------------------------------
def generate_va_figure(datasets, labels, colors, title, va_range):
    fig = go.Figure()
    for df_sub, label, color in zip(datasets, labels, colors):
        if not df_sub.empty:
            fig.add_trace(go.Scatter(x=df_sub['Defocus'], y=df_sub['VA'], mode='lines+markers', name=label, line=dict(width=2.5, color=color)))
    for d in sorted_defocus_inputs:
        fig.add_vline(x=d, line_width=1.5, line_dash="dash", line_color="#FF4B4B", opacity=0.7)
    fig.update_layout(title=title, xaxis_title="Defocus (D)", yaxis_title="VA (LogMAR)", height=320, margin=dict(l=40, r=40, t=40, b=40))
    if va_range:
        fig.update_yaxes(range=[va_range[1], va_range[0]])
    else:
        fig.update_yaxes(autorange="reversed")
    return fig

def generate_mtf_figure(datasets, labels, colors, title, tf_range):
    fig = go.Figure()
    for df_sub, label, color in zip(datasets, labels, colors):
        if not df_sub.empty and target_mtf_column in df_sub.columns:
            fig.add_trace(go.Scatter(x=df_sub['Defocus'], y=df_sub[target_mtf_column], mode='lines+markers', name=label, line=dict(width=2.5, color=color)))
    for d in sorted_defocus_inputs:
        fig.add_vline(x=d, line_width=1.5, line_dash="dash", line_color="#FF4B4B", opacity=0.7)
    fig.update_layout(title=title, xaxis_title="Defocus (D)", yaxis_title="MTF Value", height=320, margin=dict(l=40, r=40, t=40, b=40))
    if tf_range:
        fig.update_yaxes(range=tf_range)
    return fig

# Rendu des graphes
if "1)" in comparison_mode:
    for i in range(num_configs):
        with main_cols[i]:
            df_sub = filtered_dfs[i]
            iol_name = selected_iols[i]
            conf_label = chr(65+i)
            if df_sub.empty:
                st.warning(f"No data for {iol_name}.")
            else:
                st.plotly_chart(generate_va_figure([df_sub], [f"VA ({iol_name})"], [COLOR_PALETTE[i % len(COLOR_PALETTE)]], f"Visual Acuity - Config {conf_label}", global_va_range), use_container_width=True)
                st.plotly_chart(generate_mtf_figure([df_sub], [f"MTF ({iol_name})"], [COLOR_PALETTE_CYAN[i % len(COLOR_PALETTE_CYAN)]], f"MTF @ {selected_freq} LP/mm - Config {conf_label}", global_tf_range), use_container_width=True)
else:
    st.markdown("### Combined Metrics Comparison")
    col_comb_left, col_comb_right = st.columns(2)
    labels_list = [f"{chr(65+i)}: {selected_iols[i]}" for i in range(num_configs)]
    with col_comb_left:
        st.plotly_chart(generate_va_figure(filtered_dfs, labels_list, COLOR_PALETTE[:num_configs], "Visual Acuity (VA) - Superimposed", global_va_range), use_container_width=True)
    with col_comb_right:
        st.plotly_chart(generate_mtf_figure(filtered_dfs, labels_list, COLOR_PALETTE_CYAN[:num_configs], f"Through Focus (MTF) @ {selected_freq} LP/mm - Superimposed", global_tf_range), use_container_width=True)

# --------------------------------------------------------------------
# 6. TABLEAU ÉNERGIE LUMINEUSE
# --------------------------------------------------------------------
st.markdown("---")
st.header("💡 Light Energy Comparison (Through Focus Area)")

def calculate_light_energy(df_sub, freq, d_min, d_max):
    col_name = f"MTF_{freq}"
    if df_sub.empty or col_name not in df_sub.columns:
        return 0.0
    mask = (df_sub['Defocus'] >= d_min) & (df_sub['Defocus'] <= d_max)
    return df_sub.loc[mask, col_name].sum()

energy_data = []
for i in range(num_configs):
    row = {"IOL": f"Config {chr(65+i)} ({selected_iols[i]})"}
    for freq in available_frequencies:
        row[f"Light Energy at {freq} LP/mm"] = round(calculate_light_energy(filtered_dfs[i], freq, defocus_min_energy, defocus_max_energy), 4)
    energy_data.append(row)

st.caption(f"Calculated by sum of MTF values for Defocus between **{defocus_min_energy:.2f} D** and **{defocus_max_energy:.2f} D**.")
st.dataframe(pd.DataFrame(energy_data), use_container_width=True, hide_index=True)

# --------------------------------------------------------------------
# 7. MIRES USAF (OPTIMISÉES AVEC CACHE)
# --------------------------------------------------------------------
st.markdown("---")
st.header("🖼️ USAF Targets Comparison")

def render_usaf_row(selected_iol, df_filtered, label):
    st.subheader(f"Configuration {label} : **{selected_iol}**")
    if df_filtered.empty:
        st.caption("No data available to display USAF targets.")
        return
        
    img_cols = st.columns(5)
    for idx, d in enumerate(sorted_defocus_inputs):
        with img_cols[idx]:
            st.markdown(f"**{d:+.2f} D**")
            
            candidates = [
                f"{selected_iol}_{float(selected_pupil):.1f}_{selected_light}_{selected_cornea}_defocus_{d:.2f}.jpg",
                f"{selected_iol}_{selected_pupil}_{selected_light}_{selected_cornea}_defocus_{d:.2f}.jpg",
                f"{selected_iol}_{float(selected_pupil):.1f}_{selected_light}_{selected_cornea}_defocus_{float(d)}.jpg"
            ]
            
            img_loaded = None
            for filename in candidates:
                img_path = os.path.join(IMAGES_BASE_DIR, selected_iol, filename)
                img_loaded = load_image_cached(img_path)
                if img_loaded:
                    break
            
            if img_loaded:
                st.image(img_loaded, use_container_width=True)
            else:
                st.caption("⚠️ Non trouvée")
                st.code(f"{candidates[0]}", language="text")

for i in range(num_configs):
    render_usaf_row(selected_iols[i], filtered_dfs[i], chr(65+i))
    st.markdown(" ")
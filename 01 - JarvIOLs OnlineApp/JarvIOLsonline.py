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
EXCEL_FS_PATH = BASE_DIR / "Excel_FS.xlsx"  # <--- NOUVEAU PATH POUR FREQUENCY SWEEP
IMAGES_BASE_DIR = BASE_DIR / "Images"

@st.cache_data(show_spinner="Chargement des données...")
def load_iol_data(path):
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame(columns=['IOL_Model', 'Pupil_Size', 'Light_Condition', 'Cornea_Type', 'Defocus', 'VA'])

@st.cache_data(show_spinner="Chargement des données Frequency Sweep...")
def load_fs_data(path):
    """Mise en cache du fichier Excel_FS.xlsx."""
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
df_fs = load_fs_data(EXCEL_FS_PATH)  # <--- CHARGEMENT DES DONNÉES FREQUENCY SWEEP

if df.empty:
    st.error(f"Fichier Excel introuvable ou vide à l'emplacement : {EXCEL_PATH}")
    st.stop()

# Normalisation facultative du nom de la colonne modèle si orthographiée 'IOL_Mode' dans Excel_FS
if not df_fs.empty and 'IOL_Mode' in df_fs.columns:
    df_fs.rename(columns={'IOL_Mode': 'IOL_Model'}, inplace=True)

# Extraction des paramètres du premier fichier
available_models = sorted(df['IOL_Model'].unique()) if not df.empty else ["Z", "Ma"]
available_pupils = sorted(df['Pupil_Size'].unique()) if not df.empty else [3.5]
available_lights = sorted(df['Light_Condition'].unique()) if not df.empty else ["GL"]
available_corneas = sorted(df['Cornea_Type'].unique()) if not df.empty else ["ISO MC13"]

mtf_columns = [col for col in df.columns if col.startswith('MTF_')]
available_frequencies = sorted([int(col.split('_')[1]) for col in mtf_columns]) if mtf_columns else [25]

# Colonnes cibles pour le Frequency Sweep
fs_metric_columns = ["FS at distance", "FS at 1.50D", "FS at 2.5D"]
if not df_fs.empty:
    # Récupération automatique si les noms diffèrent légèrement
    found_cols = [col for col in df_fs.columns if col.startswith("FS at")]
    if found_cols:
        fs_metric_columns = found_cols

COLOR_PALETTE = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
COLOR_PALETTE_GREEN = ["#2ca02c", "#17becf", "#9467bd", "#ff7f0e", "#bcbd22", "#8c564b"]
COLOR_PALETTE_CYAN = ["#2BD9FE", "#e377c2", "#bcbd22", "#17becf", "#f7b6d2", "#c49c94"]
COLOR_PALETTE_PURPLE = ["#9467bd", "#d62728", "#1f77b4", "#ff7f0e", "#2ca02c", "#8c564b"]

# --------------------------------------------------------------------
# 3. BARRE LATÉRALE : FILTRES OPTIQUES & OPTIONS
# --------------------------------------------------------------------
st.sidebar.header("🔬 Optical Parameters")
selected_pupil = st.sidebar.selectbox("Pupil aperture (mm)", available_pupils)
selected_light = st.sidebar.selectbox("Light Condition", available_lights)
selected_cornea = st.sidebar.selectbox("Cornea model", available_corneas)

selected_freq = st.sidebar.selectbox("Spatial frequencies for Through-Focus (LP/mm)", available_frequencies)
target_mtf_column = f"MTF_{selected_freq}"

# --- NOUVELLE SECTION SIDEBAR : FREQUENCY SWEEP ---
st.sidebar.markdown("---")
st.sidebar.header("📈 Frequency Sweep Parameters")
selected_fs_metric = st.sidebar.selectbox(
    "Select MTF Frequency Sweep Curve",
    fs_metric_columns
)

st.sidebar.markdown("---")
st.sidebar.header("⚡ Light Energy Integration Range")
col_min, col_max = st.sidebar.columns(2)
default_defocus_min = float(df['Defocus'].min()) if not df.empty and 'Defocus' in df.columns else -1.0
default_defocus_max = float(df['Defocus'].max()) if not df.empty and 'Defocus' in df.columns else 5.0

with col_min:
    defocus_min_energy = st.number_input("Defocus Min (D)", value=default_defocus_min, step=0.25, format="%.2f")
with col_max:
    defocus_max_energy = st.number_input("Defocus Max (D)", value=default_defocus_max, step=0.25, format="%.2f")

# Paramètres DOF (X, Y, Z)
st.sidebar.markdown("---")
st.sidebar.header("🎯 DOF Parameters")
st.sidebar.caption("Saisissez 3 valeurs de defocus (X, Y, Z) à mettre en évidence :")
col_x, col_y, col_z = st.sidebar.columns(3)
with col_x:
    defocus_target_x = st.number_input("Defocus X (D)", value=0.00, step=0.25, format="%.2f", key="def_x")
with col_y:
    defocus_target_y = st.number_input("Defocus Y (D)", value=1.50, step=0.25, format="%.2f", key="def_y")
with col_z:
    defocus_target_z = st.number_input("Defocus Z (D)", value=2.50, step=0.25, format="%.2f", key="def_z")

target_defocus_list = [defocus_target_x, defocus_target_y, defocus_target_z]

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Comparison Mode")
comparison_mode = st.sidebar.radio(
    "Choose visualization:",
    [
        "1) Multi-Columns (Side-by-Side)", 
        "2) Superimposed (Same graph)"
    ]
)

st.sidebar.markdown("---")
st.sidebar.header("🖼️ Selection of Defocus for USAF Targets (-1 to 5 D)")
defocus_inputs = []
cols_defocus = st.sidebar.columns(5)
default_defocus = [-0.5, 0.0, 1.0, 2.0, 3.0]

for i in range(5):
    with cols_defocus[i]:
        val = st.number_input(
            f"D{i+1}", 
            min_value=-1.0, 
            max_value=5.0, 
            value=float(default_defocus[i]), 
            step=0.1, 
            format="%.2f"
        )
        defocus_inputs.append(round(val, 2))

sorted_defocus_inputs = sorted(defocus_inputs)

st.sidebar.markdown("---")
st.sidebar.header("🔢 Number of Configurations")
num_configs = st.sidebar.number_input(
    "Select number of IOLs to compare (1 - 6):",
    min_value=1,
    max_value=6,
    value=2,
    step=1
)

if st.sidebar.button("🚪 Déconnexion"):
    st.session_state["authenticated"] = False
    st.rerun()

# --------------------------------------------------------------------
# 4. INTERFACE PRINCIPALE & FILTRAGE DES DONNÉES
# --------------------------------------------------------------------
st.title("📊 IOL Comparative Analysis Tool")

main_cols = st.columns(num_configs)
selected_iols = []
filtered_dfs = []
filtered_fs_dfs = []  # <--- LISTE DES DATAFRAMES FILTRÉS POUR EXCEL_FS

for i in range(num_configs):
    with main_cols[i]:
        st.header(f"Config {chr(65+i)}")
        default_idx = min(i, len(available_models) - 1)
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
        ].sort_values(by='Defocus')
        filtered_dfs.append(df_f)

        # Filtrage fichier 2 (Frequency Sweep)
        if not df_fs.empty:
            df_fs_f = df_fs[
                (df_fs['IOL_Model'] == iol_selected) & 
                (df_fs['Pupil_Size'] == selected_pupil) & 
                (df_fs['Light_Condition'] == selected_light) & 
                (df_fs['Cornea_Type'] == selected_cornea)
            ].sort_values(by='Frequency')
            filtered_fs_dfs.append(df_fs_f)
        else:
            filtered_fs_dfs.append(pd.DataFrame())

# Calcul des plages Y globales
all_va_vals = []
all_mtfa_vals = []
all_tf_vals = []
all_fs_vals = []

for df_sub in filtered_dfs:
    if not df_sub.empty:
        if 'VA' in df_sub.columns:
            va_vals = df_sub['VA'].dropna()
            all_va_vals.extend(va_vals.tolist())
            mtfa_vals = 0.085 / (va_vals + 0.21)
            all_mtfa_vals.extend(mtfa_vals.tolist())
            
        if target_mtf_column in df_sub.columns:
            all_tf_vals.extend(df_sub[target_mtf_column].dropna().tolist())

for df_fs_sub in filtered_fs_dfs:
    if not df_fs_sub.empty and selected_fs_metric in df_fs_sub.columns:
        all_fs_vals.extend(df_fs_sub[selected_fs_metric].dropna().tolist())

global_va_range = [min(all_va_vals) - 0.05, max(all_va_vals) + 0.05] if all_va_vals else None
global_mtfa_range = [min(all_mtfa_vals) - 0.05, max(all_mtfa_vals) + 0.05] if all_mtfa_vals else None
global_tf_range = [min(all_tf_vals) - 0.05, max(all_tf_vals) + 0.05] if all_tf_vals else None
global_fs_range = [0, max(all_fs_vals) + 0.05] if all_fs_vals else [0, 1.0]

# --------------------------------------------------------------------
# 5. FONCTIONS DOF ET CRÉATION DES GRAPHIQUES
# --------------------------------------------------------------------
def calculate_dof_va(df_sub, target_threshold):
    if df_sub.empty or 'VA' not in df_sub.columns or 'Defocus' not in df_sub.columns:
        return np.nan
    best_idx = df_sub['VA'].idxmin()
    best_defocus = df_sub.loc[best_idx, 'Defocus']
    sub_pos = df_sub[df_sub['Defocus'] >= best_defocus].sort_values(by='Defocus').reset_index(drop=True)
    if sub_pos.empty:
        return np.nan
    last_valid_defocus = best_defocus
    for idx, row in sub_pos.iterrows():
        if row['VA'] <= target_threshold:
            last_valid_defocus = row['Defocus']
        else:
            break
    return round(abs(last_valid_defocus - best_defocus), 2)

def calculate_dof_mtf(df_sub, target_mtf_col, target_threshold=0.1):
    if df_sub.empty or target_mtf_col not in df_sub.columns or 'Defocus' not in df_sub.columns:
        return np.nan
    best_idx = df_sub[target_mtf_col].idxmax()
    best_defocus = df_sub.loc[best_idx, 'Defocus']
    sub_pos = df_sub[df_sub['Defocus'] >= best_defocus].sort_values(by='Defocus').reset_index(drop=True)
    if sub_pos.empty:
        return np.nan
    last_valid_defocus = best_defocus
    for idx, row in sub_pos.iterrows():
        if row[target_mtf_col] >= target_threshold:
            last_valid_defocus = row['Defocus']
        else:
            break
    return round(abs(last_valid_defocus - best_defocus), 2)

def generate_va_figure(datasets, labels, colors, title, va_range, target_defocus):
    fig = go.Figure()
    for df_sub, label, color in zip(datasets, labels, colors):
        if not df_sub.empty and 'VA' in df_sub.columns:
            fig.add_trace(go.Scatter(
                x=df_sub['Defocus'], y=df_sub['VA'], 
                mode='lines+markers', name=label,
                line=dict(width=2.5, color=color)
            ))
    for idx, d in enumerate(target_defocus):
        fig.add_vline(
            x=d, line_width=1.5, line_color="red", opacity=0.85,
            annotation_text=f"D={d:.2f}D",
            annotation_position="top left" if idx == 0 else "top right"
        )
    fig.add_hline(y=0.1, line_width=1.5, line_dash="dot", line_color="black", opacity=0.8, annotation_text="0.1 LogMAR", annotation_position="bottom right")
    fig.add_hline(y=0.2, line_width=1.5, line_dash="dot", line_color="black", opacity=0.8, annotation_text="0.2 LogMAR", annotation_position="bottom right")
    fig.update_layout(title=title, xaxis_title="Defocus (D)", yaxis_title="VA (LogMAR)", height=350, margin=dict(l=40, r=40, t=40, b=40), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    if va_range:
        fig.update_yaxes(range=[va_range[1], va_range[0]])
    else:
        fig.update_yaxes(autorange="reversed")
    return fig

def generate_mtfa_figure(datasets, labels, colors, title, mtfa_range, target_defocus):
    fig = go.Figure()
    for df_sub, label, color in zip(datasets, labels, colors):
        if not df_sub.empty and 'VA' in df_sub.columns:
            mtfa_vals = 0.085 / (df_sub['VA'] + 0.21)
            fig.add_trace(go.Scatter(
                x=df_sub['Defocus'], y=mtfa_vals, 
                mode='lines+markers', name=label,
                line=dict(width=2.5, color=color)
            ))
    for idx, d in enumerate(target_defocus):
        fig.add_vline(
            x=d, line_width=1.5, line_color="red", opacity=0.85,
            annotation_text=f"D={d:.2f}D",
            annotation_position="top left" if idx == 0 else "top right"
        )
    fig.update_layout(title=title, xaxis_title="Defocus (D)", yaxis_title="MTFa", height=350, margin=dict(l=40, r=40, t=40, b=40), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    if mtfa_range:
        fig.update_yaxes(range=mtfa_range)
    return fig

def generate_mtf_figure(datasets, labels, colors, title, tf_range, target_defocus):
    fig = go.Figure()
    for df_sub, label, color in zip(datasets, labels, colors):
        if not df_sub.empty:
            y_data = df_sub[target_mtf_column] if target_mtf_column in df_sub.columns else [0] * len(df_sub)
            fig.add_trace(go.Scatter(
                x=df_sub['Defocus'], y=y_data, 
                mode='lines+markers', name=label,
                line=dict(width=2.5, color=color)
            ))
    for idx, d in enumerate(target_defocus):
        fig.add_vline(
            x=d, line_width=1.5, line_color="red", opacity=0.85,
            annotation_text=f"D={d:.2f}D",
            annotation_position="top left" if idx == 0 else "top right"
        )
    fig.add_hline(y=0.1, line_width=1.5, line_dash="dot", line_color="black", opacity=0.8, annotation_text="0.1 MTF", annotation_position="top right")
    fig.update_layout(title=title, xaxis_title="Defocus (D)", yaxis_title="MTF Value", height=350, margin=dict(l=40, r=40, t=40, b=40), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    if tf_range:
        fig.update_yaxes(range=tf_range)
    return fig

# --- NOUVELLE FONCTION : GRAPHIQUE MTF FREQUENCY SWEEP ---
def generate_fs_figure(datasets, labels, colors, title, y_column, fs_range):
    """Génère le graphique MTF Frequency Sweep."""
    fig = go.Figure()
    for df_fs_sub, label, color in zip(datasets, labels, colors):
        if not df_fs_sub.empty and 'Frequency' in df_fs_sub.columns and y_column in df_fs_sub.columns:
            fig.add_trace(go.Scatter(
                x=df_fs_sub['Frequency'], 
                y=df_fs_sub[y_column], 
                mode='lines+markers', 
                name=label,
                line=dict(width=2.5, color=color)
            ))
    fig.add_hline(y=0.1, line_width=1.5, line_dash="dot", line_color="black", opacity=0.8, annotation_text="0.1 MTF", annotation_position="top right")
    fig.update_layout(
        title=title, 
        xaxis_title="Spatial Frequency (LP/mm)", 
        yaxis_title=f"MTF ({y_column})", 
        height=350, 
        margin=dict(l=40, r=40, t=40, b=40), 
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    if fs_range:
        fig.update_yaxes(range=fs_range)
    return fig

# Construction des tableaux de synthèse DOF
dof_va_rows = []
for i in range(num_configs):
    df_sub = filtered_dfs[i]
    iol_name = selected_iols[i]
    conf_label = f"Config {chr(65+i)} ({iol_name})"
    dof_va_rows.append({
        "Configuration IOL": conf_label,
        "DOF @ 0.1 LogMAR (D)": calculate_dof_va(df_sub, 0.1),
        "DOF @ 0.2 LogMAR (D)": calculate_dof_va(df_sub, 0.2)
    })
df_dof_va = pd.DataFrame(dof_va_rows)

dof_mtf_rows = []
for i in range(num_configs):
    df_sub = filtered_dfs[i]
    iol_name = selected_iols[i]
    conf_label = f"Config {chr(65+i)} ({iol_name})"
    dof_mtf_rows.append({
        "Configuration IOL": conf_label,
        "DOF MTF @ 0.1 (D)": calculate_dof_mtf(df_sub, target_mtf_column, 0.1)
    })
df_dof_mtf = pd.DataFrame(dof_mtf_rows)

# Affichage des graphiques et synthèses
if "1)" in comparison_mode:
    for i in range(num_configs):
        with main_cols[i]:
            df_sub = filtered_dfs[i]
            df_fs_sub = filtered_fs_dfs[i]
            iol_name = selected_iols[i]
            conf_label = chr(65+i)
            if df_sub.empty:
                st.warning(f"No data for {iol_name}.")
            else:
                # 1. Graph VA
                fig_va = generate_va_figure([df_sub], [f"VA ({iol_name})"], [COLOR_PALETTE[i % len(COLOR_PALETTE)]], f"Visual Acuity - Config {conf_label}", global_va_range, target_defocus_list)
                st.plotly_chart(fig_va, use_container_width=True)
                
                # 2. Graph TF MTFa
                fig_mtfa = generate_mtfa_figure([df_sub], [f"MTFa ({iol_name})"], [COLOR_PALETTE_GREEN[i % len(COLOR_PALETTE_GREEN)]], f"TF MTFa - Config {conf_label}", global_mtfa_range, target_defocus_list)
                st.plotly_chart(fig_mtfa, use_container_width=True)

                # 3. Graph MTF Through-Focus
                fig_tf = generate_mtf_figure([df_sub], [f"MTF ({iol_name})"], [COLOR_PALETTE_CYAN[i % len(COLOR_PALETTE_CYAN)]], f"MTF @ {selected_freq} LP/mm - Config {conf_label}", global_tf_range, target_defocus_list)
                st.plotly_chart(fig_tf, use_container_width=True)

                # 4. Graph MTF Frequency Sweep (PLACE EN DESSOUS DU THROUGH-FOCUS)
                fig_fs = generate_fs_figure([df_fs_sub], [f"FS ({iol_name})"], [COLOR_PALETTE_PURPLE[i % len(COLOR_PALETTE_PURPLE)]], f"MTF Frequency Sweep ({selected_fs_metric}) - Config {conf_label}", selected_fs_metric, global_fs_range)
                st.plotly_chart(fig_fs, use_container_width=True)

    st.markdown("### 📊 Depth of Field (DOF) Summary Tables")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.subheader("Visual Acuity (VA) DOF")
        st.dataframe(df_dof_va, use_container_width=True, hide_index=True)
    with col_t2:
        st.subheader(f"MTF Through-Focus DOF (@ {selected_freq} LP/mm)")
        st.dataframe(df_dof_mtf, use_container_width=True, hide_index=True)
else:
    st.markdown("### Combined Metrics Comparison")
    labels_list = [f"{chr(65+i)}: {selected_iols[i]}" for i in range(num_configs)]
    
    # Mode Superposé
    fig_va_combined = generate_va_figure(filtered_dfs, labels_list, COLOR_PALETTE[:num_configs], "Visual Acuity (VA) - Superimposed", global_va_range, target_defocus_list)
    st.plotly_chart(fig_va_combined, use_container_width=True)
    
    fig_mtfa_combined = generate_mtfa_figure(filtered_dfs, labels_list, COLOR_PALETTE_GREEN[:num_configs], "TF MTFa - Superimposed", global_mtfa_range, target_defocus_list)
    st.plotly_chart(fig_mtfa_combined, use_container_width=True)

    fig_tf_combined = generate_mtf_figure(filtered_dfs, labels_list, COLOR_PALETTE_CYAN[:num_configs], f"Through Focus (MTF) @ {selected_freq} LP/mm - Superimposed", global_tf_range, target_defocus_list)
    st.plotly_chart(fig_tf_combined, use_container_width=True)

    # Graphique Frequency Sweep Superposé
    fig_fs_combined = generate_fs_figure(filtered_fs_dfs, labels_list, COLOR_PALETTE_PURPLE[:num_configs], f"MTF Frequency Sweep ({selected_fs_metric}) - Superimposed", selected_fs_metric, global_fs_range)
    st.plotly_chart(fig_fs_combined, use_container_width=True)

    st.markdown("---")
    col_comb_left, col_comb_right = st.columns(2)
    with col_comb_left:
        st.subheader("📋 VA Depth of Field (DOF)")
        st.dataframe(df_dof_va, use_container_width=True, hide_index=True)
    with col_comb_right:
        st.subheader(f"📋 MTF Depth of Field (DOF @ {selected_freq} LP/mm)")
        st.dataframe(df_dof_mtf, use_container_width=True, hide_index=True)

# --------------------------------------------------------------------
# 6. TABLEAU D'ÉNERGIE LUMINEUSE
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
# 7. MIRES USAF
# --------------------------------------------------------------------
st.markdown("---")
st.header("🖼️ USAF Targets Comparison")

def render_usaf_row(selected_iol, df_filtered, label):
    st.subheader(f"Configuration {label} : **{selected_iol}**")
    if df_filtered.empty:
        st.caption("No data available to display USAF targets.")
        return
        
    img_cols = st.columns(5)
    iol_folder = IMAGES_BASE_DIR / selected_iol

    for idx, d in enumerate(sorted_defocus_inputs):
        with img_cols[idx]:
            st.markdown(f"**{d:+.2f} D**")
            
            try:
                pupil_val = float(selected_pupil)
                pupil_str = f"{pupil_val:.1f}"
            except (ValueError, TypeError):
                pupil_str = str(selected_pupil)

            light_lower = str(selected_light).lower()
            light_upper = str(selected_light).upper()

            candidates = [
                f"{selected_iol}_{pupil_str}_{light_lower}_{selected_cornea}_defocus_{d:.2f}.jpg",
                f"{selected_iol}_{pupil_str}_{light_upper}_{selected_cornea}_defocus_{d:.2f}.jpg",
                f"{selected_iol}_{pupil_str}_{light_lower}_{selected_cornea}_defocus_{d:+.2f}.jpg",
                f"{selected_iol}_{pupil_str}_{light_upper}_{selected_cornea}_defocus_{d:+.2f}.jpg",
                f"{selected_iol}_{selected_pupil}_{light_lower}_{selected_cornea}_defocus_{d:.2f}.jpg",
                f"{selected_iol}_{selected_pupil}_{light_upper}_{selected_cornea}_defocus_{d:.2f}.jpg",
            ]
            
            img_loaded = None
            for filename in candidates:
                img_path = iol_folder / filename
                img_loaded = load_image_cached(str(img_path))
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

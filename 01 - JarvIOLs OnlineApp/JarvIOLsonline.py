import os
import pandas as pd
import streamlit as st
import plotly.graph_objects as go


# ==============================================================================
# APPLICATION STREAMLIT (RUNS ON STREAMLIT CLOUD)
# ==============================================================================

def run_streamlit_app():
    st.set_page_config(page_title="IOL Performance Viewer", layout="wide")
    st.title("Analyse des performances IOL")

    # Recherche des fichiers Excel de données dans le même répertoire
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_pythonz = os.path.join(current_dir, "Excel_PythonZ.xlsx")
    file_fs = os.path.join(current_dir, "Excel_FS.xlsx")

    # Vérification des fichiers
    if not os.path.exists(file_pythonz):
        st.error(f"Fichier principal 'Excel_PythonZ.xlsx' introuvable dans : {current_dir}")
        return

    df_pythonz = pd.read_excel(file_pythonz)

    # Sidebar - Filtres existants
    st.sidebar.header("Paramètres de sélection")
    
    iol_options = df_pythonz["IOL_Mode"].unique() if "IOL_Mode" in df_pythonz.columns else []
    selected_iol = st.sidebar.selectbox("IOL Mode :", iol_options)

    pupil_options = df_pythonz["Pupil_Size"].unique() if "Pupil_Size" in df_pythonz.columns else []
    selected_pupil = st.sidebar.selectbox("Pupil Size :", pupil_options)

    light_options = df_pythonz["Light_Condition"].unique() if "Light_Condition" in df_pythonz.columns else []
    selected_light = st.sidebar.selectbox("Light Condition :", light_options)

    cornea_options = df_pythonz["Cornea_Type"].unique() if "Cornea_Type" in df_pythonz.columns else []
    selected_cornea = st.sidebar.selectbox("Cornea Type :", cornea_options)

    # Réglage min/max Y global
    st.sidebar.markdown("---")
    st.sidebar.header("Échelle de l'Axe Y")
    y_min_val = st.sidebar.number_input("Y min :", value=0.0, step=0.1)
    y_max_val = st.sidebar.number_input("Y max :", value=1.0, step=0.1)

    # --------------------------------------------------------------------------
    # PARAMÈTRES FREQUENCY SWEEP
    # --------------------------------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.header("Paramètres Frequency Sweep")

    fs_defocus_option = st.sidebar.selectbox(
        "Condition Frequency Sweep :",
        options=["at distance", "1.50D", "at 2.5D"],
        key="fs_defocus_select"
    )

    freq_max = st.sidebar.number_input(
        "Frequency max :",
        min_value=0.0,
        max_value=600.0,
        value=100.0,
        step=10.0,
        key="freq_max_input"
    )

    # --------------------------------------------------------------------------
    # AFFICHAGE DES GRAPHIQUES EXISTANTS (THROUGH FOCUS MTF)
    # --------------------------------------------------------------------------
    st.subheader("Through Focus MTF")
    
    fig_tf = go.Figure()
    if "Defocus" in df_pythonz.columns and "MTF" in df_pythonz.columns:
        df_filtered_tf = df_pythonz[
            (df_pythonz["IOL_Mode"] == selected_iol) &
            (df_pythonz["Pupil_Size"] == selected_pupil) &
            (df_pythonz["Light_Condition"] == selected_light) &
            (df_pythonz["Cornea_Type"] == selected_cornea)
        ]
        fig_tf.add_trace(go.Scatter(x=df_filtered_tf["Defocus"], y=df_filtered_tf["MTF"], mode="lines", name="MTF"))
    fig_tf.update_layout(yaxis=dict(range=[y_min_val, y_max_val]), template="plotly_white", height=400)
    st.plotly_chart(fig_tf, use_container_width=True)

    # --------------------------------------------------------------------------
    # GRAPHIQUE FREQUENCY SWEEP
    # --------------------------------------------------------------------------
    st.markdown("---")
    st.subheader("Frequency Sweep")

    if os.path.exists(file_fs):
        df_fs = pd.read_excel(file_fs)

        # Nettoyage et formatage
        df_fs["Pupil_Size_Clean"] = df_fs["Pupil_Size"].astype(str).str.replace("mm", "").str.replace(",", ".").str.strip()
        selected_pupil_clean = str(selected_pupil).replace("mm", "").replace(",", ".").strip()

        # Filtrage suivant les paramètres identiques à Excel_PythonZ
        df_fs_filtered = df_fs[
            (df_fs["IOL_Mode"].astype(str).str.strip() == str(selected_iol).strip()) &
            (df_fs["Pupil_Size_Clean"] == selected_pupil_clean) &
            (df_fs["Light_Condition"].astype(str).str.strip() == str(selected_light).strip()) &
            (df_fs["Cornea_Type"].astype(str).str.strip() == str(selected_cornea).strip())
        ]

        # Limite Frequency max (Axe X)
        df_fs_filtered = df_fs_filtered[df_fs_filtered["Frequency"] <= freq_max]

        # Mapping Axe Y
        column_mapping = {
            "at distance": "FS at distance",
            "1.50D": "FS at 1.50D",
            "at 2.5D": "FS at 2.5D"
        }
        target_y_col = column_mapping[fs_defocus_option]

        if not df_fs_filtered.empty and target_y_col in df_fs_filtered.columns:
            fig_fs = go.Figure()

            fig_fs.add_trace(
                go.Scatter(
                    x=df_fs_filtered["Frequency"],
                    y=df_fs_filtered[target_y_col],
                    mode="lines+markers",
                    name=f"FS ({fs_defocus_option})",
                    line=dict(color="#1f77b4", width=2),
                    marker=dict(size=3)
                )
            )

            fig_fs.update_layout(
                title=f"Frequency Sweep - Condition: {fs_defocus_option}",
                xaxis_title="Frequency (lp/mm)",
                yaxis_title="FS Value",
                xaxis=dict(range=[0, freq_max]),
                yaxis=dict(range=[y_min_val, y_max_val]),
                template="plotly_white",
                height=450,
                margin=dict(l=40, r=40, t=50, b=40)
            )

            st.plotly_chart(fig_fs, use_container_width=True)
        else:
            st.info("Aucune donnée correspondante trouvée dans Excel_FS.xlsx pour ces paramètres.")
    else:
        st.warning("Fichier 'Excel_FS.xlsx' introuvable dans le même répertoire.")


if __name__ == "__main__":
    run_streamlit_app()

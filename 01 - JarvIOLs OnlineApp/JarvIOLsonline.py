import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import win32com.client as win32
import pandas as pd
import streamlit as st
import plotly.graph_objects as go


# ==============================================================================
# SECTION 1 : EXTRACTION ET COMPILATION DES DONNÉES (CODE INCHANGÉ + FS ENREGISTRÉ)
# ==============================================================================

def extract_file_info(filename):
    """Extrait la condition de lumière, l'ouverture et le type de cornée."""
    light_match = re.search(r"\b(GL|WL)\b", filename, re.IGNORECASE)
    light_cond = light_match.group(1).upper() if light_match else "GL"

    aperture_match = re.search(r"(\d+(?:[.,]\d+)?\s*mm)", filename, re.IGNORECASE)
    aperture_val = aperture_match.group(1).replace(" ", "").replace(",", ".") if aperture_match else "3.5mm"

    cornea_match = re.search(r"\b(ISO\s*[^_\(\s]+)\b", filename, re.IGNORECASE)
    cornea = cornea_match.group(1).strip() if cornea_match else "Inconnu"

    return light_cond, aperture_val, cornea


def match_sheet_defocus(sheet_name, aperture_val):
    """
    Associe strictement :
    0.0  -> FS at distance
    -2.2 -> FS at 1.50D
    -3.7 -> FS at 2.5D
    """
    sheet_lower = sheet_name.lower().replace(",", ".")
    aperture_clean = aperture_val.lower().replace(",", ".")

    # S'assurer que l'onglet correspond à la pupille étudiée dans ce fichier (ex: 3.5mm)
    if "mm" in sheet_lower and aperture_clean not in sheet_lower:
        return None

    # Recherche explicite des motifs exacts dans le nom d'onglet
    if "0.0" in sheet_lower or "0t" in sheet_lower:
        return "-0.0"
    elif "-2.2" in sheet_lower:
        return "-2.2"
    elif "-3.7" in sheet_lower:
        return "-3.7"

    return None


def clean_num(val):
    """Convertit n'importe quelle cellule en float pur."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(",", ".")
    try:
        return float(val_str)
    except ValueError:
        return None


def process_excel_files():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Initialisation",
        "Veuillez choisir le dossier contenant vos fichiers Excel (.xlsx) à traiter.",
    )
    folder_path = filedialog.askdirectory(title="Sélectionner le dossier source")

    if not folder_path:
        return

    all_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    excel_files = [f for f in all_files if not f.startswith("Compilation_") and not f.startswith("Excel_FS")]

    if not excel_files:
        messagebox.showwarning("Aucun fichier", "Aucun fichier Excel (.xlsx) trouvé.")
        return

    iol_name = simpledialog.askstring("Nom de l'IOL", "Veuillez entrer le nom de l'IOL :")
    if not iol_name:
        return

    all_compiled_data = []

    print("--- DÉBUT DU TRAITEMENT ---")

    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    try:
        for file_name in excel_files:
            file_path = os.path.abspath(os.path.join(folder_path, file_name))
            try:
                light_cond, aperture_val, cornea = extract_file_info(file_name)
                
                workbook = excel_app.Workbooks.Open(file_path)
                extracted_by_defocus = {}

                for sheet in workbook.Sheets:
                    defocus_key = match_sheet_defocus(sheet.Name, aperture_val)
                    
                    if defocus_key and defocus_key not in extracted_by_defocus:
                        raw_data = sheet.Range("A2:C602").Value
                        
                        freq_vals = []
                        mean_vals = []
                        
                        if raw_data:
                            for row in raw_data:
                                val_a = clean_num(row[0])
                                val_b = clean_num(row[1])
                                val_c = clean_num(row[2])
                                
                                freq_vals.append(val_a)
                                
                                nums = [v for v in (val_b, val_c) if v is not None]
                                row_mean = sum(nums) / len(nums) if nums else None
                                mean_vals.append(row_mean)

                        extracted_by_defocus[defocus_key] = {
                            "freq": freq_vals,
                            "means": mean_vals
                        }

                workbook.Close(False)

                if not extracted_by_defocus:
                    print(f"⚠️ Aucun onglet defocus (-0.0, -2.2, -3.7) dans : {file_name}")
                    continue

                # Récupérer la colonne Frequency
                first_found_key = list(extracted_by_defocus.keys())[0]
                frequency_column = extracted_by_defocus[first_found_key]["freq"]
                num_rows = len(frequency_column)

                file_data = {
                    "IOL_Mode": [iol_name] * num_rows,
                    "Pupil_Size": [aperture_val] * num_rows,
                    "Light_Condition": [light_cond] * num_rows,
                    "Cornea_Type": [cornea] * num_rows,
                    "Frequency": frequency_column,
                    "FS at distance": extracted_by_defocus.get("-0.0", {}).get("means", [None] * num_rows),
                    "FS at 1.50D": extracted_by_defocus.get("-2.2", {}).get("means", [None] * num_rows),
                    "FS at 2.5D": extracted_by_defocus.get("-3.7", {}).get("means", [None] * num_rows),
                }

                df_file = pd.DataFrame(file_data)
                all_compiled_data.append(df_file)
                
                keys_found = list(extracted_by_defocus.keys())
                print(f"✅ {file_name} : {len(keys_found)}/3 defocus extraits {keys_found}")

            except Exception as e:
                print(f"❌ Erreur sur {file_name} : {e}")

    finally:
        excel_app.Quit()

    if all_compiled_data:
        df_final = pd.concat(all_compiled_data, ignore_index=True)
        
        # Enregistrement du fichier Excel_FS.xlsx au même endroit que le dossier sélectionné
        fs_save_path = os.path.join(folder_path, "Excel_FS.xlsx")
        df_final.to_excel(fs_save_path, index=False)
        
        save_path = filedialog.asksaveasfilename(
            title="Enregistrer la compilation",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            initialfile=f"CompilationFS_{iol_name}.xlsx",
            initialdir=folder_path,
        )
        if save_path:
            df_final.to_excel(save_path, index=False)
            messagebox.showinfo("Succès", "Fichier de compilation et Excel_FS générés et REMPLIS !")
            print(f"\n🎉 Compilation sauvegardée dans : {save_path}")
            print(f"🎉 Excel_FS sauvegardé dans : {fs_save_path}")
    else:
        messagebox.showwarning("Aucune donnée", "Aucune donnée n'a été extraite.")


# ==============================================================================
# SECTION 2 : APPLICATION STREAMLIT (AJOUT DU FREQUENCY SWEEP & SANS RESTRICTION)
# ==============================================================================

def run_streamlit_app():
    st.set_page_config(page_title="IOL Performance Viewer", layout="wide")
    st.title("Analyse des performances IOL")

    # Recherche des fichiers Excel de données
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

    # Réglage min/max Y global pour harmoniser tous les graphiques
    st.sidebar.markdown("---")
    st.sidebar.header("Échelle de l'Axe Y")
    y_min_val = st.sidebar.number_input("Y min :", value=0.0, step=0.1)
    y_max_val = st.sidebar.number_input("Y max :", value=1.0, step=0.1)

    # --------------------------------------------------------------------------
    # AJOUT POINT 2 : Menus déroulants spécifiques Frequency Sweep dans la sidebar
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
    # SECTION 3 : AFFICHAGE DES GRAPHIQUES EXISTANTS (VA, MTFa, MTF)
    # --------------------------------------------------------------------------
    
    # [Vos graphiques existants VA, MTFa et Through Focus MTF restent placés ici]
    st.subheader("Through Focus MTF")
    
    # Simulation/Exemple du graphique Through Focus MTF existant :
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
    # AJOUT POINTS 1, 3, 4, 5, 6 : NOUVEAU GRAPHIQUE FREQUENCY SWEEP EN DESSOUS
    # --------------------------------------------------------------------------
    st.markdown("---")
    st.subheader("Frequency Sweep")

    if os.path.exists(file_fs):
        df_fs = pd.read_excel(file_fs)

        # Harmonisation du nettoyage pour la correspondance stricte
        df_fs["Pupil_Size_Clean"] = df_fs["Pupil_Size"].astype(str).str.replace("mm", "").str.replace(",", ".").str.strip()
        selected_pupil_clean = str(selected_pupil).replace("mm", "").replace(",", ".").strip()

        # Filtrage suivant les paramètres identiques à Excel_PythonZ
        df_fs_filtered = df_fs[
            (df_fs["IOL_Mode"].astype(str).str.strip() == str(selected_iol).strip()) &
            (df_fs["Pupil_Size_Clean"] == selected_pupil_clean) &
            (df_fs["Light_Condition"].astype(str).str.strip() == str(selected_light).strip()) &
            (df_fs["Cornea_Type"].astype(str).str.strip() == str(selected_cornea).strip())
        ]

        # Application de la limite Frequency max sur l'axe X (0 à 600)
        df_fs_filtered = df_fs_filtered[df_fs_filtered["Frequency"] <= freq_max]

        # Mapping de la sélection du menu déroulant vers la colonne Y correspondante
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

            # Application des limites y_min / y_max identiques aux autres graphiques
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
    # Détection si exécuté en mode Streamlit ou script standard
    import sys
    if "streamlit" in sys.modules or st._is_running_with_streamlit:
        run_streamlit_app()
    else:
        process_excel_files()

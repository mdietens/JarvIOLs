import os
import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


def extract_file_info(filename):
    """Extrait la condition de lumière, l'ouverture et le type de cornée."""
    light_match = re.search(r"\b(GL|WL)\b", filename, re.IGNORECASE)
    light_cond = light_match.group(1).upper() if light_match else "GL"

    aperture_match = re.search(r"(\d+(?:[.,]\d+)?\s*mm)", filename, re.IGNORECASE)
    aperture_val = (
        aperture_match.group(1).replace(" ", "").replace(",", ".")
        if aperture_match
        else "3.5mm"
    )

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
    # Imports locaux pour ne pas bloquer Streamlit Cloud / Linux
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog
    import win32com.client as win32

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
    excel_files = [f for f in all_files if not f.startswith("Compilation_")]

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
                            "means": mean_vals,
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
                    "Nom de l'IOL": [iol_name] * num_rows,
                    "Ouverture pupillaire": [aperture_val] * num_rows,
                    "Condition de lumière": [light_cond] * num_rows,
                    "Type de cornée": [cornea] * num_rows,
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
        save_path = filedialog.asksaveasfilename(
            title="Enregistrer la compilation",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            initialfile=f"CompilationFS_{iol_name}.xlsx",
            initialdir=folder_path,
        )
        if save_path:
            df_final.to_excel(save_path, index=False)
            messagebox.showinfo("Succès", "Fichier de compilation généré et REMPLI !")
            print(f"\n🎉 Compilation sauvegardée dans : {save_path}")
    else:
        messagebox.showwarning("Aucune donnée", "Aucune donnée n'a été extraite.")


# ==============================================================================
# EXECUTION STREAMLIT
# ==============================================================================
st.set_page_config(page_title="IOL Analysis Dashboard", layout="wide")
st.title("IOL Optical Performance Analysis")

# 3) Chargement des fichiers au même emplacement
base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
path_pythonz = os.path.join(base_dir, "Excel_PythonZ.xlsx")
path_fs = os.path.join(base_dir, "Excel_FS.xlsx")

df_main = pd.read_excel(path_pythonz) if os.path.exists(path_pythonz) else pd.DataFrame()
df_fs = pd.read_excel(path_fs) if os.path.exists(path_fs) else pd.DataFrame()

# Barre latérale (Onglet à gauche)
st.sidebar.header("Filtres & Paramètres")

# Filtres IOL / Pupille / Lumière / Cornée
if not df_main.empty:
    iol_models = df_main["IOL_Model"].unique() if "IOL_Model" in df_main.columns else []
    selected_iols = st.sidebar.multiselect("Choix des IOL Models", iol_models, default=iol_models)

    pupils = df_main["Pupil_Size"].unique() if "Pupil_Size" in df_main.columns else []
    selected_pupil = st.sidebar.selectbox("Ouverture (Pupil Size)", pupils) if len(pupils) > 0 else None

    lights = df_main["Light_Condition"].unique() if "Light_Condition" in df_main.columns else []
    selected_light = st.sidebar.selectbox("Condition de lumière", lights) if len(lights) > 0 else None

    corneas = df_main["Cornea_Type"].unique() if "Cornea_Type" in df_main.columns else []
    selected_cornea = st.sidebar.selectbox("Type de cornée", corneas) if len(corneas) > 0 else None
else:
    selected_iols, selected_pupil, selected_light, selected_cornea = [], None, None, None

# 6) Gestion min-max pour l'axe Y
st.sidebar.subheader("Configuration axe Y")
y_min = st.sidebar.number_input("Y min", value=0.0, step=0.1)
y_max = st.sidebar.number_input("Y max", value=1.0, step=0.1)

# 2) Options Frequency Sweep dans l'onglet à gauche
st.sidebar.subheader("Frequency Sweep")
fs_option = st.sidebar.selectbox("Distance", ["at distance", "1.50D", "at 2.5D"])
freq_max = st.sidebar.number_input("Frequency max", min_value=0, max_value=600, value=100, step=10)

# Mappage des noms du menu vers les colonnes Excel
fs_col_map = {
    "at distance": "FS at distance",
    "1.50D": "FS at 1.50D",
    "at 2.5D": "FS at 2.5D",
}
target_fs_col = fs_col_map[fs_option]

# --- AFFICHAGE DES GRAPHES ---

# Graphs existants (VA, MTFa, Through focus MTF)
st.subheader("Visual Acuity (VA)")
st.info("Graphique VA")

st.subheader("MTFa")
st.info("Graphique MTFa")

st.subheader("Through focus MTF")
fig_mtf = go.Figure()
fig_mtf.update_layout(yaxis=dict(range=[y_min, y_max]))
st.plotly_chart(fig_mtf, use_container_width=True)

# 1) Ajout du graphique Frequency Sweep en dessous
st.subheader("Frequency Sweep")

if not df_fs.empty:
    df_fs_filtered = df_fs.copy()

    # 4) Filtrage selon IOL, Ouverture, Lumière, Cornée
    if selected_iols:
        df_fs_filtered = df_fs_filtered[df_fs_filtered["IOL_Model"].isin(selected_iols)]
    if selected_pupil is not None:
        df_fs_filtered = df_fs_filtered[df_fs_filtered["Pupil_Size"] == selected_pupil]
    if selected_light is not None:
        df_fs_filtered = df_fs_filtered[df_fs_filtered["Light_Condition"] == selected_light]
    if selected_cornea is not None:
        df_fs_filtered = df_fs_filtered[df_fs_filtered["Cornea_Type"] == selected_cornea]

    # 4) Imposer la valeur max de Frequency
    if "Frequency" in df_fs_filtered.columns:
        df_fs_filtered = df_fs_filtered[df_fs_filtered["Frequency"] <= freq_max]

    # 5) Conversion numérique si nécessaire
    if target_fs_col in df_fs_filtered.columns:
        df_fs_filtered[target_fs_col] = (
            df_fs_filtered[target_fs_col].astype(str).str.replace(",", ".").astype(float)
        )

    if not df_fs_filtered.empty and target_fs_col in df_fs_filtered.columns:
        # Création du graphe x = Frequency, y = Colonne choisie
        fig_fs = px.line(
            df_fs_filtered,
            x="Frequency",
            y=target_fs_col,
            color="IOL_Model" if "IOL_Model" in df_fs_filtered.columns else None,
            title=f"Frequency Sweep ({fs_option})",
            labels={"Frequency": "Frequency (c/mm)", target_fs_col: f"FS ({fs_option})"},
        )

        # 4 & 6) Application des bornes X (0 à freq_max) et Y (y_min à y_max)
        fig_fs.update_layout(
            xaxis=dict(range=[0, freq_max]),
            yaxis=dict(range=[y_min, y_max]),
        )

        st.plotly_chart(fig_fs, use_container_width=True)
    else:
        st.warning("Aucune donnée à afficher pour la sélection actuelle.")
else:
        st.error("Le fichier Excel_FS.xlsx n'a pas été trouvé.")


if __name__ == "__main__":
    # Ne s'exécute en Tkinter/Win32 que si lancé directement depuis un terminal Python Windows
    pass

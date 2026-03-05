import pandas as pd
import os
from tools.api import load_config

def preprocess_linkedin_data():
    config = load_config()
    
    input_filename = config["files"]["input_raw"]
    output_filename = config["files"]["input_preprocessed"]
    
    li_raw_cols = config["columns"]["linkedin_raw"]
    li_clean_col = config["columns"]["linkedin_clean"]
    activity_col = config["columns"]["activity"]

    if not os.path.exists(input_filename):
        print(f"Datei {input_filename} nicht gefunden!")
        return

    df = pd.read_excel(input_filename)
    
    # 1. Dynamisch die richtige Spalte für LinkedIn-URLs finden
    col_name = None
    for col in li_raw_cols:
        if col in df.columns:
            col_name = col
            break
            
    if not col_name:
        print(f"Keine passende LinkedIn-Spalte gefunden! Erwartet: {li_raw_cols}")
        return

    def clean_linkedin_url(url):
        if pd.isna(url): return None
        return str(url).split('?')[0].rstrip('/')

    def create_activity_url(url):
        if not url: return None
        return f"{url}/recent-activity/all/"

    # 2. LinkedIn Profil bereinigen und Beitrags-Spalte ("User social") erstellen
    df[li_clean_col] = df[col_name].apply(clean_linkedin_url)
    df[activity_col] = df[li_clean_col].apply(create_activity_url)

    # 3. Spalten anordnen (Activity direkt nach LinkedIn)
    cols = list(df.columns)
    if activity_col in cols:
        cols.remove(activity_col)
    
    if li_clean_col in cols:
        li_index = cols.index(li_clean_col)
        cols.insert(li_index + 1, activity_col)
    
    df = df[cols]

    # 4. Speichern
    df.to_excel(output_filename, index=False)
    print(f"Erfolgreich gespeichert unter: {output_filename}")
    print(df[[li_clean_col, activity_col]].head())

if __name__ == "__main__":
    preprocess_linkedin_data()
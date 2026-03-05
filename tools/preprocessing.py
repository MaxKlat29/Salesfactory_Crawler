import pandas as pd
import yaml
from pathlib import Path

# Absolute Pfade relativ zum Verzeichnis des Skripts definieren
BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config.yaml"
INPUT_FILE = BASE_DIR / "input.xlsx"
OUTPUT_FILE = BASE_DIR / "contacts_urls.xlsx"

def load_config():
    """Lädt die config.yaml, um die Spaltennamen auszulesen."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def clean_linkedin_url(url):
    """Entfernt Tracking-Parameter (alles ab dem '?') und abschließende Slashes."""
    if pd.isna(url) or not isinstance(url, str):
        return url
    clean_url = url.split('?')[0].rstrip('/')
    return clean_url

def generate_activity_url(clean_url):
    """Baut aus der sauberen Profil-URL den direkten Link zur Aktivitäten-Seite."""
    if pd.isna(clean_url) or not isinstance(clean_url, str):
        return None
    
    # Prüfen, ob es wirklich ein Profil-Link ist
    if "linkedin.com/in/" in clean_url:
        return f"{clean_url}/recent-activity/all/"
    return None

def main():
    print("Starte Preprocessing...")

    # 1. Config laden (für die möglichen Spaltennamen)
    try:
        config = load_config()
        linkedin_columns = config.get('columns', {}).get('linkedin_raw', ['LinkedIn Profile URL', 'Person Linkedin Url', 'LinkedIn'])
    except FileNotFoundError:
        print("Warnung: config.yaml nicht gefunden. Verwende Standard-Spaltennamen.")
        linkedin_columns = ['LinkedIn Profile URL', 'Person Linkedin Url', 'LinkedIn']

    # 2. Excel-Datei einlesen
    if not INPUT_FILE.exists():
        print(f"Fehler: '{INPUT_FILE.name}' wurde im Hauptverzeichnis nicht gefunden.")
        return

    try:
        df = pd.read_excel(INPUT_FILE)
    except Exception as e:
        print(f"Fehler beim Lesen der Excel-Datei: {e}")
        return

    initial_rows = len(df)

    # 3. LEERE ZEILEN ENTFERNEN (Der Fix)
    # Entfernt alle Zeilen, die in JEDER Spalte leer (NaN) sind
    df = df.dropna(how='all')
    
    # 4. Richtige LinkedIn-Spalte identifizieren
    target_col = None
    for col in linkedin_columns:
        if col in df.columns:
            target_col = col
            break
    
    if not target_col:
        print(f"Abbruch: Keine passende LinkedIn-Spalte gefunden. Gesucht wurde nach: {linkedin_columns}")
        print(f"Vorhandene Spalten in der Datei: {list(df.columns)}")
        return

    # Zusätzliche Bereinigung: Entfernt auch Zeilen, bei denen explizit die LinkedIn-URL fehlt
    df = df.dropna(subset=[target_col])
    
    # Index nach dem Löschen sauber neu aufbauen
    df = df.reset_index(drop=True)
    
    print(f"Bereinigung: {initial_rows - len(df)} leere/ungültige Zeilen entfernt.")

    # 5. URLs verarbeiten
    print(f"Extrahiere und bereinige Links aus der Spalte '{target_col}'...")
    df['LinkedIn_Clean'] = df[target_col].apply(clean_linkedin_url)
    df['LinkedIn_Activity'] = df['LinkedIn_Clean'].apply(generate_activity_url)

    # 6. Finale Datei speichern
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Erfolg! {len(df)} saubere Leads wurden in '{OUTPUT_FILE.name}' gespeichert.")

if __name__ == "__main__":
    main()
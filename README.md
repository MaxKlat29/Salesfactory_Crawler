# Salesfactory_Crawler

> Ein KI-gestütztes, lokales Automatisierungs-Tool, das LinkedIn-Profile crawlt, Leads bewertet und hochpersonalisierte Outreach-Nachrichten generiert.

Dieses Projekt wandelt rohe Kontaktlisten (z. B. aus Snov.io-Exporten) in versandfertige LinkedIn-Kampagnen um. Es umgeht gängige API-Restriktionen durch lokales Browser-Scraping und nutzt fortschrittliche LLMs, um Profil- und Beitragsdaten zu analysieren und hochgradig personalisierte Nachrichteneinstiege („Firmenwinkel") zu formulieren.

---

## 1. Systemarchitektur

Das Tool ist modular aufgebaut, um Code-Logik strikt von Inhalten (Prompts, Konfigurationen) zu trennen.

| Datei / Modul | Funktion |
| :--- | :--- |
| `config.yaml` | Zentrale Steuereinheit. Enthält API-Schlüssel, Modelleinstellungen, Dateipfade, Spaltenzuordnungen und sämtliche LLM-Prompts. |
| `tools/preprocessing.py` | Bereinigt die rohe Eingabeliste. Entfernt Tracking-Parameter von URLs und generiert die Links für die LinkedIn-Aktivitäten-Seiten. |
| `tools/crawler.py` | Physisches Bot-Modul. Nutzt `pyautogui` zur Steuerung des lokalen Browsers (Tab-Handling, Strg+A, Strg+C), um Inhalte in die Zwischenablage zu kopieren. |
| `tools/api.py` | Schnittstellenmodul zu OpenRouter. Lädt die Konfiguration und regelt den Request-/Response-Zyklus mit dem LLM. |
| `run.py` | Der Haupt-Orchestrator. Iteriert durch die Kontaktliste, koordiniert Crawler und LLM, sendet Statusmeldungen und speichert die finalen Ergebnisse. |

---

## 2. Setup & Installation

### 2.1 Lokale Umgebung vorbereiten

**Voraussetzungen:** Python 3.9+, ein Standard-Browser mit aktivem LinkedIn-Login.

```bash
# Virtuelle Umgebung erstellen
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate

# Abhängigkeiten installieren
pip install pandas openpyxl pyautogui pyperclip openai pyyaml tqdm
```

---

### 2.2 OpenRouter API einrichten

1. **Account erstellen** → [openrouter.ai](https://openrouter.ai)
2. **Guthaben aufladen** → unter „Credits" (empfohlen: 5–10 $)
3. **API-Key generieren** → unter „Keys" → „Create Key" → Schlüssel sofort kopieren
4. **Modell:** Standardmäßig ist `qwen/qwen3-max` hinterlegt – exzellentes Preis-Leistungs-Verhältnis für Copywriting

---

### 2.3 `config.yaml` anpassen

```yaml
openrouter:
  base_url: "https://openrouter.ai/api/v1"
  api_key: "sk-or-v1-HIER-DEIN-KEY-EINTRAGEN"
  model: "qwen/qwen3-max"
```

---

## 3. Konfigurations-Parameter (`config.yaml`)

| Kategorie | Parameter | Beschreibung |
| :--- | :--- | :--- |
| `settings` | `sleep_sec` | Pausenzeit in Sekunden zwischen Kontakten. Verhindert Rate-Limits des Browsers. |
| `settings` | `overwrite` | `true`: Überschreibt bestehende Nachrichten. `false`: Überspringt bereits verarbeitete Kontakte. |
| `settings` | `enable_company_angle` | Aktiviert einen separaten LLM-Aufruf zur Analyse spezifischer Unternehmensherausforderungen. |
| `settings` | `company_angle_cache` | Speichert Firmenwinkel im Arbeitsspeicher, um API-Kosten bei redundanten Kontakten zu sparen. |
| `columns` | `linkedin_raw` | Array möglicher Spaltennamen für die anfängliche LinkedIn-URL aus dem Snov.io-Export. |
| `prompts` | `static_body` | Der unveränderliche Standard-Pitch, der am Ende jeder personalisierten Nachricht angehängt wird. |

---

## 4. Ausführungs-Workflow

### Schritt 1 — Daten bereitstellen
Exportiere die Kontaktliste aus Snov.io und lege sie als `input.xlsx` ins Hauptverzeichnis.

### Schritt 2 — Preprocessing ausführen
```bash
python tools/preprocessing.py
```
Erzeugt die bereinigte Arbeitsdatei `contacts_urls.xlsx`.

### Schritt 3 — Agenten starten
```bash
python run.py
```

---

### Verhalten während der Ausführung

| Phase | Beschreibung |
| :--- | :--- |
| **Initialisierung** | Der Agent sendet eine WhatsApp-Statusmeldung mit Zeitstempel. |
| **⚠️ Hands-Off-Regel** | Da `pyautogui` den Browser physisch steuert, dürfen Maus und Tastatur **nicht** bedient werden, solange ein aktiver Crawl-Vorgang läuft. |
| **Quality Gate** | Nach den ersten **3 generierten Nachrichten** pausiert das Skript zur manuellen Kontrolle im Terminal. |
| **Batch-Verarbeitung** | Nach Bestätigung mit `Enter` wird die restliche Liste vollautomatisch abgearbeitet. |

---

## 5. Output & Ergebnisse

Die Resultate werden fortlaufend in `contacts_out.xlsx` gespeichert. Die Datei wird um folgende Spalten erweitert:

| Spalte | Beschreibung |
| :--- | :--- |
| **Matching Punkte** | Fachliche Bewertung (0–100) basierend auf relevanter Expertise (z. B. KI & Compliance). |
| **LinkedIn Aktivität Punkte** | Bewertung (0–100) der Beitrags- und Interaktionsfrequenz des Leads. |
| **Firmenwinkel** | Durch das LLM extrahierter, individueller Kontext-Satz zum Unternehmen. |
| **Vorlage** | Finale, absendebereite Nachricht aus personalisiertem Einstieg + `static_body`. |

---

## 6. Aufbau der `input.xlsx`

Das Tool verarbeitet Standard-Exporte von Lead-Plattformen wie Snov.io oder Apollo. Entscheidend ist, dass die Spaltennamen für die LinkedIn-URLs mit deiner `config.yaml` übereinstimmen.

### Beispiel-Struktur (Header in Zeile 1)

| First Name | Last Name | Job Title | Company Name | LinkedIn Profile URL |
| :--- | :--- | :--- | :--- | :--- |
| Max | Mustermann | Head of Sales | Musterfirma GmbH | `https://www.linkedin.com/in/max-mustermann-123/` |
| Anna | Schmidt | CEO & Founder | TechNova Solutions | `https://www.linkedin.com/in/anna-schmidt-tech/` |
| Lukas | Weber | Director Business Development | Weber Consulting | `https://www.linkedin.com/in/lukas-weber-bd/` |

### Wichtige Regeln

| Regel | Details |
| :--- | :--- |
| **LinkedIn-Spalte ist Pflicht** | Der Crawler braucht die Profil-URL. Der Spaltenname (z. B. `LinkedIn Profile URL`, `Person Linkedin Url` oder `LinkedIn`) muss exakt im Array `linkedin_raw` in der `config.yaml` eingetragen sein. |
| **Personalisierungs-Daten** | `First Name`, `Last Name` und `Company Name` sind für das LLM essenziell – der Agent nutzt diese Felder für die Anrede und den Firmenbezug im Pitch. |
| **Keine leeren Zeilen** | Komplett leere Zeilen zwischen Leads können dazu führen, dass Pandas die Verarbeitung abbricht. |
| **Dateiformat** | Zwingend als `.xlsx` (Excel-Arbeitsmappe) speichern – **nicht** als `.csv`. |

---

## 7. Troubleshooting

| Fehlerbild | Ursache | Lösung |
| :--- | :--- | :--- |
| LLM-Output enthält `"FEHLER"` | API-Limit erreicht oder fehlerhafter API-Key. | OpenRouter-Guthaben und Key in `config.yaml` prüfen. |
| Crawl-Ergebnis ist leer oder fehlerhaft | LinkedIn hat die Seite nicht schnell genug geladen. | `sleep_after_open` in `tools/crawler.py` erhöhen. |
| Zwischenablage enthält falschen Text | Maus/Tastatur während des Crawls manuell bedient. | Eingabegeräte während aktivem Crawl-Vorgang nicht verwenden. |
| Spalte nicht gefunden (Preprocessing) | Abweichende Spaltennamen im Snov.io-Export. | Exakten Spaltennamen in `config.yaml` unter `columns → linkedin_raw` eintragen. |
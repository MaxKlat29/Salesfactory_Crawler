from __future__ import annotations

from pathlib import Path
import re
import time
import sys
from typing import Optional, Dict

import pandas as pd
from tqdm import tqdm

from tools.crawler import crawl
from tools.api import call, load_config


# ========= Konfiguration laden =========
CONFIG = load_config()

INPUT_XLSX = CONFIG["files"]["input_preprocessed"]
OUTPUT_XLSX = CONFIG["files"]["output"]
SLEEP_SEC = CONFIG["settings"]["sleep_sec"]
OVERWRITE = CONFIG["settings"]["overwrite"]

ENABLE_COMPANY_ANGLE = CONFIG["settings"]["enable_company_angle"]
COMPANY_ANGLE_CACHE = CONFIG["settings"]["company_angle_cache"]

ENABLE_MATCHING_SCORE = CONFIG["settings"]["enable_matching_score"]
MATCHING_SCORE_COL = CONFIG["columns"]["matching_score"]
MATCHING_RAW_COL = CONFIG["columns"]["matching_raw"]

ENABLE_ACTIVITY_SCORE = CONFIG["settings"]["enable_activity_score"]
ACTIVITY_SCORE_COL = CONFIG["columns"]["activity_score"]
ACTIVITY_RAW_COL = CONFIG["columns"]["activity_raw"]

TOTAL_SCORE_COL = CONFIG["columns"]["total_score"]
TEMPLATE_COL = CONFIG["columns"]["template"]
COMPANY_ANGLE_COL = CONFIG["columns"]["company_angle"]
LI_CLEAN_COL = CONFIG["columns"]["linkedin_clean"]
ACTIVITY_COL = CONFIG["columns"]["activity"]

FORBIDDEN_CHARS = ("–", "—", ":", ";", "•")
URL_RE = re.compile(r"(https?://[^\s)]+)")

# ========= Prompts =========
STATIC_BODY = CONFIG["prompts"]["static_body"]
SYSTEM_GUARDRAILS_OPENING = CONFIG["prompts"]["system_guardrails_opening"]
USER_PROMPT_OPENING = CONFIG["prompts"]["user_prompt_opening"]
COMPANY_ANGLE_GUARDRAILS = CONFIG["prompts"]["company_angle_guardrails"]
MATCHING_SCORE_PROMPT = CONFIG["prompts"]["matching_score_prompt"]
ACTIVITY_SCORE_PROMPT = CONFIG["prompts"]["activity_score_prompt"]


# ========= Company Normalization =========
_LEGAL_SUFFIXES = {
    "ag", "gmbh", "kg", "kgaa", "se", "sa", "s.a.", "sarl", "s.r.l.", "srl",
    "plc", "ltd", "llc", "inc", "corp", "co", "company", "limited", "bv", "nv",
    "oy", "ab", "sas", "spa", "pte", "pte.", "ohg", "ug", "e.v.", "ev", "stiftung", "foundation"
}

_GENERIC_TAIL_WORDS = {
    "versicherung", "versicherungs", "insurance", "bank", "gruppe", "group", "holding", "holdings",
    "deutschland", "germany", "europe", "europa", "beratungs", "vertriebs"
}

_SPECIAL_SPOKEN = {
    "allianz": "bei der Allianz",
    "r+v": "bei der R und V",
    "rv": "bei der R und V",
}


def _norm_company_spaces(s: str) -> str:
    return re.sub(r"\s{2,}", " ", (s or "").strip())


def company_short_name(company: str) -> str:
    c = _norm_company_spaces(company)
    if not c:
        return ""

    low = c.lower()
    if "allianz" in low:
        return "Allianz"
    if "r+v" in c:
        return "R+V"

    c = re.sub(r"\([^)]*\)", " ", c)
    c = _norm_company_spaces(c)

    parts = c.split()
    while parts:
        tail = parts[-1].strip(".,").lower()
        if tail in _LEGAL_SUFFIXES:
            parts = parts[:-1]
            continue
        break
    c = _norm_company_spaces(" ".join(parts))
    if not c:
        return ""

    parts = c.split()
    while len(parts) >= 2:
        tail = parts[-1].strip(".,").lower()
        if tail in _GENERIC_TAIL_WORDS:
            parts = parts[:-1]
            continue
        break
    c = _norm_company_spaces(" ".join(parts))

    parts = c.split()
    if len(parts) >= 3:
        second = parts[1].lower()
        if second in {"bank", "group", "gruppe", "holding"}:
            return _norm_company_spaces(" ".join(parts[:2]))
        if parts[1].lower() == "group":
            return _norm_company_spaces(" ".join(parts[:2]))
        return _norm_company_spaces(parts[0])

    return c


def company_spoken_phrase(company_short: str) -> str:
    name = _norm_company_spaces(company_short)
    if not name:
        return ""

    key = name.lower()
    if key in _SPECIAL_SPOKEN:
        return _SPECIAL_SPOKEN[key]

    if name.isupper() and 2 <= len(name) <= 8:
        return f"bei der {name}"

    if " " in name:
        return f"bei der {name}"

    return f"bei {name}"


def enforce_company_spoken(text: str, company_full: str, company_short: str, company_spoken: str) -> str:
    if not text:
        return text
    if not company_spoken:
        return text

    t = text
    if company_full:
        t = t.replace(company_full, company_short or company_full)

    if company_short:
        cs = re.escape(company_short)
        t = re.sub(rf"\bbei\s+der\s+{cs}\b", company_spoken, t)
        t = re.sub(rf"\bbei\s+dem\s+{cs}\b", company_spoken, t)
        t = re.sub(rf"\bbei\s+den\s+{cs}\b", company_spoken, t)
        t = re.sub(rf"\bbei\s+{cs}\b", company_spoken, t)

    return t


# ========= Prompt Builders =========
def make_opening_prompt(
    dump: str,
    first_name: str,
    last_name: str,
    company_spoken: str,
    role: str,
    company_angle: str,
) -> str:
    meta = (
        "Kontakt Infos\n"
        f"Vorname {first_name}\n"
        f"Nachname {last_name}\n"
        f"Unternehmen Umgang {company_spoken}\n"
        f"Rolle {role}\n\n"
    )
    angle = ""
    if company_angle and company_angle.strip() and company_angle.strip().upper() != "LEER":
        angle = f"Optionaler Kontext Satz\n{company_angle.strip()}\n\n"

    return (
        f"{SYSTEM_GUARDRAILS_OPENING}\n"
        f"{USER_PROMPT_OPENING}\n\n"
        f"{meta}"
        f"{angle}"
        "LinkedIn Kontext Dump\n"
        f"{dump}"
    )


def make_company_angle_prompt(company_spoken: str, role: str, dump: str) -> str:
    return (
        f"{COMPANY_ANGLE_GUARDRAILS}\n\n"
        "Aufgabe\n"
        "Erzeuge einen einzigen Satz als optionalen Kontext.\n"
        "Der Satz klingt wie eine menschliche Einordnung und zeigt, dass du das Profil oder den Beitrag verstanden hast.\n"
        "Kein Fragezeichen.\n"
        "Keine externen Fakten, keine Zahlen.\n\n"
        f"Firma Umgang {company_spoken}\n"
        f"Rolle {role}\n\n"
        "LinkedIn Kontext Dump\n"
        f"{dump}"
    )


def make_matching_prompt(dump: str) -> str:
    return MATCHING_SCORE_PROMPT.format(text=dump or "")


def make_activity_prompt(dump: str) -> str:
    return ACTIVITY_SCORE_PROMPT.format(text=dump or "")


# ========= Utilities =========
def to_url(value) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, float):
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.startswith("http://") or s.startswith("https://"):
        return s
    m = URL_RE.search(s)
    return m.group(1) if m else None


def pick_first_nonempty(row: pd.Series, cols: list[str]) -> str:
    for c in cols:
        if c in row and pd.notna(row[c]):
            s = str(row[c]).strip()
            if s:
                return s
    return ""


def infer_company_from_dump(dump: str) -> str:
    if not dump:
        return ""
    patterns = [
        r"Unternehmen\s*[:\n]\s*([^\n\r]{2,80})",
        r"Company\s*[:\n]\s*([^\n\r]{2,80})",
        r"\bbei\s+([A-Z][A-Za-z0-9&.,\s]{2,60})\b",
        r"\bat\s+([A-Z][A-Za-z0-9&.,\s]{2,60})\b",
    ]
    for p in patterns:
        m = re.search(p, dump, flags=re.IGNORECASE)
        if m:
            cand = re.sub(r"\s{2,}", " ", m.group(1)).strip()
            if 2 <= len(cand) <= 80 and "linkedin" not in cand.lower():
                return cand
    return ""


def collapse_whitespace(txt: str) -> str:
    if not isinstance(txt, str):
        return ""
    lines = txt.splitlines()
    cleaned = [" ".join(line.split()) for line in lines]
    return "\n".join(cleaned).strip()


def sanitize_text(txt: str) -> str:
    if not txt:
        return ""

    txt = txt.replace("\r\n", "\n").replace("\r", "\n")

    for ch in FORBIDDEN_CHARS:
        txt = txt.replace(ch, "")

    txt = txt.replace("–", " ").replace("—", " ")
    txt = re.sub(r"(?<=\w)-(?=\w)", " ", txt)
    txt = re.sub(r"\s-\s", " ", txt)
    txt = re.sub(r"[ \t]{2,}", " ", txt)
    txt = re.sub(r"\s+,", ",", txt)
    txt = re.sub(r",\s{2,}", ", ", txt)
    txt = re.sub(r",\s*,+", ",", txt)

    txt = "\n".join([l.strip() for l in txt.splitlines()]).strip()
    return txt


def normalize_two_sentence_body_no_question(body: str) -> str:
    b = " ".join((body or "").split()).strip()

    parts = re.split(r"(?<=[.!?])\s+", b)
    parts = [p.strip() for p in parts if p.strip()]

    s1 = ""
    for p in parts:
        if "linkedin" in p.lower():
            s1 = p
            break
    if not s1:
        s1 = "Auf LinkedIn bin ich ueber Ihren Beitrag oder Ihr Profil gestolpert."

    s2 = ""
    for p in parts:
        if p != s1 and "linkedin" not in p.lower():
            s2 = p
            break
    if not s2:
        s2 = "Gerade die Uebersetzung von Anforderungen in saubere Prozesse wirkt in der Praxis oft wie der Knackpunkt."

    s1 = s1.replace("?", ".")
    s2 = s2.replace("?", ".")

    if s1 and s1[-1] not in ".!":
        s1 += "."
    if s2 and s2[-1] not in ".!":
        s2 += "."

    return f"{s1} {s2}"


def enforce_message_format(txt: str, last_name: str, company_full: str, company_short: str, company_spoken: str) -> str:
    if not txt:
        return ""

    txt = txt.replace("\r\n", "\n").replace("\r", "\n").strip()

    if "\n" not in txt and "," in txt:
        left, right = txt.split(",", 1)
        greeting = left.strip() + ","
        body = right.strip()
        txt = f"{greeting}\n\n{body}"

    lines = txt.splitlines()
    if lines and "," in lines[0]:
        left, right = lines[0].split(",", 1)
        greeting = left.strip() + ","
        after = right.strip()
        tail = [l.strip() for l in lines[1:] if l.strip()]
        body = " ".join([after] + tail).strip() if after else " ".join(tail).strip()
    else:
        body = " ".join([l.strip() for l in lines if l.strip()]).strip()
        greeting = f"Hallo Herr {last_name}," if last_name else "Hallo,"

    body = enforce_company_spoken(body, company_full=company_full, company_short=company_short, company_spoken=company_spoken)
    body = sanitize_text(body)
    body = normalize_two_sentence_body_no_question(body)
    body = enforce_company_spoken(body, company_full=sanitize_text(company_full), company_short=company_short, company_spoken=company_spoken)

    return f"{greeting.strip()}\n\n{body.strip()}"


def safe_call_llm(prompt: str, context: str, retries: int = 2) -> str:
    last_err = None
    for _ in range(retries + 1):
        try:
            ans = call(prompt, context)
            return (ans or "").strip()
        except Exception as e:
            last_err = e
            time.sleep(0.8)
    return f"[FEHLER] {type(last_err).__name__}: {last_err}"


# ========= Score Parsing =========
_SCORE_PATTERNS = [
    re.compile(r"\bSCORE\s*=\s*(\d{1,3})\s*%?\b", re.IGNORECASE),
    re.compile(r"\b(\d{1,3})\s*%\b"),
    re.compile(r"\b(\d{1,3})\b"),
]

_ACTIVITY_PATTERNS = [
    re.compile(r"\bACTIVITY\s*=\s*(\d{1,3})\b", re.IGNORECASE),
    re.compile(r"\b(\d{1,3})\b"),
]


def extract_points_from_score(text: str) -> Optional[int]:
    if not text:
        return None
    t = text.strip()
    for pat in _SCORE_PATTERNS:
        m = pat.search(t)
        if m:
            try:
                v = int(m.group(1))
            except Exception:
                continue
            if 0 <= v <= 100:
                return v
    return None


def extract_points_from_activity(text: str) -> Optional[int]:
    if not text:
        return None
    t = text.strip()
    for pat in _ACTIVITY_PATTERNS:
        m = pat.search(t)
        if m:
            try:
                v = int(m.group(1))
            except Exception:
                continue
            if 0 <= v <= 100:
                return v
    return None


def print_contact_preview(idx: int, row: pd.Series, url1: str | None, url2: str | None, text: str):
    print("\n" + "=" * 80)
    print(f"PREVIEW #{idx + 1}")
    print("-" * 80)
    for col in ["Name", "Vorname", "Nachname", "Firma", "Unternehmen", "Position", "Rolle", "E-Mail", "Email"]:
        if col in row and pd.notna(row[col]):
            print(f"{col}: {row[col]}")

    if TOTAL_SCORE_COL in row and pd.notna(row[TOTAL_SCORE_COL]) and str(row[TOTAL_SCORE_COL]).strip():
        print(f"{TOTAL_SCORE_COL}: {row[TOTAL_SCORE_COL]}")
    if MATCHING_SCORE_COL in row and pd.notna(row[MATCHING_SCORE_COL]) and str(row[MATCHING_SCORE_COL]).strip():
        print(f"{MATCHING_SCORE_COL}: {row[MATCHING_SCORE_COL]}")
    if ACTIVITY_SCORE_COL in row and pd.notna(row[ACTIVITY_SCORE_COL]) and str(row[ACTIVITY_SCORE_COL]).strip():
        print(f"{ACTIVITY_SCORE_COL}: {row[ACTIVITY_SCORE_COL]}")

    print(f"LinkedIn: {url1 or '-'}")
    print(f"{ACTIVITY_COL}: {url2 or '-'}")
    print("-" * 80)
    print("VORSCHLAG NACHRICHT:\n")
    print(text)
    print("=" * 80 + "\n")


def wait_for_user_confirmation():
    if sys.stdin is None or sys.stdin.closed:
        print("[Hinweis] Kein interaktives stdin verfügbar – fahre fort.")
        return

    if not sys.stdin.isatty():
        print(
            "[Hinweis] Skript läuft nicht in einem interaktiven Terminal.\n"
            "Abbruch mit STRG+C möglich, sonst läuft es automatisch weiter."
        )
        try:
            for i in range(10, 0, -1):
                print(f"  Fahre in {i} Sekunden fort ...", end="\r", flush=True)
                time.sleep(1)
            print("  Fahre jetzt fort.           ")
        except KeyboardInterrupt:
            print("\nAbbruch durch Benutzer.")
            raise SystemExit(1)
        return

    while True:
        try:
            ans = input(
                "Bitte prüfe die oben angezeigten Nachrichten.\n"
                "Enter = weiter, 'n' = abbrechen: "
            ).strip().lower()
        except EOFError:
            print("\n[Hinweis] EOF bei input() – fahre fort.")
            return

        if ans in ("", "j", "y", "ja", "yes"):
            return
        if ans in ("n", "no", "nein"):
            print("Abbruch durch Benutzer.")
            raise SystemExit(1)

        print("Bitte Enter, 'j' oder 'n' eingeben.\n")


# ========= Hauptlogik =========
def process_excel(
    input_xlsx: str | Path = INPUT_XLSX,
    output_xlsx: str | Path = OUTPUT_XLSX,
    overwrite: bool = OVERWRITE,
):
    input_xlsx = Path(input_xlsx)
    if not input_xlsx.exists():
        raise FileNotFoundError(f"Eingabedatei nicht gefunden: {input_xlsx}")

    df = pd.read_excel(input_xlsx)

    for col in [LI_CLEAN_COL, ACTIVITY_COL]:
        if col not in df.columns:
            raise ValueError(f"Spalte fehlt: {col}")

    if TEMPLATE_COL not in df.columns:
        df[TEMPLATE_COL] = ""

    if ENABLE_MATCHING_SCORE and MATCHING_SCORE_COL not in df.columns:
        df[MATCHING_SCORE_COL] = ""
    if ENABLE_MATCHING_SCORE and MATCHING_RAW_COL not in df.columns:
        df[MATCHING_RAW_COL] = ""

    if ENABLE_ACTIVITY_SCORE and ACTIVITY_SCORE_COL not in df.columns:
        df[ACTIVITY_SCORE_COL] = ""
    if ENABLE_ACTIVITY_SCORE and ACTIVITY_RAW_COL not in df.columns:
        df[ACTIVITY_RAW_COL] = ""

    if (ENABLE_MATCHING_SCORE or ENABLE_ACTIVITY_SCORE) and TOTAL_SCORE_COL not in df.columns:
        df[TOTAL_SCORE_COL] = ""

    if ENABLE_COMPANY_ANGLE and COMPANY_ANGLE_COL not in df.columns:
        df[COMPANY_ANGLE_COL] = ""

    indices_to_process: list[int] = []
    for idx, row in df.iterrows():
        existing = row.get(TEMPLATE_COL, "")
        if overwrite or not (isinstance(existing, str) and existing.strip()):
            indices_to_process.append(idx)

    total = len(indices_to_process)
    if total == 0:
        print("Nichts zu tun – alle Vorlagen sind bereits gefüllt und OVERWRITE ist False.")
        return

    print(f"Starte Verarbeitung von {total} Kontakten ...")

    company_angle_cache: Dict[str, str] = {}

    def _clear_scores(idx: int):
        if ENABLE_MATCHING_SCORE:
            df.at[idx, MATCHING_SCORE_COL] = ""
            df.at[idx, MATCHING_RAW_COL] = ""
        if ENABLE_ACTIVITY_SCORE:
            df.at[idx, ACTIVITY_SCORE_COL] = ""
            df.at[idx, ACTIVITY_RAW_COL] = ""
        df.at[idx, TOTAL_SCORE_COL] = ""

    def generate_for_row(idx: int) -> str:
        row = df.iloc[idx]

        url1 = to_url(row.get(LI_CLEAN_COL))
        url2 = to_url(row.get(ACTIVITY_COL))

        if not url1 and not url2:
            df.at[idx, TEMPLATE_COL] = "[Kein verwertbarer Link vorhanden]"
            _clear_scores(idx)
            return df.at[idx, TEMPLATE_COL]

        if not url1:
            url1 = url2
        if not url2:
            url2 = url1

        try:
            dump = crawl(url1, url2) or ""
        except Exception as e:
            df.at[idx, TEMPLATE_COL] = f"[FEHLER Crawl] {type(e).__name__}: {e}"
            _clear_scores(idx)
            return df.at[idx, TEMPLATE_COL]

        dump = (dump or "").strip()

        first_name = pick_first_nonempty(row, ["Vorname", "First Name", "FirstName"])
        last_name = pick_first_nonempty(row, ["Nachname", "Last Name", "LastName"])
        role = pick_first_nonempty(row, ["Position", "Rolle", "Title"])

        company_full = pick_first_nonempty(row, ["Firma", "Unternehmen", "Company"])
        if not company_full:
            company_full = infer_company_from_dump(dump)

        company_short = company_short_name(company_full) if company_full else ""
        company_spoken = company_spoken_phrase(company_short) if company_short else ""

        matching_points: Optional[int] = None
        activity_points: Optional[int] = None

        if ENABLE_MATCHING_SCORE:
            mp = make_matching_prompt(dump=dump)
            raw = safe_call_llm(mp, context=dump, retries=2)
            raw_clean = collapse_whitespace(raw)
            matching_points = extract_points_from_score(raw_clean)
            df.at[idx, MATCHING_RAW_COL] = raw_clean
            df.at[idx, MATCHING_SCORE_COL] = matching_points if matching_points is not None else ""

        if ENABLE_ACTIVITY_SCORE:
            ap = make_activity_prompt(dump=dump)
            araw = safe_call_llm(ap, context=dump, retries=2)
            araw_clean = collapse_whitespace(araw)
            activity_points = extract_points_from_activity(araw_clean)
            df.at[idx, ACTIVITY_RAW_COL] = araw_clean
            df.at[idx, ACTIVITY_SCORE_COL] = activity_points if activity_points is not None else ""

        total_points = 0
        if matching_points is not None:
            total_points += matching_points
        if activity_points is not None:
            total_points += activity_points

        df.at[idx, TOTAL_SCORE_COL] = (
            total_points if (matching_points is not None or activity_points is not None) else ""
        )

        company_angle = ""
        if ENABLE_COMPANY_ANGLE and company_spoken:
            key = company_spoken.strip().lower()
            if COMPANY_ANGLE_CACHE and key in company_angle_cache:
                company_angle = company_angle_cache[key]
            else:
                p = make_company_angle_prompt(company_spoken=company_spoken, role=role, dump=dump)
                ca = safe_call_llm(p, context=dump, retries=2)
                ca = collapse_whitespace(ca)
                ca = sanitize_text(ca)
                ca = ca.replace("?", ".")
                if ca.strip().upper() == "LEER":
                    ca = ""
                company_angle = ca
                if COMPANY_ANGLE_CACHE:
                    company_angle_cache[key] = company_angle

            df.at[idx, COMPANY_ANGLE_COL] = company_angle

        prompt = make_opening_prompt(
            dump=dump,
            first_name=first_name,
            last_name=last_name,
            company_spoken=company_spoken,
            role=role,
            company_angle=company_angle,
        )

        opening = safe_call_llm(prompt, context=dump, retries=2)
        opening = collapse_whitespace(opening)

        opening = enforce_company_spoken(opening, company_full=company_full, company_short=company_short, company_spoken=company_spoken)
        opening = sanitize_text(opening)
        opening = enforce_message_format(
            opening,
            last_name=last_name,
            company_full=company_full,
            company_short=company_short,
            company_spoken=company_spoken,
        )

        full_msg = f"{opening}\n\n{STATIC_BODY}"
        df.at[idx, TEMPLATE_COL] = full_msg
        return full_msg

    preview_count = min(3, total)
    for i in range(preview_count):
        idx = indices_to_process[i]
        msg = generate_for_row(idx)

        row = df.iloc[idx]
        url1 = to_url(row.get(LI_CLEAN_COL))
        url2 = to_url(row.get(ACTIVITY_COL))
        if not url1:
            url1 = url2
        if not url2:
            url2 = url1

        print_contact_preview(idx, row, url1, url2, msg)

    if preview_count > 0:
        wait_for_user_confirmation()

    if total > preview_count:
        for idx in tqdm(indices_to_process[preview_count:], desc="Kontakte", unit="row"):
            _ = generate_for_row(idx)
            time.sleep(SLEEP_SEC)

    if TEMPLATE_COL in df.columns:
        cols = list(df.columns)

        score_cols = [TOTAL_SCORE_COL, MATCHING_SCORE_COL, ACTIVITY_SCORE_COL]
        raw_cols = [MATCHING_RAW_COL, ACTIVITY_RAW_COL]

        for c in score_cols + raw_cols:
            if c in cols:
                cols.remove(c)

        insert_pos = cols.index(TEMPLATE_COL)

        for c in reversed(score_cols):
            if c in df.columns:
                cols.insert(insert_pos, c)

        insert_after_vorlage = cols.index(TEMPLATE_COL) + 1
        for i, c in enumerate(raw_cols):
            if c in df.columns:
                cols.insert(insert_after_vorlage + i, c)

        df = df[cols]

    out = Path(output_xlsx)
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out, index=False)
    print(f"Fertig. Output: {out.resolve()}")


if __name__ == "__main__":
    process_excel()
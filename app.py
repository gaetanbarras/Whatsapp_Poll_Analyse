from __future__ import annotations

import hmac
import io
import json
import re
import sqlite3
import unicodedata
from contextlib import closing
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from streamlit.errors import StreamlitSecretNotFoundError


ANALYSIS_MODE_FREE = "Analyse libre"
ANALYSIS_MODE_CATEGORIZED = "Analyse catégorisée"
MANUAL_CATEGORY_OPTIONS = ["Automatique", "Présent", "Absent", "Autre", "Pas de réponse"]
DB_PATH = Path(__file__).resolve().parent / ".app_data" / "whatsapp_poll.db"
DEFAULT_LOCAL_SESSION_NAME = "Session locale"
CURRENT_EVENT_SESSION_KEY = "current_event_name"
NEW_EVENT_OPTION = "__new_event__"
PENDING_EVENT_SELECTION_KEY = "pending_selected_event_option"
APP_PASSWORD_SECRET_KEY = "app_password"
AUTHENTICATION_STATE_KEY = "password_authenticated"
MEMBER_BASE_EDITOR_COLUMNS = ["Prénom", "Nom", "Nom complet", "Téléphone", "Âge", "Sexe", "Fonction"]
MEMBER_BASE_MAPPING = {
    "first_name": "Prénom",
    "last_name": "Nom",
    "full_name": "Nom complet",
    "phone": "Téléphone",
    "age": "Âge",
    "gender": "Sexe",
    "function": "Fonction",
}

REFERENCE_COLUMN_ALIASES = {
    "first_name": ["prénom", "prenom", "first name", "firstname", "given name"],
    "last_name": ["nom", "lastname", "last name", "surname", "family name"],
    "full_name": ["nom complet", "full name", "fullname", "display name", "name"],
    "phone": [
        "téléphone portable",
        "telephone portable",
        "portable",
        "mobile",
        "phone",
        "telephone",
        "téléphone",
        "gsm",
        "natel",
        "tel",
    ],
    "gender": ["sexe", "sex", "genre", "civilité", "civilite"],
    "function": ["fonction", "instrument", "section", "role", "rôle", "poste"],
    "age": ["âge", "age", "date de naissance", "annee de naissance", "année de naissance", "birth date", "birth year"],
}

WHATSAPP_COLUMN_ALIASES = {
    "name": ["name", "nom", "display name", "participant"],
    "phone": ["phone", "phone number", "telephone", "téléphone", "mobile", "portable"],
}

RESULT_COLUMN_LABELS = {
    "prenom": "Prénom",
    "nom": "Nom",
    "nom_complet": "Nom complet",
    "age": "Âge",
    "sexe": "Sexe",
    "fonction": "Fonction",
    "invite": "Invité",
    "telephone_reference_original": "Téléphone original référence",
    "telephone_reference_normalise": "Téléphone normalisé référence",
    "nom_whatsapp": "Nom WhatsApp",
    "telephone_whatsapp_original": "Téléphone original WhatsApp",
    "telephone_whatsapp_normalise": "Téléphone normalisé WhatsApp",
    "reponses_brutes_detectees": "Réponses brutes détectées",
    "categorie_automatique": "Catégorie automatique",
    "categorie_manuelle": "Catégorie manuelle",
    "categorie_finale": "Catégorie finale",
    "statut_reponse": "Statut réponse",
    "commentaire_suivi": "Commentaire suivi",
    "commentaire_anomalie": "Commentaire anomalie",
}

RESULT_COLUMN_ORDER = list(RESULT_COLUMN_LABELS.keys())

UNKNOWN_COLUMN_ORDER = [
    "nom_whatsapp",
    "telephone_whatsapp_original",
    "telephone_whatsapp_normalise",
    "reponses_brutes_detectees",
    "categorie_finale",
    "statut_reponse",
    "commentaire_anomalie",
]

ANOMALY_COLUMN_ORDER = [
    "type_anomalie",
    "nom_concerne",
    "telephone_original",
    "telephone_normalise",
    "detail",
]


def normalize_label(value: Any) -> str:
    """Normalise un libellé pour faciliter les détections automatiques."""
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-zA-Z0-9]+", " ", text).strip().lower()


def clean_value(value: Any) -> str:
    """Retourne une valeur texte nettoyée."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def append_message(current: str, message: str) -> str:
    """Ajoute un message dans une cellule texte sans écraser l'existant."""
    current = clean_value(current)
    message = clean_value(message)
    if not message:
        return current
    if not current:
        return message
    return f"{current} | {message}"


def bool_to_label(value: Any) -> str:
    """Affiche un booléen en français."""
    return "Oui" if bool(value) else "Non"


def normalize_function_name(value: Any) -> str:
    """Normalise l'affichage de la fonction pour les synthèses."""
    return clean_value(value) or "Non renseigné"


def compute_years_since(birth_date: datetime, now: datetime) -> str:
    """Retourne l'âge entier à partir d'une date de naissance."""
    years = now.year - birth_date.year - ((now.month, now.day) < (birth_date.month, birth_date.day))
    return str(max(years, 0))


def clean_age_value(value: Any) -> str:
    """Convertit une colonne d'âge, d'année ou de date en âge lisible."""
    text = clean_value(value)
    if not text:
        return ""

    now = datetime.now()
    if re.fullmatch(r"\d{1,3}", text):
        return text

    if re.fullmatch(r"\d{4}", text):
        year = int(text)
        if 1900 <= year <= now.year:
            return str(now.year - year)

    if re.fullmatch(r"\d{5}", text):
        try:
            excel_birth_date = pd.Timestamp("1899-12-30") + pd.to_timedelta(int(text), unit="D")
            if 1900 <= excel_birth_date.year <= now.year:
                return compute_years_since(excel_birth_date.to_pydatetime(), now)
        except (TypeError, ValueError):
            pass

    for fmt in (
        "%d.%m.%Y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%d.%m.%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M",
    ):
        try:
            birth_date = datetime.strptime(text, fmt)
            if 1900 <= birth_date.year <= now.year:
                return compute_years_since(birth_date, now)
        except ValueError:
            continue

    for dayfirst in (True, False):
        birth_date = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if pd.notna(birth_date) and 1900 <= int(birth_date.year) <= now.year:
            return compute_years_since(birth_date.to_pydatetime(), now)

    return text


def normalize_phone(phone: Any) -> tuple[str, bool, str]:
    """Convertit un numéro suisse vers le format interne 41791234567."""
    original = clean_value(phone)
    if not original:
        return "", False, "Numéro vide"

    digits = re.sub(r"\D+", "", original)
    if not digits:
        return "", False, "Aucun chiffre détecté"

    normalized = digits
    if digits.startswith("00410"):
        normalized = "41" + digits[5:]
    elif digits.startswith("0041"):
        normalized = "41" + digits[4:]
    elif digits.startswith("410"):
        normalized = "41" + digits[3:]
    elif digits.startswith("41"):
        normalized = digits
    elif digits.startswith("0"):
        normalized = "41" + digits[1:]
    elif len(digits) == 9:
        normalized = "41" + digits
    else:
        return digits, False, "Format suisse non reconnu"

    if normalized.startswith("410"):
        normalized = "41" + normalized[3:]

    if not normalized.startswith("41"):
        return normalized, False, "Indicatif suisse manquant"

    if len(normalized) != 11:
        return normalized, False, "Longueur inattendue"

    return normalized, True, ""


def build_member_key(full_name: str, first_name: str, last_name: str, phone_normalized: str, phone_original: str) -> str:
    """Construit une clé stable de membre pour la persistance locale."""
    identity = clean_value(full_name) or f"{clean_value(first_name)} {clean_value(last_name)}".strip()
    phone_part = clean_value(phone_normalized) or re.sub(r"\D+", "", clean_value(phone_original))
    identity_part = normalize_label(identity) or "membre"
    phone_part = phone_part or "sans-telephone"
    return f"{identity_part}|{phone_part}"


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie les noms de colonnes et supprime les colonnes vides."""
    cleaned = df.copy()
    cleaned.columns = [clean_value(column) for column in cleaned.columns]
    cleaned = cleaned.fillna("")
    keep_columns: list[str] = []
    for column in cleaned.columns:
        series = cleaned[column].astype(str).str.strip()
        normalized_column = normalize_label(column)
        is_empty_column = normalized_column == "" and (series == "").all()
        is_unnamed_empty = normalized_column.startswith("unnamed") and (series == "").all()
        if is_empty_column or is_unnamed_empty:
            continue
        if (series == "").all():
            continue
        keep_columns.append(column)
    return cleaned.loc[:, keep_columns]


def score_dataframe_quality(df: pd.DataFrame) -> int:
    """Donne un score simple pour choisir la meilleure lecture CSV."""
    non_empty_headers = sum(1 for column in df.columns if normalize_label(column))
    if df.empty:
        return non_empty_headers * 100
    row_has_value = df.astype(str).apply(lambda row: row.str.strip().ne("").any(), axis=1)
    return non_empty_headers * 100 + min(int(row_has_value.sum()), 50)


def ensure_storage() -> None:
    """Crée le dossier de stockage local si besoin."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)


def get_db_connection() -> sqlite3.Connection:
    """Retourne une connexion SQLite locale."""
    ensure_storage()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def ensure_table_column(connection: sqlite3.Connection, table_name: str, column_name: str, definition: str) -> None:
    """Ajoute une colonne SQLite si elle n'existe pas encore."""
    existing_columns = {
        row["name"]
        for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    }
    if column_name not in existing_columns:
        connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")


def initialize_database() -> None:
    """Initialise la base locale de l'application."""
    with closing(get_db_connection()) as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS reference_sources (
                source_name TEXT PRIMARY KEY,
                reference_name TEXT,
                reference_bytes BLOB,
                reference_separator TEXT,
                settings_json TEXT NOT NULL DEFAULT '{}',
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS event_sessions (
                event_name TEXT PRIMARY KEY,
                reference_source_name TEXT,
                reference_name TEXT,
                reference_bytes BLOB,
                reference_separator TEXT,
                whatsapp_name TEXT,
                whatsapp_bytes BLOB,
                whatsapp_separator TEXT,
                settings_json TEXT NOT NULL DEFAULT '{}',
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS members (
                member_key TEXT PRIMARY KEY,
                phone_normalized TEXT,
                first_name TEXT,
                last_name TEXT,
                full_name TEXT,
                phone_original TEXT,
                age_text TEXT NOT NULL DEFAULT '',
                sex TEXT NOT NULL DEFAULT '',
                function_name TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS event_member_overrides (
                event_name TEXT NOT NULL,
                member_key TEXT NOT NULL,
                invite INTEGER NOT NULL DEFAULT 0,
                manual_category TEXT NOT NULL DEFAULT '',
                follow_up_note TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (event_name, member_key)
            )
            """
        )

        ensure_table_column(connection, "members", "age_text", "TEXT NOT NULL DEFAULT ''")
        ensure_table_column(connection, "members", "sex", "TEXT NOT NULL DEFAULT ''")
        ensure_table_column(connection, "members", "function_name", "TEXT NOT NULL DEFAULT ''")

        connection.execute("CREATE INDEX IF NOT EXISTS idx_members_phone_normalized ON members(phone_normalized)")
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_event_member_overrides_event ON event_member_overrides(event_name)"
        )

        legacy_rows = connection.execute("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'saved_state'").fetchall()
        has_events = connection.execute("SELECT COUNT(*) AS count FROM event_sessions").fetchone()["count"]
        if legacy_rows and not has_events:
            legacy = connection.execute("SELECT * FROM saved_state ORDER BY updated_at DESC LIMIT 1").fetchone()
            if legacy:
                connection.execute(
                    """
                    INSERT OR IGNORE INTO event_sessions (
                        event_name,
                        reference_name,
                        reference_bytes,
                        reference_separator,
                        whatsapp_name,
                        whatsapp_bytes,
                        whatsapp_separator,
                        settings_json,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        "Session migrée",
                        legacy["reference_name"],
                        legacy["reference_bytes"],
                        legacy["reference_separator"],
                        legacy["whatsapp_name"],
                        legacy["whatsapp_bytes"],
                        legacy["whatsapp_separator"],
                        legacy["settings_json"],
                        legacy["updated_at"],
                    ),
                )
        connection.commit()


def list_saved_events() -> list[dict[str, str]]:
    """Retourne la liste des événements enregistrés."""
    with closing(get_db_connection()) as connection:
        rows = connection.execute(
            """
            SELECT event_name, reference_source_name, updated_at
            FROM event_sessions
            ORDER BY updated_at DESC, event_name ASC
            """
        ).fetchall()
    return [dict(row) for row in rows]


def list_reference_sources() -> list[dict[str, str]]:
    """Retourne la liste des sources membres enregistrées."""
    with closing(get_db_connection()) as connection:
        rows = connection.execute(
            """
            SELECT source_name, reference_name, updated_at
            FROM reference_sources
            ORDER BY source_name ASC
            """
        ).fetchall()
    return [dict(row) for row in rows]


def load_event_snapshot(event_name: str | None) -> dict[str, Any] | None:
    """Charge un événement sauvegardé."""
    if not clean_value(event_name):
        return None

    with closing(get_db_connection()) as connection:
        row = connection.execute(
            "SELECT * FROM event_sessions WHERE event_name = ?",
            (clean_value(event_name),),
        ).fetchone()

    if row is None:
        return None

    settings = {}
    try:
        settings = json.loads(row["settings_json"] or "{}")
    except json.JSONDecodeError:
        settings = {}

    return {
        "event_name": row["event_name"],
        "reference_source_name": row["reference_source_name"],
        "reference_name": row["reference_name"],
        "reference_bytes": row["reference_bytes"],
        "reference_separator": row["reference_separator"] or "Auto",
        "whatsapp_name": row["whatsapp_name"],
        "whatsapp_bytes": row["whatsapp_bytes"],
        "whatsapp_separator": row["whatsapp_separator"] or "Auto",
        "settings": settings,
        "updated_at": row["updated_at"],
    }


def load_saved_snapshot() -> dict[str, Any] | None:
    """Charge la session locale la plus récente."""
    with closing(get_db_connection()) as connection:
        row = connection.execute(
            """
            SELECT event_name
            FROM event_sessions
            ORDER BY updated_at DESC, event_name ASC
            LIMIT 1
            """
        ).fetchone()

    if row is None:
        return None

    return load_event_snapshot(row["event_name"])


def load_reference_source(source_name: str | None) -> dict[str, Any] | None:
    """Charge une source membres enregistrée."""
    if not clean_value(source_name):
        return None

    with closing(get_db_connection()) as connection:
        row = connection.execute(
            "SELECT * FROM reference_sources WHERE source_name = ?",
            (clean_value(source_name),),
        ).fetchone()

    if row is None:
        return None

    settings = {}
    try:
        settings = json.loads(row["settings_json"] or "{}")
    except json.JSONDecodeError:
        settings = {}

    return {
        "source_name": row["source_name"],
        "reference_name": row["reference_name"],
        "reference_bytes": row["reference_bytes"],
        "reference_separator": row["reference_separator"] or "Auto",
        "settings": settings,
        "updated_at": row["updated_at"],
    }


def save_reference_source(source_name: str, reference_source: dict[str, Any], settings: dict[str, Any] | None = None) -> None:
    """Enregistre ou met à jour une source membres réutilisable."""
    normalized_name = clean_value(source_name)
    if not normalized_name:
        raise ValueError("Le nom de la source membres est obligatoire.")

    settings_payload = settings or {}
    with closing(get_db_connection()) as connection:
        connection.execute(
            """
            INSERT INTO reference_sources (
                source_name,
                reference_name,
                reference_bytes,
                reference_separator,
                settings_json,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(source_name) DO UPDATE SET
                reference_name = excluded.reference_name,
                reference_bytes = excluded.reference_bytes,
                reference_separator = excluded.reference_separator,
                settings_json = excluded.settings_json,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                normalized_name,
                reference_source.get("name"),
                reference_source.get("bytes"),
                reference_source.get("separator", "Auto"),
                json.dumps(settings_payload, ensure_ascii=False),
            ),
        )
        connection.commit()


def save_event_session(
    event_name: str,
    reference_source_name: str | None,
    reference_source: dict[str, Any],
    whatsapp_source: dict[str, Any],
    settings: dict[str, Any],
) -> None:
    """Sauvegarde un événement nommé."""
    normalized_event_name = clean_value(event_name)
    if not normalized_event_name:
        raise ValueError("Le nom de l'événement est obligatoire.")

    with closing(get_db_connection()) as connection:
        connection.execute(
            """
            INSERT INTO event_sessions (
                event_name,
                reference_source_name,
                reference_name,
                reference_bytes,
                reference_separator,
                whatsapp_name,
                whatsapp_bytes,
                whatsapp_separator,
                settings_json,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(event_name) DO UPDATE SET
                reference_source_name = excluded.reference_source_name,
                reference_name = excluded.reference_name,
                reference_bytes = excluded.reference_bytes,
                reference_separator = excluded.reference_separator,
                whatsapp_name = excluded.whatsapp_name,
                whatsapp_bytes = excluded.whatsapp_bytes,
                whatsapp_separator = excluded.whatsapp_separator,
                settings_json = excluded.settings_json,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                normalized_event_name,
                clean_value(reference_source_name),
                reference_source.get("name"),
                reference_source.get("bytes"),
                reference_source.get("separator", "Auto"),
                whatsapp_source.get("name"),
                whatsapp_source.get("bytes"),
                whatsapp_source.get("separator", "Auto"),
                json.dumps(settings, ensure_ascii=False),
            ),
        )
        connection.commit()


def save_session_snapshot(
    reference_source: dict[str, Any],
    whatsapp_source: dict[str, Any],
    settings: dict[str, Any],
    event_name: str | None = None,
) -> str:
    """Sauvegarde la session locale courante."""
    target_event_name = clean_value(event_name) or DEFAULT_LOCAL_SESSION_NAME
    save_event_session(target_event_name, None, reference_source, whatsapp_source, settings)
    return target_event_name


def delete_event_session(event_name: str) -> None:
    """Supprime un événement sauvegardé."""
    with closing(get_db_connection()) as connection:
        connection.execute("DELETE FROM event_sessions WHERE event_name = ?", (clean_value(event_name),))
        connection.execute("DELETE FROM event_member_overrides WHERE event_name = ?", (clean_value(event_name),))
        connection.commit()


def rename_event_session(current_event_name: str, new_event_name: str) -> str:
    """Renomme un événement sauvegardé et conserve ses données associées."""
    normalized_current_name = clean_value(current_event_name)
    normalized_new_name = clean_value(new_event_name)
    if not normalized_current_name:
        raise ValueError("Le nom actuel de l'événement est obligatoire.")
    if not normalized_new_name:
        raise ValueError("Le nouveau nom de l'événement est obligatoire.")
    if normalized_current_name == normalized_new_name:
        raise ValueError("Saisissez un nouveau nom différent de l'actuel.")

    with closing(get_db_connection()) as connection:
        current_row = connection.execute(
            "SELECT 1 FROM event_sessions WHERE event_name = ?",
            (normalized_current_name,),
        ).fetchone()
        if current_row is None:
            raise ValueError("L'événement à renommer est introuvable.")

        target_row = connection.execute(
            "SELECT 1 FROM event_sessions WHERE event_name = ?",
            (normalized_new_name,),
        ).fetchone()
        if target_row is not None:
            raise ValueError("Un événement porte déjà ce nom.")

        connection.execute(
            """
            UPDATE event_sessions
            SET event_name = ?, updated_at = CURRENT_TIMESTAMP
            WHERE event_name = ?
            """,
            (normalized_new_name, normalized_current_name),
        )
        connection.execute(
            "UPDATE event_member_overrides SET event_name = ? WHERE event_name = ?",
            (normalized_new_name, normalized_current_name),
        )
        connection.commit()

    return normalized_new_name


def clear_saved_snapshot(event_name: str | None = None) -> None:
    """Supprime la session locale courante."""
    target_event_name = clean_value(event_name)
    if not target_event_name:
        snapshot = load_saved_snapshot()
        target_event_name = snapshot.get("event_name", "") if snapshot else ""
    if target_event_name:
        delete_event_session(target_event_name)


def get_empty_member_base_frame() -> pd.DataFrame:
    """Retourne une trame vide pour l'éditeur de la base membres."""
    return pd.DataFrame(columns=MEMBER_BASE_EDITOR_COLUMNS)


def build_member_base_editor_frame(reference_data: pd.DataFrame) -> pd.DataFrame:
    """Convertit une table interne de référence vers la base membres éditable."""
    if reference_data.empty:
        return get_empty_member_base_frame()

    editor = reference_data.copy()
    for column in ["prenom", "nom", "nom_complet", "telephone_reference_original", "age", "sexe", "fonction"]:
        if column not in editor.columns:
            editor[column] = ""

    return editor[
        ["prenom", "nom", "nom_complet", "telephone_reference_original", "age", "sexe", "fonction"]
    ].rename(
        columns={
            "prenom": "Prénom",
            "nom": "Nom",
            "nom_complet": "Nom complet",
            "telephone_reference_original": "Téléphone",
            "age": "Âge",
            "sexe": "Sexe",
            "fonction": "Fonction",
        }
    ).fillna("")


def get_member_registry_summary() -> dict[str, Any]:
    """Retourne un résumé simple de la base membres locale."""
    with closing(get_db_connection()) as connection:
        row = connection.execute(
            """
            SELECT COUNT(*) AS member_count, MAX(updated_at) AS updated_at
            FROM members
            """
        ).fetchone()

    return {
        "count": int(row["member_count"] or 0),
        "updated_at": clean_value(row["updated_at"]),
    }


def load_member_base() -> pd.DataFrame:
    """Charge la base membres locale dans un format éditable."""
    with closing(get_db_connection()) as connection:
        rows = connection.execute(
            """
            SELECT
                first_name,
                last_name,
                full_name,
                phone_original,
                age_text,
                sex,
                function_name
            FROM members
            ORDER BY
                CASE
                    WHEN TRIM(full_name) <> '' THEN LOWER(TRIM(full_name))
                    ELSE LOWER(TRIM(first_name || ' ' || last_name))
                END,
                phone_normalized
            """
        ).fetchall()

    if not rows:
        return get_empty_member_base_frame()

    return pd.DataFrame([dict(row) for row in rows]).rename(
        columns={
            "first_name": "Prénom",
            "last_name": "Nom",
            "full_name": "Nom complet",
            "phone_original": "Téléphone",
            "age_text": "Âge",
            "sex": "Sexe",
            "function_name": "Fonction",
        }
    ).fillna("")


def save_member_base(reference_data: pd.DataFrame) -> dict[str, int]:
    """Remplace la base membres locale à partir d'une référence normalisée."""
    payload = []
    for _, row in reference_data.iterrows():
        member_key = clean_value(row.get("member_key"))
        if not member_key:
            continue
        payload.append(
            (
                member_key,
                clean_value(row.get("telephone_reference_normalise")),
                clean_value(row.get("prenom")),
                clean_value(row.get("nom")),
                clean_value(row.get("nom_complet")),
                clean_value(row.get("telephone_reference_original")),
                clean_value(row.get("age")),
                clean_value(row.get("sexe")),
                clean_value(row.get("fonction")),
            )
        )

    with closing(get_db_connection()) as connection:
        connection.execute("DELETE FROM members")
        if payload:
            connection.executemany(
                """
                INSERT INTO members (
                    member_key,
                    phone_normalized,
                    first_name,
                    last_name,
                    full_name,
                    phone_original,
                    age_text,
                    sex,
                    function_name,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """,
                payload,
            )
        connection.commit()

    return {
        "count": len(payload),
        "invalid_count": int((~reference_data["telephone_reference_valide"]).sum()) if not reference_data.empty else 0,
        "duplicate_count": int(reference_data.loc[reference_data["doublon_reference"] > 1, "member_key"].nunique())
        if not reference_data.empty
        else 0,
    }


def save_member_base_from_editor(editor_df: pd.DataFrame) -> dict[str, int]:
    """Enregistre la base membres locale depuis la table éditable."""
    candidate = editor_df.copy() if not editor_df.empty else get_empty_member_base_frame()
    for column in MEMBER_BASE_EDITOR_COLUMNS:
        if column not in candidate.columns:
            candidate[column] = ""
    candidate = candidate.loc[:, MEMBER_BASE_EDITOR_COLUMNS].fillna("")
    candidate = candidate.loc[
        candidate.apply(
            lambda row: any(clean_value(row.get(column)) for column in MEMBER_BASE_EDITOR_COLUMNS),
            axis=1,
        )
    ].reset_index(drop=True)
    if candidate.empty:
        return save_member_base(pd.DataFrame(columns=[
            "member_key",
            "telephone_reference_valide",
            "doublon_reference",
        ]))

    reference_data = build_reference_dataframe(candidate, MEMBER_BASE_MAPPING)
    return save_member_base(reference_data)


def sync_reference_members(reference_data: pd.DataFrame, event_name: str) -> pd.DataFrame:
    """Synchronise la liste des membres avec la base locale et recharge les overrides de l'événement."""
    result = reference_data.copy()
    result["invite"] = False
    result["categorie_manuelle"] = ""
    result["commentaire_suivi"] = ""

    if result.empty:
        return result

    rows = result.to_dict("records")
    with closing(get_db_connection()) as connection:
        for row in rows:
            connection.execute(
                """
                INSERT INTO members (
                    member_key,
                    phone_normalized,
                    first_name,
                    last_name,
                    full_name,
                    phone_original,
                    age_text,
                    sex,
                    function_name,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(member_key) DO UPDATE SET
                    phone_normalized = excluded.phone_normalized,
                    first_name = excluded.first_name,
                    last_name = excluded.last_name,
                    full_name = excluded.full_name,
                    phone_original = excluded.phone_original,
                    age_text = CASE WHEN excluded.age_text <> '' THEN excluded.age_text ELSE members.age_text END,
                    sex = CASE WHEN excluded.sex <> '' THEN excluded.sex ELSE members.sex END,
                    function_name = CASE WHEN excluded.function_name <> '' THEN excluded.function_name ELSE members.function_name END,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (
                    row["member_key"],
                    row["telephone_reference_normalise"],
                    row["prenom"],
                    row["nom"],
                    row["nom_complet"],
                    row["telephone_reference_original"],
                    row["age"],
                    row["sexe"],
                    row["fonction"],
                ),
            )
        connection.commit()

        keys = result["member_key"].dropna().astype(str).tolist()
        placeholders = ",".join("?" for _ in keys)
        member_query = f"SELECT member_key, age_text, sex, function_name FROM members WHERE member_key IN ({placeholders})"
        override_query = (
            f"SELECT member_key, invite, manual_category, follow_up_note "
            f"FROM event_member_overrides WHERE event_name = ? AND member_key IN ({placeholders})"
        )
        stored_rows = connection.execute(member_query, keys).fetchall() if keys else []
        override_rows = connection.execute(override_query, [clean_value(event_name), *keys]).fetchall() if keys else []

    if not stored_rows:
        return result

    stored = pd.DataFrame([dict(row) for row in stored_rows])
    stored = stored.rename(
        columns={
            "age_text": "age_db",
            "sex": "sexe_db",
            "function_name": "fonction_db",
        }
    )
    overrides = pd.DataFrame([dict(row) for row in override_rows]) if override_rows else pd.DataFrame(columns=["member_key", "invite", "manual_category", "follow_up_note"])
    overrides = overrides.rename(
        columns={
            "invite": "invite_db",
            "manual_category": "categorie_manuelle_db",
            "follow_up_note": "commentaire_suivi_db",
        }
    )

    result = result.merge(stored, on="member_key", how="left")
    result = result.merge(overrides, on="member_key", how="left")
    result["age_db"] = result["age_db"].fillna("").map(clean_value)
    result["sexe_db"] = result["sexe_db"].fillna("").map(clean_value)
    result["fonction_db"] = result["fonction_db"].fillna("").map(clean_value)
    result["age"] = result["age_db"].where(result["age_db"].ne(""), result["age"])
    result["sexe"] = result["sexe_db"].where(result["sexe_db"].ne(""), result["sexe"])
    result["fonction"] = result["fonction_db"].where(result["fonction_db"].ne(""), result["fonction"])
    result["invite"] = result["invite_db"].fillna(0).astype(int).astype(bool)
    result["categorie_manuelle"] = result["categorie_manuelle_db"].fillna("").map(clean_value)
    result["commentaire_suivi"] = result["commentaire_suivi_db"].fillna("").map(clean_value)

    return result.drop(
        columns=[
            "age_db",
            "sexe_db",
            "fonction_db",
            "invite_db",
            "categorie_manuelle_db",
            "commentaire_suivi_db",
        ],
        errors="ignore",
    )


def save_member_updates(event_name: str, updates: pd.DataFrame) -> None:
    """Enregistre les modifications manuelles du suivi pour un événement donné."""
    if updates.empty:
        return

    payload = []
    for _, row in updates.iterrows():
        member_key = clean_value(row.get("member_key"))
        if not member_key:
            continue
        manual_category = clean_value(row.get("Décision manuelle"))
        if manual_category == "Automatique":
            manual_category = ""
        payload.append(
            (
                clean_value(event_name),
                1 if bool(row.get("Invité")) else 0,
                manual_category,
                clean_value(row.get("Commentaire suivi")),
                member_key,
            )
        )

    with closing(get_db_connection()) as connection:
        connection.executemany(
            """
            INSERT INTO event_member_overrides (
                event_name,
                invite,
                manual_category,
                follow_up_note,
                member_key,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(event_name, member_key) DO UPDATE SET
                invite = excluded.invite,
                manual_category = excluded.manual_category,
                follow_up_note = excluded.follow_up_note,
                updated_at = CURRENT_TIMESTAMP
            """,
            payload,
        )
        connection.commit()


@st.cache_data(show_spinner=False)
def read_csv_bytes(file_bytes: bytes, separator_choice: str) -> tuple[pd.DataFrame, dict[str, str]]:
    """Charge un CSV avec plusieurs encodages et séparateurs possibles."""
    separators = [";", ",", "\t"]
    if separator_choice == "Auto":
        candidate_separators: list[str | None] = [None, *separators]
    else:
        candidate_separators = [separator_choice, *[sep for sep in separators if sep != separator_choice]]

    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
    best_result: tuple[pd.DataFrame, dict[str, str]] | None = None
    best_score = -1
    errors: list[str] = []

    for encoding in encodings:
        for separator in candidate_separators:
            try:
                buffer = io.BytesIO(file_bytes)
                if separator is None:
                    df = pd.read_csv(
                        buffer,
                        sep=None,
                        engine="python",
                        dtype=str,
                        encoding=encoding,
                        index_col=False,
                    )
                else:
                    df = pd.read_csv(
                        buffer,
                        sep=separator,
                        dtype=str,
                        encoding=encoding,
                        index_col=False,
                    )
                cleaned = clean_dataframe(df)
                score = score_dataframe_quality(cleaned)
                if score > best_score:
                    best_score = score
                    best_result = (
                        cleaned,
                        {
                            "encoding": encoding,
                            "separator": "auto" if separator is None else separator,
                            "file_type": "CSV",
                        },
                    )
            except Exception as exc:
                errors.append(f"{encoding} / {separator or 'auto'}: {exc}")

    if best_result is None:
        error_message = "Impossible de lire le fichier CSV."
        if errors:
            error_message = f"{error_message} Détails: {errors[0]}"
        raise ValueError(error_message)

    return best_result


@st.cache_data(show_spinner=False)
def load_reference_file(file_bytes: bytes, filename: str, separator_choice: str) -> tuple[pd.DataFrame, dict[str, str]]:
    """Charge le fichier de référence depuis Excel ou CSV."""
    lower_name = filename.lower()
    if lower_name.endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
        return clean_dataframe(df), {"file_type": "Excel", "sheet": "Première feuille"}
    if lower_name.endswith(".csv"):
        return read_csv_bytes(file_bytes, separator_choice)
    raise ValueError("Le fichier de référence doit être au format Excel ou CSV.")


@st.cache_data(show_spinner=False)
def load_whatsapp_votes(file_bytes: bytes, separator_choice: str) -> tuple[pd.DataFrame, dict[str, str]]:
    """Charge l'export WhatsApp au format CSV."""
    return read_csv_bytes(file_bytes, separator_choice)


def score_column_name(column_name: str, aliases: list[str], expected_field: str) -> int:
    """Calcule un score de similarité entre un nom de colonne et une liste d'alias."""
    normalized = normalize_label(column_name)
    if not normalized:
        return -1

    best_score = 0
    for index, alias in enumerate(aliases):
        normalized_alias = normalize_label(alias)
        if normalized == normalized_alias:
            best_score = max(best_score, 100 - index)
        elif normalized_alias in normalized:
            best_score = max(best_score, 75 - index)

    if expected_field == "phone":
        if any(token in normalized for token in ["portable", "mobile", "gsm", "natel"]):
            best_score += 15
        if any(token in normalized for token in ["fax", "professionnel", "professional"]):
            best_score -= 20

    return best_score


def guess_column(columns: list[str], aliases: dict[str, list[str]], field_name: str) -> str | None:
    """Tente de proposer automatiquement une colonne."""
    best_column = None
    best_score = 0
    for column in columns:
        score = score_column_name(column, aliases[field_name], field_name)
        if score > best_score:
            best_score = score
            best_column = column
    return best_column


def detect_poll_columns(df: pd.DataFrame, name_column: str | None, phone_column: str | None) -> list[str]:
    """Détecte les colonnes de réponses du sondage."""
    ignored = {clean_value(name_column), clean_value(phone_column)}
    return [column for column in df.columns if clean_value(column) not in ignored]


def is_selected_marker(value: Any) -> bool:
    """Détecte une case cochée dans l'export WhatsApp."""
    return clean_value(value).lower() == "x"


def extract_selected_options(row: pd.Series, poll_columns: list[str]) -> list[str]:
    """Retourne toutes les options marquées X pour une ligne donnée."""
    return [column for column in poll_columns if is_selected_marker(row.get(column, ""))]


def merge_unique_values(values: list[Any], preferred_order: list[str] | None = None) -> list[str]:
    """Fusionne des valeurs sans doublons en conservant un ordre lisible."""
    flattened: list[str] = []
    for value in values:
        if isinstance(value, list):
            flattened.extend(clean_value(item) for item in value if clean_value(item))
        else:
            cleaned = clean_value(value)
            if cleaned:
                flattened.append(cleaned)

    unique_values: list[str] = []
    seen: set[str] = set()

    if preferred_order:
        for item in preferred_order:
            cleaned = clean_value(item)
            if cleaned and cleaned in flattened and cleaned not in seen:
                unique_values.append(cleaned)
                seen.add(cleaned)

    for item in flattened:
        if item not in seen:
            unique_values.append(item)
            seen.add(item)

    return unique_values


def build_reference_dataframe(reference_df: pd.DataFrame, mapping: dict[str, str | None]) -> pd.DataFrame:
    """Construit la table de référence normalisée à partir du mapping choisi."""
    first_name_series = (
        reference_df[mapping["first_name"]].map(clean_value)
        if mapping.get("first_name")
        else pd.Series("", index=reference_df.index)
    )
    last_name_series = (
        reference_df[mapping["last_name"]].map(clean_value)
        if mapping.get("last_name")
        else pd.Series("", index=reference_df.index)
    )
    if mapping.get("full_name"):
        full_name_series = reference_df[mapping["full_name"]].map(clean_value)
    else:
        full_name_series = (first_name_series + " " + last_name_series).str.replace(r"\s+", " ", regex=True).str.strip()
        full_name_series = full_name_series.where(
            full_name_series.ne(""),
            first_name_series.where(first_name_series.ne(""), last_name_series),
        )

    phone_series = reference_df[mapping["phone"]].map(clean_value)
    normalized = phone_series.apply(normalize_phone)
    gender_series = (
        reference_df[mapping["gender"]].map(clean_value)
        if mapping.get("gender")
        else pd.Series("", index=reference_df.index)
    )
    age_series = (
        reference_df[mapping["age"]].map(clean_age_value)
        if mapping.get("age")
        else pd.Series("", index=reference_df.index)
    )
    function_series = (
        reference_df[mapping["function"]].map(clean_value)
        if mapping.get("function")
        else pd.Series("", index=reference_df.index)
    )

    result = pd.DataFrame(
        {
            "prenom": first_name_series,
            "nom": last_name_series,
            "nom_complet": full_name_series,
            "age": age_series,
            "sexe": gender_series,
            "fonction": function_series,
            "telephone_reference_original": phone_series,
            "telephone_reference_normalise": normalized.map(lambda item: item[0]),
            "telephone_reference_valide": normalized.map(lambda item: item[1]),
            "telephone_reference_detail": normalized.map(lambda item: item[2]),
            "anomalie_reference": "",
        }
    )

    result["member_key"] = result.apply(
        lambda row: build_member_key(
            row["nom_complet"],
            row["prenom"],
            row["nom"],
            row["telephone_reference_normalise"],
            row["telephone_reference_original"],
        ),
        axis=1,
    )

    invalid_mask = ~result["telephone_reference_valide"]
    result.loc[invalid_mask, "anomalie_reference"] = result.loc[invalid_mask, "telephone_reference_detail"]

    duplicate_counts = (
        result.loc[result["telephone_reference_valide"], "telephone_reference_normalise"].value_counts().to_dict()
    )
    result["doublon_reference"] = result["telephone_reference_normalise"].map(duplicate_counts).fillna(0).astype(int)
    duplicate_mask = result["doublon_reference"] > 1
    result.loc[duplicate_mask, "anomalie_reference"] = result.loc[duplicate_mask, "anomalie_reference"].map(
        lambda current: append_message(current, "Numéro présent plusieurs fois dans le fichier de référence")
    )

    return result


def is_summary_vote_row(row: pd.Series, name_column: str | None, poll_columns: list[str]) -> bool:
    """Ignore les lignes récapitulatives de type total présentes dans certains exports."""
    name_value = clean_value(row.get(name_column, "")) if name_column else ""
    poll_values = [clean_value(row.get(column, "")) for column in poll_columns]
    if name_value:
        return False
    has_marker = any(is_selected_marker(value) for value in poll_values)
    numeric_or_blank = all(not value or re.fullmatch(r"\d+([.,]\d+)?", value) for value in poll_values)
    return numeric_or_blank and not has_marker


def build_whatsapp_vote_rows(
    votes_df: pd.DataFrame,
    mapping: dict[str, str | None],
    poll_columns: list[str],
) -> pd.DataFrame:
    """Prépare toutes les lignes de vote WhatsApp avant regroupement."""
    name_series = votes_df[mapping["name"]].map(clean_value) if mapping.get("name") else pd.Series("", index=votes_df.index)
    phone_series = votes_df[mapping["phone"]].map(clean_value)
    normalized = phone_series.apply(normalize_phone)
    selected_options = votes_df.apply(lambda row: extract_selected_options(row, poll_columns), axis=1)
    summary_mask = votes_df.apply(lambda row: is_summary_vote_row(row, mapping.get("name"), poll_columns), axis=1)

    result = pd.DataFrame(
        {
            "nom_whatsapp": name_series,
            "telephone_whatsapp_original": phone_series,
            "telephone_whatsapp_normalise": normalized.map(lambda item: item[0]),
            "telephone_whatsapp_valide": normalized.map(lambda item: item[1]),
            "telephone_whatsapp_detail": normalized.map(lambda item: item[2]),
            "reponses_detectees_liste": selected_options,
            "reponses_brutes_detectees": selected_options.map(lambda item: " | ".join(item)),
            "ligne_resume": summary_mask,
            "numero_ligne_source": pd.Series(range(2, len(votes_df) + 2), index=votes_df.index),
            "anomalie_whatsapp": "",
        }
    )

    invalid_mask = ~result["telephone_whatsapp_valide"] & ~result["ligne_resume"]
    result.loc[invalid_mask, "anomalie_whatsapp"] = result.loc[invalid_mask, "telephone_whatsapp_detail"]
    return result


def aggregate_whatsapp_votes(
    vote_rows: pd.DataFrame,
    poll_columns: list[str],
    duplicate_strategy: str,
) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    """Regroupe les votes WhatsApp par numéro de téléphone normalisé."""
    valid_votes = vote_rows.loc[vote_rows["telephone_whatsapp_valide"] & ~vote_rows["ligne_resume"]].copy()
    invalid_votes = vote_rows.loc[~vote_rows["telephone_whatsapp_valide"] & ~vote_rows["ligne_resume"]].copy()
    ignored_summary_rows = int(vote_rows["ligne_resume"].sum())

    aggregated_rows: list[dict[str, Any]] = []
    if valid_votes.empty:
        aggregated = pd.DataFrame(
            columns=[
                "nom_whatsapp",
                "telephone_whatsapp_original",
                "telephone_whatsapp_normalise",
                "reponses_detectees_liste",
                "reponses_brutes_detectees",
                "doublon_whatsapp",
                "anomalie_whatsapp",
            ]
        )
        return aggregated, invalid_votes, ignored_summary_rows

    for phone_number, group in valid_votes.groupby("telephone_whatsapp_normalise", sort=False):
        vote_count = len(group)
        names = merge_unique_values(group["nom_whatsapp"].tolist())
        original_phones = merge_unique_values(group["telephone_whatsapp_original"].tolist())

        if duplicate_strategy == "Garder la dernière réponse":
            latest_row = group.iloc[-1]
            selected_options = list(latest_row["reponses_detectees_liste"])
            anomaly = ""
            if vote_count > 1:
                anomaly = f"{vote_count} votes WhatsApp détectés, dernière réponse conservée"
        else:
            selected_options = merge_unique_values(group["reponses_detectees_liste"].tolist(), preferred_order=poll_columns)
            anomaly = ""
            if vote_count > 1:
                anomaly = f"{vote_count} votes WhatsApp détectés, réponses fusionnées"

        if not selected_options:
            anomaly = append_message(anomaly, "Aucune option cochée")

        aggregated_rows.append(
            {
                "nom_whatsapp": " | ".join(names),
                "telephone_whatsapp_original": " | ".join(original_phones),
                "telephone_whatsapp_normalise": phone_number,
                "reponses_detectees_liste": selected_options,
                "reponses_brutes_detectees": " | ".join(selected_options),
                "doublon_whatsapp": vote_count,
                "anomalie_whatsapp": anomaly,
            }
        )

    return pd.DataFrame(aggregated_rows), invalid_votes, ignored_summary_rows


def categorize_response(
    selected_options: list[str],
    has_vote: bool,
    analysis_mode: str,
    category_mapping: dict[str, set[str]],
) -> str:
    """Attribue une catégorie métier à une réponse ou laisse le texte brut en mode libre."""
    if not has_vote:
        return "Pas de réponse"

    if analysis_mode == ANALYSIS_MODE_FREE:
        return " | ".join(selected_options) if selected_options else "Réponse vide"

    if not selected_options:
        return "Autre"

    present_hits = any(option in category_mapping["present"] for option in selected_options)
    absent_hits = any(option in category_mapping["absent"] for option in selected_options)
    other_hits = any(option in category_mapping["other"] for option in selected_options)
    mapped_options = category_mapping["present"] | category_mapping["absent"] | category_mapping["other"]
    unmapped_hits = any(option not in mapped_options for option in selected_options)

    if present_hits and absent_hits:
        return "Conflit"
    if present_hits:
        return "Présent"
    if absent_hits:
        return "Absent"
    if other_hits or unmapped_hits:
        return "Autre"
    return "Autre"


def merge_votes_with_reference(
    reference_data: pd.DataFrame,
    aggregated_votes: pd.DataFrame,
    analysis_mode: str,
    category_mapping: dict[str, set[str]],
) -> pd.DataFrame:
    """Fusionne la référence, les votes et les overrides manuels."""
    merged = reference_data.merge(
        aggregated_votes,
        how="left",
        left_on="telephone_reference_normalise",
        right_on="telephone_whatsapp_normalise",
    )

    for column in [
        "nom_whatsapp",
        "telephone_whatsapp_original",
        "telephone_whatsapp_normalise",
        "reponses_brutes_detectees",
        "anomalie_whatsapp",
    ]:
        if column not in merged.columns:
            merged[column] = ""
        merged[column] = merged[column].fillna("")

    if "reponses_detectees_liste" not in merged.columns:
        merged["reponses_detectees_liste"] = [[] for _ in range(len(merged))]
    merged["reponses_detectees_liste"] = merged["reponses_detectees_liste"].apply(
        lambda value: value if isinstance(value, list) else []
    )

    if "doublon_whatsapp" not in merged.columns:
        merged["doublon_whatsapp"] = 0
    merged["doublon_whatsapp"] = merged["doublon_whatsapp"].fillna(0).astype(int)

    merged["a_vote"] = merged["telephone_whatsapp_normalise"].ne("")
    merged["statut_reponse_source"] = merged.apply(
        lambda row: "Téléphone invalide"
        if not row["telephone_reference_valide"]
        else ("A répondu" if row["a_vote"] else "Pas de réponse"),
        axis=1,
    )
    merged["categorie_automatique"] = merged.apply(
        lambda row: categorize_response(
            row["reponses_detectees_liste"],
            bool(row["a_vote"]),
            analysis_mode,
            category_mapping,
        ),
        axis=1,
    )

    merged["categorie_manuelle"] = merged["categorie_manuelle"].fillna("").map(clean_value)
    merged["categorie_finale"] = merged.apply(
        lambda row: row["categorie_manuelle"] if row["categorie_manuelle"] else row["categorie_automatique"],
        axis=1,
    )
    merged["statut_reponse"] = merged["statut_reponse_source"]
    merged.loc[
        merged["categorie_manuelle"].ne("") & merged["statut_reponse_source"].eq("Pas de réponse"),
        "statut_reponse",
    ] = "Mis à jour manuellement"
    merged.loc[
        merged["categorie_manuelle"].ne("") & merged["statut_reponse_source"].eq("A répondu"),
        "statut_reponse",
    ] = "A répondu (ajusté manuellement)"

    merged["commentaire_anomalie"] = merged["anomalie_reference"].fillna("")
    merged["commentaire_anomalie"] = merged.apply(
        lambda row: append_message(row["commentaire_anomalie"], row["anomalie_whatsapp"]),
        axis=1,
    )
    merged["invite"] = merged["invite"].fillna(False).astype(bool)
    merged["commentaire_suivi"] = merged["commentaire_suivi"].fillna("").map(clean_value)

    return merged


def build_unknown_votes(
    aggregated_votes: pd.DataFrame,
    reference_data: pd.DataFrame,
    analysis_mode: str,
    category_mapping: dict[str, set[str]],
) -> pd.DataFrame:
    """Construit la liste des numéros présents dans WhatsApp mais absents de la référence."""
    if aggregated_votes.empty:
        return pd.DataFrame(columns=UNKNOWN_COLUMN_ORDER)

    known_numbers = set(reference_data.loc[reference_data["telephone_reference_valide"], "telephone_reference_normalise"].tolist())
    unknown = aggregated_votes.loc[~aggregated_votes["telephone_whatsapp_normalise"].isin(known_numbers)].copy()
    if unknown.empty:
        return pd.DataFrame(columns=UNKNOWN_COLUMN_ORDER)

    unknown["categorie_finale"] = unknown["reponses_detectees_liste"].apply(
        lambda options: categorize_response(options, True, analysis_mode, category_mapping)
    )
    unknown["statut_reponse"] = "Numéro inconnu"
    unknown["commentaire_anomalie"] = unknown["anomalie_whatsapp"].fillna("")
    return unknown


def build_anomalies(
    reference_data: pd.DataFrame,
    aggregated_votes: pd.DataFrame,
    invalid_votes: pd.DataFrame,
) -> pd.DataFrame:
    """Consolide les doublons et autres anomalies détectées."""
    rows: list[dict[str, str]] = []

    for _, row in reference_data.loc[~reference_data["telephone_reference_valide"]].iterrows():
        rows.append(
            {
                "type_anomalie": "Participant - téléphone invalide",
                "nom_concerne": row["nom_complet"] or f"{row['prenom']} {row['nom']}".strip(),
                "telephone_original": row["telephone_reference_original"],
                "telephone_normalise": row["telephone_reference_normalise"],
                "detail": row["telephone_reference_detail"],
            }
        )

    for _, row in reference_data.loc[reference_data["doublon_reference"] > 1].iterrows():
        rows.append(
            {
                "type_anomalie": "Participant - doublon de téléphone",
                "nom_concerne": row["nom_complet"] or f"{row['prenom']} {row['nom']}".strip(),
                "telephone_original": row["telephone_reference_original"],
                "telephone_normalise": row["telephone_reference_normalise"],
                "detail": f"Numéro partagé par {row['doublon_reference']} participants",
            }
        )

    if not aggregated_votes.empty:
        for _, row in aggregated_votes.loc[aggregated_votes["doublon_whatsapp"] > 1].iterrows():
            rows.append(
                {
                    "type_anomalie": "WhatsApp - doublon de vote",
                    "nom_concerne": row["nom_whatsapp"],
                    "telephone_original": row["telephone_whatsapp_original"],
                    "telephone_normalise": row["telephone_whatsapp_normalise"],
                    "detail": row["anomalie_whatsapp"],
                }
            )

    if not invalid_votes.empty:
        for _, row in invalid_votes.iterrows():
            rows.append(
                {
                    "type_anomalie": "WhatsApp - téléphone invalide",
                    "nom_concerne": row["nom_whatsapp"],
                    "telephone_original": row["telephone_whatsapp_original"],
                    "telephone_normalise": row["telephone_whatsapp_normalise"],
                    "detail": row["telephone_whatsapp_detail"],
                }
            )

    return pd.DataFrame(rows, columns=ANOMALY_COLUMN_ORDER)


def prepare_export_dataframe(df: pd.DataFrame, ordered_columns: list[str]) -> pd.DataFrame:
    """Prépare un DataFrame lisible pour l'affichage et l'export."""
    export_df = df.copy() if not df.empty else pd.DataFrame(columns=ordered_columns)
    for column in ordered_columns:
        if column not in export_df.columns:
            export_df[column] = ""
    export_df = export_df.loc[:, ordered_columns]
    if "invite" in export_df.columns:
        export_df["invite"] = export_df["invite"].fillna(False).map(bool_to_label)
    export_df["categorie_manuelle"] = export_df["categorie_manuelle"].map(lambda value: clean_value(value) or "Automatique")
    return export_df.rename(columns=RESULT_COLUMN_LABELS).fillna("")


def prepare_unknown_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Prépare la table des numéros inconnus."""
    export_df = df.copy() if not df.empty else pd.DataFrame(columns=UNKNOWN_COLUMN_ORDER)
    for column in UNKNOWN_COLUMN_ORDER:
        if column not in export_df.columns:
            export_df[column] = ""
    return export_df.loc[:, UNKNOWN_COLUMN_ORDER].rename(
        columns={
            "nom_whatsapp": "Nom WhatsApp",
            "telephone_whatsapp_original": "Téléphone original WhatsApp",
            "telephone_whatsapp_normalise": "Téléphone normalisé WhatsApp",
            "reponses_brutes_detectees": "Réponses brutes détectées",
            "categorie_finale": "Catégorie finale",
            "statut_reponse": "Statut réponse",
            "commentaire_anomalie": "Commentaire anomalie",
        }
    ).fillna("")


def prepare_anomalies_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Prépare la table des anomalies."""
    export_df = df.copy() if not df.empty else pd.DataFrame(columns=ANOMALY_COLUMN_ORDER)
    for column in ANOMALY_COLUMN_ORDER:
        if column not in export_df.columns:
            export_df[column] = ""
    return export_df.loc[:, ANOMALY_COLUMN_ORDER].rename(
        columns={
            "type_anomalie": "Type anomalie",
            "nom_concerne": "Nom concerné",
            "telephone_original": "Téléphone original",
            "telephone_normalise": "Téléphone normalisé",
            "detail": "Détail",
        }
    ).fillna("")


def build_response_option_summary(aggregated_votes: pd.DataFrame, poll_columns: list[str]) -> pd.DataFrame:
    """Compte le nombre de votes par option cochée après dédoublonnage."""
    if not poll_columns:
        return pd.DataFrame(columns=["option_reponse", "nombre_reponses"])

    rows: list[dict[str, Any]] = []
    for option in poll_columns:
        count = 0
        if not aggregated_votes.empty and "reponses_detectees_liste" in aggregated_votes.columns:
            count = int(
                aggregated_votes["reponses_detectees_liste"].apply(
                    lambda selected: option in selected if isinstance(selected, list) else False
                ).sum()
            )
        rows.append(
            {
                "option_reponse": option,
                "nombre_reponses": count,
            }
        )

    return pd.DataFrame(rows)


def prepare_response_option_summary_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Prépare la synthèse des options de réponse pour l'affichage et l'export."""
    export_df = df.copy() if not df.empty else pd.DataFrame(columns=["option_reponse", "nombre_reponses"])
    for column in ["option_reponse", "nombre_reponses"]:
        if column not in export_df.columns:
            export_df[column] = ""
    if "nombre_reponses" in export_df.columns:
        export_df["nombre_reponses"] = export_df["nombre_reponses"].fillna(0).astype(int)
    return export_df.rename(
        columns={
            "option_reponse": "Option de réponse",
            "nombre_reponses": "Nombre de réponses",
        }
    )


def ensure_response_option_summary(results: dict[str, Any]) -> None:
    """Complète les anciens résultats en mémoire avec la synthèse des options."""
    if "option_summary_internal" in results and "option_summary_export" in results:
        return

    poll_columns = results.get("poll_columns", [])
    frames: list[pd.DataFrame] = []
    for key in ["all_internal", "unknown_internal"]:
        dataframe = results.get(key)
        if isinstance(dataframe, pd.DataFrame) and "reponses_detectees_liste" in dataframe.columns:
            frames.append(dataframe[["reponses_detectees_liste"]].copy())

    combined_votes = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["reponses_detectees_liste"])
    option_summary = build_response_option_summary(combined_votes, poll_columns)
    results["option_summary_internal"] = option_summary
    results["option_summary_export"] = prepare_response_option_summary_dataframe(option_summary)


def build_report_pdf(results: dict[str, Any]) -> bytes:
    """Génère un rapport PDF synthétique pour le suivi métier."""
    ensure_response_option_summary(results)
    final_df = results["all_internal"].copy()
    metrics = results["metrics"]
    option_summary_df = results["option_summary_export"].copy()

    present_df = final_df.loc[
        final_df["categorie_finale"] == "Présent",
        ["nom", "prenom", "reponses_brutes_detectees", "fonction", "invite"],
    ].copy()
    absent_df = final_df.loc[
        final_df["categorie_finale"] == "Absent",
        ["nom", "prenom", "reponses_brutes_detectees", "fonction", "invite"],
    ].copy()
    no_response_df = final_df.loc[
        final_df["categorie_finale"] == "Pas de réponse",
        ["nom", "prenom", "reponses_brutes_detectees", "fonction", "invite"],
    ].copy()

    present_df = present_df.sort_values(by=["nom", "prenom"], na_position="last").reset_index(drop=True)
    absent_df = absent_df.sort_values(by=["nom", "prenom"], na_position="last").reset_index(drop=True)
    no_response_df = no_response_df.sort_values(by=["nom", "prenom"], na_position="last").reset_index(drop=True)

    function_base = final_df.copy()
    function_base["fonction_affichee"] = function_base["fonction"].map(normalize_function_name)
    all_functions = sorted(function_base["fonction_affichee"].dropna().astype(str).unique().tolist())
    function_summary = pd.DataFrame({"fonction_affichee": all_functions})
    if function_summary.empty:
        function_summary = pd.DataFrame(columns=["fonction_affichee"])
    function_summary["Membres présents"] = function_summary["fonction_affichee"].map(
        function_base.loc[function_base["categorie_finale"] == "Présent"].groupby("fonction_affichee").size().to_dict()
    ).fillna(0).astype(int)
    function_summary["Invités présents"] = function_summary["fonction_affichee"].map(
        function_base.loc[function_base["invite"] & function_base["categorie_finale"].eq("Présent")]
        .groupby("fonction_affichee")
        .size()
        .to_dict()
    ).fillna(0).astype(int)
    function_summary["Invités absents"] = function_summary["fonction_affichee"].map(
        function_base.loc[function_base["invite"] & function_base["categorie_finale"].eq("Absent")]
        .groupby("fonction_affichee")
        .size()
        .to_dict()
    ).fillna(0).astype(int)
    function_summary["Présents (avec invités)"] = (
        function_summary["Membres présents"] + function_summary["Invités présents"]
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    story: list[Any] = []

    story.append(Paragraph("Rapport simple de suivi du sondage", styles["Title"]))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(escape(f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}"), styles["Normal"]))
    story.append(Paragraph("Source du rapport: ensemble complet des participants analysés.", styles["Normal"]))
    story.append(Spacer(1, 0.4 * cm))

    summary_data = [
        ["Indicateur", "Valeur"],
        ["Participants suivis", str(metrics["participants_total"])],
        ["Présents", str(metrics["present_count"])],
        ["Présents (avec invités)", str(metrics["present_with_invites_count"])],
        ["Absents", str(metrics["absent_count"])],
        ["Sans réponse", str(metrics["no_response_count"])],
        ["Invités présents", str(metrics["invite_present_count"])],
        ["Invités absents", str(metrics["invite_absent_count"])],
    ]
    summary_table = Table(summary_data, colWidths=[9 * cm, 4 * cm], repeatRows=1)
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0BEC5")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(summary_table)
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph("Synthèse par option de réponse", styles["Heading2"]))
    story.append(Paragraph("Comptage des votes WhatsApp valides après dédoublonnage.", styles["Normal"]))
    story.append(Spacer(1, 0.1 * cm))
    if option_summary_df.empty:
        story.append(Paragraph("Aucune option de réponse détectée.", styles["Normal"]))
    else:
        option_rows = [["Option de réponse", "Nombre de réponses"]]
        for _, row in option_summary_df.iterrows():
            option_rows.append(
                [
                    Paragraph(escape(clean_value(row["Option de réponse"]) or "-"), styles["BodyText"]),
                    str(int(row["Nombre de réponses"])),
                ]
            )
        option_table = Table(option_rows, colWidths=[12.5 * cm, 3.0 * cm], repeatRows=1)
        option_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8F5E9")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CFD8DC")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("PADDING", (0, 0), (-1, -1), 5),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        story.append(option_table)
    story.append(Spacer(1, 0.35 * cm))

    def append_table_section(
        title: str,
        total: int,
        dataframe: pd.DataFrame,
        columns: list[str],
        headers: list[str],
        col_widths: list[float] | None = None,
    ) -> None:
        story.append(Paragraph(escape(title), styles["Heading2"]))
        story.append(Paragraph(escape(f"Total: {total}"), styles["Normal"]))
        story.append(Spacer(1, 0.1 * cm))
        if dataframe.empty:
            story.append(Paragraph("Aucune ligne dans cette section.", styles["Normal"]))
            story.append(Spacer(1, 0.3 * cm))
            return
        table_rows = [headers]
        for _, row in dataframe.iterrows():
            values = []
            for column in columns:
                value = row[column]
                if column == "invite":
                    values.append(bool_to_label(value))
                else:
                    values.append(Paragraph(escape(clean_value(value) or "-"), styles["BodyText"]))
            table_rows.append(values)
        table = Table(table_rows, colWidths=col_widths, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ECEFF1")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CFD8DC")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("PADDING", (0, 0), (-1, -1), 5),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        story.append(table)
        story.append(Spacer(1, 0.35 * cm))

    append_table_section(
        "Liste des présents",
        len(present_df),
        present_df,
        ["nom", "prenom", "reponses_brutes_detectees", "fonction", "invite"],
        ["Nom", "Prénom", "Réponse sélectionnée", "Fonction", "Invité"],
        [3.0 * cm, 3.0 * cm, 6.8 * cm, 3.2 * cm, 1.5 * cm],
    )
    append_table_section(
        "Liste des absents",
        len(absent_df),
        absent_df,
        ["nom", "prenom", "reponses_brutes_detectees", "fonction", "invite"],
        ["Nom", "Prénom", "Réponse sélectionnée", "Fonction", "Invité"],
        [3.0 * cm, 3.0 * cm, 6.8 * cm, 3.2 * cm, 1.5 * cm],
    )
    append_table_section(
        "Liste des sans réponse",
        len(no_response_df),
        no_response_df,
        ["nom", "prenom", "reponses_brutes_detectees", "fonction", "invite"],
        ["Nom", "Prénom", "Réponse sélectionnée", "Fonction", "Invité"],
        [3.0 * cm, 3.0 * cm, 6.8 * cm, 3.2 * cm, 1.5 * cm],
    )

    story.append(Paragraph("Synthèse par fonction", styles["Heading2"]))
    if function_summary.empty:
        story.append(Paragraph("Aucune synthèse disponible.", styles["Normal"]))
    else:
        summary_rows = [[
            "Fonction",
            "Membres présents",
            "Invités présents",
            "Invités absents",
            "Présents (avec invités)",
        ]]
        for _, row in function_summary.iterrows():
            summary_rows.append([
                normalize_function_name(row["fonction_affichee"]),
                str(int(row.get("Membres présents", 0))),
                str(int(row.get("Invités présents", 0))),
                str(int(row.get("Invités absents", 0))),
                str(int(row.get("Présents (avec invités)", 0))),
            ])
        function_table = Table(summary_rows, colWidths=[6.2 * cm, 2.4 * cm, 2.4 * cm, 2.4 * cm, 3.0 * cm], repeatRows=1)
        function_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FFF3CD")),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CFD8DC")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("PADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        story.append(function_table)

    doc.build(story)
    return buffer.getvalue()


def build_exports(results: dict[str, Any]) -> tuple[bytes, bytes]:
    """Construit les exports Excel et CSV."""
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        results["all_export"].to_excel(writer, index=False, sheet_name="Tous")
        results["present_export"].to_excel(writer, index=False, sheet_name="Presents")
        results["absent_export"].to_excel(writer, index=False, sheet_name="Absents")
        results["other_export"].to_excel(writer, index=False, sheet_name="Autres")
        results["no_response_export"].to_excel(writer, index=False, sheet_name="Sans_reponse")
        results["unknown_export"].to_excel(writer, index=False, sheet_name="Inconnus")
        results["anomalies_export"].to_excel(writer, index=False, sheet_name="Anomalies")
        results["option_summary_export"].to_excel(writer, index=False, sheet_name="Synthese_options")

    csv_bytes = results["all_export"].to_csv(index=False).encode("utf-8-sig")
    return excel_buffer.getvalue(), csv_bytes


def build_template_excel() -> bytes:
    """Génère un modèle simple de fichier de référence."""
    template = pd.DataFrame(
        [
            {
                "Prénom": "Jean",
                "Nom": "Dupont",
                "Nom complet": "Jean Dupont",
                "Téléphone": "079 123 45 67",
                "Sexe": "Homme",
                "Fonction": "Fifre",
            }
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        template.to_excel(writer, index=False, sheet_name="Participants")
    return buffer.getvalue()


def build_member_base_source(member_base_df: pd.DataFrame) -> dict[str, Any]:
    """Construit une pseudo-source de référence à partir de la base membres locale."""
    export_df = member_base_df.copy() if not member_base_df.empty else get_empty_member_base_frame()
    for column in MEMBER_BASE_EDITOR_COLUMNS:
        if column not in export_df.columns:
            export_df[column] = ""
    export_df = export_df.loc[:, MEMBER_BASE_EDITOR_COLUMNS].fillna("")
    export_df = export_df.loc[
        export_df.apply(
            lambda row: any(clean_value(row.get(column)) for column in MEMBER_BASE_EDITOR_COLUMNS),
            axis=1,
        )
    ].reset_index(drop=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Membres")

    return {
        "name": "base_membres.xlsx",
        "bytes": buffer.getvalue(),
        "separator": "Auto",
        "origin": "Base membres locale",
    }


def filter_dataframe(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Filtre un tableau sur une recherche texte globale."""
    if df.empty or not query:
        return df
    lowered_query = query.lower()
    mask = df.astype(str).apply(lambda column: column.str.lower().str.contains(lowered_query, na=False))
    return df.loc[mask.any(axis=1)]


def show_preview(title: str, df: pd.DataFrame, metadata: dict[str, str], origin: str) -> None:
    """Affiche un aperçu simple d'un fichier importé."""
    st.markdown(f"**{title}**")
    if df.empty:
        st.warning("Le fichier est vide.")
        return
    details = [f"{len(df)} lignes", f"{len(df.columns)} colonnes", origin]
    if metadata.get("file_type"):
        details.append(metadata["file_type"])
    if metadata.get("encoding"):
        details.append(f"encodage {metadata['encoding']}")
    if metadata.get("separator"):
        details.append(f"séparateur {metadata['separator']}")
    st.caption(" | ".join(details))
    st.dataframe(df.head(10), use_container_width=True, hide_index=True)
    st.caption(f"Colonnes détectées: {', '.join(df.columns)}")


def process_files(
    reference_df: pd.DataFrame,
    votes_df: pd.DataFrame,
    reference_mapping: dict[str, str | None],
    whatsapp_mapping: dict[str, str | None],
    analysis_mode: str,
    category_mapping: dict[str, set[str]],
    duplicate_strategy: str,
    event_name: str,
) -> dict[str, Any]:
    """Exécute tout le pipeline d'analyse."""
    reference_data = build_reference_dataframe(reference_df, reference_mapping)
    reference_data = sync_reference_members(reference_data, event_name)
    poll_columns = detect_poll_columns(votes_df, whatsapp_mapping.get("name"), whatsapp_mapping.get("phone"))
    vote_rows = build_whatsapp_vote_rows(votes_df, whatsapp_mapping, poll_columns)
    aggregated_votes, invalid_votes, ignored_summary_rows = aggregate_whatsapp_votes(
        vote_rows,
        poll_columns,
        duplicate_strategy,
    )
    option_summary = build_response_option_summary(aggregated_votes, poll_columns)
    merged = merge_votes_with_reference(reference_data, aggregated_votes, analysis_mode, category_mapping)
    merged = merged.sort_values(by=["nom_complet", "prenom", "nom"], na_position="last").reset_index(drop=True)

    unknown_votes = build_unknown_votes(aggregated_votes, reference_data, analysis_mode, category_mapping)
    unknown_votes = unknown_votes.sort_values(by=["nom_whatsapp", "telephone_whatsapp_normalise"], na_position="last").reset_index(drop=True)

    anomalies = build_anomalies(reference_data, aggregated_votes, invalid_votes)
    anomalies = anomalies.sort_values(by=["type_anomalie", "nom_concerne"], na_position="last").reset_index(drop=True)

    present_df = merged.loc[merged["categorie_finale"] == "Présent"].copy()
    absent_df = merged.loc[merged["categorie_finale"] == "Absent"].copy()
    no_response_df = merged.loc[merged["categorie_finale"] == "Pas de réponse"].copy()
    other_df = merged.loc[~merged["categorie_finale"].isin(["Présent", "Absent", "Pas de réponse"])].copy()

    duplicate_reference_numbers = int(reference_data.loc[reference_data["doublon_reference"] > 1, "telephone_reference_normalise"].nunique())
    duplicate_whatsapp_numbers = int(aggregated_votes.loc[aggregated_votes["doublon_whatsapp"] > 1, "telephone_whatsapp_normalise"].nunique()) if not aggregated_votes.empty else 0
    invite_present_count = int((merged["invite"] & merged["categorie_finale"].eq("Présent")).sum())
    invite_absent_count = int((merged["invite"] & merged["categorie_finale"].eq("Absent")).sum())

    metrics = {
        "participants_total": len(merged),
        "responded_count": int((merged["statut_reponse_source"] == "A répondu").sum()),
        "present_count": int((merged["categorie_finale"] == "Présent").sum()),
        "absent_count": int((merged["categorie_finale"] == "Absent").sum()),
        "other_count": int((~merged["categorie_finale"].isin(["Présent", "Absent", "Pas de réponse"])).sum()),
        "no_response_count": int((merged["categorie_finale"] == "Pas de réponse").sum()),
        "unknown_count": len(unknown_votes),
        "duplicate_count": duplicate_reference_numbers + duplicate_whatsapp_numbers,
        "invalid_reference_count": int((~reference_data["telephone_reference_valide"]).sum()),
        "invalid_whatsapp_count": len(invalid_votes),
        "ignored_summary_rows": ignored_summary_rows,
        "invite_count": int(merged["invite"].sum()),
        "invite_present_count": invite_present_count,
        "invite_absent_count": invite_absent_count,
        "present_with_invites_count": int((merged["categorie_finale"] == "Présent").sum()) + invite_present_count,
        "manual_updates_count": int(merged["categorie_manuelle"].ne("").sum()),
    }

    results = {
        "all_internal": merged,
        "unknown_internal": unknown_votes,
        "anomalies_internal": anomalies,
        "all_export": prepare_export_dataframe(merged, RESULT_COLUMN_ORDER),
        "present_export": prepare_export_dataframe(present_df, RESULT_COLUMN_ORDER),
        "absent_export": prepare_export_dataframe(absent_df, RESULT_COLUMN_ORDER),
        "other_export": prepare_export_dataframe(other_df, RESULT_COLUMN_ORDER),
        "no_response_export": prepare_export_dataframe(no_response_df, RESULT_COLUMN_ORDER),
        "unknown_export": prepare_unknown_dataframe(unknown_votes),
        "anomalies_export": prepare_anomalies_dataframe(anomalies),
        "option_summary_internal": option_summary,
        "option_summary_export": prepare_response_option_summary_dataframe(option_summary),
        "metrics": metrics,
        "poll_columns": poll_columns,
    }
    return results


def get_column_options(columns: list[str], allow_empty: bool = True) -> list[str]:
    """Construit la liste d'options pour les widgets de mapping."""
    return ["-- Aucun --", *columns] if allow_empty else columns


def translate_selection(value: str) -> str | None:
    """Traduit l'option d'interface vers une valeur exploitable."""
    return None if value == "-- Aucun --" else value


def prepare_select_default(key: str, options: list[str], preferred: str | None) -> str:
    """Initialise proprement une valeur de selectbox dans Streamlit."""
    desired = preferred if preferred in options else options[0]
    current = st.session_state.get(key)
    if current not in options:
        st.session_state[key] = desired
    return st.session_state[key]


def prepare_multiselect_default(key: str, options: list[str], preferred: list[str]) -> list[str]:
    """Initialise proprement une valeur de multiselect dans Streamlit."""
    cleaned = [value for value in preferred if value in options]
    current = st.session_state.get(key)
    if not isinstance(current, list):
        st.session_state[key] = cleaned
    else:
        st.session_state[key] = [value for value in current if value in options]
    return st.session_state[key]


def clear_event_ui_state() -> None:
    """Nettoie les données temporaires liées à l'événement courant."""
    for key in [
        "results",
        "export_excel",
        "export_csv",
        "export_pdf",
        "current_settings",
        CURRENT_EVENT_SESSION_KEY,
        "ui_event_token",
        "wa_name",
        "wa_phone",
        "analysis_mode",
        "present_options",
        "absent_options",
        "other_options",
        "duplicate_strategy",
        "whatsapp_separator",
        "follow_up_editor",
    ]:
        st.session_state.pop(key, None)


def queue_event_selection(event_name: str) -> None:
    """Programme la sélection d'un événement au prochain rerun Streamlit."""
    st.session_state[PENDING_EVENT_SELECTION_KEY] = event_name


def build_settings_payload(
    reference_mapping: dict[str, str | None],
    whatsapp_mapping: dict[str, str | None],
    reference_separator: str,
    whatsapp_separator: str,
    analysis_mode: str,
    category_mapping: dict[str, set[str]],
    duplicate_strategy: str,
) -> dict[str, Any]:
    """Sérialise la configuration courante."""
    return {
        "reference_mapping": reference_mapping,
        "whatsapp_mapping": whatsapp_mapping,
        "reference_separator": reference_separator,
        "whatsapp_separator": whatsapp_separator,
        "analysis_mode": analysis_mode,
        "category_mapping": {
            "present": sorted(category_mapping["present"]),
            "absent": sorted(category_mapping["absent"]),
            "other": sorted(category_mapping["other"]),
        },
        "duplicate_strategy": duplicate_strategy,
    }


def refresh_analysis(
    event_name: str,
    reference_source: dict[str, Any],
    whatsapp_source: dict[str, Any],
    reference_df: pd.DataFrame,
    votes_df: pd.DataFrame,
    reference_mapping: dict[str, str | None],
    whatsapp_mapping: dict[str, str | None],
    analysis_mode: str,
    category_mapping: dict[str, set[str]],
    duplicate_strategy: str,
    reference_separator: str,
    whatsapp_separator: str,
) -> None:
    """Recalcule les résultats, met à jour les exports et sauvegarde l'état."""
    results = process_files(
        reference_df,
        votes_df,
        reference_mapping,
        whatsapp_mapping,
        analysis_mode,
        category_mapping,
        duplicate_strategy,
        event_name,
    )
    excel_bytes, csv_bytes = build_exports(results)
    pdf_bytes = build_report_pdf(results)
    settings_payload = build_settings_payload(
        reference_mapping,
        whatsapp_mapping,
        reference_separator,
        whatsapp_separator,
        analysis_mode,
        category_mapping,
        duplicate_strategy,
    )
    saved_event_name = save_session_snapshot(reference_source, whatsapp_source, settings_payload, event_name)
    st.session_state["results"] = results
    st.session_state["export_excel"] = excel_bytes
    st.session_state["export_csv"] = csv_bytes
    st.session_state["export_pdf"] = pdf_bytes
    st.session_state["current_settings"] = settings_payload
    st.session_state[CURRENT_EVENT_SESSION_KEY] = saved_event_name


def build_follow_up_editor_frame(data: pd.DataFrame) -> pd.DataFrame:
    """Prépare la table éditable de suivi manuel."""
    editor = data.copy()
    editor["Décision manuelle"] = editor["categorie_manuelle"].map(lambda value: clean_value(value) or "Automatique")
    editor["Invité"] = editor["invite"].fillna(False).astype(bool)
    editor["Commentaire suivi"] = editor["commentaire_suivi"].fillna("")
    editor = editor[
        [
            "member_key",
            "nom_complet",
            "age",
            "fonction",
            "sexe",
            "telephone_reference_original",
            "categorie_automatique",
            "Décision manuelle",
            "categorie_finale",
            "statut_reponse",
            "Invité",
            "Commentaire suivi",
        ]
    ].rename(
        columns={
            "nom_complet": "Nom complet",
            "age": "Âge",
            "fonction": "Fonction",
            "sexe": "Sexe",
            "telephone_reference_original": "Téléphone",
            "categorie_automatique": "Catégorie automatique",
            "categorie_finale": "Catégorie finale",
            "statut_reponse": "Statut réponse",
        }
    )
    return editor.set_index("member_key")


def get_configured_app_password() -> str:
    """Retourne le mot de passe configuré dans les secrets Streamlit."""
    try:
        password = st.secrets.get(APP_PASSWORD_SECRET_KEY, "")
    except StreamlitSecretNotFoundError:
        return ""
    return password.strip() if isinstance(password, str) else str(password).strip()


def require_app_password() -> None:
    """Bloque l'accès à l'application tant que le mot de passe n'est pas validé."""
    configured_password = get_configured_app_password()
    if not configured_password:
        st.title("Suivi des réponses à des sondages WhatsApp")
        st.error(
            "Mot de passe non configuré. Ajoutez `app_password` dans `.streamlit/secrets.toml` "
            "en local, ou dans `App settings > Secrets` sur Streamlit Community Cloud."
        )
        st.stop()

    if st.session_state.get(AUTHENTICATION_STATE_KEY, False):
        return

    st.title("Accès protégé")
    st.write("Saisissez le mot de passe pour ouvrir l'application.")
    with st.form("password_gate"):
        password = st.text_input("Mot de passe", type="password")
        submitted = st.form_submit_button("Ouvrir l'application", use_container_width=True)

    if submitted:
        if hmac.compare_digest(password, configured_password):
            st.session_state[AUTHENTICATION_STATE_KEY] = True
            st.rerun()
        st.error("Mot de passe incorrect.")

    st.stop()


def main() -> None:
    """Point d'entrée de l'application Streamlit."""
    st.set_page_config(page_title="Suivi des sondages WhatsApp", layout="wide")
    require_app_password()
    initialize_database()

    saved_events = list_saved_events()
    saved_event_names = [row["event_name"] for row in saved_events]
    member_summary = get_member_registry_summary()
    stored_member_base = load_member_base()

    event_options = [NEW_EVENT_OPTION, *saved_event_names]
    pending_event_selection = st.session_state.pop(PENDING_EVENT_SELECTION_KEY, None)
    if pending_event_selection in event_options:
        st.session_state["selected_event_option"] = pending_event_selection
    preferred_event = st.session_state.get("selected_event_option")
    if preferred_event not in event_options:
        preferred_event = saved_event_names[0] if saved_event_names else NEW_EVENT_OPTION
    prepare_select_default("selected_event_option", event_options, preferred_event)

    st.title("Suivi des réponses à des sondages WhatsApp")
    st.write(
        "Conservez une base membres locale, créez plusieurs événements dans SQLite et rechargez chaque vote WhatsApp "
        "depuis une liste déroulante."
    )

    with st.sidebar:
        if st.button("Verrouiller l'application", use_container_width=True):
            st.session_state[AUTHENTICATION_STATE_KEY] = False
            st.rerun()

        st.divider()
        st.header("Base membres")
        if member_summary["count"]:
            st.caption(
                f"{member_summary['count']} membre(s) en base"
                + (f" | mise à jour {member_summary['updated_at']}" if member_summary["updated_at"] else "")
            )
        else:
            st.info("Aucune base membres enregistrée pour le moment.")

        st.download_button(
            "Télécharger un modèle Excel",
            data=build_template_excel(),
            file_name="modele_reference.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        reference_file = st.file_uploader(
            "Importer une liste de membres",
            type=["xlsx", "xls", "xlsm", "csv"],
            key="member_base_upload",
            help="Cette liste alimente la base membres commune à tous les événements.",
        )

        reference_separator_options = ["Auto", ";", ",", "\t"]
        prepare_select_default("member_reference_separator", reference_separator_options, "Auto")
        reference_separator = st.selectbox(
            "Séparateur de la liste membres",
            options=reference_separator_options,
            key="member_reference_separator",
            help="Utile seulement si la liste membres est un CSV.",
        )

        st.header("Événement")
        selected_event_option = st.selectbox(
            "Événement enregistré",
            options=event_options,
            key="selected_event_option",
            format_func=lambda option: "Nouvel événement" if option == NEW_EVENT_OPTION else option,
        )

        if selected_event_option == NEW_EVENT_OPTION:
            snapshot = None
            current_event_name = clean_value(
                st.text_input(
                    "Nom du nouvel événement",
                    key="new_event_name",
                    placeholder="Ex: ATFVR",
                )
            )
        else:
            current_event_name = selected_event_option
            snapshot = load_event_snapshot(current_event_name)
            st.caption(
                f"Dernière mise à jour: {snapshot.get('updated_at', '')}"
                if snapshot
                else "Aucun instantané trouvé pour cet événement."
            )
            if snapshot:
                rename_event_name = clean_value(
                    st.text_input(
                        "Renommer l'événement",
                        value=current_event_name,
                        key=f"rename_event_name__{current_event_name}",
                    )
                )
                rename_col, delete_col = st.columns(2)
                if rename_col.button("Renommer", use_container_width=True):
                    try:
                        renamed_event_name = rename_event_session(current_event_name, rename_event_name)
                    except ValueError as exc:
                        st.error(str(exc))
                    else:
                        queue_event_selection(renamed_event_name)
                        st.session_state[CURRENT_EVENT_SESSION_KEY] = renamed_event_name
                        st.rerun()
                if delete_col.button("Supprimer cet événement", use_container_width=True):
                    delete_event_session(current_event_name)
                    clear_event_ui_state()
                    queue_event_selection(NEW_EVENT_OPTION)
                    st.rerun()

        snapshot_settings = snapshot.get("settings", {}) if snapshot else {}

        whatsapp_file = st.file_uploader(
            "Export WhatsApp de l'événement",
            type=["csv"],
            key="event_whatsapp_upload",
            help="Si rien n'est importé, l'application recharge le dernier CSV sauvegardé pour cet événement.",
        )

        whatsapp_separator_options = ["Auto", ",", ";", "\t"]
        whatsapp_separator_default = snapshot_settings.get(
            "whatsapp_separator",
            snapshot.get("whatsapp_separator", "Auto") if snapshot else "Auto",
        )
        prepare_select_default("whatsapp_separator", whatsapp_separator_options, whatsapp_separator_default)
        whatsapp_separator = st.selectbox(
            "Séparateur du CSV WhatsApp",
            options=whatsapp_separator_options,
            key="whatsapp_separator",
        )

        duplicate_options = ["Fusionner les réponses", "Garder la dernière réponse"]
        duplicate_default = snapshot_settings.get("duplicate_strategy", "Fusionner les réponses")
        prepare_select_default("duplicate_strategy", duplicate_options, duplicate_default)
        duplicate_strategy = st.radio(
            "Gestion des doublons de vote",
            options=duplicate_options,
            key="duplicate_strategy",
        )

    ui_event_token = current_event_name or NEW_EVENT_OPTION
    previous_ui_event_token = st.session_state.get("ui_event_token")
    if previous_ui_event_token != ui_event_token:
        for key in [
            "results",
            "export_excel",
            "export_csv",
            "export_pdf",
            "current_settings",
            "wa_name",
            "wa_phone",
            "analysis_mode",
            "present_options",
            "absent_options",
            "other_options",
            "duplicate_strategy",
            "whatsapp_separator",
            "follow_up_editor",
        ]:
            st.session_state.pop(key, None)
        st.session_state["ui_event_token"] = ui_event_token

    uploaded_reference_df = pd.DataFrame()
    uploaded_reference_metadata: dict[str, str] = {}
    reference_errors: list[str] = []
    if reference_file is not None:
        try:
            uploaded_reference_df, uploaded_reference_metadata = load_reference_file(
                reference_file.getvalue(),
                reference_file.name,
                reference_separator,
            )
        except Exception as exc:
            reference_errors.append(f"Base membres: {exc}")

    whatsapp_source: dict[str, Any] | None = None
    if whatsapp_file is not None:
        whatsapp_source = {
            "name": whatsapp_file.name,
            "bytes": whatsapp_file.getvalue(),
            "separator": whatsapp_separator,
            "origin": "Fichier importé",
        }
    elif snapshot and snapshot.get("whatsapp_bytes"):
        whatsapp_source = {
            "name": snapshot.get("whatsapp_name") or f"{current_event_name}.csv",
            "bytes": snapshot.get("whatsapp_bytes"),
            "separator": whatsapp_separator,
            "origin": "Événement enregistré",
        }

    votes_df = pd.DataFrame()
    votes_metadata: dict[str, str] = {}
    file_errors = list(reference_errors)
    if whatsapp_source:
        try:
            votes_df, votes_metadata = load_whatsapp_votes(
                whatsapp_source["bytes"],
                whatsapp_source["separator"],
            )
        except Exception as exc:
            file_errors.append(f"WhatsApp: {exc}")

    if file_errors:
        for message in file_errors:
            st.error(message)

    st.subheader("1. Aperçu des données")
    preview_left, preview_right = st.columns(2)
    with preview_left:
        if reference_file is not None and not uploaded_reference_df.empty:
            show_preview("Liste de membres importée", uploaded_reference_df, uploaded_reference_metadata, "Fichier importé")
        elif not stored_member_base.empty:
            show_preview("Base membres enregistrée", stored_member_base, {"file_type": "Base locale"}, "Base locale")
        else:
            st.info("Importez une liste de membres pour alimenter la base locale.")
    with preview_right:
        if whatsapp_source is None:
            st.info("Importez un CSV WhatsApp ou sélectionnez un événement déjà enregistré.")
        else:
            show_preview("Vote WhatsApp", votes_df, votes_metadata, whatsapp_source["origin"])

    saved_whatsapp_mapping = snapshot_settings.get("whatsapp_mapping", {})
    member_base_candidate = stored_member_base.copy()
    member_base_validation_errors: list[str] = []

    st.subheader("2. Base membres")
    if reference_file is None:
        st.info(
            "La base membres locale sert de référence commune à tous les événements. "
            "Vous pouvez la modifier directement ci-dessous."
        )
    elif uploaded_reference_df.empty:
        st.warning("Le fichier membres importé est vide.")
    else:
        reference_columns = list(uploaded_reference_df.columns)
        reference_suggestions = {
            "first_name": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "first_name"),
            "last_name": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "last_name"),
            "full_name": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "full_name"),
            "phone": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "phone"),
            "age": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "age"),
            "gender": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "gender"),
            "function": guess_column(reference_columns, REFERENCE_COLUMN_ALIASES, "function"),
        }
        reference_options = get_column_options(reference_columns)
        reference_phone_options = get_column_options(reference_columns, allow_empty=False)

        prepare_select_default("ref_first_name", reference_options, reference_suggestions["first_name"])
        prepare_select_default("ref_last_name", reference_options, reference_suggestions["last_name"])
        prepare_select_default("ref_full_name", reference_options, reference_suggestions["full_name"] or "-- Aucun --")
        prepare_select_default("ref_phone", reference_phone_options, reference_suggestions["phone"] or reference_phone_options[0])
        prepare_select_default("ref_age", reference_options, reference_suggestions["age"] or "-- Aucun --")
        prepare_select_default("ref_gender", reference_options, reference_suggestions["gender"] or "-- Aucun --")
        prepare_select_default("ref_function", reference_options, reference_suggestions["function"] or "-- Aucun --")

        st.write("Mappez les colonnes du fichier membres avant de l'intégrer à la base locale.")
        map_col_a, map_col_b = st.columns(2)
        with map_col_a:
            first_name_selection = st.selectbox("Colonne prénom", options=reference_options, key="ref_first_name")
            last_name_selection = st.selectbox("Colonne nom", options=reference_options, key="ref_last_name")
            full_name_selection = st.selectbox(
                "Colonne nom complet",
                options=reference_options,
                key="ref_full_name",
                help="Facultatif. Si vide, le nom complet sera reconstruit à partir du prénom et du nom.",
            )
            age_selection = st.selectbox(
                "Colonne âge / année / date de naissance",
                options=reference_options,
                key="ref_age",
                help="Facultatif. Les dates de naissance, années et âges bruts sont convertis en âge lisible quand c'est possible.",
            )
        with map_col_b:
            phone_selection = st.selectbox("Colonne téléphone", options=reference_phone_options, key="ref_phone")
            gender_selection = st.selectbox(
                "Colonne sexe",
                options=reference_options,
                key="ref_gender",
                help="Facultatif mais utile pour les rapports.",
            )
            function_selection = st.selectbox(
                "Colonne fonction / instrument",
                options=reference_options,
                key="ref_function",
                help="Exemples: Fifre, Tambour, Banneret, etc.",
            )

        reference_mapping = {
            "first_name": translate_selection(first_name_selection),
            "last_name": translate_selection(last_name_selection),
            "full_name": translate_selection(full_name_selection),
            "phone": phone_selection,
            "age": translate_selection(age_selection),
            "gender": translate_selection(gender_selection),
            "function": translate_selection(function_selection),
        }
        if not reference_mapping.get("phone"):
            member_base_validation_errors.append("Sélectionnez la colonne téléphone du fichier membres.")
        if reference_mapping.get("full_name") is None and not (
            reference_mapping.get("first_name") or reference_mapping.get("last_name")
        ):
            member_base_validation_errors.append(
                "Choisissez soit une colonne nom complet, soit au moins prénom ou nom pour la base membres."
            )

        if member_base_validation_errors:
            for message in member_base_validation_errors:
                st.error(message)
        else:
            member_base_candidate = build_member_base_editor_frame(
                build_reference_dataframe(uploaded_reference_df, reference_mapping)
            )
            st.success(f"{len(member_base_candidate)} membre(s) préparé(s) depuis le fichier importé.")

    editor_key_suffix = "saved"
    if reference_file is not None:
        editor_key_suffix = "_".join(
            [
                clean_value(reference_file.name) or "upload",
                clean_value(st.session_state.get("ref_phone")),
                clean_value(st.session_state.get("ref_full_name")),
                clean_value(st.session_state.get("ref_first_name")),
                clean_value(st.session_state.get("ref_last_name")),
                clean_value(st.session_state.get("ref_age")),
                clean_value(st.session_state.get("ref_gender")),
                clean_value(st.session_state.get("ref_function")),
            ]
        )
    member_base_editor = st.data_editor(
        member_base_candidate if not member_base_candidate.empty else get_empty_member_base_frame(),
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key=f"member_base_editor_{editor_key_suffix}",
    )

    save_member_base_disabled = bool(member_base_validation_errors)
    if st.button(
        "Enregistrer la base membres",
        use_container_width=True,
        disabled=save_member_base_disabled,
    ):
        member_stats = save_member_base_from_editor(member_base_editor)
        st.success(f"Base membres enregistrée ({member_stats['count']} membre(s)).")
        if member_stats["invalid_count"]:
            st.warning(f"{member_stats['invalid_count']} téléphone(s) invalides ont été conservés dans la base.")
        if member_stats["duplicate_count"]:
            st.warning(f"{member_stats['duplicate_count']} ligne(s) présentent un doublon de téléphone.")
        st.rerun()

    active_member_base = member_base_editor.copy() if not member_base_editor.empty else get_empty_member_base_frame()
    for column in MEMBER_BASE_EDITOR_COLUMNS:
        if column not in active_member_base.columns:
            active_member_base[column] = ""
    active_member_base = active_member_base.loc[:, MEMBER_BASE_EDITOR_COLUMNS].fillna("")
    active_member_base = active_member_base.loc[
        active_member_base.apply(
            lambda row: any(clean_value(row.get(column)) for column in MEMBER_BASE_EDITOR_COLUMNS),
            axis=1,
        )
    ].reset_index(drop=True)

    reference_df = active_member_base.copy()
    reference_mapping = MEMBER_BASE_MAPPING.copy()
    reference_source = build_member_base_source(active_member_base) if not active_member_base.empty else None

    whatsapp_columns = list(votes_df.columns)
    whatsapp_suggestions = {
        "name": guess_column(whatsapp_columns, WHATSAPP_COLUMN_ALIASES, "name"),
        "phone": guess_column(whatsapp_columns, WHATSAPP_COLUMN_ALIASES, "phone"),
    }

    st.subheader("3. Détection des colonnes du sondage WhatsApp")
    if votes_df.empty:
        st.info("Le mapping WhatsApp sera disponible après import du CSV.")
        whatsapp_mapping = {"name": None, "phone": None}
        poll_columns: list[str] = []
    else:
        whatsapp_name_options = get_column_options(whatsapp_columns)
        whatsapp_phone_options = get_column_options(whatsapp_columns, allow_empty=False)
        prepare_select_default(
            "wa_name",
            whatsapp_name_options,
            saved_whatsapp_mapping.get("name") or whatsapp_suggestions["name"] or "-- Aucun --",
        )
        prepare_select_default(
            "wa_phone",
            whatsapp_phone_options,
            saved_whatsapp_mapping.get("phone") or whatsapp_suggestions["phone"] or whatsapp_phone_options[0],
        )

        whatsapp_mapping = {
            "name": translate_selection(st.selectbox("Colonne du nom WhatsApp", options=whatsapp_name_options, key="wa_name")),
            "phone": st.selectbox("Colonne du téléphone WhatsApp", options=whatsapp_phone_options, key="wa_phone"),
        }
        poll_columns = detect_poll_columns(votes_df, whatsapp_mapping.get("name"), whatsapp_mapping.get("phone"))
        if poll_columns:
            st.success(f"{len(poll_columns)} colonnes de réponse détectées automatiquement.")
            st.write(", ".join(poll_columns))
        else:
            st.warning("Aucune colonne de réponse n'a été détectée après exclusion du nom et du téléphone.")

    st.subheader("4. Configuration métier")
    analysis_options = [ANALYSIS_MODE_FREE, ANALYSIS_MODE_CATEGORIZED]
    prepare_select_default("analysis_mode", analysis_options, snapshot_settings.get("analysis_mode", ANALYSIS_MODE_CATEGORIZED))
    analysis_mode = st.radio("Mode d'analyse", options=analysis_options, horizontal=True, key="analysis_mode")

    saved_category_mapping = snapshot_settings.get("category_mapping", {"present": [], "absent": [], "other": []})
    present_options: list[str] = []
    absent_options: list[str] = []
    other_options: list[str] = []
    if not poll_columns:
        st.info("Les groupes métier seront disponibles dès qu'au moins une colonne de réponse est détectée.")
    elif analysis_mode == ANALYSIS_MODE_CATEGORIZED:
        prepare_multiselect_default("present_options", poll_columns, saved_category_mapping.get("present", []))
        prepare_multiselect_default("absent_options", poll_columns, saved_category_mapping.get("absent", []))
        prepare_multiselect_default("other_options", poll_columns, saved_category_mapping.get("other", []))
        present_options = st.multiselect("Réponses à classer comme présent", options=poll_columns, key="present_options")
        absent_options = st.multiselect("Réponses à classer comme absent", options=poll_columns, key="absent_options")
        other_options = st.multiselect("Réponses à classer comme autre", options=poll_columns, key="other_options")
        st.info("Les réponses non mappées iront dans la catégorie Autre.")
    else:
        st.info("En analyse libre, l'application conserve les réponses brutes sans interprétation métier.")

    category_mapping = {
        "present": set(present_options),
        "absent": set(absent_options),
        "other": set(other_options),
    }

    can_process = bool(
        current_event_name
        and reference_source
        and whatsapp_source
        and not reference_df.empty
        and not votes_df.empty
        and not file_errors
        and not member_base_validation_errors
    )
    process_button = st.sidebar.button(
        "Traiter l'événement",
        type="primary",
        use_container_width=True,
        disabled=not can_process,
    )

    validation_errors: list[str] = []
    if not current_event_name:
        validation_errors.append("Saisissez un nom d'événement.")
    if reference_df.empty:
        validation_errors.append("La base membres est vide. Importez-la ou complétez-la avant le traitement.")
    if whatsapp_source is None or votes_df.empty:
        validation_errors.append("Importez un CSV WhatsApp pour l'événement sélectionné.")
    if votes_df.empty and whatsapp_source is not None:
        validation_errors.append("Le CSV WhatsApp chargé est vide.")
    if not whatsapp_mapping.get("phone"):
        validation_errors.append("Sélectionnez la colonne téléphone du fichier WhatsApp.")
    if whatsapp_source and not poll_columns:
        validation_errors.append("Aucune colonne de réponse WhatsApp n'a été détectée.")

    if not can_process:
        if not current_event_name:
            st.warning("Renseignez le nom du nouvel événement pour pouvoir le sauvegarder.")
        elif reference_df.empty:
            st.warning("Constituez d'abord la base membres commune avant d'analyser un événement.")
        elif whatsapp_source is None:
            st.warning("Importez un CSV WhatsApp ou choisissez un événement déjà enregistré.")

    auto_process = bool(snapshot and can_process and "results" not in st.session_state)
    if process_button or auto_process:
        if validation_errors:
            for message in validation_errors:
                st.error(message)
        else:
            with st.spinner("Traitement de l'événement en cours..."):
                save_member_base_from_editor(active_member_base)
                refresh_analysis(
                    current_event_name,
                    reference_source,
                    whatsapp_source,
                    reference_df,
                    votes_df,
                    reference_mapping,
                    whatsapp_mapping,
                    analysis_mode,
                    category_mapping,
                    duplicate_strategy,
                    reference_separator,
                    whatsapp_separator,
                )
            if process_button:
                st.success(f"Événement « {current_event_name} » analysé et sauvegardé.")
            elif auto_process:
                st.info(f"L'événement « {current_event_name} » a été rechargé automatiquement depuis SQLite.")

    if "results" not in st.session_state:
        st.info("Configurez la base membres et le vote WhatsApp puis lancez le traitement pour afficher les résultats.")
        return

    results = st.session_state["results"]
    had_option_summary = "option_summary_internal" in results and "option_summary_export" in results
    ensure_response_option_summary(results)
    if not had_option_summary:
        export_excel, export_csv = build_exports(results)
        st.session_state["export_excel"] = export_excel
        st.session_state["export_csv"] = export_csv
    metrics = results["metrics"]
    current_pdf_bytes = build_report_pdf(results)
    st.session_state["export_pdf"] = current_pdf_bytes

    with st.sidebar:
        st.header("Exports")
        st.caption("Le traitement et les modifications manuelles sont sauvegardés dans SQLite.")
        st.download_button(
            "Exporter le résultat en Excel",
            data=st.session_state["export_excel"],
            file_name="suivi_sondage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Exporter le résultat en CSV",
            data=st.session_state["export_csv"],
            file_name="suivi_sondage.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.download_button(
            "Télécharger le rapport PDF",
            data=current_pdf_bytes,
            file_name="rapport_sondage.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    if metrics["ignored_summary_rows"]:
        st.info(f"{metrics['ignored_summary_rows']} ligne(s) récapitulative(s) de type total ont été ignorées.")
    if metrics["invalid_reference_count"]:
        st.warning(f"{metrics['invalid_reference_count']} téléphone(s) invalides dans le fichier de référence.")
    if metrics["invalid_whatsapp_count"]:
        st.warning(f"{metrics['invalid_whatsapp_count']} ligne(s) WhatsApp avec téléphone invalide ont été isolées dans les anomalies.")
    if metrics["manual_updates_count"]:
        st.success(f"{metrics['manual_updates_count']} participant(s) possèdent déjà une mise à jour manuelle enregistrée.")

    st.subheader("5. Résultats globaux")
    metric_columns = st.columns(4)
    metric_columns[0].metric("Participants", metrics["participants_total"])
    metric_columns[1].metric("A répondu", metrics["responded_count"])
    metric_columns[2].metric("Présents", metrics["present_count"])
    metric_columns[3].metric("Absents", metrics["absent_count"])

    metric_columns = st.columns(4)
    metric_columns[0].metric("Autres réponses", metrics["other_count"])
    metric_columns[1].metric("Sans réponse", metrics["no_response_count"])
    metric_columns[2].metric("Numéros inconnus", metrics["unknown_count"])
    metric_columns[3].metric("Doublons détectés", metrics["duplicate_count"])

    metric_columns = st.columns(4)
    metric_columns[0].metric("Invités cochés", metrics["invite_count"])
    metric_columns[1].metric("Mises à jour manuelles", metrics["manual_updates_count"])
    metric_columns[2].metric("Téléphones invalides réf.", metrics["invalid_reference_count"])
    metric_columns[3].metric("Téléphones invalides WhatsApp", metrics["invalid_whatsapp_count"])

    st.subheader("6. Synthèse des options de réponse")
    st.caption("Comptage des votes WhatsApp valides après dédoublonnage.")
    st.dataframe(results["option_summary_export"], use_container_width=True, hide_index=True)

    st.subheader("7. Suivi manuel et relance")
    st.write(
        "Utilisez ce tableau pour marquer une décision manuelle, cocher les invités et garder une note de suivi. "
        "Les modifications sont sauvegardées en base locale et réappliquées automatiquement au prochain lancement."
    )
    follow_up_filter_options = [
        "Tous",
        "Sans réponse",
        "Présents",
        "Absents",
        "Autres réponses",
    ]
    prepare_select_default("follow_up_filter", follow_up_filter_options, "Sans réponse")
    follow_up_filter = st.selectbox("Filtre de suivi", options=follow_up_filter_options, key="follow_up_filter")
    follow_up_search = st.text_input("Recherche dans le suivi manuel", placeholder="Nom, fonction, téléphone...")

    follow_up_df = results["all_internal"].copy()
    if follow_up_filter == "Sans réponse":
        follow_up_df = follow_up_df.loc[follow_up_df["categorie_finale"] == "Pas de réponse"]
    elif follow_up_filter == "Présents":
        follow_up_df = follow_up_df.loc[follow_up_df["categorie_finale"] == "Présent"]
    elif follow_up_filter == "Absents":
        follow_up_df = follow_up_df.loc[follow_up_df["categorie_finale"] == "Absent"]
    elif follow_up_filter == "Autres réponses":
        follow_up_df = follow_up_df.loc[~follow_up_df["categorie_finale"].isin(["Présent", "Absent", "Pas de réponse"])]

    if follow_up_search:
        follow_up_df = filter_dataframe(
            follow_up_df[
                [
                    "member_key",
                    "nom_complet",
                    "fonction",
                    "sexe",
                    "telephone_reference_original",
                    "categorie_automatique",
                    "categorie_finale",
                    "statut_reponse",
                    "commentaire_suivi",
                ]
            ].assign(invite=follow_up_df["invite"].map(bool_to_label)),
            follow_up_search,
        )
        follow_up_df = results["all_internal"].loc[results["all_internal"]["member_key"].isin(follow_up_df["member_key"])]

    editor_df = build_follow_up_editor_frame(follow_up_df)
    edited_follow_up = st.data_editor(
        editor_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Décision manuelle": st.column_config.SelectboxColumn(
                "Décision manuelle",
                options=MANUAL_CATEGORY_OPTIONS,
                help="Choisissez une décision pour remplacer la catégorie automatique.",
            ),
            "Invité": st.column_config.CheckboxColumn(
                "Invité",
                help="Cochez si cette personne doit être comptée dans le suivi des invités.",
            ),
            "Commentaire suivi": st.column_config.TextColumn("Commentaire suivi"),
        },
        disabled=[
            "Nom complet",
            "Fonction",
            "Sexe",
            "Téléphone",
            "Catégorie automatique",
            "Catégorie finale",
            "Statut réponse",
        ],
        key="follow_up_editor",
    )
    if st.button("Enregistrer les modifications du suivi", use_container_width=True):
        updates = edited_follow_up.reset_index()[["member_key", "Décision manuelle", "Invité", "Commentaire suivi"]]
        save_member_updates(current_event_name, updates)
        refresh_analysis(
            current_event_name,
            reference_source,
            whatsapp_source,
            reference_df,
            votes_df,
            reference_mapping,
            whatsapp_mapping,
            analysis_mode,
            category_mapping,
            duplicate_strategy,
            reference_separator,
            whatsapp_separator,
        )
        st.success("Modifications enregistrées en base locale.")
        st.rerun()

    st.subheader("8. Tableaux détaillés")
    search_query = st.text_input("Recherche texte dans les tableaux de résultat", placeholder="Nom, téléphone, réponse...")
    tabs = st.tabs(
        [
            "Tous les participants",
            "Présents",
            "Absents",
            "Autres réponses",
            "Sans réponse",
            "Numéros inconnus",
            "Doublons / anomalies",
        ]
    )

    with tabs[0]:
        st.dataframe(filter_dataframe(results["all_export"], search_query), use_container_width=True, hide_index=True)
    with tabs[1]:
        st.dataframe(filter_dataframe(results["present_export"], search_query), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(filter_dataframe(results["absent_export"], search_query), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(filter_dataframe(results["other_export"], search_query), use_container_width=True, hide_index=True)
    with tabs[4]:
        st.dataframe(filter_dataframe(results["no_response_export"], search_query), use_container_width=True, hide_index=True)
    with tabs[5]:
        st.dataframe(filter_dataframe(results["unknown_export"], search_query), use_container_width=True, hide_index=True)
    with tabs[6]:
        st.dataframe(filter_dataframe(results["anomalies_export"], search_query), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()

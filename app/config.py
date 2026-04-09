from __future__ import annotations

import os
from dataclasses import dataclass
from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    telegram_bot_token: str
    google_service_account_json: str
    google_sheet_id: str
    google_drive_output_folder_id: str
    template_docx_path: str
    timezone: str = "Asia/Bangkok"
    admin_telegram_ids: tuple[int, ...] = ()


def _parse_admin_ids(raw: str | None) -> tuple[int, ...]:
    if not raw:
        return ()
    values = []
    for item in raw.split(","):
        item = item.strip()
        if item:
            values.append(int(item))
    return tuple(values)


settings = Settings(
    telegram_bot_token=os.getenv("TELEGRAM_BOT_TOKEN", ""),
    google_service_account_json=os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json"),
    google_sheet_id=os.getenv("GOOGLE_SHEET_ID", ""),
    google_drive_output_folder_id=os.getenv("GOOGLE_DRIVE_OUTPUT_FOLDER_ID", ""),
    template_docx_path=os.getenv("TEMPLATE_DOCX_PATH", "templates/site_letter_template.docx"),
    timezone=os.getenv("TIMEZONE", "Asia/Bangkok"),
    admin_telegram_ids=_parse_admin_ids(os.getenv("ADMIN_TELEGRAM_IDS")),
)


def validate_settings() -> None:
    missing = []
    for key in [
        "telegram_bot_token",
        "google_service_account_json",
        "google_sheet_id",
        "google_drive_output_folder_id",
        "template_docx_path",
    ]:
        if not getattr(settings, key):
            missing.append(key)
    if missing:
        raise RuntimeError(f"Missing required settings: {', '.join(missing)}")

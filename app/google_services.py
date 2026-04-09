from __future__ import annotations

import json
import os
from dataclasses import dataclass
from typing import Any

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


@dataclass
class GoogleClients:
    gc: gspread.Client
    spreadsheet: gspread.Spreadsheet
    drive: Any


def build_clients(service_account_json: str, sheet_id: str) -> GoogleClients:
    creds = Credentials.from_service_account_file(service_account_json, scopes=SCOPES)
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(sheet_id)
    drive = build("drive", "v3", credentials=creds)
    return GoogleClients(gc=gc, spreadsheet=spreadsheet, drive=drive)


class SheetRepository:
    def __init__(self, spreadsheet: gspread.Spreadsheet):
        self.spreadsheet = spreadsheet
        self.requests_log = spreadsheet.worksheet("requests_log")

    def append_request(self, payload: dict[str, Any]) -> None:
        headers = self.requests_log.row_values(1)
        row = [payload.get(h, "") for h in headers]
        self.requests_log.append_row(row, value_input_option="USER_ENTERED")

    def update_request_by_id(self, request_id: str, updates: dict[str, Any]) -> None:
        records = self.requests_log.get_all_records()
        headers = self.requests_log.row_values(1)
        target_row_index = None
        for idx, row in enumerate(records, start=2):
            if row.get("request_id") == request_id:
                target_row_index = idx
                break
        if not target_row_index:
            raise ValueError(f"request_id not found: {request_id}")
        for key, value in updates.items():
            if key not in headers:
                continue
            col = headers.index(key) + 1
            self.requests_log.update_cell(target_row_index, col, value)

    def list_recent_by_user(self, telegram_user_id: int, limit: int = 5) -> list[dict[str, Any]]:
        records = self.requests_log.get_all_records()
        filtered = [r for r in records if str(r.get("telegram_user_id", "")) == str(telegram_user_id)]
        return filtered[-limit:][::-1]


class DriveRepository:
    def __init__(self, drive_service: Any, output_folder_id: str):
        self.drive = drive_service
        self.output_folder_id = output_folder_id

    def upload_file(self, local_path: str, mime_type: str) -> dict[str, str]:
        file_metadata = {
            "name": os.path.basename(local_path),
            "parents": [self.output_folder_id],
        }
        media = MediaFileUpload(local_path, mimetype=mime_type, resumable=False)
        result = (
            self.drive.files()
            .create(body=file_metadata, media_body=media, fields="id, webViewLink")
            .execute()
        )
        return {"file_id": result["id"], "web_link": result.get("webViewLink", "")}


def staff_to_json(staff_items: list[dict[str, str]]) -> str:
    return json.dumps(staff_items, ensure_ascii=False)

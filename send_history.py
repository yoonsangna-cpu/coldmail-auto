"""
발송 이력 관리 모듈 (Google Sheets 저장)
- 각 사용자의 Google Drive에 "콜드메일_발송이력" 스프레드시트를 자동 생성
- 성공적으로 발송된 이메일 주소를 기록
- 다음 발송 시 이미 보낸 이메일을 자동 건너뛰기
"""

from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SPREADSHEET_TITLE = "콜드메일_발송이력"
SHEET_NAME = "발송이력"


# ─────────────────────────────────────────────────────────
# 내부 헬퍼
# ─────────────────────────────────────────────────────────

def _get_sheets_service(credentials: Credentials):
    """Google Sheets API 서비스 객체를 생성한다."""
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def _get_drive_service(credentials: Credentials):
    """Google Drive API 서비스 객체를 생성한다."""
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def _find_spreadsheet(credentials: Credentials) -> str | None:
    """
    사용자의 Drive에서 기존 발송이력 스프레드시트를 찾는다.

    Returns:
        spreadsheet_id 또는 None
    """
    try:
        drive = _get_drive_service(credentials)
        query = (
            f"name = '{SPREADSHEET_TITLE}' "
            f"and mimeType = 'application/vnd.google-apps.spreadsheet' "
            f"and trashed = false"
        )
        result = drive.files().list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            pageSize=1,
        ).execute()
        files = result.get("files", [])
        if files:
            return files[0]["id"]
        return None
    except HttpError:
        return None


def _create_spreadsheet(credentials: Credentials) -> str:
    """
    발송이력 스프레드시트를 새로 생성한다.

    Returns:
        새 spreadsheet_id
    """
    sheets = _get_sheets_service(credentials)
    body = {
        "properties": {"title": SPREADSHEET_TITLE},
        "sheets": [{
            "properties": {"title": SHEET_NAME},
        }],
    }
    spreadsheet = sheets.spreadsheets().create(body=body).execute()
    spreadsheet_id = spreadsheet["spreadsheetId"]

    # 헤더 행 추가
    sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{SHEET_NAME}!A1:B1",
        valueInputOption="RAW",
        body={"values": [["이메일", "발송시각"]]},
    ).execute()

    return spreadsheet_id


def _get_or_create_spreadsheet(credentials: Credentials) -> str:
    """기존 스프레드시트를 찾거나, 없으면 새로 생성한다."""
    spreadsheet_id = _find_spreadsheet(credentials)
    if spreadsheet_id:
        return spreadsheet_id
    return _create_spreadsheet(credentials)


# ─────────────────────────────────────────────────────────
# 공개 API — credentials를 인자로 받음
# ─────────────────────────────────────────────────────────

def get_sent_emails(credentials: Credentials) -> set[str]:
    """
    발송된 이메일 주소 집합을 반환한다.

    Args:
        credentials: Google OAuth Credentials 객체
    """
    try:
        spreadsheet_id = _get_or_create_spreadsheet(credentials)
        sheets = _get_sheets_service(credentials)
        result = sheets.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_NAME}!A:A",
        ).execute()
        values = result.get("values", [])
        # 첫 행(헤더)은 제외, 소문자로 통일
        emails = set()
        for row in values[1:]:
            if row:
                emails.add(row[0].strip().lower())
        return emails
    except (HttpError, Exception):
        return set()


def get_sent_count(credentials: Credentials) -> int:
    """총 발송 이력 건수를 반환한다."""
    return len(get_sent_emails(credentials))


def add_sent_emails_batch(credentials: Credentials, emails: list[str]) -> None:
    """
    여러 이메일을 한꺼번에 이력에 추가한다.

    Args:
        credentials: Google OAuth Credentials 객체
        emails: 성공 발송된 이메일 주소 리스트
    """
    if not emails:
        return
    try:
        spreadsheet_id = _get_or_create_spreadsheet(credentials)
        sheets = _get_sheets_service(credentials)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = [[email.strip().lower(), now] for email in emails]
        sheets.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_NAME}!A:B",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": rows},
        ).execute()
    except (HttpError, Exception):
        pass


def add_sent_email(credentials: Credentials, email: str) -> None:
    """성공 발송된 이메일 1건을 이력에 추가한다."""
    add_sent_emails_batch(credentials, [email])


def clear_history(credentials: Credentials) -> None:
    """
    발송 이력을 초기화한다 (시트 내용 삭제 후 헤더만 남김).
    """
    try:
        spreadsheet_id = _find_spreadsheet(credentials)
        if not spreadsheet_id:
            return
        sheets = _get_sheets_service(credentials)

        # 시트 ID 가져오기
        meta = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == SHEET_NAME:
                sheet_id = sheet["properties"]["sheetId"]
                break

        if sheet_id is None:
            return

        # 2행부터 전부 삭제 (헤더 유지)
        total_rows = 0
        result = sheets.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_NAME}!A:A",
        ).execute()
        total_rows = len(result.get("values", []))

        if total_rows <= 1:
            return  # 헤더밖에 없음

        requests = [{
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": total_rows,
                }
            }
        }]
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ).execute()
    except (HttpError, Exception):
        pass

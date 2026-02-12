"""
발송 이력 관리 모듈 (Google Sheets 저장)
- 각 사용자의 Google Drive에 "콜드메일_발송이력" 스프레드시트를 자동 생성
- 성공적으로 발송된 이메일 주소를 기록
- 다음 발송 시 이미 보낸 이메일을 자동 건너뛰기
"""

from __future__ import annotations

import logging
from datetime import datetime, date
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

log = logging.getLogger(__name__)

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
# API 접근성 검증
# ─────────────────────────────────────────────────────────

def verify_sheets_access(credentials: Credentials) -> tuple[bool, str]:
    """
    Google Sheets / Drive API 접근 가능 여부를 확인한다.
    로그인 직후 한 번 호출하여 API 활성화 상태를 진단.

    Returns:
        (성공 여부, 메시지)
    """
    # 1. Drive API 확인
    try:
        drive = _get_drive_service(credentials)
        drive.files().list(pageSize=1, fields="files(id)").execute()
    except HttpError as e:
        if e.resp.status == 403:
            return False, "Google Drive API가 활성화되지 않았습니다. Google Cloud Console → API 라이브러리에서 'Google Drive API'를 사용 설정해주세요."
        return False, f"Drive API 오류: {e.reason if hasattr(e, 'reason') else e}"
    except Exception as e:
        return False, f"Drive API 접근 실패: {e}"

    # 2. Sheets API 확인
    try:
        sheets = _get_sheets_service(credentials)
        # 기존 스프레드시트가 있으면 읽기, 없으면 생성 시도
        spreadsheet_id = _find_spreadsheet(credentials)
        if spreadsheet_id:
            sheets.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{SHEET_NAME}!A1:A1",
            ).execute()
        else:
            # 새 스프레드시트 생성 시도
            _create_spreadsheet(credentials)
    except HttpError as e:
        if e.resp.status == 403:
            return False, "Google Sheets API가 활성화되지 않았습니다. Google Cloud Console → API 라이브러리에서 'Google Sheets API'를 사용 설정해주세요."
        return False, f"Sheets API 오류: {e.reason if hasattr(e, 'reason') else e}"
    except Exception as e:
        return False, f"Sheets API 접근 실패: {e}"

    return True, "정상"


# ─────────────────────────────────────────────────────────
# 공개 API — credentials를 인자로 받음
# ─────────────────────────────────────────────────────────

def get_sent_emails(credentials: Credentials) -> set[str]:
    """
    발송된 이메일 주소 집합을 반환한다.

    Args:
        credentials: Google OAuth Credentials 객체

    Raises:
        HttpError: API 호출 실패 시
    """
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


def get_sent_count(credentials: Credentials) -> int:
    """총 발송 이력 건수를 반환한다."""
    return len(get_sent_emails(credentials))


def get_today_sent_count(credentials: Credentials) -> int:
    """오늘 발송한 건수를 Google Sheets에서 계산하여 반환한다."""
    try:
        spreadsheet_id = _get_or_create_spreadsheet(credentials)
        sheets = _get_sheets_service(credentials)
        result = sheets.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{SHEET_NAME}!A:B",
        ).execute()
        values = result.get("values", [])
        today_str = date.today().isoformat()  # "YYYY-MM-DD"
        count = 0
        for row in values[1:]:  # 헤더 제외
            if len(row) >= 2 and row[1].startswith(today_str):
                count += 1
        return count
    except Exception:
        return 0


def add_sent_emails_batch(credentials: Credentials, emails: list[str]) -> tuple[bool, str]:
    """
    여러 이메일을 한꺼번에 이력에 추가한다.

    Args:
        credentials: Google OAuth Credentials 객체
        emails: 성공 발송된 이메일 주소 리스트

    Returns:
        (성공 여부, 메시지)
    """
    if not emails:
        return True, ""
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
        return True, f"{len(emails)}건 이력 저장 완료"
    except HttpError as e:
        msg = f"이력 저장 실패 (Sheets API {e.resp.status}): {e.reason if hasattr(e, 'reason') else e}"
        log.error(msg)
        return False, msg
    except Exception as e:
        msg = f"이력 저장 실패: {e}"
        log.error(msg)
        return False, msg


def add_sent_email(credentials: Credentials, email: str) -> tuple[bool, str]:
    """성공 발송된 이메일 1건을 이력에 추가한다."""
    return add_sent_emails_batch(credentials, [email])


def clear_history(credentials: Credentials) -> None:
    """
    발송 이력을 초기화한다 (시트 내용 삭제 후 헤더만 남김).
    """
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

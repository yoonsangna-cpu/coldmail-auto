"""
Google OAuth2 인증 + Gmail API 모듈
- Google 계정으로 로그인 (OAuth2)
- Gmail API를 통한 메일 발송
- 토큰 관리 (직렬화/역직렬화)
- redirect_uri 자동 감지
"""

from __future__ import annotations

import os
import base64
import mimetypes
import streamlit as st
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import formataddr

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# 로컬 개발 시 HTTP 허용 (프로덕션에서는 자동으로 HTTPS 사용)
if os.environ.get("STREAMLIT_SERVER_ADDRESS", "localhost") == "localhost":
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

# OAuth 스코프: 사용자 정보 + Gmail 발송 + 설정 읽기(서명) + Sheets + Drive(앱 파일만)
SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]


# ─────────────────────────────────────────────────────────
# 앱 URL 자동 감지 + OAuth 설정 통합 관리
# ─────────────────────────────────────────────────────────

def detect_app_url() -> str:
    """
    현재 앱의 URL을 자동 감지한다.
    Streamlit Cloud, 로컬 등 환경에 따라 적절한 URL을 반환한다.
    항상 trailing slash를 제거하여 redirect_uri_mismatch를 방지한다.
    """
    try:
        headers = st.context.headers
        host = headers.get("Host", "localhost:8501")
        proto = headers.get("X-Forwarded-Proto", None)
        if proto:
            scheme = proto
        elif host.startswith("localhost") or host.startswith("127.0.0.1"):
            scheme = "http"
        else:
            scheme = "https"
        return f"{scheme}://{host}".rstrip("/")
    except Exception:
        return "http://localhost:8501"


def _get_oauth_config() -> dict | None:
    """
    OAuth 설정을 가져온다.
    1순위: 세션에 사용자가 직접 입력한 설정
    2순위: secrets.toml (또는 Streamlit Cloud Secrets)
    둘 다 없으면 None 반환.
    """
    # 1. 세션에 사용자 입력 설정이 있는지 확인
    if hasattr(st, "session_state") and st.session_state.get("user_oauth_config"):
        return st.session_state.user_oauth_config

    # 2. st.secrets에서 가져오기
    try:
        return {
            "client_id": st.secrets["google"]["client_id"],
            "client_secret": st.secrets["google"]["client_secret"],
            "redirect_uri": st.secrets["google"]["redirect_uri"].rstrip("/"),
        }
    except Exception:
        return None


def _get_redirect_uri() -> str:
    """
    현재 앱의 redirect_uri를 결정한다.
    OAuth 설정에 값이 있으면 사용하고, 없으면 앱 URL을 자동 감지한다.
    항상 trailing slash를 제거하여 redirect_uri_mismatch를 방지한다.
    """
    config = _get_oauth_config()
    if config and config.get("redirect_uri"):
        return config["redirect_uri"].rstrip("/")
    return detect_app_url()


# ─────────────────────────────────────────────────────────
# OAuth2 흐름
# ─────────────────────────────────────────────────────────

def _get_client_config() -> dict:
    """OAuth 설정을 읽어 client config를 생성한다. (세션 또는 secrets.toml)"""
    config = _get_oauth_config()
    if not config:
        raise ValueError("OAuth 설정이 없습니다. client_id와 client_secret을 입력해주세요.")
    redirect_uri = config["redirect_uri"]
    return {
        "web": {
            "client_id": config["client_id"],
            "client_secret": config["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [redirect_uri],
        }
    }


def _create_flow() -> Flow:
    """OAuth2 Flow 객체를 생성한다."""
    redirect_uri = _get_redirect_uri()
    return Flow.from_client_config(
        _get_client_config(),
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )


def get_authorization_url() -> tuple[str, str]:
    """
    Google OAuth 인증 URL을 생성한다.

    Returns:
        (auth_url, state) - 인증 URL과 CSRF 방지용 state 값
    """
    flow = _create_flow()
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return auth_url, state


def exchange_code_for_credentials(code: str) -> Credentials:
    """
    Google이 반환한 authorization code를 credential로 교환한다.

    Args:
        code: Google OAuth 콜백에서 받은 authorization code

    Returns:
        Google OAuth Credentials 객체
    """
    flow = _create_flow()
    flow.fetch_token(code=code)
    return flow.credentials


# ─────────────────────────────────────────────────────────
# 사용자 정보
# ─────────────────────────────────────────────────────────

def get_user_info(credentials: Credentials) -> dict:
    """
    Google 사용자 정보를 가져온다.

    Returns:
        {"email": "...", "name": "...", "picture": "..."}
    """
    service = build("oauth2", "v2", credentials=credentials)
    user_info = service.userinfo().get().execute()
    return user_info


# ─────────────────────────────────────────────────────────
# 토큰 직렬화/역직렬화 (세션 저장용)
# ─────────────────────────────────────────────────────────

def credentials_to_dict(credentials: Credentials) -> dict:
    """Credentials 객체를 딕셔너리로 변환 (세션 저장용)."""
    return {
        "token": credentials.token,
        "refresh_token": credentials.refresh_token,
        "token_uri": credentials.token_uri,
        "client_id": credentials.client_id,
        "client_secret": credentials.client_secret,
        "scopes": list(credentials.scopes) if credentials.scopes else list(SCOPES),
    }


def credentials_from_dict(cred_dict: dict) -> Credentials:
    """딕셔너리에서 Credentials 객체를 복원한다."""
    return Credentials(**cred_dict)


# ─────────────────────────────────────────────────────────
# Gmail API 메일 발송
# ─────────────────────────────────────────────────────────

def get_gmail_service(credentials: Credentials):
    """Gmail API 서비스 객체를 생성한다."""
    return build("gmail", "v1", credentials=credentials)


def get_gmail_signature(service, email: str) -> str:
    """
    Gmail에 설정된 서명(HTML)을 가져온다.

    Args:
        service: Gmail API 서비스 객체
        email: 사용자 Gmail 주소

    Returns:
        서명 HTML 문자열 (없으면 빈 문자열)
    """
    try:
        result = service.users().settings().sendAs().get(
            userId="me",
            sendAsEmail=email,
        ).execute()
        return result.get("signature", "")
    except HttpError:
        return ""
    except Exception:
        return ""


def send_email(
    service,
    to_email: str,
    subject: str,
    body: str,
    from_email: str = "",
    from_name: str = "",
    signature_html: str = "",
    attachments: list = None,
) -> tuple[bool, str]:
    """
    Gmail API를 통해 이메일을 발송한다.

    Args:
        service: Gmail API 서비스 객체
        to_email: 수신자 이메일
        subject: 메일 제목
        body: 메일 본문 (plain text)
        from_email: 발신자 이메일
        from_name: 발신자 이름
        signature_html: Gmail 서명 HTML (비어있으면 서명 없이 발송)
        attachments: 첨부 파일 리스트 (Streamlit UploadedFile 객체)

    Returns:
        (성공 여부, 메시지)
    """
    try:
        # 첨부파일이 있으면 mixed, 없으면 alternative
        if attachments:
            msg = MIMEMultipart("mixed")
            msg_body = MIMEMultipart("alternative")
        else:
            msg = MIMEMultipart("alternative")
            msg_body = msg

        msg["To"] = to_email
        msg["Subject"] = subject

        if from_name and from_email:
            # 한글 등 non-ASCII 이름을 RFC 2047로 인코딩
            msg["From"] = formataddr((from_name, from_email))
        elif from_email:
            msg["From"] = from_email

        # HTML 본문 생성 (줄바꿈 → <br>)
        html_body = body.replace("\n", "<br>")

        # 서명이 있으면 본문 뒤에 추가
        if signature_html:
            html_body = (
                f"{html_body}"
                f'<br><br>'
                f'<div style="border-top: 1px solid #ccc; padding-top: 8px; margin-top: 8px;">'
                f'{signature_html}'
                f'</div>'
            )

        # Plain text 버전
        msg_body.attach(MIMEText(body, "plain", "utf-8"))
        # HTML 버전 (서명 포함)
        msg_body.attach(MIMEText(html_body, "html", "utf-8"))

        # 첨부파일이 있으면 본문을 mixed에 넣고 파일 추가
        if attachments:
            msg.attach(msg_body)
            for file in attachments:
                file.seek(0)
                file_data = file.read()
                file_name = file.name

                mime_type, _ = mimetypes.guess_type(file_name)
                if mime_type is None:
                    mime_type = "application/octet-stream"
                main_type, sub_type = mime_type.split("/", 1)

                attachment = MIMEBase(main_type, sub_type)
                attachment.set_payload(file_data)
                encoders.encode_base64(attachment)
                attachment.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=file_name,
                )
                msg.attach(attachment)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")

        result = service.users().messages().send(
            userId="me",
            body={"raw": raw},
        ).execute()

        return True, f"발송 성공 (ID: {result.get('id', '')})"

    except HttpError as e:
        error_reason = e.reason if hasattr(e, "reason") else str(e)
        if e.resp.status == 403:
            return False, f"권한 오류: Gmail 발송 권한이 없습니다. ({error_reason})"
        elif e.resp.status == 429:
            return False, "발송 한도 초과: 잠시 후 다시 시도해주세요."
        else:
            return False, f"Gmail API 오류 ({e.resp.status}): {error_reason}"
    except Exception as e:
        return False, f"발송 오류: {e}"


def check_secrets_configured() -> tuple[bool, str]:
    """
    OAuth 설정이 있는지 확인한다.
    세션에 사용자 입력 값이 있거나 secrets.toml에 설정이 있으면 True.
    placeholder 값이 아닌 실제 값이 있는지도 검증한다.

    Returns:
        (설정 여부, 메시지)
    """
    config = _get_oauth_config()
    if not config:
        return False, "missing"

    client_id = config.get("client_id", "")
    client_secret = config.get("client_secret", "")
    redirect_uri = config.get("redirect_uri", "")

    if "YOUR_CLIENT_ID" in client_id or "YOUR_CLIENT_SECRET" in client_secret:
        return False, "placeholder"

    if not client_id or not client_secret or not redirect_uri:
        return False, "empty"

    return True, "설정 확인됨"

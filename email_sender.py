"""
Gmail 메일 발송 모듈
- Gmail 앱 비밀번호 인증 (SMTP SSL / STARTTLS)
- 연결 테스트
- 단일/대량 메일 발송
- 발송 간격 제어
"""

import smtplib
import ssl
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import Optional, Callable

SMTP_SERVER = "smtp.gmail.com"
SMTP_SSL_PORT = 465
SMTP_TLS_PORT = 587


def _clean_password(app_password: str) -> str:
    """앱 비밀번호에서 공백을 제거한다. (Google이 4자리씩 공백으로 표시하므로)"""
    return app_password.replace(" ", "").strip()


def _create_ssl_context() -> ssl.SSLContext:
    """호환성 높은 SSL 컨텍스트를 생성한다."""
    ctx = ssl.create_default_context()
    return ctx


def _connect_smtp(email: str, app_password: str, timeout: int = 15) -> smtplib.SMTP_SSL:
    """
    Gmail SMTP에 연결하고 로그인한다.
    SSL(465) 시도 후 실패하면 STARTTLS(587)로 폴백한다.

    Returns:
        로그인된 SMTP 연결 객체
    """
    password = _clean_password(app_password)
    ctx = _create_ssl_context()

    # 1차 시도: SMTP_SSL (포트 465)
    try:
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_SSL_PORT, timeout=timeout, context=ctx)
        server.ehlo()
        server.login(email, password)
        return server
    except smtplib.SMTPAuthenticationError:
        raise  # 인증 오류는 그대로 전파
    except Exception:
        pass  # SSL 연결 실패 시 STARTTLS로 폴백

    # 2차 시도: STARTTLS (포트 587)
    server = smtplib.SMTP(SMTP_SERVER, SMTP_TLS_PORT, timeout=timeout)
    server.ehlo()
    server.starttls(context=ctx)
    server.ehlo()
    server.login(email, password)
    return server


def test_connection(email: str, app_password: str) -> tuple[bool, str]:
    """
    Gmail SMTP 연결을 테스트한다.

    Args:
        email: Gmail 주소
        app_password: 앱 비밀번호

    Returns:
        (성공 여부, 메시지)
    """
    try:
        server = _connect_smtp(email, app_password, timeout=15)
        server.quit()
        return True, "Gmail 연결 성공!"
    except smtplib.SMTPAuthenticationError as e:
        code, msg = e.args
        detail = msg.decode("utf-8", errors="replace") if isinstance(msg, bytes) else str(msg)
        return False, (
            f"인증 실패 (코드 {code}): Gmail 주소 또는 앱 비밀번호를 확인해주세요.\n\n"
            f"**확인사항:**\n"
            f"- Google 계정에서 **2단계 인증**이 활성화되어 있나요?\n"
            f"- **앱 비밀번호**를 사용하고 계신가요? (일반 비밀번호는 사용 불가)\n"
            f"- 앱 비밀번호 공백 없이 16자리를 정확히 입력했나요?\n\n"
            f"상세: {detail}"
        )
    except smtplib.SMTPException as e:
        return False, f"SMTP 오류: {e}"
    except ConnectionRefusedError:
        return False, "연결 거부: 네트워크 방화벽이 SMTP 포트를 차단하고 있을 수 있습니다."
    except TimeoutError:
        return False, "연결 시간 초과: 네트워크 상태를 확인해주세요."
    except Exception as e:
        return False, f"연결 오류 ({type(e).__name__}): {e}"


def send_single_email(
    smtp_connection: smtplib.SMTP_SSL,
    from_email: str,
    from_name: str,
    to_email: str,
    subject: str,
    body: str,
    is_html: bool = False,
) -> tuple[bool, str]:
    """
    단일 메일을 발송한다.

    Args:
        smtp_connection: 이미 연결된 SMTP 객체
        from_email: 발신자 이메일
        from_name: 발신자 이름
        to_email: 수신자 이메일
        subject: 메일 제목
        body: 메일 본문
        is_html: HTML 형식 여부

    Returns:
        (성공 여부, 메시지)
    """
    try:
        msg = MIMEMultipart("alternative")
        msg["From"] = f"{from_name} <{from_email}>" if from_name else from_email
        msg["To"] = to_email
        msg["Subject"] = subject

        content_type = "html" if is_html else "plain"
        # 줄바꿈을 <br>로 변환 (plain text인 경우에도 HTML로 감싸서 줄바꿈 유지)
        if not is_html:
            html_body = body.replace("\n", "<br>")
            msg.attach(MIMEText(body, "plain", "utf-8"))
            msg.attach(MIMEText(html_body, "html", "utf-8"))
        else:
            msg.attach(MIMEText(body, "html", "utf-8"))

        smtp_connection.send_message(msg)
        return True, "발송 성공"
    except smtplib.SMTPRecipientsRefused:
        return False, "수신자 주소 오류"
    except smtplib.SMTPException as e:
        return False, f"SMTP 오류: {e}"
    except Exception as e:
        return False, f"발송 오류: {e}"


def send_bulk_emails(
    from_email: str,
    app_password: str,
    from_name: str,
    email_list: list[dict],
    delay_seconds: float = 3.0,
    progress_callback: Optional[Callable] = None,
) -> list[dict]:
    """
    대량 메일을 발송한다.

    Args:
        from_email: 발신자 이메일
        app_password: 앱 비밀번호
        from_name: 발신자 이름
        email_list: [{"to": 수신자, "subject": 제목, "body": 본문, ...}, ...]
        delay_seconds: 발송 간격 (초)
        progress_callback: 진행 상황 콜백 fn(current, total, result_dict)

    Returns:
        [{"to": 수신자, "success": 성공여부, "message": 결과메시지, "time": 발송시각}, ...]
    """
    results = []
    total = len(email_list)

    try:
        server = _connect_smtp(from_email, app_password, timeout=30)

        try:
            for i, email_data in enumerate(email_list):
                to_email = email_data["to"]
                subject = email_data["subject"]
                body = email_data["body"]

                success, message = send_single_email(
                    smtp_connection=server,
                    from_email=from_email,
                    from_name=from_name,
                    to_email=to_email,
                    subject=subject,
                    body=body,
                )

                result = {
                    "to": to_email,
                    "success": success,
                    "message": message,
                    "time": time.strftime("%H:%M:%S"),
                    "note": email_data.get("note", ""),
                }
                results.append(result)

                if progress_callback:
                    progress_callback(i + 1, total, result)

                # 마지막 메일이 아닐 때만 대기
                if i < total - 1:
                    time.sleep(delay_seconds)
        finally:
            server.quit()

    except smtplib.SMTPAuthenticationError:
        # 인증 실패 시 모든 나머지 메일을 실패 처리
        for email_data in email_list[len(results):]:
            results.append({
                "to": email_data["to"],
                "success": False,
                "message": "인증 실패",
                "time": time.strftime("%H:%M:%S"),
                "note": "",
            })
    except Exception as e:
        # 연결 오류 시 나머지 메일 실패 처리
        for email_data in email_list[len(results):]:
            results.append({
                "to": email_data["to"],
                "success": False,
                "message": f"연결 오류: {e}",
                "time": time.strftime("%H:%M:%S"),
                "note": "",
            })

    return results

"""
콜드메일 자동발송 - Streamlit 웹 앱
Google OAuth2 로그인 + Gmail API 발송
"""

import streamlit as st
import pandas as pd
import time

from send_history import (
    add_sent_emails_batch,
    get_sent_emails,
    get_sent_count,
    clear_history,
)

from excel_parser import read_excel, get_column_names, analyze_data, get_row_data
from template_engine import extract_variables, render_email, get_empty_variables
from google_auth import (
    check_secrets_configured,
    get_authorization_url,
    exchange_code_for_credentials,
    get_user_info,
    credentials_to_dict,
    credentials_from_dict,
    get_gmail_service,
    get_gmail_signature,
    send_email,
)

# ─────────────────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="콜드메일 자동발송",
    page_icon="✉️",
    layout="wide",
)

# ─────────────────────────────────────────────────────────
# 커스텀 CSS
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Google 로그인 버튼 (공식 스타일) */
    .google-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 12px;
        width: 100%;
        padding: 10px 16px;
        background-color: #ffffff;
        color: #3c4043;
        border: 1px solid #dadce0;
        border-radius: 8px;
        text-decoration: none;
        font-weight: 500;
        font-size: 14px;
        font-family: 'Google Sans', Roboto, Arial, sans-serif;
        cursor: pointer;
        transition: background-color 0.2s, box-shadow 0.2s;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .google-btn:hover {
        background-color: #f7f8f8;
        box-shadow: 0 1px 3px rgba(0,0,0,0.16);
        color: #3c4043;
        text-decoration: none;
    }
    .google-btn:active {
        background-color: #e8eaed;
    }

    /* 로그인된 사용자 프로필 카드 */
    .user-profile-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 20px 16px;
        text-align: center;
        color: white;
        margin-bottom: 8px;
    }
    .user-profile-card img {
        width: 56px;
        height: 56px;
        border-radius: 50%;
        border: 3px solid rgba(255,255,255,0.6);
        margin-bottom: 8px;
    }
    .user-profile-card .user-name {
        font-size: 16px;
        font-weight: 600;
        margin: 4px 0 2px 0;
    }
    .user-profile-card .user-email {
        font-size: 12px;
        opacity: 0.85;
    }
    .user-profile-card .badge {
        display: inline-block;
        background: rgba(255,255,255,0.25);
        border-radius: 12px;
        padding: 2px 10px;
        font-size: 11px;
        margin-top: 8px;
    }

    /* 설정 가이드 스타일 */
    .setup-step {
        background-color: #f8f9fa;
        border-left: 3px solid #4285f4;
        padding: 12px 16px;
        margin: 8px 0;
        border-radius: 0 8px 8px 0;
        font-size: 13px;
    }
    .setup-step strong {
        color: #4285f4;
    }

    /* 엑셀 드래그앤드랍 업로드 영역 */
    .upload-area [data-testid="stFileUploader"] {
        background: transparent;
    }
    .upload-area [data-testid="stFileUploader"] section {
        border: 2px dashed #b0bec5;
        border-radius: 12px;
        padding: 40px 20px;
        background-color: #fafbfc;
        transition: all 0.2s ease;
    }
    .upload-area [data-testid="stFileUploader"] section:hover {
        border-color: #4A90D9;
        background-color: #f0f6ff;
    }
    .upload-area [data-testid="stFileUploader"] section > div {
        display: flex;
        flex-direction: column;
        align-items: center;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# 세션 상태 초기화
# ─────────────────────────────────────────────────────────
DEFAULT_STATE = {
    "gmail_connected": False,
    "gmail_email": "",
    "google_credentials": None,
    "user_info": None,
    "gmail_sender_name": "",
    "df": None,
    "subject_template": "",
    "body_template": "",
    "empty_handling": "defaults",
    "defaults_map": {},
    "alt_subject": "",
    "alt_body": "",
    "email_column": "",
    "column_mapping": {},
    "send_delay": 3,
    "send_results": [],
    "sending_done": False,
    "gmail_signature": "",
    "use_signature": True,
    "attachments": [],                 # 첨부 파일 목록
    "daily_limit": 500,               # 일일 발송 한도 (사용자 설정)
    "daily_sent_count": 0,            # 오늘 발송한 건수
    "daily_sent_date": "",            # 마지막 발송 날짜 (YYYY-MM-DD)
}

for key, default_val in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

# ── 일일 발송 카운터 날짜 리셋 ──
import datetime
_today = datetime.date.today().isoformat()
if st.session_state.daily_sent_date != _today:
    st.session_state.daily_sent_count = 0
    st.session_state.daily_sent_date = _today

# ── 한도 헬퍼 ──
def _get_daily_limit() -> int:
    return st.session_state.daily_limit

def _get_remaining() -> int:
    return max(0, _get_daily_limit() - st.session_state.daily_sent_count)

# ── 인증 헬퍼 (send_history 호출용) ──
def _get_credentials():
    """세션에 저장된 credentials를 복원한다. 로그인 안 된 경우 None."""
    cred_dict = st.session_state.get("google_credentials")
    if cred_dict:
        return credentials_from_dict(cred_dict)
    return None

# ─────────────────────────────────────────────────────────
# Google OAuth 콜백 처리
# ─────────────────────────────────────────────────────────
query_params = st.query_params

# Google OAuth 에러 응답 처리 (예: 테스트 사용자 미등록, 권한 거부 등)
if "error" in query_params:
    error_code = query_params.get("error", "unknown")
    error_messages = {
        "access_denied": "접근이 거부되었습니다. Google Cloud 프로젝트가 '테스트' 모드인 경우, OAuth 동의 화면 → Audience에서 사용할 Google 계정을 테스트 사용자로 추가해주세요.",
        "invalid_client": "OAuth 클라이언트 정보가 올바르지 않습니다. client_id와 client_secret을 확인해주세요.",
        "redirect_uri_mismatch": "리디렉션 URI가 일치하지 않습니다. Google Cloud Console의 승인된 리디렉션 URI와 secrets의 redirect_uri가 동일한지 확인해주세요.",
        "invalid_scope": "요청한 OAuth 스코프가 올바르지 않습니다. API가 모두 활성화되어 있는지 확인해주세요.",
    }
    st.session_state.login_error = error_messages.get(
        error_code,
        f"Google 인증 오류가 발생했습니다. (오류 코드: {error_code})"
    )
    st.query_params.clear()
    st.rerun()

if "code" in query_params and not st.session_state.gmail_connected:
    try:
        code = query_params["code"]
        credentials = exchange_code_for_credentials(code)
        user_info = get_user_info(credentials)
        st.session_state.google_credentials = credentials_to_dict(credentials)
        st.session_state.user_info = user_info
        st.session_state.gmail_connected = True
        st.session_state.gmail_email = user_info.get("email", "")
        st.session_state.gmail_sender_name = user_info.get("name", "")

        # Gmail 서명 자동 가져오기
        try:
            gmail_svc = get_gmail_service(credentials)
            sig = get_gmail_signature(gmail_svc, user_info.get("email", ""))
            st.session_state.gmail_signature = sig
        except Exception:
            st.session_state.gmail_signature = ""
    except Exception as e:
        error_str = str(e)
        if "redirect_uri_mismatch" in error_str.lower() or "redirect" in error_str.lower():
            st.session_state.login_error = "리디렉션 URI가 일치하지 않습니다. Google Cloud Console의 '승인된 리디렉션 URI'와 secrets.toml의 redirect_uri가 현재 앱 URL과 동일한지 확인해주세요."
        elif "invalid_grant" in error_str.lower():
            st.session_state.login_error = "인증 코드가 만료되었거나 이미 사용되었습니다. 다시 로그인해주세요."
        else:
            st.session_state.login_error = f"로그인 처리 중 오류: {error_str}"
    st.query_params.clear()
    st.rerun()

# ─────────────────────────────────────────────────────────
# 사이드바: Google 로그인 + 발송 설정
# ─────────────────────────────────────────────────────────
with st.sidebar:
    if not st.session_state.gmail_connected:
        # ── 로그인 전 ──
        st.header("📧 Gmail 연동")

        secrets_ok, secrets_msg = check_secrets_configured()

        if not secrets_ok:
            # ── OAuth 미설정: 가이드 표시 ──
            st.markdown("""
            <div style="text-align: center; padding: 16px 0;">
                <div style="font-size: 48px; margin-bottom: 8px;">🔐</div>
                <div style="color: #5f6368; font-size: 13px;">
                    Google 로그인을 사용하려면<br/>
                    OAuth 설정이 필요합니다
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("🔧 설정 방법 (5분 소요)", expanded=True):
                # 현재 앱 URL을 동적으로 결정
                from google_auth import _get_redirect_uri
                current_redirect = _get_redirect_uri()

                st.markdown(f"""
**1단계: Google Cloud 프로젝트**
- [Google Cloud Console](https://console.cloud.google.com/) 접속
- 새 프로젝트 생성

**2단계: API 활성화**
- [API 라이브러리](https://console.cloud.google.com/apis/library)에서 아래 3개 검색 후 사용 설정:
  - **Gmail API**
  - **Google Sheets API**
  - **Google Drive API**

**3단계: OAuth 동의 화면**
- [Auth Platform](https://console.cloud.google.com/auth/overview) → Branding 설정
- Audience → **외부** 선택 → ADD USERS로 테스트 계정 추가

**4단계: OAuth 클라이언트 생성**
- [Clients](https://console.cloud.google.com/auth/clients) → CREATE CLIENT
- 유형: **웹 애플리케이션**
- 승인된 리디렉션 URI: `{current_redirect}`

**5단계: Secrets 설정**
""")
                st.code(f"""[google]
client_id = "발급받은_ID.apps.googleusercontent.com"
client_secret = "발급받은_SECRET"
redirect_uri = "{current_redirect}"
""", language="toml")
                st.caption("로컬: `.streamlit/secrets.toml` 수정 후 새로고침\nStreamlit Cloud: 앱 설정 → Secrets에 입력")

        else:
            # ── OAuth 설정 완료: 로그인 버튼 ──
            try:
                auth_url, state = get_authorization_url()
                st.session_state.oauth_state = state

                st.markdown("""
                <div style="text-align: center; padding: 24px 0 16px 0;">
                    <div style="font-size: 48px; margin-bottom: 8px;">✉️</div>
                    <div style="color: #5f6368; font-size: 13px; margin-bottom: 20px;">
                        Google 계정으로 로그인하여<br/>Gmail 발송 기능을 사용하세요
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # 공식 Google 로그인 버튼 스타일
                st.markdown(
                    f"""
                    <a href="{auth_url}" target="_self" class="google-btn">
                        <svg width="18" height="18" viewBox="0 0 48 48">
                            <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                            <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                            <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
                            <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
                        </svg>
                        Google로 로그인
                    </a>
                    """,
                    unsafe_allow_html=True,
                )

                st.markdown("<br>", unsafe_allow_html=True)
                st.caption("🔒 로그인 시 Gmail 발송 권한만 요청합니다.  \n비밀번호는 저장되지 않습니다.")

            except Exception as e:
                st.error(f"OAuth 설정 오류: {e}")
                st.info("💡 OAuth 설정을 확인해주세요. client_id, client_secret, redirect_uri가 올바른지 확인이 필요합니다.")

        # 로그인 에러 표시
        if "login_error" in st.session_state:
            st.error(f"⚠️ 로그인 실패: {st.session_state.login_error}")
            del st.session_state.login_error

            # 에러 발생 시 문제 해결 가이드 표시
            with st.expander("🔍 문제 해결 가이드", expanded=True):
                from google_auth import _get_redirect_uri
                current_redirect = _get_redirect_uri()
                st.markdown(f"""
**자주 발생하는 문제:**

1. **테스트 사용자 미등록** (가장 흔함)
   - [Google Cloud Console → OAuth 동의 화면 → Audience](https://console.cloud.google.com/auth/audience)
   - **ADD USERS**로 로그인할 Google 계정 추가

2. **리디렉션 URI 불일치**
   - [Google Cloud Console → 사용자 인증 정보 → OAuth 클라이언트](https://console.cloud.google.com/auth/clients)
   - 승인된 리디렉션 URI에 `{current_redirect}` 등록 확인
   - secrets의 redirect_uri도 동일한지 확인

3. **API 미활성화**
   - [API 라이브러리](https://console.cloud.google.com/apis/library)에서 아래 API 활성화 확인:
   - Gmail API, Google Sheets API, Google Drive API
""")

    else:
        # ── 로그인 후: 프로필 표시 ──
        user_info = st.session_state.user_info or {}
        user_name = user_info.get("name", "사용자")
        user_email = user_info.get("email", "")
        user_picture = user_info.get("picture", "")

        # 프로필 카드
        profile_img = f'<img src="{user_picture}" />' if user_picture else '<div style="width:56px;height:56px;border-radius:50%;background:rgba(255,255,255,0.3);margin:0 auto 8px auto;display:flex;align-items:center;justify-content:center;font-size:24px;">👤</div>'

        st.markdown(f"""
        <div class="user-profile-card">
            {profile_img}
            <div class="user-name">{user_name}</div>
            <div class="user-email">{user_email}</div>
            <div class="badge">✅ Gmail 연동됨</div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("로그아웃", use_container_width=True, type="secondary"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        # ── Gmail 서명 설정 ──
        st.divider()
        st.subheader("✍️ Gmail 서명")

        if st.session_state.gmail_signature:
            st.session_state.use_signature = st.toggle(
                "서명 자동 첨부",
                value=st.session_state.use_signature,
                help="ON 시 Gmail에 설정된 서명이 모든 메일 하단에 자동 추가됩니다",
            )

            with st.expander("서명 미리보기", expanded=False):
                st.markdown(
                    st.session_state.gmail_signature,
                    unsafe_allow_html=True,
                )
        else:
            st.caption("Gmail에 설정된 서명이 없습니다.")
            st.caption("Gmail 설정에서 서명을 추가하면\n여기에 자동으로 표시됩니다.")

    st.divider()
    st.header("⚙️ 발송 설정")

    st.session_state.gmail_sender_name = st.text_input(
        "발신자 이름",
        value=st.session_state.gmail_sender_name,
        placeholder="홍길동 / ABC컴퍼니",
    )

    st.session_state.send_delay = st.number_input(
        "발송 간격 (초)",
        min_value=1,
        max_value=30,
        value=st.session_state.send_delay,
        help="스팸 방지를 위해 최소 3초 이상 권장",
    )

    # ── 일일 발송 한도 ──
    st.divider()
    st.header("📊 일일 발송 한도")

    st.session_state.daily_limit = st.number_input(
        "일일 최대 발송 건수",
        min_value=1,
        max_value=2000,
        value=st.session_state.daily_limit,
        step=50,
        help="개인 Gmail은 500건/일, Workspace는 2,000건/일이 Google 제한입니다. 초과 시 계정이 24시간 차단될 수 있습니다.",
        key="daily_limit_input",
    )
    st.caption("⚠️ 개인 Gmail: 최대 500건 / Workspace: 최대 2,000건")

    daily_limit = _get_daily_limit()
    sent_today = st.session_state.daily_sent_count
    remaining = _get_remaining()
    usage_pct = sent_today / daily_limit if daily_limit > 0 else 0

    # 잔여 한도 프로그레스 바
    st.progress(min(usage_pct, 1.0))

    col_l, col_r = st.columns(2)
    with col_l:
        st.metric("오늘 발송", f"{sent_today}건")
    with col_r:
        st.metric("잔여 한도", f"{remaining}건")

    if remaining == 0:
        st.error("🚫 오늘 발송 한도를 모두 소진했습니다.\n24시간 후 자동 초기화됩니다.")
    elif usage_pct >= 0.8:
        st.warning(f"⚠️ 한도의 {int(usage_pct * 100)}%를 사용했습니다. 남은 {remaining}건만 발송 가능합니다.")

    # ── 발송 이력 관리 ──
    if st.session_state.gmail_connected:
        st.divider()
        st.header("📋 발송 이력")
        creds = _get_credentials()
        if creds:
            total_history = get_sent_count(creds)
            st.metric("누적 발송 완료", f"{total_history}건")
            if total_history > 0:
                st.caption("이미 발송한 수신자는 자동으로 건너뜁니다.\n이력은 내 Google Drive 시트에 저장됩니다.")
                if st.button("🗑️ 이력 초기화", key="sidebar_clear_history", help="이력을 초기화하면 같은 수신자에게 다시 발송할 수 있습니다"):
                    clear_history(creds)
                    st.rerun()
            else:
                st.caption("아직 발송 이력이 없습니다.")


# ─────────────────────────────────────────────────────────
# 메인 영역: 4단계 탭
# ─────────────────────────────────────────────────────────
st.title("✉️ 콜드메일 자동발송")

tab1, tab2, tab3, tab4 = st.tabs([
    "📝 Step 1: 메일 작성",
    "📂 Step 2: 엑셀 업로드",
    "👁️ Step 3: 미리보기",
    "🚀 Step 4: 발송",
])


# ─────────────────────────────────────────────────────────
# Step 1: 메일 작성
# ─────────────────────────────────────────────────────────
with tab1:
    st.subheader("📝 메일 스크립트 작성")
    st.info("💡 변수를 넣고 싶은 곳에 **{변수명}** 을 입력하세요.  \n예: {회사명}, {담당자}, {직책}")

    subject_template = st.text_input(
        "메일 제목",
        value=st.session_state.subject_template,
        placeholder="예: {회사명} 협업 제안드립니다 - ABC컴퍼니",
        key="subject_input",
    )
    st.session_state.subject_template = subject_template

    body_template = st.text_area(
        "메일 본문",
        value=st.session_state.body_template,
        height=300,
        placeholder="예:\n안녕하세요, {담당자}님.\n\n{회사명}의 {직책}님께 협업을 제안드리고자 연락드립니다.\n\n저희 ABC컴퍼니는 ...",
        key="body_input",
    )
    st.session_state.body_template = body_template

    # 사용된 변수 감지
    all_text = subject_template + " " + body_template
    used_vars = extract_variables(all_text)

    if used_vars:
        st.success(f"🔖 **사용된 변수 목록:** {', '.join(['{' + v + '}' for v in used_vars])}")
    elif subject_template or body_template:
        st.warning("⚠️ 아직 변수가 없습니다. {변수명} 형태로 변수를 입력해보세요.")

    # 빈 데이터 처리
    st.divider()
    st.subheader("⚠️ 빈 데이터 처리")
    st.caption("엑셀 데이터에 빈 값이 있을 때 어떻게 처리할지 설정합니다.")

    empty_handling = st.radio(
        "빈 값 처리 방식",
        options=["defaults", "alt_template"],
        format_func=lambda x: "기본값으로 대체" if x == "defaults" else "빈 값이 있는 행은 별도 템플릿 사용",
        index=0 if st.session_state.empty_handling == "defaults" else 1,
        key="empty_handling_radio",
    )
    st.session_state.empty_handling = empty_handling

    if empty_handling == "defaults":
        if used_vars:
            st.caption("각 변수의 기본값을 설정하세요. (비워두면 해당 부분이 빈 채로 남습니다)")
            defaults_map = {}
            cols = st.columns(min(len(used_vars), 3))
            for i, var in enumerate(used_vars):
                with cols[i % len(cols)]:
                    default_val = st.text_input(
                        f"{{{var}}} 기본값",
                        value=st.session_state.defaults_map.get(var, ""),
                        key=f"default_{var}",
                    )
                    defaults_map[var] = default_val
            st.session_state.defaults_map = defaults_map
    else:
        st.caption("빈 값이 포함된 수신자에게 보낼 별도 제목/본문을 작성하세요.")
        alt_subject = st.text_input(
            "대체 메일 제목",
            value=st.session_state.alt_subject,
            placeholder="예: 협업 제안드립니다 - ABC컴퍼니",
            key="alt_subject_input",
        )
        st.session_state.alt_subject = alt_subject

        alt_body = st.text_area(
            "대체 메일 본문",
            value=st.session_state.alt_body,
            height=200,
            placeholder="빈 변수 없이 작성하거나, 다른 변수를 사용하세요.",
            key="alt_body_input",
        )
        st.session_state.alt_body = alt_body

    # 파일 첨부
    st.divider()
    st.subheader("📎 파일 첨부")
    st.caption("모든 수신자에게 동일한 파일이 첨부됩니다. 최대 25MB (Gmail 제한)")

    attached_files = st.file_uploader(
        "첨부할 파일을 선택하세요",
        accept_multiple_files=True,
        help="PDF, 이미지, 문서 등 다양한 파일을 첨부할 수 있습니다. Gmail 제한: 총 25MB",
        key="attachment_uploader",
    )

    if attached_files:
        st.session_state.attachments = attached_files
        total_size = sum(f.size for f in attached_files)
        size_mb = total_size / (1024 * 1024)

        if size_mb > 25:
            st.error(f"⚠️ 총 첨부 파일 크기가 {size_mb:.1f}MB입니다. Gmail 제한(25MB)을 초과합니다.")
        else:
            file_info = ", ".join([f"**{f.name}** ({f.size/1024:.0f}KB)" for f in attached_files])
            st.success(f"📎 첨부 파일 {len(attached_files)}개: {file_info}  \n총 크기: {size_mb:.1f}MB")
    else:
        st.session_state.attachments = []


# ─────────────────────────────────────────────────────────
# Step 2: 엑셀 업로드
# ─────────────────────────────────────────────────────────
with tab2:
    st.subheader("📂 엑셀 파일 업로드")

    st.markdown("""
    <div style="text-align:center; padding:8px 0 4px 0; color:#78909c;">
        <span style="font-size:36px;">📎</span><br>
        <span style="font-size:14px;">파일을 아래 영역에 <b>드래그앤드랍</b>하거나 <b>Browse files</b>를 클릭하세요</span><br>
        <span style="font-size:12px; color:#aaa;">.xlsx, .xls 지원</span>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="upload-area">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "엑셀 파일을 드래그앤드랍 또는 선택하세요",
            type=["xlsx", "xls"],
            help="수신자 목록이 담긴 .xlsx 또는 .xls 파일을 업로드하세요",
            label_visibility="collapsed",
        )
        st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        try:
            df = read_excel(uploaded_file)
            st.session_state.df = df

            st.success(f"✅ 파일 업로드 완료: **{uploaded_file.name}** (총 {len(df)}건)")

            # 데이터 미리보기
            st.subheader("📋 데이터 미리보기")

            def highlight_empty(val):
                if pd.isna(val) or str(val).strip() == "":
                    return "background-color: #FFF3CD; color: #856404;"
                return ""

            styled_df = df.head(10).style.map(highlight_empty)
            st.dataframe(styled_df, use_container_width=True)

            if len(df) > 10:
                st.caption(f"... 외 {len(df) - 10}건 더 있음")

            # 변수 연결
            st.divider()
            st.subheader("🔗 변수 연결")

            columns = get_column_names(df)

            # 이메일 열 자동 추측
            email_guess_idx = 0
            for i, col in enumerate(columns):
                col_lower = col.lower()
                if "email" in col_lower or "이메일" in col_lower or "메일" in col_lower or "mail" in col_lower:
                    email_guess_idx = i
                    break

            email_column = st.selectbox(
                "📮 수신 이메일 열 선택 (필수)",
                options=columns,
                index=email_guess_idx,
                key="email_col_select",
            )
            st.session_state.email_column = email_column

            # 템플릿 변수와 엑셀 열 매핑
            used_vars = extract_variables(
                st.session_state.subject_template + " " + st.session_state.body_template
            )

            if used_vars:
                st.caption("템플릿 변수와 엑셀 열을 연결하세요. (같은 이름이면 자동 매칭됩니다)")

                mapping = {}
                for var in used_vars:
                    auto_idx = 0
                    matched = False
                    for i, col in enumerate(columns):
                        if col == var:
                            auto_idx = i
                            matched = True
                            break

                    col1, col2 = st.columns([3, 1])
                    with col1:
                        selected_col = st.selectbox(
                            f"{{{var}}}",
                            options=columns,
                            index=auto_idx,
                            key=f"map_{var}",
                        )
                        mapping[var] = selected_col
                    with col2:
                        if matched:
                            st.markdown("<br>", unsafe_allow_html=True)
                            st.success("✅ 자동 매칭")
                        else:
                            st.markdown("<br>", unsafe_allow_html=True)
                            st.warning("🔧 수동 선택")

                st.session_state.column_mapping = mapping
            else:
                st.info("💡 Step 1에서 먼저 {변수명}이 포함된 메일 스크립트를 작성해주세요.")

            # 데이터 분석
            if used_vars and email_column:
                st.divider()
                st.subheader("📊 데이터 요약")

                mapped_vars = [st.session_state.column_mapping.get(v, v) for v in used_vars]
                analysis = analyze_data(df, mapped_vars, email_column)

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("전체 데이터", f"{analysis['total']}건")
                with col2:
                    st.metric("정상 데이터", f"{analysis['complete']}건")
                with col3:
                    st.metric("⚠️ 빈 값 포함", f"{analysis['has_empty']}건")
                with col4:
                    st.metric("❌ 이메일 없음", f"{analysis['no_email']}건")

                if analysis["empty_details"]:
                    with st.expander(f"⚠️ 빈 값이 있는 행 상세 ({analysis['has_empty']}건)", expanded=False):
                        empty_df = pd.DataFrame(analysis["empty_details"])
                        empty_df.columns = ["엑셀 행", "이메일", "빈 변수"]
                        st.dataframe(empty_df, use_container_width=True)

        except Exception as e:
            st.error(f"❌ 파일 처리 중 오류: {e}")
    else:
        st.info("💡 엑셀 파일(.xlsx, .xls)을 업로드해주세요.")


# ─────────────────────────────────────────────────────────
# Step 3: 미리보기
# ─────────────────────────────────────────────────────────
with tab3:
    st.subheader("👁️ 발송 미리보기")

    df = st.session_state.df
    subject_t = st.session_state.subject_template
    body_t = st.session_state.body_template
    email_col = st.session_state.email_column
    col_mapping = st.session_state.column_mapping

    if df is None:
        st.info("💡 Step 2에서 먼저 엑셀 파일을 업로드해주세요.")
    elif not subject_t or not body_t:
        st.info("💡 Step 1에서 먼저 메일 스크립트를 작성해주세요.")
    elif not email_col:
        st.info("💡 Step 2에서 이메일 열을 선택해주세요.")
    else:
        used_vars = extract_variables(subject_t + " " + body_t)

        # 유효한 행만 필터링
        valid_indices = []
        for idx in range(len(df)):
            row_data = get_row_data(df, idx)
            email_val = row_data.get(email_col, "")
            if email_val and email_val != "nan":
                valid_indices.append(idx)

        if not valid_indices:
            st.error("❌ 발송 가능한 데이터가 없습니다. 이메일 열을 확인해주세요.")
        else:
            total_valid = len(valid_indices)
            st.write(f"**전체 {total_valid}건** 미리보기")

            # 미리보기 네비게이션
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                preview_idx = st.number_input(
                    "미리보기 번호",
                    min_value=1,
                    max_value=total_valid,
                    value=1,
                    step=1,
                    key="preview_nav",
                    label_visibility="collapsed",
                )
                st.caption(f"◀ {preview_idx} / {total_valid} ▶")

            # 현재 행 데이터
            current_row_idx = valid_indices[preview_idx - 1]
            row_data = get_row_data(df, current_row_idx)

            mapped_data = {}
            for var in used_vars:
                mapped_col = col_mapping.get(var, var)
                mapped_data[var] = row_data.get(mapped_col, "")

            defaults = st.session_state.defaults_map if st.session_state.empty_handling == "defaults" else {}
            alt_sub = st.session_state.alt_subject if st.session_state.empty_handling == "alt_template" else None
            alt_bod = st.session_state.alt_body if st.session_state.empty_handling == "alt_template" else None

            rendered = render_email(
                subject_template=subject_t,
                body_template=body_t,
                data=mapped_data,
                defaults=defaults,
                alt_subject_template=alt_sub,
                alt_body_template=alt_bod,
            )

            to_email = row_data.get(email_col, "")

            # 메일 미리보기 카드
            st.markdown("---")
            st.markdown(f"**To:** {to_email}")
            st.markdown(f"**Subject:** {rendered['subject']}")
            if rendered["used_alt"]:
                st.caption("📌 별도 템플릿 적용됨")

            # 첨부파일 표시
            if st.session_state.attachments:
                att_names = ", ".join([f"📎 {f.name}" for f in st.session_state.attachments])
                st.markdown(f"**Attachments:** {att_names}")

            empty_vars = get_empty_variables(mapped_data, used_vars)
            if empty_vars:
                st.warning(f"⚠️ 빈 값 변수: {', '.join(empty_vars)}")

            st.markdown("---")

            # 본문 + 서명 미리보기
            body_html = rendered['body'].replace(chr(10), '<br>')
            sig_html = ""
            if st.session_state.use_signature and st.session_state.gmail_signature:
                sig_html = (
                    f'<div style="border-top: 1px solid #ccc; padding-top: 8px; margin-top: 16px;">'
                    f'{st.session_state.gmail_signature}'
                    f'</div>'
                )

            preview_html = f'<div style="background-color:#f8f9fa;padding:20px;border-radius:8px;border:1px solid #dee2e6;line-height:1.8;">{body_html}{sig_html}</div>'
            st.markdown(preview_html, unsafe_allow_html=True)

            # 빈 값 요약
            st.divider()
            mapped_vars_for_analysis = [col_mapping.get(v, v) for v in used_vars]
            analysis = analyze_data(df, mapped_vars_for_analysis, email_col)

            if analysis["empty_details"]:
                st.subheader(f"⚠️ 빈 값이 있는 메일 ({analysis['has_empty']}건)")

                empty_summary = []
                for detail in analysis["empty_details"]:
                    handling = ""
                    if st.session_state.empty_handling == "defaults":
                        parts = []
                        for ev in detail["empty_vars"]:
                            var_name = ev
                            for v, mapped_col in col_mapping.items():
                                if mapped_col == ev:
                                    var_name = v
                                    break
                            default_val = st.session_state.defaults_map.get(var_name, "")
                            if default_val:
                                parts.append(f'→ "{default_val}"')
                            else:
                                parts.append("→ (빈 채로)")
                        handling = ", ".join(parts)
                    else:
                        handling = "별도 템플릿 적용"

                    empty_summary.append({
                        "수신자": detail["email"],
                        "빈 변수": ", ".join(detail["empty_vars"]),
                        "처리 방식": handling,
                    })

                st.dataframe(pd.DataFrame(empty_summary), use_container_width=True)

            st.divider()
            st.success(
                f"✅ **발송 준비 완료**  \n"
                f"- 정상 발송: {analysis['complete']}건  \n"
                f"- 기본값/대체 적용: {analysis['has_empty']}건  \n"
                f"- 발송 제외 (이메일 없음): {analysis['no_email']}건"
            )


# ─────────────────────────────────────────────────────────
# Step 4: 발송
# ─────────────────────────────────────────────────────────
with tab4:
    st.subheader("🚀 메일 발송")

    df = st.session_state.df
    subject_t = st.session_state.subject_template
    body_t = st.session_state.body_template
    email_col = st.session_state.email_column
    col_mapping = st.session_state.column_mapping

    # 사전 조건 체크
    can_send = True
    if not st.session_state.gmail_connected:
        st.warning("⚠️ 사이드바에서 Google 로그인을 먼저 완료해주세요.")
        can_send = False
    if df is None:
        st.warning("⚠️ Step 2에서 엑셀 파일을 업로드해주세요.")
        can_send = False
    if not subject_t or not body_t:
        st.warning("⚠️ Step 1에서 메일 스크립트를 작성해주세요.")
        can_send = False
    if not email_col:
        st.warning("⚠️ Step 2에서 이메일 열을 선택해주세요.")
        can_send = False

    if can_send:
        used_vars = extract_variables(subject_t + " " + body_t)

        # 이미 발송한 이메일 목록 로드
        creds = _get_credentials()
        already_sent = get_sent_emails(creds) if creds else set()

        # 발송 대상 이메일 목록 생성 (이미 보낸 건 제외)
        email_list = []
        skipped_already_sent = 0
        total_all = 0

        for idx in range(len(df)):
            row_data = get_row_data(df, idx)
            to_email = row_data.get(email_col, "")

            if not to_email or to_email == "nan":
                continue

            total_all += 1

            # 이미 보낸 이메일은 건너뛰기
            if to_email.strip().lower() in already_sent:
                skipped_already_sent += 1
                continue

            mapped_data = {}
            for var in used_vars:
                mapped_col = col_mapping.get(var, var)
                mapped_data[var] = row_data.get(mapped_col, "")

            defaults = st.session_state.defaults_map if st.session_state.empty_handling == "defaults" else {}
            alt_sub = st.session_state.alt_subject if st.session_state.empty_handling == "alt_template" else None
            alt_bod = st.session_state.alt_body if st.session_state.empty_handling == "alt_template" else None

            rendered = render_email(
                subject_template=subject_t,
                body_template=body_t,
                data=mapped_data,
                defaults=defaults,
                alt_subject_template=alt_sub,
                alt_body_template=alt_bod,
            )

            empty_vars_list = get_empty_variables(mapped_data, used_vars)
            note = ""
            if rendered["used_alt"]:
                note = "별도 템플릿"
            elif empty_vars_list:
                note = "기본값 적용"

            email_list.append({
                "to": to_email,
                "subject": rendered["subject"],
                "body": rendered["body"],
                "note": note,
            })

        total = len(email_list)
        delay = st.session_state.send_delay
        remaining = _get_remaining()
        daily_limit = _get_daily_limit()

        # 이미 보낸 건수 안내
        if skipped_already_sent > 0:
            st.success(
                f"✅ **이전에 발송 완료된 {skipped_already_sent}건은 자동으로 건너뜁니다.**  \n"
                f"전체 {total_all}건 중 **미발송 {total}건**만 발송합니다."
            )

        if total == 0 and skipped_already_sent > 0:
            st.info("🎉 모든 수신자에게 이미 발송 완료되었습니다!")
            st.caption(f"총 발송 이력: {get_sent_count(creds) if creds else 0}건")

            if st.button("🗑️ 발송 이력 초기화", help="이력을 초기화하면 모든 수신자에게 다시 발송할 수 있습니다"):
                if creds:
                    clear_history(creds)
                st.rerun()

        elif total > 0:
            # 실제 발송 가능 건수 계산
            actual_send_count = min(total, remaining)
            est_time = actual_send_count * delay
            est_min = est_time // 60
            est_sec = est_time % 60

            # 발송 정보 표시
            attach_info = ""
            if st.session_state.attachments:
                att_count = len(st.session_state.attachments)
                att_size = sum(f.size for f in st.session_state.attachments) / (1024 * 1024)
                attach_info = f"  \n📎 첨부 파일: {att_count}개 ({att_size:.1f}MB)"

            st.info(
                f"📮 **발송 대상: {total}건**  \n"
                f"⏱️ 예상 소요 시간: 약 {est_min}분 {est_sec}초 (간격 {delay}초 기준)"
                f"{attach_info}"
            )

            # ── 한도 관련 경고 ──
            limit_blocked = False

            if remaining == 0:
                st.error(
                    f"🚫 **오늘 일일 한도({daily_limit}건)를 모두 소진했습니다.**  \n"
                    f"24시간 후 자동 초기화됩니다. 내일 다시 시도해주세요."
                )
                limit_blocked = True
            elif total > remaining:
                st.warning(
                    f"⚠️ **발송 대상({total}건)이 잔여 한도({remaining}건)를 초과합니다.**  \n"
                    f"앞에서부터 **{remaining}건만 발송**하고 자동 중단됩니다.  \n"
                    f"나머지는 내일 같은 파일로 발송하면 자동으로 이어서 보냅니다."
                )

            # 발송 시작 버튼
            if not st.session_state.sending_done and not limit_blocked:
                if st.button("✉️ 발송 시작", type="primary", use_container_width=True):
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    log_container = st.empty()

                    results = []
                    success_count = 0
                    fail_count = 0
                    skipped_count = 0
                    sent_emails_batch = []  # 이번에 성공한 이메일 모아서 일괄 저장

                    try:
                        # Gmail API 서비스 생성
                        credentials = credentials_from_dict(st.session_state.google_credentials)
                        gmail_service = get_gmail_service(credentials)

                        # 서명 HTML 결정
                        sig_html = ""
                        if st.session_state.use_signature and st.session_state.gmail_signature:
                            sig_html = st.session_state.gmail_signature

                        for i, email_data in enumerate(email_list):
                            # ── 한도 도달 시 자동 중단 ──
                            if _get_remaining() <= 0:
                                for remaining_data in email_list[i:]:
                                    results.append({
                                        "시간": time.strftime("%H:%M:%S"),
                                        "수신자": remaining_data["to"],
                                        "상태": "⏸️ 한도초과",
                                        "메모": f"일일 한도 {daily_limit}건 도달 → 내일 자동 이어서 발송",
                                    })
                                    skipped_count += 1
                                break

                            success, message = send_email(
                                service=gmail_service,
                                to_email=email_data["to"],
                                subject=email_data["subject"],
                                body=email_data["body"],
                                from_email=st.session_state.gmail_email,
                                from_name=st.session_state.gmail_sender_name,
                                signature_html=sig_html,
                                attachments=st.session_state.attachments if st.session_state.attachments else None,
                            )

                            result = {
                                "시간": time.strftime("%H:%M:%S"),
                                "수신자": email_data["to"],
                                "상태": "✅ 성공" if success else "❌ 실패",
                                "메모": email_data.get("note", "") if success else message,
                            }
                            results.append(result)

                            if success:
                                success_count += 1
                                st.session_state.daily_sent_count += 1
                                sent_emails_batch.append(email_data["to"])
                                # 10건마다 이력 저장 (중간 크래시 대비)
                                if len(sent_emails_batch) >= 10:
                                    add_sent_emails_batch(creds, sent_emails_batch)
                                    sent_emails_batch = []
                            else:
                                fail_count += 1

                            # 진행률 업데이트
                            progress = (i + 1) / total
                            progress_bar.progress(progress)

                            remaining_now = _get_remaining()
                            status_text.markdown(
                                f"**진행:** {i + 1} / {total} ({int(progress * 100)}%)  |  "
                                f"✅ 성공: {success_count}  |  ❌ 실패: {fail_count}  |  "
                                f"📊 잔여 한도: {remaining_now}건"
                            )

                            log_container.dataframe(
                                pd.DataFrame(results),
                                use_container_width=True,
                            )

                            # 대기 (마지막 메일 제외)
                            if i < total - 1:
                                time.sleep(delay)

                    except Exception as e:
                        st.error(f"❌ 발송 중 오류 발생: {e}")
                        for email_data in email_list[len(results):]:
                            results.append({
                                "시간": time.strftime("%H:%M:%S"),
                                "수신자": email_data["to"],
                                "상태": "❌ 실패",
                                "메모": str(e),
                            })
                            fail_count += 1

                    # 남은 이력 일괄 저장
                    if sent_emails_batch and creds:
                        add_sent_emails_batch(creds, sent_emails_batch)

                    # 발송 완료
                    st.session_state.send_results = results
                    st.session_state.sending_done = True

                    st.balloons()

                    summary = f"🎉 **발송 완료!**  \n- ✅ 성공: {success_count}건  \n- ❌ 실패: {fail_count}건"
                    if skipped_count > 0:
                        summary += f"  \n- ⏸️ 한도초과 스킵: {skipped_count}건 (내일 이어서 발송 가능)"
                    summary += f"  \n- 📋 총 발송 이력: {get_sent_count(creds) if creds else '?'}건"
                    st.success(summary)

        # 발송 완료 후 결과 표시
        if st.session_state.sending_done and st.session_state.send_results:
            st.divider()
            st.subheader("📋 발송 결과")

            results_df = pd.DataFrame(st.session_state.send_results)
            st.dataframe(results_df, use_container_width=True)

            csv = results_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="📥 결과 CSV 다운로드",
                data=csv,
                file_name=f"발송결과_{time.strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🔄 이어서 발송 준비", use_container_width=True, help="같은 엑셀로 미발송분을 이어서 발송합니다"):
                    st.session_state.sending_done = False
                    st.session_state.send_results = []
                    st.rerun()
            with col_btn2:
                if st.button("🗑️ 발송 이력 초기화", use_container_width=True, help="이력을 초기화하면 모든 수신자에게 다시 발송합니다"):
                    creds = _get_credentials()
                    if creds:
                        clear_history(creds)
                    st.session_state.sending_done = False
                    st.session_state.send_results = []
                    st.rerun()

            creds_for_count = _get_credentials()
            st.caption(f"📋 총 누적 발송 이력: {get_sent_count(creds_for_count) if creds_for_count else 0}건")

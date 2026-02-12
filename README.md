# ✉️ 콜드메일 자동발송

엑셀 파일의 데이터를 바탕으로 변수가 치환된 콜드 메일을 Gmail로 자동 발송하는 웹 앱입니다.
**Google 계정으로 로그인**하여 별도 비밀번호 설정 없이 바로 사용할 수 있습니다.

## 주요 기능

- **Google 로그인**: "Google로 로그인" 버튼 클릭만으로 Gmail 연동 완료
- **엑셀 데이터 연동**: `.xlsx` / `.xls` 파일을 업로드하면 자동으로 데이터를 파싱
- **변수 치환 템플릿**: `{회사명}`, `{담당자}` 같은 변수를 엑셀 열과 매칭하여 자동 치환
- **빈 데이터 처리**: 빈 값이 있는 경우 기본값 대체 또는 별도 템플릿 적용
- **발송 미리보기**: 실제 발송 전 수신자별 최종 메일 확인
- **실시간 발송 로그**: 진행률, 성공/실패 여부를 실시간 확인

## 사전 준비: Google Cloud 프로젝트 설정

### 1. Google Cloud 프로젝트 생성

1. [Google Cloud Console](https://console.cloud.google.com/) 접속
2. 새 프로젝트 생성 (또는 기존 프로젝트 선택)

### 2. Gmail API 활성화

1. [API 라이브러리](https://console.cloud.google.com/apis/library) 이동
2. "Gmail API" 검색 후 **사용 설정**

### 3. OAuth 동의 화면 설정

1. [Google Auth Platform](https://console.cloud.google.com/auth/overview) → **Branding** 메뉴
2. 앱 이름, 사용자 지원 이메일 등 입력
3. **Audience** 메뉴 → **ADD USERS** → 테스트할 Gmail 주소 추가
4. 앱을 다른 사람도 사용하게 하려면 **Published** 상태로 변경

### 4. OAuth 클라이언트 생성

1. [Clients](https://console.cloud.google.com/auth/clients) 메뉴 → **CREATE CLIENT**
2. 애플리케이션 유형: **웹 애플리케이션**
3. 승인된 리디렉션 URI 추가:
   - 로컬 개발: `http://localhost:8501`
   - Streamlit Cloud: `https://your-app.streamlit.app`
4. **만들기** 클릭 → **Client ID**와 **Client Secret** 복사

### 5. secrets.toml 설정

`.streamlit/secrets.toml` 파일에 다음을 입력합니다:

```toml
[google]
client_id = "YOUR_CLIENT_ID.apps.googleusercontent.com"
client_secret = "YOUR_CLIENT_SECRET"
redirect_uri = "http://localhost:8501"
```

> ⚠️ 이 파일은 `.gitignore`에 포함되어 있어 Git에 커밋되지 않습니다.

## 로컬 실행 방법

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. 앱 실행
streamlit run app.py
```

브라우저에서 `http://localhost:8501` 이 자동으로 열립니다.

## 사용 방법

### Step 1: Google 로그인 (사이드바)
- "Google로 로그인" 버튼 클릭
- Google 계정 선택 및 Gmail 발송 권한 승인

### Step 2: 메일 작성
- 메일 제목과 본문 작성
- 변수를 넣고 싶은 곳에 `{변수명}` 입력 (예: `{회사명}`, `{담당자}`)
- 빈 데이터 처리 방식 설정

### Step 3: 엑셀 업로드
- 수신자 목록이 담긴 엑셀 파일 업로드
- 이메일 열 선택 + 변수-엑셀 열 매칭

### Step 4: 미리보기
- 각 수신자별 최종 메일 내용 확인

### Step 5: 발송
- "발송 시작" 클릭 → 실시간 진행률 확인 → 결과 CSV 다운로드

## 엑셀 파일 형식 예시

| 이메일 | 회사명 | 담당자 | 직책 |
|--------|--------|--------|------|
| kim@company.com | 삼성전자 | 김철수 | 과장 |
| lee@startup.io | 토스 | 이영희 | 팀장 |

## Streamlit Cloud 배포

1. GitHub에 코드 푸시 (`.streamlit/secrets.toml` 제외)
2. [Streamlit Cloud](https://share.streamlit.io/) 접속 → GitHub 연동
3. 앱 설정의 **Secrets**에 `secrets.toml` 내용 입력
4. `redirect_uri`를 배포된 앱 URL로 변경
5. Google Cloud Console에서도 리디렉션 URI 추가

## 기술 스택

- Python 3.10+
- Streamlit 1.42+
- Google OAuth2 + Gmail API
- pandas + openpyxl

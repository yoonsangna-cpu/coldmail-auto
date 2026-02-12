# 콜드메일 자동발송

엑셀 파일의 데이터를 바탕으로 변수가 치환된 콜드 메일을 Gmail로 자동 발송하는 웹 앱입니다.
**Google 계정으로 로그인**하여 별도 비밀번호 설정 없이 바로 사용할 수 있습니다.

## 주요 기능

- **Google 소셜 로그인**: "Google로 로그인" 버튼 클릭만으로 Gmail 연동 완료
- **엑셀 데이터 연동**: `.xlsx` / `.xls` 파일을 드래그앤드랍으로 업로드
- **변수 치환 템플릿**: `{회사명}`, `{담당자}` 같은 변수를 엑셀 열과 매칭하여 자동 치환
- **빈 데이터 처리**: 빈 값이 있는 경우 기본값 대체 또는 별도 템플릿 적용
- **Gmail 서명 연동**: Gmail에 설정된 서명을 자동으로 가져와 메일에 첨부
- **발송 이력 관리**: Google Sheets에 이력 저장, 이미 보낸 수신자는 자동 건너뛰기
- **일일 한도 관리**: 사용자가 설정한 일일 발송 한도 초과 시 자동 중단
- **실시간 발송 로그**: 진행률, 성공/실패 여부를 실시간 확인 + CSV 다운로드

## 로컬 실행 방법

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. 앱 실행
streamlit run app.py
```

브라우저에서 `http://localhost:8501`이 자동으로 열립니다.

## Streamlit Cloud 배포 (권장)

### 1단계: Google Cloud 프로젝트 설정

1. [Google Cloud Console](https://console.cloud.google.com/) 접속 → 새 프로젝트 생성
2. [API 라이브러리](https://console.cloud.google.com/apis/library)에서 아래 3개 API 활성화:
   - **Gmail API**
   - **Google Sheets API**
   - **Google Drive API**
3. [OAuth 동의 화면](https://console.cloud.google.com/auth/overview) → Branding 설정
   - Audience → **외부** 선택
   - **ADD USERS**로 테스트 계정 추가 (100명 이상이면 Google 검증 필요)
4. [OAuth 클라이언트 생성](https://console.cloud.google.com/auth/clients) → CREATE CLIENT
   - 유형: **웹 애플리케이션**
   - 승인된 리디렉션 URI에 **두 개** 추가:
     - `http://localhost:8501` (로컬 개발용)
     - `https://your-app.streamlit.app` (배포 URL - 아래에서 확인)
5. **Client ID**와 **Client Secret** 복사

### 2단계: GitHub에 코드 올리기

이 리포지토리를 Fork하거나, 직접 GitHub에 push합니다.

> `.streamlit/secrets.toml`은 `.gitignore`에 포함되어 있어 Git에 커밋되지 않습니다.

### 3단계: Streamlit Cloud 배포

1. [share.streamlit.io](https://share.streamlit.io/) 접속 → GitHub 계정 연결
2. **New app** → 이 리포지토리 선택 → `app.py` 지정
3. **Advanced settings** → **Secrets**에 아래 내용 입력:

```toml
[google]
client_id = "발급받은_CLIENT_ID.apps.googleusercontent.com"
client_secret = "발급받은_CLIENT_SECRET"
redirect_uri = "https://your-app.streamlit.app"
```

4. **Deploy** 클릭
5. 배포 완료 후 실제 앱 URL(예: `https://coldmail-auto.streamlit.app`)을 확인
6. Google Cloud Console → OAuth 클라이언트 → 승인된 리디렉션 URI에 실제 앱 URL 추가
7. Streamlit Cloud Secrets의 `redirect_uri`도 실제 앱 URL로 수정

### 4단계: 완료

이제 배포된 URL을 공유하면 누구나 자신의 Google 계정으로 로그인해서 사용할 수 있습니다.

## 사용 방법

1. **Google 로그인** (사이드바) → Google 계정 선택 및 권한 승인
2. **메일 작성** → 제목/본문에 `{변수명}` 형태로 변수 삽입
3. **엑셀 업로드** → 수신자 목록 파일 업로드 + 변수-열 매칭
4. **미리보기** → 수신자별 최종 메일 내용 확인
5. **발송** → "발송 시작" 클릭 → 실시간 진행률 확인 → 결과 CSV 다운로드

## 엑셀 파일 형식 예시

| 이메일 | 회사명 | 담당자 | 직책 |
|--------|--------|--------|------|
| kim@company.com | 삼성전자 | 김철수 | 과장 |
| lee@startup.io | 토스 | 이영희 | 팀장 |

## 대량 발송 (수천 건)

- 사이드바에서 **일일 발송 한도**를 설정 (개인 Gmail: 500건/일 권장)
- 한도 도달 시 자동 중단, 나머지는 다음 날 같은 파일로 이어서 발송
- **발송 이력**이 Google Sheets에 자동 저장되어 이미 보낸 수신자는 건너뜀
- 각 사용자의 Google Drive에 "콜드메일_발송이력" 시트가 자동 생성됨

## 기술 스택

- Python 3.10+
- Streamlit
- Google OAuth2 + Gmail API + Sheets API + Drive API
- pandas + openpyxl

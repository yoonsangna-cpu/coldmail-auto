"""
엑셀 파일 파싱 모듈
- 엑셀 파일 읽기 (xlsx, xls)
- 열 이름 추출
- 빈 데이터 감지 및 요약
"""

import pandas as pd
from io import BytesIO
from typing import Optional


def read_excel(file: BytesIO, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    업로드된 엑셀 파일을 DataFrame으로 변환한다.

    Args:
        file: Streamlit file_uploader에서 받은 파일 객체
        sheet_name: 읽을 시트 이름 (None이면 첫 번째 시트)

    Returns:
        pandas DataFrame
    """
    try:
        df = pd.read_excel(file, sheet_name=sheet_name or 0, engine="openpyxl")
        # 열 이름 공백 정리
        df.columns = [str(col).strip() for col in df.columns]
        return df
    except Exception as e:
        raise ValueError(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")


def get_sheet_names(file: BytesIO) -> list[str]:
    """엑셀 파일의 시트 이름 목록을 반환한다."""
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
        return xls.sheet_names
    except Exception as e:
        raise ValueError(f"시트 목록을 읽는 중 오류가 발생했습니다: {e}")


def get_column_names(df: pd.DataFrame) -> list[str]:
    """DataFrame의 열 이름 목록을 반환한다."""
    return list(df.columns)


def analyze_data(df: pd.DataFrame, used_variables: list[str], email_column: str) -> dict:
    """
    데이터의 빈 값 상태를 분석한다.

    Args:
        df: 엑셀 데이터 DataFrame
        used_variables: 템플릿에서 사용된 변수 목록
        email_column: 이메일 주소가 있는 열 이름

    Returns:
        {
            "total": 전체 행 수,
            "complete": 모든 변수가 채워진 행 수,
            "has_empty": 빈 값이 있는 행 수,
            "no_email": 이메일이 없는 행 수,
            "empty_details": [{"row": 행번호, "email": 이메일, "empty_vars": [빈 변수들]}]
        }
    """
    total = len(df)
    no_email = 0
    complete = 0
    has_empty = 0
    empty_details = []

    # 분석 대상 열 (사용된 변수 중 실제로 존재하는 열)
    check_columns = [v for v in used_variables if v in df.columns]

    for idx, row in df.iterrows():
        email = str(row.get(email_column, "")).strip()

        # 이메일이 비어있는 경우
        if not email or email == "nan":
            no_email += 1
            continue

        # 사용된 변수 중 빈 값이 있는지 확인
        empty_vars = []
        for col in check_columns:
            val = row.get(col, "")
            if pd.isna(val) or str(val).strip() == "":
                empty_vars.append(col)

        if empty_vars:
            has_empty += 1
            empty_details.append({
                "row": idx + 2,  # 엑셀 행 번호 (헤더=1행, 데이터 시작=2행)
                "email": email,
                "empty_vars": empty_vars,
            })
        else:
            complete += 1

    return {
        "total": total,
        "complete": complete,
        "has_empty": has_empty,
        "no_email": no_email,
        "empty_details": empty_details,
    }


def get_row_data(df: pd.DataFrame, row_idx: int) -> dict[str, str]:
    """
    특정 행의 데이터를 {열이름: 값} 딕셔너리로 반환한다.
    NaN 값은 빈 문자열로 변환한다.
    """
    row = df.iloc[row_idx]
    data = {}
    for col in df.columns:
        val = row[col]
        if pd.isna(val):
            data[col] = ""
        else:
            data[col] = str(val).strip()
    return data

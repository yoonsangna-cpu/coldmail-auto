"""
메일 템플릿 변수 치환 엔진
- {변수명} 형태의 변수 감지
- 엑셀 데이터로 변수 치환
- 빈 값 기본값 대체 로직
"""

import re
from typing import Optional


# {변수명} 패턴 매칭 정규식
VARIABLE_PATTERN = re.compile(r"\{([^{}]+)\}")


def extract_variables(text: str) -> list[str]:
    """
    텍스트에서 {변수명} 형태의 변수를 모두 추출한다.

    Args:
        text: 제목 또는 본문 텍스트

    Returns:
        고유한 변수명 리스트 (등장 순서 유지)
    """
    found = VARIABLE_PATTERN.findall(text)
    # 중복 제거하면서 순서 유지
    seen = set()
    result = []
    for var in found:
        var = var.strip()
        if var and var not in seen:
            seen.add(var)
            result.append(var)
    return result


def render_template(
    template: str,
    data: dict[str, str],
    defaults: Optional[dict[str, str]] = None,
) -> str:
    """
    템플릿의 변수를 데이터로 치환한다.

    Args:
        template: {변수명}이 포함된 텍스트
        data: {변수명: 값} 딕셔너리
        defaults: {변수명: 기본값} 딕셔너리 (빈 값일 때 사용)

    Returns:
        변수가 치환된 텍스트
    """
    if defaults is None:
        defaults = {}

    def replacer(match):
        var_name = match.group(1).strip()
        value = data.get(var_name, "")

        # 값이 비어있으면 기본값 사용
        if not value:
            value = defaults.get(var_name, "")

        return value

    return VARIABLE_PATTERN.sub(replacer, template)


def render_email(
    subject_template: str,
    body_template: str,
    data: dict[str, str],
    defaults: Optional[dict[str, str]] = None,
    alt_subject_template: Optional[str] = None,
    alt_body_template: Optional[str] = None,
) -> dict[str, str]:
    """
    한 수신자에 대해 최종 이메일 제목/본문을 생성한다.

    빈 값이 있고 별도 템플릿이 지정된 경우, 대체 템플릿을 사용한다.

    Args:
        subject_template: 메일 제목 템플릿
        body_template: 메일 본문 템플릿
        data: 해당 행의 {변수명: 값} 딕셔너리
        defaults: 빈 값 기본값 딕셔너리
        alt_subject_template: 빈 값 행용 대체 제목 템플릿
        alt_body_template: 빈 값 행용 대체 본문 템플릿

    Returns:
        {"subject": 최종 제목, "body": 최종 본문, "used_alt": 대체 템플릿 사용 여부}
    """
    if defaults is None:
        defaults = {}

    # 사용된 변수 추출
    all_vars = set(extract_variables(subject_template) + extract_variables(body_template))

    # 빈 값이 있는지 확인
    has_empty = any(not data.get(v, "") for v in all_vars if v in data or v in defaults)

    # 빈 값이 있고 별도 템플릿이 있으면 대체 템플릿 사용
    if has_empty and alt_subject_template and alt_body_template:
        subject = render_template(alt_subject_template, data, defaults)
        body = render_template(alt_body_template, data, defaults)
        return {"subject": subject, "body": body, "used_alt": True}

    # 기본 템플릿 사용
    subject = render_template(subject_template, data, defaults)
    body = render_template(body_template, data, defaults)
    return {"subject": subject, "body": body, "used_alt": False}


def get_empty_variables(data: dict[str, str], variables: list[str]) -> list[str]:
    """데이터에서 빈 값인 변수 목록을 반환한다."""
    return [v for v in variables if not data.get(v, "")]

"""샘플 질문 목록. UI에서 빠르게 선택할 수 있도록 제공한다."""

from typing import NamedTuple


class SampleQuestion(NamedTuple):
    category: str
    question: str


SAMPLE_QUESTIONS: list[SampleQuestion] = [
    SampleQuestion("HR",    "연차는 며칠 전까지 신청해야 해?"),
    SampleQuestion("HR",    "연차 잔여일수는 어디서 확인해?"),
    SampleQuestion("재무",  "출장비 정산할 때 영수증은 어떻게 제출해?"),
    SampleQuestion("재무",  "출장비 정산 기한이 언제까지야?"),
    SampleQuestion("복지",  "복지포인트 잔액은 어디서 확인해?"),
    SampleQuestion("복지",  "경조사 지원금 신청 방법 알려줘."),
    SampleQuestion("IT",    "신규 입사자 장비 신청 절차 알려줘."),
    SampleQuestion("IT",    "보안 교육은 언제까지 이수해야 해?"),
    SampleQuestion("정책",  "재택근무 신청은 어떻게 해?"),
    SampleQuestion("정책",  "사내 복장 규정이 어떻게 돼?"),
    # 근거 없는 질문 – 제한 답변 검증용
    SampleQuestion("검증",  "회사 주식 상장 일정이 언제야?"),
]


def questions_by_category() -> dict[str, list[str]]:
    result: dict[str, list[str]] = {}
    for sq in SAMPLE_QUESTIONS:
        result.setdefault(sq.category, []).append(sq.question)
    return result

from __future__ import annotations

"""
RAG 파이프라인.
query() 하나로 검색 → 컨텍스트 조립 → LLM 답변 생성까지 처리한다.
"""

import logging
from collections.abc import Generator
from typing import Any

from src.config import settings
from src.llm import chat
from src.vector_store import search

logger = logging.getLogger(__name__)

# 근거가 너무 없을 때 기준 (score < threshold 이면 근거 부족으로 판단)
_MIN_SCORE_THRESHOLD = 0.30
# context에 넣을 최대 글자 수 (너무 길면 LLM context window 초과)
_MAX_CONTEXT_CHARS = 4000


def query(
    user_question: str,
    top_k: int | None = None,
    stream: bool = False,
) -> dict[str, Any]:
    """
    user_question: 사용자 자연어 질문
    top_k:         검색 결과 수 (None이면 settings 기본값)
    stream:        True이면 answer가 Generator

    반환값:
    {
        "answer":  str | Generator,   # LLM 생성 답변
        "sources": list[dict],        # 출처 정보
        "grounded": bool,             # 근거 있음 여부
    }
    """
    top_k = top_k or settings.top_k
    hits = search(user_question, top_k=top_k)

    # 근거 없음 판단
    strong_hits = [h for h in hits if h["score"] >= _MIN_SCORE_THRESHOLD]
    if not strong_hits:
        no_context_answer = "제공된 문서에서 확인할 수 없습니다. 관련 문서가 인덱싱되어 있는지 확인해 주세요."
        return {
            "answer":   no_context_answer,
            "sources":  hits,   # 약한 hit도 UI에 표시해 참고하도록
            "grounded": False,
        }

    context = _build_context(strong_hits)
    answer  = chat(user_question, context, stream=stream)

    return {
        "answer":   answer,
        "sources":  strong_hits,
        "grounded": True,
    }


def _build_context(hits: list[dict[str, Any]]) -> str:
    """
    검색 결과 → LLM 프롬프트에 넣을 컨텍스트 문자열.
    [1], [2] 번호를 붙여 LLM이 출처 번호를 답변에 사용할 수 있게 한다.
    """
    parts = []
    total_chars = 0

    for i, hit in enumerate(hits, start=1):
        header  = f"[{i}] 파일: {hit['filename']} | {hit['location']} | 카테고리: {hit['category']}"
        snippet = hit["snippet"]

        entry = f"{header}\n{snippet}"
        if total_chars + len(entry) > _MAX_CONTEXT_CHARS:
            break
        parts.append(entry)
        total_chars += len(entry)

    return "\n\n".join(parts)

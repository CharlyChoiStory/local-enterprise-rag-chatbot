from __future__ import annotations

"""
로컬 LLM 어댑터.
Ollama chat API만 사용한다. 외부 AI SDK 없음.
"""

import logging
from collections.abc import Generator
from typing import Any

import ollama

from src.config import settings

logger = logging.getLogger(__name__)

_SYSTEM_PROMPT = """\
너는 회사 내부 문서 기반 업무지원 챗봇이다.
아래 제공된 문서 근거만 사용해서 한국어로 답변하라.
근거에 없는 내용은 추측하지 말고 "제공된 문서에서 확인할 수 없습니다"라고 답하라.
답변에는 [1], [2] 같은 출처 번호를 붙여라.
연차, 출장비, 복지포인트, 회사 내규 질문은 실무자가 바로 이해할 수 있게 간결하게 답하라.\
"""


def chat(
    user_message: str,
    context: str,
    model: str | None = None,
    stream: bool = False,
) -> str | Generator[str, None, None]:
    """
    user_message: 사용자 질문
    context:      RAG로 검색한 문서 조각들을 이어 붙인 문자열
    stream:       True이면 텍스트 청크를 yield하는 Generator 반환
    """
    model = model or settings.ollama_chat_model

    user_content = f"[참고 문서]\n{context}\n\n[질문]\n{user_message}"

    messages: list[dict[str, Any]] = [
        {"role": "system",    "content": _SYSTEM_PROMPT},
        {"role": "user",      "content": user_content},
    ]

    if stream:
        return _stream_chat(model, messages)
    else:
        return _blocking_chat(model, messages)


def _blocking_chat(model: str, messages: list[dict]) -> str:
    try:
        resp = ollama.chat(model=model, messages=messages)
        return resp.message.content.strip()
    except Exception as exc:
        logger.error("LLM chat 실패: %s", exc)
        raise


def _stream_chat(model: str, messages: list[dict]) -> Generator[str, None, None]:
    try:
        for chunk in ollama.chat(model=model, messages=messages, stream=True):
            token = chunk.message.content
            if token:
                yield token
    except Exception as exc:
        logger.error("LLM stream 실패: %s", exc)
        raise

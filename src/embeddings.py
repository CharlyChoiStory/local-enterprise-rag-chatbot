from __future__ import annotations

"""
로컬 임베딩 어댑터.
Ollama embedding API만 사용한다. 외부 AI SDK 없음.
"""

import logging
from typing import Any

import ollama

from src.config import settings

logger = logging.getLogger(__name__)

# Ollama에 한 번에 보낼 최대 텍스트 수 (메모리 안전)
_BATCH_SIZE = 32


def embed_texts(texts: list[str], model: str | None = None) -> list[list[float]]:
    """
    텍스트 목록을 임베딩 벡터 목록으로 변환한다.
    Ollama embed API를 배치로 호출한다.
    """
    model = model or settings.ollama_embedding_model
    vectors: list[list[float]] = []

    for i in range(0, len(texts), _BATCH_SIZE):
        batch = texts[i : i + _BATCH_SIZE]
        try:
            resp: Any = ollama.embed(model=model, input=batch)
            # ollama>=0.2: resp.embeddings 는 list[list[float]]
            vectors.extend(resp.embeddings)
        except Exception as exc:
            logger.error("embed_texts 실패 (batch %d): %s", i // _BATCH_SIZE, exc)
            raise

    return vectors


def embed_query(query: str, model: str | None = None) -> list[float]:
    """단일 쿼리를 임베딩 벡터로 변환한다."""
    model = model or settings.ollama_embedding_model
    try:
        resp: Any = ollama.embed(model=model, input=[query])
        return resp.embeddings[0]
    except Exception as exc:
        logger.error("embed_query 실패: %s", exc)
        raise

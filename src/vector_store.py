from __future__ import annotations

"""
Chroma 로컬 벡터 스토어.
upsert_chunks()  : chunk 목록을 임베딩하고 저장한다.
search()         : 쿼리와 가까운 chunk top-k를 반환한다.
collection_info(): 현재 collection 상태를 반환한다.
"""

import logging
from typing import Any

import chromadb
from chromadb.config import Settings as ChromaSettings

from src.config import settings
from src.embeddings import embed_query, embed_texts

logger = logging.getLogger(__name__)

# Chroma는 metadata value로 None을 허용하지 않으므로 빈 문자열로 대체
_NONE_PLACEHOLDER = ""

# 한 번에 upsert할 chunk 수 (Chroma 권장)
_UPSERT_BATCH = 100


def _get_collection() -> chromadb.Collection:
    client = chromadb.PersistentClient(
        path=str(settings.chroma_persist_dir),
        settings=ChromaSettings(anonymized_telemetry=False),
    )
    return client.get_or_create_collection(
        name=settings.chroma_collection_name,
        metadata={"hnsw:space": "cosine"},
    )


def upsert_chunks(chunks: list[dict[str, Any]]) -> int:
    """
    chunk 목록을 임베딩해 Chroma에 upsert한다.
    이미 존재하는 chunk_id는 덮어쓴다.
    반환값: 저장된 chunk 수
    """
    if not chunks:
        return 0

    collection = _get_collection()
    stored = 0

    for i in range(0, len(chunks), _UPSERT_BATCH):
        batch = chunks[i : i + _UPSERT_BATCH]
        texts      = [c["text"] for c in batch]
        ids        = [c["chunk_id"] for c in batch]
        metadatas  = [_safe_metadata(c) for c in batch]

        try:
            embeddings = embed_texts(texts)
            collection.upsert(
                ids=ids,
                embeddings=embeddings,
                documents=texts,
                metadatas=metadatas,
            )
            stored += len(batch)
            logger.info("upsert %d/%d chunks", stored, len(chunks))
        except Exception as exc:
            logger.error("upsert 실패 (batch %d): %s", i // _UPSERT_BATCH, exc)
            raise

    return stored


def search(query: str, top_k: int | None = None) -> list[dict[str, Any]]:
    """
    쿼리와 코사인 유사도가 높은 chunk top-k를 반환한다.

    반환값:
    [
      {
        "chunk_id":    str,
        "document_id": str,
        "filename":    str,
        "file_type":   str,
        "category":    str,
        "location":    str,
        "chunk_index": int,
        "score":       float,   # 0~1, 높을수록 유사
        "snippet":     str,     # 앞 300자
      },
      ...
    ]
    """
    top_k = top_k or settings.top_k
    collection = _get_collection()

    if collection.count() == 0:
        return []

    query_vec = embed_query(query)
    results = collection.query(
        query_embeddings=[query_vec],
        n_results=min(top_k, collection.count()),
        include=["documents", "metadatas", "distances"],
    )

    hits = []
    for doc, meta, dist in zip(
        results["documents"][0],
        results["metadatas"][0],
        results["distances"][0],
    ):
        # cosine distance → similarity
        score = max(0.0, 1.0 - dist)
        hits.append({
            "chunk_id":    meta.get("chunk_id", ""),
            "document_id": meta.get("document_id", ""),
            "filename":    meta.get("filename", ""),
            "file_type":   meta.get("file_type", ""),
            "category":    meta.get("category", ""),
            "location":    meta.get("location", ""),
            "chunk_index": int(meta.get("chunk_index", 0)),
            "score":       round(score, 4),
            "snippet":     doc[:300],
        })

    return hits


def collection_info() -> dict[str, Any]:
    """현재 collection의 문서 수를 반환한다."""
    try:
        col = _get_collection()
        return {"count": col.count(), "name": col.name}
    except Exception as exc:
        return {"count": 0, "name": settings.chroma_collection_name, "error": str(exc)}


def _safe_metadata(chunk: dict[str, Any]) -> dict[str, Any]:
    """Chroma metadata는 str/int/float/bool만 허용 – None 제거."""
    allowed = {
        "chunk_id", "document_id", "filename", "file_path",
        "file_type", "category", "location", "chunk_index",
    }
    meta = {}
    for k in allowed:
        v = chunk.get(k, _NONE_PLACEHOLDER)
        if v is None:
            v = _NONE_PLACEHOLDER
        meta[k] = v
    return meta

"""
텍스트 청킹.
parse_document() 반환값을 받아 검색 가능한 chunk list로 변환한다.
"""

from typing import Any

# 너무 짧은 chunk는 노이즈 – 기본 50자 미만 제외
_MIN_CHUNK_LEN = 50


def chunk_document(
    parsed: dict[str, Any],
    chunk_size: int = 1600,
    overlap: int = 250,
) -> list[dict[str, Any]]:
    """
    parsed: parse_document() 반환값
    반환값:
    [
      {
        "chunk_id":    str,   # "{document_id}_{chunk_index}"
        "document_id": str,
        "filename":    str,
        "file_path":   str,
        "file_type":   str,
        "category":    str,
        "location":    str,   # 첫 번째 섹션의 location
        "chunk_index": int,
        "text":        str,
      },
      ...
    ]
    """
    if parsed.get("error") or not parsed.get("sections"):
        return []

    doc_id   = parsed["document_id"]
    filename = parsed["filename"]
    file_path = parsed["file_path"]
    file_type = parsed["file_type"]
    category  = parsed["category"]

    chunks: list[dict[str, Any]] = []
    chunk_index = 0

    for section in parsed["sections"]:
        location = section["location"]
        text     = section["text"]

        # 섹션 하나가 chunk_size보다 짧으면 그냥 통째로
        if len(text) <= chunk_size:
            if len(text) >= _MIN_CHUNK_LEN:
                chunks.append(_make_chunk(
                    doc_id, filename, file_path, file_type, category,
                    location, chunk_index, text,
                ))
                chunk_index += 1
            continue

        # 슬라이딩 윈도우
        start = 0
        while start < len(text):
            end  = min(start + chunk_size, len(text))
            part = text[start:end]
            if len(part) >= _MIN_CHUNK_LEN:
                chunks.append(_make_chunk(
                    doc_id, filename, file_path, file_type, category,
                    location, chunk_index, part,
                ))
                chunk_index += 1
            if end >= len(text):
                break
            start = end - overlap  # overlap 만큼 되돌려 연속성 확보

    return chunks


def chunk_documents(
    parsed_list: list[dict[str, Any]],
    chunk_size: int = 1600,
    overlap: int = 250,
) -> list[dict[str, Any]]:
    """여러 문서를 한 번에 청킹한다."""
    result = []
    for parsed in parsed_list:
        result.extend(chunk_document(parsed, chunk_size, overlap))
    return result


def _make_chunk(
    doc_id: str,
    filename: str,
    file_path: str,
    file_type: str,
    category: str,
    location: str,
    chunk_index: int,
    text: str,
) -> dict[str, Any]:
    return {
        "chunk_id":    f"{doc_id}_{chunk_index}",
        "document_id": doc_id,
        "filename":    filename,
        "file_path":   file_path,
        "file_type":   file_type,
        "category":    category,
        "location":    location,
        "chunk_index": chunk_index,
        "text":        text,
    }

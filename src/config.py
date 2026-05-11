from __future__ import annotations

"""
환경 설정 로딩.
.env 파일 또는 환경변수에서 값을 읽어 pydantic-settings로 검증합니다.
외부 AI API key 항목은 없습니다.
"""

from pathlib import Path

from pydantic import field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

# 프로젝트 루트 (src/ 한 단계 위)
_PROJECT_ROOT = Path(__file__).parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=_PROJECT_ROOT / ".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # Ollama
    ollama_base_url: str = "http://localhost:11434"
    ollama_chat_model: str = "llama3.1:8b"
    ollama_embedding_model: str = "bge-m3"

    # 경로
    source_docs_dir: Path = _PROJECT_ROOT / "data" / "source_docs"
    chroma_persist_dir: Path = _PROJECT_ROOT / "data" / "chroma"

    # Chroma
    chroma_collection_name: str = "company_documents"

    # RAG 파라미터
    top_k: int = 5
    chunk_size: int = 1600
    chunk_overlap: int = 250

    @field_validator("source_docs_dir", "chroma_persist_dir", mode="before")
    @classmethod
    def resolve_path(cls, v: str | Path) -> Path:
        p = Path(v)
        # 상대경로는 프로젝트 루트 기준으로 변환
        if not p.is_absolute():
            p = _PROJECT_ROOT / p
        p.mkdir(parents=True, exist_ok=True)
        return p


# 싱글턴
settings = Settings()

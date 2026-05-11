from __future__ import annotations

"""
로컬 파일 기반 기업 내 반복 질문 자동응답 RAG 챗봇 (POC)
실행: streamlit run app.py
"""

import streamlit as st
from pathlib import Path

from src.config import settings
from src.sample_questions import SAMPLE_QUESTIONS, questions_by_category

# ── 페이지 기본 설정 ──────────────────────────────────────────
st.set_page_config(
    page_title="사내 문서 챗봇",
    page_icon="💬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 전역 CSS (Slack 스타일) ───────────────────────────────────
st.markdown(
    """
    <style>
    .stApp { background-color: #1a1d21; color: #d1d2d3; }
    [data-testid="stSidebar"] { background-color: #19171d; }
    [data-testid="stSidebar"] * { color: #cfc3de !important; }

    .msg-user {
        background: #1164a3; color: #fff;
        border-radius: 8px 8px 2px 8px;
        padding: 10px 14px; margin: 6px 0 6px 60px;
        font-size: 0.95rem; line-height: 1.5;
    }
    .msg-bot {
        background: #2b2d30; color: #d1d2d3;
        border-radius: 8px 8px 8px 2px;
        padding: 10px 14px; margin: 6px 60px 6px 0;
        font-size: 0.95rem; line-height: 1.5;
    }
    .msg-name {
        font-size: 0.75rem; font-weight: 600;
        margin-bottom: 3px; opacity: 0.7;
    }
    .source-box {
        background: #222529; border-left: 3px solid #1164a3;
        border-radius: 4px; padding: 8px 12px; margin-top: 6px;
        font-size: 0.8rem; color: #a8a9aa;
    }
    .stTextInput > div > div > input {
        background-color: #222529; color: #d1d2d3;
        border: 1px solid #3d3f42; border-radius: 8px;
    }
    .stButton > button {
        background-color: #1164a3; color: #fff;
        border: none; border-radius: 6px;
    }
    .stButton > button:hover { background-color: #0d4f87; }
    hr { border-color: #3d3f42; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── 세션 상태 초기화 ──────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []
if "indexed" not in st.session_state:
    st.session_state.indexed = False
if "index_count" not in st.session_state:
    st.session_state.index_count = 0
if "pending_question" not in st.session_state:
    st.session_state.pending_question = ""


# ── 헬퍼 함수 ────────────────────────────────────────────────
def render_message(msg: dict) -> None:
    if msg["role"] == "user":
        st.markdown(
            f'<div class="msg-name" style="text-align:right;color:#7ecff5;">나</div>'
            f'<div class="msg-user">{msg["content"]}</div>',
            unsafe_allow_html=True,
        )
    else:
        sources_html = ""
        if msg.get("sources"):
            items = "".join(
                f"<div>[{i+1}] 📄 <b>{s.get('filename','')}</b> "
                f"({s.get('file_type','').upper()}) · {s.get('category','')} · {s.get('location','')} "
                f"· 유사도 {s.get('score', 0):.2f}<br>"
                f"<span style='opacity:.8'>{s.get('snippet','')[:120]}…</span></div>"
                for i, s in enumerate(msg["sources"])
            )
            sources_html = f'<div class="source-box">📎 <b>출처</b><br>{items}</div>'

        grounded = msg.get("grounded", True)
        badge = "" if grounded else ' <span style="color:#e87040;font-size:0.75rem;">⚠ 근거 부족</span>'
        st.markdown(
            f'<div class="msg-name" style="color:#7ecff5;">🤖 사내 문서 챗봇{badge}</div>'
            f'<div class="msg-bot">{msg["content"]}</div>'
            f'{sources_html}',
            unsafe_allow_html=True,
        )


def run_indexing(docs_dir: str, chunk_size: int, chunk_overlap: int) -> tuple[int, list[str]]:
    """파싱 → 청킹 → 임베딩 → Chroma upsert. (chunk 수, 오류 파일 목록) 반환."""
    from src.parsers import scan_documents
    from src.chunker import chunk_documents
    from src.vector_store import upsert_chunks

    parsed_list = scan_documents(docs_dir)
    errors = [p["filename"] for p in parsed_list if p.get("error")]
    valid  = [p for p in parsed_list if not p.get("error")]

    chunks = chunk_documents(valid, chunk_size=chunk_size, overlap=chunk_overlap)
    stored = upsert_chunks(chunks)
    return stored, errors


def run_rag(question: str, top_k: int) -> dict:
    from src.rag import query as rag_query
    return rag_query(question, top_k=top_k, stream=False)


# ── 사이드바 ──────────────────────────────────────────────────
with st.sidebar:
    st.title("💬 사내 문서 챗봇")
    st.caption("로컬 파일 기반 · 외부 AI 없음")
    st.divider()

    st.subheader("⚙️ 설정")
    docs_dir = st.text_input(
        "문서 폴더",
        value=str(settings.source_docs_dir),
        help="PDF/TXT/DOCX/PPTX가 있는 폴더 경로",
    )
    chat_model = st.text_input(
        "Chat 모델", value=settings.ollama_chat_model,
        help="Ollama에서 pull한 채팅 모델",
    )
    embed_model = st.text_input(
        "Embedding 모델", value=settings.ollama_embedding_model,
        help="Ollama에서 pull한 임베딩 모델",
    )
    top_k = st.slider("검색 결과 수 (Top-K)", 1, 10, settings.top_k)

    st.divider()
    st.subheader("📂 인덱싱")

    docs_path = Path(docs_dir)
    file_count = (
        sum(1 for ext in ("*.pdf", "*.txt", "*.docx", "*.pptx") for _ in docs_path.rglob(ext))
        if docs_path.exists() else 0
    )
    st.caption(f"감지된 파일: **{file_count}개**")

    if st.button("🔄 인덱싱 실행", use_container_width=True):
        with st.spinner("파싱 → 청킹 → 임베딩 → 저장 중…"):
            try:
                stored, errs = run_indexing(
                    docs_dir,
                    chunk_size=settings.chunk_size,
                    chunk_overlap=settings.chunk_overlap,
                )
                st.session_state.indexed = True
                st.session_state.index_count = stored
                if errs:
                    st.warning(f"파싱 실패 파일 {len(errs)}개: {', '.join(errs)}")
                st.success(f"완료! {stored}개 chunk 저장됨")
            except Exception as exc:
                st.error(f"인덱싱 오류: {exc}")

    if st.session_state.indexed:
        st.success(f"인덱싱 완료 ({st.session_state.index_count}개 chunk)")
    else:
        st.warning("아직 인덱싱되지 않았습니다.")

    # 컬렉션 현황
    try:
        from src.vector_store import collection_info
        info = collection_info()
        st.caption(f"Chroma: {info.get('count', 0)}개 chunk 저장됨")
    except Exception:
        pass

    st.divider()

    # 샘플 질문 빠른 선택
    st.subheader("💡 샘플 질문")
    for sq in SAMPLE_QUESTIONS:
        label = f"[{sq.category}] {sq.question}"
        if st.button(label, use_container_width=True, key=f"sq_{sq.question}"):
            st.session_state.pending_question = sq.question

    st.divider()
    st.caption(f"Ollama: `{settings.ollama_base_url}`")
    st.caption("v1.0 – POC 완성")


# ── 메인 채팅 영역 ────────────────────────────────────────────
st.markdown("### 💬 채팅")
st.caption("사내 문서 기반으로 질문에 답변합니다. 인덱싱 후 질문하세요.")
st.divider()

chat_area = st.container()
with chat_area:
    if not st.session_state.messages:
        st.markdown(
            '<div class="msg-bot" style="opacity:.6;">'
            "안녕하세요! 연차, 출장비, 복지포인트, IT 장비 신청 등 사내 문서 관련 질문을 입력하세요.<br>"
            "왼쪽 사이드바에서 <b>인덱싱 실행</b>을 먼저 누르거나 샘플 질문을 선택하세요."
            "</div>",
            unsafe_allow_html=True,
        )
    for msg in st.session_state.messages:
        render_message(msg)

st.divider()

# ── 입력창 ────────────────────────────────────────────────────
with st.form(key="chat_form", clear_on_submit=True):
    col_input, col_send = st.columns([9, 1])
    with col_input:
        default_q = st.session_state.pop("pending_question", "") if "pending_question" in st.session_state else ""
        user_input = st.text_input(
            label="질문",
            value=default_q,
            placeholder="예) 연차는 며칠 전까지 신청해야 해?",
            label_visibility="collapsed",
        )
    with col_send:
        submitted = st.form_submit_button("전송", use_container_width=True)

# 샘플 질문 버튼 클릭 시 자동 전송 처리
if st.session_state.get("pending_question"):
    auto_q = st.session_state.pending_question
    st.session_state.pending_question = ""
    submitted = True
    user_input = auto_q

if submitted and user_input and user_input.strip():
    question = user_input.strip()
    st.session_state.messages.append({"role": "user", "content": question, "sources": []})

    if not st.session_state.indexed:
        reply = {"answer": "⚠️ 먼저 왼쪽 사이드바에서 **인덱싱 실행**을 눌러주세요.", "sources": [], "grounded": False}
    else:
        with st.spinner("문서 검색 중…"):
            try:
                reply = run_rag(question, top_k=top_k)
            except Exception as exc:
                reply = {"answer": f"❌ 오류: {exc}", "sources": [], "grounded": False}

    st.session_state.messages.append({
        "role":     "bot",
        "content":  reply["answer"],
        "sources":  reply.get("sources", []),
        "grounded": reply.get("grounded", True),
    })
    st.rerun()

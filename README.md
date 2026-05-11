# 로컬 파일 기반 기업 내 반복 질문 자동응답 RAG 챗봇 (POC)

퍼블릭 AI API 없이, 로컬 PC에서만 동작하는 사내 문서 기반 챗봇 POC입니다.  
연차 신청, 출장비 정산, 복지포인트, IT 장비 신청 등 반복 질문에 자동 답변합니다.

---

## 기술 스택

| 구성 요소 | 도구 |
|---|---|
| UI | Streamlit |
| LLM | Ollama (`llama3.1:8b` / `qwen2.5:7b`) |
| Embedding | Ollama (`bge-m3` / `nomic-embed-text`) |
| Vector DB | Chroma (로컬 persistent) |
| 문서 파서 | PyMuPDF, python-docx, python-pptx, charset-normalizer |
| 설정 | python-dotenv, pydantic-settings |

> 외부 AI API(OpenAI, Anthropic, Google 등)를 사용하지 않습니다.

---

## 사전 준비

### 1. Ollama 설치

```bash
# macOS
brew install ollama

# 또는 공식 설치 스크립트
curl -fsSL https://ollama.com/install.sh | sh
```

### 2. 모델 다운로드

```bash
# 채팅용 LLM (둘 중 하나 선택, llama3.1:8b 권장)
ollama pull llama3.1:8b
ollama pull qwen2.5:7b

# 임베딩 모델 (둘 중 하나 선택, bge-m3 권장)
ollama pull bge-m3
ollama pull nomic-embed-text
```

### 3. Ollama 서버 실행 확인

```bash
ollama serve
# 또는 백그라운드로 이미 실행 중인지 확인
curl http://localhost:11434
```

---

## 설치

```bash
# 1. 저장소 클론 또는 폴더로 이동
cd local-rag-poc

# 2. Python 가상환경 생성 (Python 3.11 이상 권장)
python3 -m venv .venv
source .venv/bin/activate      # macOS/Linux
# .venv\Scripts\activate       # Windows

# 3. 의존성 설치
pip install -r requirements.txt

# 4. 환경 설정 파일 복사
cp .env.example .env
# 필요 시 .env 안의 모델명/경로 수정
```

---

## 실행

```bash
streamlit run app.py
```

브라우저에서 `http://localhost:8501` 접속

---

## 문서 추가 방법

`data/source_docs/` 하위 폴더에 파일을 넣고 앱에서 **인덱싱 실행** 버튼을 누릅니다.

```
data/source_docs/
  hr/         ← 인사/연차/복지 관련 문서
  finance/    ← 출장비/경비 관련 문서
  welfare/    ← 복지포인트/경조사 문서
  it/         ← IT 장비/보안 관련 문서
  policy/     ← 회사 내규/업무 매뉴얼
```

지원 형식: `.pdf`, `.txt`, `.docx`, `.pptx`

---

## 폴더 구조

```
local-rag-poc/
  app.py                  # Streamlit 메인 앱
  requirements.txt
  .env.example
  README.md
  data/
    source_docs/          # 원본 문서 (로컬 전용)
    chroma/               # Chroma vector DB 저장 위치
  src/
    config.py             # 환경 설정 로딩
    parsers.py            # PDF/TXT/DOCX/PPTX 파서
    chunker.py            # 텍스트 chunking
    embeddings.py         # 로컬 embedding 어댑터
    vector_store.py       # Chroma 연동
    llm.py                # 로컬 LLM 어댑터
    rag.py                # RAG 파이프라인
    sample_questions.py   # 샘플 질문 목록
```

---

## 보안 원칙

- 외부 AI API 호출 없음
- 원본 문서는 `data/source_docs/`에만 보관
- Vector DB는 `data/chroma/`에 로컬 저장
- `.env`에 외부 AI API key 항목 없음

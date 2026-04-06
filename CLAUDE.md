# HWP-AI AutoFill 프로젝트

## 프로젝트 개요
웹 채팅 인터페이스에서 LLM과 대화하면서 정부지원사업 계획서 HWP/HWPX
템플릿의 표 구조를 인식하고, 서식을 유지한 채 자동으로 내용을 채워넣는 시스템.

## 시스템 아키텍처
```
[React 웹 채팅 UI] <-WebSocket-> [FastAPI 서버] <-COM-> [한/글 프로그램]
|                                |
파일 업로드/다운로드           [LLM 라우터]
실시간 진행 상황                /   |   \
문서 미리보기           Claude  GPT  DeepSeek  로컬LLM
|
[MCP 서버] <-> Claude Desktop / Cursor
```

## 접근 경로 (3가지)
1. **웹 UI**: `hwp-ai serve` → http://localhost:8080
2. **CLI**: `hwp-ai generate template.hwp -p 예비창업패키지 -c 회사명`
3. **MCP**: Claude Desktop/Cursor에서 직접 도구 호출

## 기술 전략
- 주력: pyhwpx COM 자동화 — 한/글 프로그램 직접 제어, 서식 100% 보존
- 보조: python-hwpx — HWPX XML 직접 조작 (COM 불가 시 폴백)
- 멀티 LLM: Claude, GPT-4o, DeepSeek, Qwen3, 로컬 모델 등 자유 전환
- 대화형: 웹 채팅으로 실시간 문서 편집, 수정 요청, 되돌리기 가능
- 셀 단위 생성: 표 셀별로 독립 프롬프트 호출 (서식 유지 핵심)

## 디렉토리 구조
```
hwphelper/
├── CLAUDE.md
├── pyproject.toml
├── .env.example / .env
├── src/                        # 백엔드 (Python)
│   ├── hwp_engine/             # HWP COM 자동화 엔진
│   │   ├── com_controller.py       # COM 컨트롤러 (컨텍스트 매니저)
│   │   ├── table_reader.py         # 표 구조 읽기 + 데이터클래스
│   │   ├── cell_classifier.py      # 셀 분류 (LABEL/EMPTY/PREFILLED/PLACEHOLDER)
│   │   ├── cell_writer.py          # 셀 텍스트 삽입 (서식 보존)
│   │   ├── field_manager.py        # 누름틀 필드 관리
│   │   ├── document_manager.py     # 세션 관리 + 스냅샷 undo/redo
│   │   ├── schema_generator.py     # JSON 스키마 생성 (LLM용)
│   │   └── hwpx_fallback.py        # python-hwpx 폴백
│   ├── ai/                     # LLM 연동
│   │   ├── llm_router.py          # 멀티 LLM 라우터 (Anthropic/OpenAI/호환)
│   │   ├── chat_agent.py          # 대화형 에이전트 (도구 호출 루프)
│   │   ├── cell_generator.py      # 셀 단위 콘텐츠 생성 (동시성)
│   │   ├── prompt_builder.py      # 프롬프트 템플릿
│   │   ├── tool_definitions.py    # LLM 도구 10개 정의
│   │   └── rag_engine.py          # ChromaDB RAG 파이프라인
│   ├── validator/              # 서식 검증
│   │   └── format_checker.py      # 과제별 규정 검사 + 자동 교정
│   ├── api/                    # FastAPI
│   │   ├── routes.py              # REST API 11개 엔드포인트
│   │   ├── websocket_handler.py   # WebSocket 채팅 핸들러
│   │   └── schemas.py             # Pydantic 모델
│   ├── utils/
│   │   └── debug_utils.py         # 디버깅 유틸리티
│   ├── server.py               # FastAPI 앱 (lifespan)
│   ├── cli.py                  # Typer CLI (8개 명령어)
│   └── mcp_server.py           # MCP 서버 (8개 도구 + 2개 리소스)
├── frontend/                   # React (Vite + Tailwind)
│   └── src/
│       ├── App.tsx                 # 3단 레이아웃 (사이드바/채팅/헤더)
│       ├── components/             # 9개 컴포넌트
│       ├── hooks/                  # useWebSocket, useChat
│       ├── lib/api.ts              # REST API 클라이언트
│       └── types/index.ts          # TypeScript 타입
├── config/
│   ├── llm_config.yaml         # 7개 LLM 모델 설정
│   └── format_rules.yaml       # 5개 정부과제 서식 규정
├── scripts/
│   ├── setup_check.py          # 환경 진단 (19개 항목)
│   ├── build.py                # 프로덕션 빌드
│   └── quick_fill.py           # 원라이너 자동 채우기
└── tests/                      # pytest (185+ 테스트)
```

## Windows 환경 설정

### 1. 한/글 COM 자동화 설정
1. 한/글 2022 이상 설치 확인
2. 한/글 실행 → 도구 → 환경 설정 → 기타
   → "파일 스크립트 매크로 실행 허용" 체크
3. COM 연결 테스트:
   ```
   python -c "import win32com.client; h=win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject'); print('OK'); h.Quit()"
   ```

### 2. 설치 및 실행
```bash
# 1. 의존성 설치
uv sync
cd frontend && npm install && npm run build && cd ..

# 2. .env 설정 (최소 하나의 LLM API 키)
cp .env.example .env
# .env 파일 편집하여 API 키 입력

# 3. 환경 진단
uv run python scripts/setup_check.py

# 4. 실행 (3가지 방식)
hwp-ai serve                    # 웹 UI
hwp-ai generate t.hwp -p 과제 -c 회사  # CLI
python -m src.mcp_server        # MCP (Claude Desktop)
```

### 3. .env 파일 (최소 하나의 API 키만 있으면 동작)
```env
ANTHROPIC_API_KEY=sk-ant-...         # Claude (권장)
OPENAI_API_KEY=sk-...                # GPT-4o (선택)
DEEPSEEK_API_KEY=sk-...              # DeepSeek (선택)
LOCAL_LLM_BASE_URL=http://localhost:8000/v1  # 로컬 (선택)
```

## 주의사항
- COM 정리 필수: `with HwpController() as ctrl:` 또는 try/finally로 hwp.Quit()
- 절대 경로: COM API는 상대 경로 인식 불가 (os.path.abspath 자동 변환)
- 백그라운드 실행: `HwpController(visible=False)`
- 보안 팝업 우회: RegisterModule 자동 호출
- 셀 단위 생성: 표 셀별 독립 LLM 호출 (서식 유지 핵심)
- 누름틀 우선: put_field_text()가 가장 안정적인 텍스트 삽입 방법
- 스냅샷 저장: 문서 수정 전 반드시 되돌리기용 스냅샷
- 프로세스 충돌: 한/글 COM 인스턴스는 세션당 1개만 유지

## 코딩 컨벤션
- Python: 3.11+, type hints, ruff, pytest, structlog
- TypeScript: strict mode, ESLint
- API: RESTful + WebSocket (채팅은 WS 필수)
- COM 객체: 컨텍스트 매니저(with문) 사용
- 비동기: asyncio (백엔드), React hooks (프론트엔드)
- 테스트: COM 의존 테스트는 @pytest.mark.skipif(not HAS_HWP) 로 보호

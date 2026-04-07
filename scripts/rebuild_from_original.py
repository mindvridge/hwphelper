"""원본에서 새로 복사하여 전체 재작성.

원본 표 매핑 (확인 완료):
  표[4]  8열 — 현황 (창업아이템명, 산출물, 직업, 기업명, 팀)
  표[6]  5열 — 개요 요약 (명칭/범주/아이템개요/문제인식/실현가능성/성장전략/팀구성/이미지)
  표[17] 1셀 — 1. 문제인식 본문 (※ 안내문 → 본문으로 교체)
  표[19] 1셀 — 2. 실현가능성 본문
  표[20] 4열 — 개발 일정표
  표[22] 3열 — 1단계 사업비 (get_cell_addr=False)
  표[24] 3열 — 2단계 사업비 (get_cell_addr=False)
  표[26] 1셀 — 3. 성장전략 본문
  표[27] 4열 — 사업화 로드맵
  표[29] 1셀 — 4. 팀구성 본문
  표[30] 5열 — 팀원
  표[31] 5열 — 파트너
"""
import sys
import os
import shutil

sys.stdout.reconfigure(encoding="utf-8")
os.chdir("c:/hwphelper")

src = "[별첨 1] 2026년 예비창업패키지 사업계획서 양식.hwp"
dst = "[별첨 1] 2026년 예비창업패키지 사업계획서_작성본.hwp"
shutil.copy2(src, dst)
print(f"원본 복사: {dst}")

from src.hwp_engine.com_controller import HwpController
from src.hwp_engine.cell_writer import CellWriter

ctrl = HwpController(visible=True)
ctrl.connect()
ctrl.open(dst)
w = CellWriter(ctrl)
hwp = ctrl.hwp


# ===== 표[4]: 현황 (셀 주소 작동) =====
print("\n[표4] 현황")
w.write_cell(4, 0, 3, 'AI 기반 정부과제 문서 자동작성 솔루션 "DocuMind"')
w.write_cell(4, 1, 3, "웹 애플리케이션(1개), AI 엔진 API(1개)")
w.write_cell(4, 2, 3, "소프트웨어 개발자")
w.write_cell(4, 2, 6, "주식회사 도큐마인드")
w.write_cell(4, 5, 1, "CTO")
w.write_cell(4, 5, 2, "AI 엔진 개발")
w.write_cell(4, 5, 4, "컴퓨터공학 석사, NLP 연구 경력 5년")
w.write_cell(4, 5, 7, "완료")
w.write_cell(4, 6, 1, "디자이너")
w.write_cell(4, 6, 2, "UX/UI 설계")
w.write_cell(4, 6, 4, "시각디자인 학사, SaaS 디자인 경력 3년")
w.write_cell(4, 6, 7, "예정('26.9)")
print("  OK")

# ===== 표[17]: 1. 문제인식 본문 (1셀) =====
print("\n[표17] 문제인식 본문")
w.write_cell(17, 0, 0, """1-1. 국내 정부지원사업 시장 현황

국내 정부지원사업은 연간 4,000건 이상 공고되며, 2025년 기준 중소벤처기업부만 약 3.2조원 규모의 예산을 집행하고 있다. 예비창업패키지, TIPS, 창업도약패키지 등 주요 사업의 신청 건수는 매년 15~20% 증가하고 있어, 사업계획서 작성 수요는 지속적으로 확대되고 있다.

1-2. 사업계획서 작성의 핵심 문제점

① 높은 시간·비용 부담: 사업계획서 1건 작성에 평균 2~4주 소요, 전문 컨설팅 건당 200~500만원
② 서식 준수의 어려움: 한/글(HWP) 지정 양식, 표 구조·글꼴·줄간격 등 서식 규정 위반 시 감점/탈락
③ 항목별 작성 가이드 부족: 과제마다 평가 기준 상이, 작성 수준 가이드 미흡
④ 기존 AI 도구 한계: ChatGPT 등 범용 AI는 HWP 표 구조 인식·편집 불가, 서식 보존 불가능

1-3. DocuMind의 필요성

HWP 문서의 표 구조를 자동 인식하고, 서식을 100% 보존하면서 LLM이 항목별 맞춤 콘텐츠를 생성·삽입하는 전문 솔루션이 필요하다. 이를 통해 작성 시간 80% 단축, 서식 오류 완전 제거, 평가 기준 최적화된 사업계획서 작성이 가능하다.""")
print("  OK")

# ===== 표[19]: 2. 실현가능성 본문 (1셀) =====
print("\n[표19] 실현가능성 본문")
w.write_cell(19, 0, 0, """2-1. 핵심 기술 및 개발 계획

DocuMind의 핵심 기술은 세 가지로 구성된다.
① HWP COM 자동화 엔진: pyhwpx 기반 한/글 표·셀 구조 직접 제어, 원본 서식 100% 보존
② 멀티 LLM 라우터: Claude, GPT-4o, DeepSeek 등 항목 특성별 최적 모델 자동 선택
③ RAG 파이프라인: ChromaDB 기반 과제별 공고문·평가기준·우수사례 실시간 반영

2-2. 차별성 및 경쟁력

- HWP 네이티브 지원: 국내 유일 HWP/HWPX 표 구조 직접 조작 솔루션
- 대화형 편집: 웹 채팅에서 자연어 수정 요청 가능
- 스냅샷 undo/redo: 수정 전 자동 백업, 언제든 복원
- 서식 검증 자동화: 과제별 글꼴·줄간격·분량 규정 자동 검사 및 교정

2-3. 정부지원금 사용 계획

1단계(0~4개월): 핵심 엔진 개발 및 MVP 구축
2단계(5~8개월): 베타 서비스 런칭 및 사용자 피드백 반영""")
print("  OK")

# ===== 표[20]: 개발 일정 (4열, 셀 주소 작동) =====
print("\n[표20] 개발 일정")
w.write_cell(20, 1, 1, "AI 엔진 핵심 모듈 개발")
w.write_cell(20, 1, 2, "26.09 ~ 26.11")
w.write_cell(20, 1, 3, "LLM 라우터, 프롬프트 빌더, RAG 파이프라인 구축")
w.write_cell(20, 2, 1, "HWP 자동화 엔진 개발")
w.write_cell(20, 2, 2, "26.09 ~ 26.12")
w.write_cell(20, 2, 3, "COM 컨트롤러, 표 읽기/쓰기, 서식 검증 모듈")
w.write_cell(20, 3, 1, "웹 UI/UX 개발")
w.write_cell(20, 3, 2, "26.11 ~ 27.01")
w.write_cell(20, 3, 3, "React 채팅 인터페이스, 문서 미리보기, 대시보드")
w.write_cell(20, 4, 1, "MVP 출시 및 베타 테스트")
w.write_cell(20, 4, 2, "27.02 ~ 27.04")
w.write_cell(20, 4, 3, "베타 사용자 50명 모집, 피드백 수집 및 반영")
print("  OK")

# ===== 표[22],[24]: 사업비 — find_replace로 별도 처리 =====
# (스크립트 끝나고 fix_budget_tables.py로 처리)

# ===== 표[26]: 3. 성장전략 본문 (1셀) =====
print("\n[표26] 성장전략 본문")
w.write_cell(26, 0, 0, """3-1. 경쟁사 분석

① 전문 컨설팅 업체: 건당 200~500만원, 높은 비용과 2~3주 소요
② 범용 AI(ChatGPT 등): 한/글 서식 지원 불가, 표 구조 인식·편집 불가능
③ 문서 템플릿 서비스: 단순 양식 제공, 내용 생성 기능 없음

DocuMind는 HWP 네이티브 지원 + AI 콘텐츠 생성 + 서식 자동 검증을 결합한 유일한 솔루션으로, 비용 95% 절감, 시간 80% 단축의 가치를 제공한다.

3-2. 비즈니스 모델

① SaaS 구독: Basic(월 9.9만원, 월 3건), Pro(월 29.9만원, 무제한)
② 건별 과금: 비구독 사용자 건당 3만원
③ B2B 라이선스: 대학 창업지원단·BI센터 연간 500만원~

3-3. 사업 로드맵

- 2026 Q4: MVP 출시, 얼리어답터 50명 확보
- 2027 Q1~Q2: 정식 서비스 런칭, 월 유료 사용자 200명
- 2027 Q3~Q4: B2B 진출(대학 5곳), MAU 1,000명
- 2028 상반기: Pre-A 투자유치(5억원)""")
print("  OK")

# ===== 표[27]: 사업화 로드맵 (4열, 셀 주소 작동) =====
print("\n[표27] 사업화 로드맵")
w.write_cell(27, 1, 1, "클로즈드 베타 운영")
w.write_cell(27, 1, 2, "27년 Q1")
w.write_cell(27, 1, 3, "얼리어답터 50명 대상 무료 베타, 피드백 수집")
w.write_cell(27, 2, 1, "정식 서비스 런칭")
w.write_cell(27, 2, 2, "27년 Q2")
w.write_cell(27, 2, 3, "유료 전환, 마케팅 캠페인 시작")
w.write_cell(27, 3, 1, "B2B 영업 시작")
w.write_cell(27, 3, 2, "27년 Q3")
w.write_cell(27, 3, 3, "대학 창업지원단 5곳 파일럿 계약")
w.write_cell(27, 4, 1, "Pre-A 투자유치")
w.write_cell(27, 4, 2, "28년 Q1")
w.write_cell(27, 4, 3, "VC 미팅, IR 자료 준비, 5억원 유치 목표")
print("  OK")

# ===== 표[29]: 4. 팀구성 본문 (1셀) =====
print("\n[표29] 팀구성 본문")
w.write_cell(29, 0, 0, """4-1. 대표자 역량

대표자는 풀스택 소프트웨어 개발 10년 경력을 보유하고 있으며, Python, React, FastAPI 등 핵심 기술 스택에 깊은 전문성을 갖추고 있다. 정부과제 사업계획서를 직접 20건 이상 작성·수행한 경험이 있어, 작성의 pain point와 평가 기준을 정확히 이해하고 있다.

4-2. 팀원 구성 및 역할

① CTO: 컴퓨터공학 석사, NLP/LLM 연구 경력 5년. AI 엔진 아키텍처 설계 및 최적화
② UX 디자이너: 시각디자인 학사, SaaS 제품 디자인 경력 3년. 채팅 UI/UX 설계

4-3. 업무 파트너

① 한국정보화진흥원: 전문 멘토링 및 네트워킹 지원
② AWS Korea: 클라우드 인프라 크레딧 지원""")
print("  OK")

# ===== 표[30]: 팀원 (5열, 셀 주소 작동) =====
print("\n[표30] 팀원")
w.write_cell(30, 1, 1, "CTO")
w.write_cell(30, 1, 2, "AI 엔진 개발 총괄")
w.write_cell(30, 1, 3, "컴퓨터공학 석사, NLP 연구 경력 5년, 논문 3편")
w.write_cell(30, 1, 4, "완료('26.09)")
w.write_cell(30, 2, 1, "디자이너")
w.write_cell(30, 2, 2, "UX/UI 설계")
w.write_cell(30, 2, 3, "시각디자인 학사, SaaS 디자인 경력 3년")
w.write_cell(30, 2, 4, "예정('26.12)")
print("  OK")

# ===== 표[31]: 파트너 (5열, 셀 주소 작동) =====
print("\n[표31] 파트너")
w.write_cell(31, 1, 1, "한국정보화진흥원")
w.write_cell(31, 1, 2, "창업 멘토링, 네트워킹")
w.write_cell(31, 1, 3, "전문 멘토 매칭, 투자자 연결")
w.write_cell(31, 1, 4, "26.09")
w.write_cell(31, 2, 1, "AWS Korea")
w.write_cell(31, 2, 2, "클라우드 인프라 운영")
w.write_cell(31, 2, 3, "스타트업 크레딧 프로그램 참여")
w.write_cell(31, 2, 4, "26.10")
print("  OK")

# ===== 사업비: find_replace =====
print("\n[표22,24] 사업비 — find_replace")
pset = hwp.HParameterSet.HFindReplace


def do_replace(old, new):
    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.FindString = old
    pset.ReplaceString = new
    pset.IgnoreMessage = 1
    pset.FindType = 1
    r = hwp.HAction.Execute("AllReplace", pset.HSet)
    status = "OK" if r else "skip"
    print(f"  {status}: {old[:40]} -> {new[:40]}")


do_replace("DMD소켓 구입(00개×0000원)", "클라우드 서버(AWS) 월 50만원 x 4개월")
do_replace("전원IC류 구입(00개×000원)", "LLM API 호출 비용 (Claude, GPT-4o)")
do_replace("시금형제작 외주용역(OOO제품 .... 플라스틱금형제작)", "UX/UI 디자인 외주 용역")
do_replace("국내 OO전시회 참가비(부스 임차 등 포함)", "도메인, SSL, SaaS 도구 구독료")

ctrl.save()
print("\n=== 전체 저장 완료 ===")
ctrl.quit()

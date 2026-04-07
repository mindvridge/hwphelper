"""팀 구성 섹션 수정: 표[29] 헤더 복원 + 본문 재배치."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
os.chdir("c:/hwphelper")

from src.hwp_engine.com_controller import HwpController
from src.hwp_engine.cell_writer import CellWriter
from src.hwp_engine.table_reader import TableReader

dst = "[별첨 1] 2026년 예비창업패키지 사업계획서_작성본.hwp"
ctrl = HwpController(visible=True)
ctrl.connect()
ctrl.open(dst)
w = CellWriter(ctrl)
hwp = ctrl.hwp


def write_direct(table_idx: int, moves: int, text: str) -> None:
    """표 진입 후 moves번 TableRightCell로 이동하고 텍스트 쓰기."""
    hwp.get_into_nth_table(table_idx)
    for _ in range(moves):
        hwp.TableRightCell()
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Delete")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


# 현재 작성본 구조:
# 표[28] = 본문 안내 (원본 표[29]) — 여기에 팀구성 본문을 넣어야 함
# 표[29] = 팀원표 (원본 표[30]) — [0,0]에 본문이 잘못 들어감, 헤더 복원 필요
# 표[30] = 파트너표 (원본 표[31])

# 1. 표[28] (1셀짜리)에 팀구성 본문 넣기
print("=== 표[28] 팀구성 본문 ===")
team_text = """4-1. 대표자 역량

대표자는 풀스택 소프트웨어 개발 10년 경력을 보유하고 있으며, Python, React, FastAPI 등 본 프로젝트의 핵심 기술 스택에 깊은 전문성을 갖추고 있다. 정부과제 사업계획서를 직접 20건 이상 작성·수행한 경험이 있어, 사업계획서 작성의 pain point와 평가 기준을 정확히 이해하고 있다. 특히 한/글 COM 자동화, LLM API 연동, 웹 서비스 아키텍처 설계에서의 실무 역량은 DocuMind의 핵심 기술 개발에 직결된다.

4-2. 팀원 구성 및 역할

핵심 개발팀 3인 체제로 구성하며, 각 팀원의 전문 분야가 상호 보완적이다.
① CTO (김OO): 컴퓨터공학 석사, NLP/LLM 연구 경력 5년. AI 엔진 아키텍처 설계 및 모델 최적화 담당
② UX 디자이너 (이OO): 시각디자인 학사, SaaS 제품 디자인 경력 3년. 채팅 UI/UX 설계 및 사용성 테스트 담당

4-3. 업무 파트너 및 협력 네트워크

① 한국정보화진흥원: 전문 멘토링 및 네트워킹 지원
② AWS: 클라우드 인프라 크레딧 지원 (스타트업 프로그램 참여)
③ 법무법인: 법인 설립, 특허 출원, 계약서 검토 등 법률 자문"""
w.write_cell(28, 0, 0, team_text)
print("  완료")

# 2. 표[29] 헤더 복원 — [0,0]="구분"으로 되돌리기
print("\n=== 표[29] 헤더 복원 + 팀원 데이터 ===")
w.write_cell(29, 0, 0, "구분")

# 팀원 데이터 재입력
w.write_cell(29, 1, 1, "CTO")
w.write_cell(29, 1, 2, "AI 엔진 개발 총괄")
w.write_cell(29, 1, 3, "컴퓨터공학 석사, NLP 연구 경력 5년, 논문 3편")
w.write_cell(29, 1, 4, "완료('26.09)")
w.write_cell(29, 2, 1, "디자이너")
w.write_cell(29, 2, 2, "UX/UI 설계")
w.write_cell(29, 2, 3, "시각디자인 학사, SaaS 디자인 경력 3년")
w.write_cell(29, 2, 4, "예정('26.12)")
print("  완료")

# 3. 표[30] 파트너 재입력
print("\n=== 표[30] 파트너 데이터 ===")
w.write_cell(30, 1, 1, "한국정보화진흥원")
w.write_cell(30, 1, 2, "창업 멘토링, 네트워킹")
w.write_cell(30, 1, 3, "전문 멘토 매칭, 투자자 연결")
w.write_cell(30, 1, 4, "26.09")
w.write_cell(30, 2, 1, "AWS Korea")
w.write_cell(30, 2, 2, "클라우드 인프라 운영")
w.write_cell(30, 2, 3, "스타트업 크레딧 프로그램 참여")
w.write_cell(30, 2, 4, "26.10")
print("  완료")

ctrl.save()
print("\n저장 완료")
ctrl.quit()

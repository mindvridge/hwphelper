"""사업비 표 수정 — HAction FindReplace 사용."""
import sys
import os
import shutil

sys.stdout.reconfigure(encoding="utf-8")
os.chdir("c:/hwphelper")

# 원본에서 다시 복사
src = "[별첨 1] 2026년 예비창업패키지 사업계획서_작성본.hwp"

from src.hwp_engine.com_controller import HwpController

ctrl = HwpController(visible=True)
ctrl.connect()
ctrl.open(src)
hwp = ctrl.hwp


def do_replace(old: str, new: str) -> bool:
    """HAction 기반 찾아바꾸기. 표 내부까지 검색."""
    pset = hwp.HParameterSet.HFindReplace
    hwp.HAction.GetDefault("AllReplace", pset.HSet)
    pset.FindString = old
    pset.ReplaceString = new
    pset.IgnoreMessage = 1
    pset.FindType = 1  # 전체 범위
    result = hwp.HAction.Execute("AllReplace", pset.HSet)
    print(f"  '{old[:30]}' -> '{new[:30]}': {result}")
    return result


print("=== 사업비 산출근거 바꾸기 ===")

# 1단계+2단계 공통 (양쪽 표 모두 동일 원본 텍스트)
replacements = [
    # 산출근거
    ("DMD소켓 구입(00개×0000원)", "클라우드 서버(AWS) 월 50만원 x 4개월"),
    ("전원IC류 구입(00개×000원)", "LLM API 호출 비용 (Claude, GPT-4o)"),
    ("시금형제작 외주용역(OOO제품 .... 플라스틱금형제작)", "UX/UI 디자인 외주 용역"),
    ("국내 OO전시회 참가비(부스 임차 등 포함)", "도메인, SSL, SaaS 도구 구독료"),
]

for old, new in replacements:
    do_replace(old, new)

# 금액 바꾸기 (1단계+2단계 동일)
# 원본: 3,000,000 / 7,000,000 / 10,000,000 / 1,000,000
# → 2,000,000 / 3,000,000 / 8,000,000 / 2,000,000
print("\n=== 금액 바꾸기 ===")
do_replace("10,000,000", "8,000,000")   # 먼저 10M → 8M (순서 중요)
do_replace("7,000,000", "3,000,000")
do_replace("3,000,000", "2,000,000")
# 1,000,000 → 2,000,000으로 바꾸면 위 결과와 겹칠 수 있으므로 skip

ctrl.save()
print("\n저장 완료")
ctrl.quit()

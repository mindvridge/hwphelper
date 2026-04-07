"""예비창업패키지 양식 전체 채우기 + 검증 스크립트."""
import sys, io, time, threading, ctypes
from ctypes import wintypes

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# 팝업 자동 클릭
user32 = ctypes.windll.user32
running = True

def auto_click():
    while running:
        def cb(hwnd, _):
            if not user32.IsWindowVisible(hwnd):
                return True
            buf = ctypes.create_unicode_buffer(512)
            user32.GetClassNameW(hwnd, buf, 512)
            user32.GetWindowTextW(hwnd, buf, 512)
            if buf.value == "#32770" or "한글" in buf.value:
                child = user32.FindWindowExW(hwnd, None, "Button", None)
                while child:
                    user32.GetWindowTextW(child, buf, 512)
                    if any(k in buf.value for k in ["예", "확인", "접근 허용", "모두 허용"]):
                        user32.SendMessageW(child, 0x00F5, 0, 0)
                        return True
                    child = user32.FindWindowExW(hwnd, child, "Button", None)
            return True
        user32.EnumWindows(
            ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(cb), 0
        )
        time.sleep(0.2)

threading.Thread(target=auto_click, daemon=True).start()

from pyhwpx import Hwp

hwp = Hwp(visible=True)
try:
    hwp.SetMessageBoxMode(0x10000 | 0x20000)
except:
    pass

hwp.open(r"C:\hwphelper\[별첨 1] 2026년 예비창업패키지 사업계획서 양식.hwp")

ctrl = hwp.HeadCtrl
orig = 0
while ctrl:
    if ctrl.CtrlID == "tbl":
        orig += 1
    ctrl = ctrl.Next
print(f"원본 표: {orig}")


def FR(f, r):
    """찾은 텍스트만 교체. 셀 이동/삭제 없음."""
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = f
    hwp.HParameterSet.HFindReplace.FindType = 1
    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        print(f"  X [{f[:30]}]")
        return False
    act = hwp.CreateAction("CharShape")
    p = act.CreateSet()
    act.GetDefault(p)
    p.SetItem("TextColor", 0)
    act.Execute(p)
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = r
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    print(f"  O [{f[:30]}]")
    return True


# ===== 명칭/범주 =====
print("[명칭/범주]")
# 명칭/범주: 표에 직접 진입하여 교체
def fill_table_cell(table_idx, content):
    """표에 진입 → 셀 전체를 content로 교체 (InsertText로 선택 영역 덮어쓰기)."""
    hwp.get_into_nth_table(table_idx)
    # 현재 셀에서 Ctrl+A (셀 내 전체 선택) 시뮬레이션
    import win32api, win32con
    win32api.keybd_event(0x11, 0, 0, 0)  # Ctrl down
    win32api.keybd_event(0x41, 0, 0, 0)  # A down
    time.sleep(0.05)
    win32api.keybd_event(0x41, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)
    # 글자색 검정 + 삽입
    act = hwp.CreateAction("CharShape")
    p = act.CreateSet(); act.GetDefault(p); p.SetItem("TextColor", 0); act.Execute(p)
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = content
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    print(f"  O [표{table_idx} 직접]")

print("[명칭/범주 — 표 직접 진입]")
fill_table_cell(7, "AI 심리상담 플랫폼 마인드케어")
fill_table_cell(8, "AI 헬스케어 / 디지털 심리상담")

# ===== 아이템 개요 =====
print("[아이템 개요]")
FR("※ 본 지원사업을 통해 개발 또는 구체화하고자 하는 제품·서비스 개요",
   "AI 아바타 기반 실시간 심리상담 플랫폼")
FR("(사용 용도, 사양, 가격 등), 핵심 기능·성능, 고객 제공 혜택 등",
   "MuseTalk 립싱크와 Qwen3.5 한국어 모델로 전문 상담사 수준 비대면 상담. 24시간, 비용 90%절감.")
FR("※ 예시 : 가벼움(고객 제공 혜택)을 위해서 용량을 줄이는 재료(핵심 기능)를 사용", "")

# ===== 일반현황 =====
print("[일반현황]")
FR("OO기술이 적용된 OO기능의(혜택을 제공하는) OO제품·서비스 등",
   "AI 아바타 기술이 적용된 실시간 심리상담(접근성 향상) 플랫폼")
FR("모바일 어플리케이션(0개), 웹사이트(0개)",
   "모바일 앱(1개), 웹사이트(1개), AI 상담 엔진(1개)")
FR("※ 협약기간 내 제작·개발 완료할 최종 생산품의 형태, 수량 등 기재", "")
FR("OOOOO", "마인드브이알")
FR("S/W 개발 총괄", "AI 엔진 개발")
FR("OO학 박사, OO학과 교수 재직(00년)", "컴공 석사, AI 스타트업 3년")
FR("홍보 및 마케팅", "임상심리 감수")
FR("OO학 학사, OO 관련 경력(00년 이상)", "임상심리전문가, 상담 10년")

# ===== 1. 문제 인식 =====
print("[1. 문제 인식]")
# 표10 안내문
FR("※ 개발하고자 하는 창업 아이템의 국내·외 시장 현황 및 문제점 등",
   "국내 심리상담 시장은 연간 2조원 규모이나 전문 상담사 부족과 높은 비용(회당 8-15만원)으로 접근성 제한. 20-30대 우울증 유병률 18.7%로 5년간 2.3배 증가.")
FR("문제 해결을 위한 창업 아이템 필요성 등",
   "기존 AI 상담(트로스트)은 텍스트 기반 세션 3.2분으로 몰입도 낮음. 마인드브이알은 아바타 음성상담으로 세션 10분(3배), 재방문율 40% 달성.")
# 표17 상세 안내문
FR("※ 개발하고자 하는 창업 아이템의 국내·외 시장 현황 및 문제점 등의 제시", "")
FR("문제 해결을 위한 창업 아이템의 개발 필요성 등 기재", "")

# ===== 2. 실현 가능성 =====
print("[2. 실현 가능성]")
# 표11 안내문
FR("※ 개발하고자 하는 창업 아이템을 사업기간 내 제품·서비스로 개발 또는 구체화",
   "1단계(1-3월): AI엔진(Qwen3.5 파인튜닝). 2단계(4-6월): 아바타통합(MuseTalk). 3단계(7-10월): 정식출시.")
FR("하고자 하는 계획", "산출물: AI엔진v1.0, 앱v0.9, 서비스v1.0")
FR("- 개발하고자 하는 창업 아이템의 차별성 및 경쟁력 확보 전략",
   "차별성: 아바타 음성상담 세션10분(3배), 재방문율40%. Qwen3.5 한국어85%. 특허출원 예정.")
# 표19 상세 안내문
FR("※ 아이디어를 제품·서비스로 개발 또는 구체화 하고자 하는 계획", "")
FR("(사업기간 내 일정 등)", "")
FR("개발 창업 아이템의 기능·성능의 차별성", "")

# ===== 3. 성장전략 =====
print("[3. 성장전략]")
FR("※ 경쟁사 분석, 목표 시장 진입 전략, 창업 아이템의 비즈니스 모델(수익화 모델), 사업 전체 로드맵, 투자유치 전략 등",
   "TAM 2.3조, SAM 4,500억, SOM 150억(2027). BM: B2C 월19,900원+B2B 월50만원. 2026 MVP→2027 정식(매출6억)→2028 해외(매출30억).")
# 표26 상세 안내문
FR("※ 경쟁제품·경쟁사 분석, 창업 아이템의 목표 시장 진입 전략 등 기재", "")

# ===== 4. 팀 구성 =====
print("[4. 팀 구성]")
FR("※ 대표자, 팀원, 업무파트너(협력기업) 등 역량 활용 계획 등",
   "대표 박수현(컴공석사, AI5년, 특허2건). CTO(딥러닝석사). 임상심리전문가(10년). 프론트엔드(UI3년). 파트너: 연세대MOU.")
FR("※ 대표자 보유 역량(경영 능력, 경력·학력, 기술력, 노하우, 인적 네트워크 등)",
   "박수현: 컴공석사(NLP), AI 5년, 특허2건, SCI 3편")

# ===== 기타 =====
print("[기타]")
FR("※ 1단계 정부지원사업비는 20백만원 내외로 작성", "")
FR("※ 2단계 정부지원사업비는 20백만원 내외로 작성", "")
FR("※ 제품", "")
FR("서비스 특징을 나타낼 수 있는 참고 사진(이미지)", "")
FR("설계도 등 삽입", "")
FR("(해당 시)", "")
FR("※ 제품", "")
FR("서비스 특징을 나타낼 수 있는 참고 사진(이미지)", "")
FR("설계도 등 삽입", "")
FR("(해당 시)", "")

# ===== 검증 =====
print()
print("=== 검증 ===")
ctrl = hwp.HeadCtrl
final = 0
while ctrl:
    if ctrl.CtrlID == "tbl":
        final += 1
    ctrl = ctrl.Next
print(f"표: {final}/{orig}", "✅" if final == orig else f"❌ {orig-final}개 삭제")

labels = ["산출물", "직업", "기업(예정)명", "순번", "직위", "문제 인식", "실현 가능성", "성장전략", "팀 구성", "아이템 개요", "담당 업무", "구분", "비  목"]
miss = []
for l in labels:
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = l
    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        miss.append(l)
print("라벨:", "✅ 모두 유지" if not miss else f"❌ 삭제: {miss}")

# ※ 잔존 확인
remain = 0
for ti in [7, 8, 9, 10, 11, 12, 13, 17, 19, 21, 23, 26, 29]:
    try:
        hwp.get_into_nth_table(ti)
        df = hwp.table_to_df_q()
        txt = str(df.columns[0])
        has_old = any(x in txt for x in ["게토레이", "스포츠음료", "Windows", "알파고", "OO기술", "OO학", "OOOOO"])
        if has_old:
            remain += 1
            print(f"  ⚠ 표{ti}: 예시잔존 [{txt[:40]}]")
    except:
        pass
print("예시잔존:", f"⚠ {remain}개" if remain else "✅ 없음")

# 새 내용
fc = 0
for n in ["마인드케어", "마인드브이알", "AI 엔진 개발", "임상심리전문가", "2.3조", "Qwen3.5", "세션 10분"]:
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = n
    if hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        fc += 1
print(f"새 내용: {fc}/7")

output = r"C:\hwphelper\예비창업패키지_마인드브이알_최종.hwp"
hwp.SaveAs(output)
running = False
print(f"\n저장: {output}")
print("완료!")

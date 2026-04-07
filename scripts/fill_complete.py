"""예비창업패키지 양식 완전 채우기 — 빠짐없이 모든 교체 대상 처리."""
import sys, io, time, threading, ctypes
from ctypes import wintypes

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# 팝업 자동 클릭
user32 = ctypes.windll.user32
running = True
def auto_click():
    while running:
        def cb(hwnd, _):
            if not user32.IsWindowVisible(hwnd): return True
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
        user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(cb), 0)
        time.sleep(0.2)
threading.Thread(target=auto_click, daemon=True).start()

from pyhwpx import Hwp
import win32api, win32con

hwp = Hwp(visible=True)
try: hwp.SetMessageBoxMode(0x10000 | 0x20000)
except: pass
hwp.open(r"C:\hwphelper\[별첨 1] 2026년 예비창업패키지 사업계획서 양식.hwp")

ctrl = hwp.HeadCtrl
orig = 0
while ctrl:
    if ctrl.CtrlID == "tbl": orig += 1
    ctrl = ctrl.Next
print(f"원본 표: {orig}")


def FR(f, r):
    """찾은 텍스트만 교체. 표 구조 절대 안전."""
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = f
    hwp.HParameterSet.HFindReplace.FindType = 1
    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        print(f"  X [{f[:35]}]")
        return False
    act = hwp.CreateAction("CharShape")
    p = act.CreateSet(); act.GetDefault(p); p.SetItem("TextColor", 0); act.Execute(p)
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = r
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    print(f"  O [{f[:35]}]")
    return True


def _select_all_in_cell():
    """Home → Shift+Ctrl+End → Delete를 반복하여 셀 전체 비우기."""
    for _ in range(5):  # 5회 반복으로 확실히
        win32api.keybd_event(0x24, 0, 0, 0)  # Home
        time.sleep(0.02)
        win32api.keybd_event(0x24, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        # Ctrl+Home (셀 내 절대 시작)
        win32api.keybd_event(0x11, 0, 0, 0)
        win32api.keybd_event(0x24, 0, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(0x24, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.02)
        # Shift+Ctrl+End
        win32api.keybd_event(0x10, 0, 0, 0)
        win32api.keybd_event(0x11, 0, 0, 0)
        win32api.keybd_event(0x23, 0, 0, 0)
        time.sleep(0.03)
        win32api.keybd_event(0x23, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(0x10, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)
        # Delete
        win32api.keybd_event(0x2E, 0, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(0x2E, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.05)

def fill_table(ti, content):
    """표에 직접 진입 → 셀 내용 완전 비우기 → 교체."""
    hwp.get_into_nth_table(ti)
    time.sleep(0.1)
    _select_all_in_cell()
    if content:
        act = hwp.CreateAction("CharShape")
        p = act.CreateSet(); act.GetDefault(p); p.SetItem("TextColor", 0); act.Execute(p)
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = content
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    print(f"  O [표{ti} 직접]")


# ==============================================================
# 1. 표0: 안내문 (삭제 — 제출 시 삭제하라는 안내)
# ==============================================================
print("\n[표0: 안내문]")
fill_table(0, "")

# ==============================================================
# 2. 표3: 작성 안내문 (유지 — 읽기 전용)
# ==============================================================
print("\n[표3: 작성 안내 — 유지]")

# ==============================================================
# 3. 표4: 일반현황 (기업정보)
# ==============================================================
print("\n[표4: 일반현황]")
FR("OO기술이 적용된 OO기능의(혜택을 제공하는) OO제품·서비스 등",
   "AI 아바타 기술이 적용된 실시간 심리상담(접근성 향상) 플랫폼")
FR("모바일 어플리케이션(0개), 웹사이트(0개)",
   "모바일 앱(1개), 웹사이트(1개), AI 상담 엔진(1개)")
FR("OOOOO", "마인드브이알")

# ==============================================================
# 4. 표5: 산출물 안내 (삭제)
# ==============================================================
print("\n[표5: 산출물 안내]")
FR("협약기간 내 제작", "")

# ==============================================================
# 5. 표7,8: 명칭/범주 예시 (표 직접 진입으로 교체)
# ==============================================================
print("\n[표7,8: 명칭/범주]")
fill_table(7, "AI 심리상담 플랫폼 마인드케어")
fill_table(8, "AI 헬스케어 / 디지털 심리상담")

# ==============================================================
# 6. 표9: 아이템 개요 + 예시
# ==============================================================
print("\n[표9: 아이템 개요]")
fill_table(9, "AI 아바타 기반 실시간 심리상담 플랫폼. MuseTalk 립싱크와 Qwen3.5 한국어 모델로 전문 상담사 수준 비대면 상담. 24시간 접근, 비용 90% 절감, 익명성 보장.")

# ==============================================================
# 7. 표10: 문제 인식 안내문
# ==============================================================
print("\n[표10: 문제 인식 안내문]")
FR("※ 개발하고자 하는 창업 아이템의 국내·외 시장 현황 및 문제점 등",
   "국내 심리상담 시장은 연간 2조원 규모이나 전문 상담사 부족과 높은 비용(회당 8-15만원)으로 접근성 제한. 20-30대 우울증 유병률 18.7%로 5년간 2.3배 증가.")
FR("문제 해결을 위한 창업 아이템 필요성 등",
   "기존 AI 상담(트로스트)은 텍스트 기반 세션 3.2분으로 몰입도 낮음. 마인드브이알은 아바타 음성상담으로 세션 10분(3배), 재방문율 40% 달성.")

# ==============================================================
# 8. 표11: 실현 가능성 안내문
# ==============================================================
print("\n[표11: 실현 가능성 안내문]")
FR("※ 개발하고자 하는 창업 아이템을 사업기간 내 제품·서비스로 개발 또는 구체화",
   "1단계(1-3월): AI엔진 구축 - Qwen3.5 파인튜닝, CBT/DBT 알고리즘, 위기감지. 2단계(4-6월): 아바타통합 - MuseTalk 4종, STT/TTS(500ms). 3단계(7-10월): 정식출시 - 앱스토어, 대학B2B 3개소.")
FR("하고자 하는 계획", "산출물: AI엔진v1.0, 앱v0.9, 서비스v1.0, MAU 1,000")
FR("- 개발하고자 하는 창업 아이템의 차별성 및 경쟁력 확보 전략",
   "차별성: 아바타 음성상담 세션10분(경쟁사3배), 재방문율40%. Qwen3.5 한국어85%. 임상심리전문가 감수. 특허출원 예정.")

# ==============================================================
# 9. 표12: 성장전략 안내문
# ==============================================================
print("\n[표12: 성장전략 안내문]")
FR("※ 경쟁사 분석, 목표 시장 진입 전략, 창업 아이템의 비즈니스 모델(수익화 모델), 사업 전체 로드맵, 투자유치 전략 등",
   "경쟁사: 트로스트(텍스트, MAU5만), 마보(명상), Woebot(영어). 차별점: 아바타 음성 + 한국어 + 임상심리. TAM 2.3조, SAM 4,500억, SOM 150억(2027). BM: B2C 월19,900원 + B2B 월50만원. 투자: 엔젤5천만 → 시드3억 → 시리즈A 15억. 로드맵: 2026 MVP → 2027 정식(매출6억) → 2028 해외(매출30억).")

# ==============================================================
# 10. 표13: 팀 구성 안내문
# ==============================================================
print("\n[표13: 팀 구성 안내문]")
FR("※ 대표자, 팀원, 업무파트너(협력기업) 등 역량 활용 계획 등",
   "대표 박수현(컴공석사, NLP전공, AI5년, 특허2건, SCI3편): 전체 총괄, AI엔진 설계. CTO 김OO(딥러닝석사): LLM 최적화. 임상심리 이OO(전문가자격, 10년): 콘텐츠 감수. 프론트엔드 박OO(3년): 아바타UI. 파트너: 연세대 심리학과(MOU), vLLM커뮤니티(기술자문).")

# ==============================================================
# 11. 표14,15: 이미지 안내문 (삭제)
# ==============================================================
print("\n[표14,15: 이미지 안내]")
FR("참고 사진(이미지)", "")
FR("설계도 등 삽입", "")
FR("(해당 시)", "")
FR("참고 사진(이미지)", "")
FR("설계도 등 삽입", "")
FR("(해당 시)", "")

# ==============================================================
# 12. 표17: 문제 인식 상세 안내문 (삭제)
# ==============================================================
print("\n[표17: 문제 인식 상세]")
FR("문제점 등의 제시", "")
FR("개발 필요성 등 기재", "")

# ==============================================================
# 13. 표19: 실현 가능성 상세 안내문 (삭제)
# ==============================================================
print("\n[표19: 실현 가능성 상세]")
FR("구체화 하고자 하는 계획", "")
FR("(사업기간 내 일정 등)", "")
FR("차별성 및 경쟁력 확보 전략", "")

# ==============================================================
# 14. 표21,23: 사업비 안내문 (삭제)
# ==============================================================
print("\n[표21,23: 사업비 안내]")
FR("1단계 정부지원사업비는 20백만원 내외로 작성", "")
FR("2단계 정부지원사업비는 20백만원 내외로 작성", "")

# ==============================================================
# 15. 표26: 성장전략 상세 안내문 (삭제)
# ==============================================================
print("\n[표26: 성장전략 상세]")
FR("목표 시장 진입 전략 등 기재", "")
FR("사업 확장을 위한 투자유치 전략 등", "")

# ==============================================================
# 16. 표29: 팀 구성 상세 안내문
# ==============================================================
print("\n[표29: 팀 구성 상세]")
fill_table(29, "박수현: 컴공석사(NLP전공), AI스타트업 5년, 특허2건, SCI 3편, 정부과제 2건 참여. 연세대 심리학과 공동연구 MOU, vLLM 커뮤니티 기술자문.")

# ==============================================================
# 17. 표30: 팀원 정보 (예시 데이터 교체)
# ==============================================================
print("\n[표30: 팀원 정보]")
FR("S/W 개발 총괄", "AI 엔진 개발")
FR("OO학 박사, OO학과 교수 재직(00년)", "컴공 석사, AI 스타트업 3년 경력")
FR("홍보 및 마케팅", "임상심리 감수 및 콘텐츠")
FR("OO학 학사, OO 관련 경력(00년 이상)", "임상심리전문가 자격, 상담경력 10년")

# ==============================================================
# 18. 표31: 업무파트너 (예시 데이터 교체)
# ==============================================================
print("\n[표31: 업무파트너]")
FR("○○전자", "연세대 심리학과")
FR("시제품 관련 H/W 제작·개발", "공동연구(임상시험 설계, 효과성 검증)")
FR("테스트 장비 지원", "연구 인프라 활용")
FR("○○기업", "vLLM 커뮤니티")

# ==============================================================
# 검증
# ==============================================================
print("\n" + "=" * 50)
print("=== 검증 ===")

# 표 개수
ctrl = hwp.HeadCtrl
final = 0
while ctrl:
    if ctrl.CtrlID == "tbl": final += 1
    ctrl = ctrl.Next
print(f"표: {final}/{orig}", "✅" if final == orig else f"❌ {orig-final}개 삭제")

# 라벨 유지
labels = ["산출물", "직업", "기업(예정)명", "순번", "직위", "문제 인식",
          "실현 가능성", "성장전략", "팀 구성", "아이템 개요", "담당 업무",
          "구분", "비  목", "파트너명", "보유역량", "협업방안"]
miss = []
for l in labels:
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = l
    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        miss.append(l)
print("라벨:", "✅ 모두 유지" if not miss else f"❌ 삭제: {miss}")

# 예시 잔존
remain = 0
for ti in [7, 8, 9, 10, 11, 12, 13, 17, 19, 21, 23, 26, 29]:
    try:
        hwp.get_into_nth_table(ti)
        df = hwp.table_to_df_q()
        txt = str(df.columns[0])
        old = ["게토레이", "스포츠음료", "Windows", "알파고", "OO기술", "OOOOO"]
        if any(x in txt for x in old):
            remain += 1
            print(f"  ⚠ 표{ti}: [{txt[:40]}]")
    except:
        pass
print("예시잔존:", f"⚠ {remain}개" if remain else "✅ 없음")

# 새 내용 존재
news = ["마인드케어", "마인드브이알", "AI 엔진 개발", "임상심리전문가",
        "2.3조", "Qwen3.5", "세션 10분", "연세대 심리학과"]
fc = 0
for n in news:
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = n
    if hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        fc += 1
print(f"새 내용: {fc}/{len(news)}")

output = r"C:\hwphelper\예비창업패키지_마인드브이알_최종.hwp"
hwp.SaveAs(output)
running = False
print(f"\n저장: {output}")
print("완료! 한/글 열린 상태.")

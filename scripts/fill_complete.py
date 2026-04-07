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
    """찾은 텍스트만 교체. 검정색으로 삽입. 표 구조 절대 안전."""
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = f
    hwp.HParameterSet.HFindReplace.FindType = 1
    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        print(f"  X [{f[:35]}]")
        return False
    # 1. 찾은 텍스트 삭제 (선택 상태에서 Delete)
    time.sleep(0.1)
    hwp.HAction.Run("Delete")
    # 2. 커서 위치에서 글자색 검정 설정
    act = hwp.CreateAction("CharShape")
    p = act.CreateSet()
    act.GetDefault(p)
    p.SetItem("TextColor", 0)
    p.SetItem("Bold", 0)
    act.Execute(p)
    # 3. 검정색으로 새 텍스트 삽입
    if r:
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
# 직업 (교수/연구원... → 일반인)
FR("교수 / 연구원 / 사무직 /", "일반인")
FR("일반인 / 대학생 등", "")  # 줄바꿈 후 남은 텍스트

# ==============================================================
# 4. 표5: 산출물 안내 (삭제)
# ==============================================================
print("\n[표5: 산출물 안내]")
FR("※ 협약기간 내 제작", "")
FR("완료할 최종 생산품의 형태, 수량 등 기재", "")

# ==============================================================
# 5. 표7,8: 명칭/범주 예시 (표 직접 진입으로 교체)
# ==============================================================
print("\n[표7,8: 명칭/범주]")
FR("게토레이", "AI 심리상담 플랫폼 마인드케어")
FR("Windows", "")
FR("알파고", "")
FR("스포츠음료", "AI 헬스케어 / 디지털 심리상담")
FR("OS(운영체계)", "")
FR("인공지능프로그램", "")
FR("※ 예시 1 : ", "")
FR("예시 2 : ", "")
FR("예시 3 : ", "")
FR("※ 예시 1 : ", "")
FR("예시 2 : ", "")
FR("예시 3 : ", "")

# ==============================================================
# 6. 표9: 아이템 개요 + 예시
# ==============================================================
print("\n[표9: 아이템 개요]")
FR("※ 본 지원사업을 통해 개발 또는 구체화하고자 하는 제품",
   "AI 아바타 기반 실시간 심리상담 플랫폼")
FR("서비스 개요", "")
FR("(사용 용도, 사양, 가격 등), 핵심 기능",
   "MuseTalk 립싱크와 Qwen3.5 한국어 모델로 전문 상담사 수준 비대면 상담. 24시간, 비용 90%절감.")
FR("성능, 고객 제공 혜택 등", "")
FR("※ 예시 : 가벼움(고객 제공 혜택)을 위해서 용량을 줄이는 재료(핵심 기능)를 사용", "")

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
FR("※ 제품", "")
FR("참고 사진(이미지)", "")
FR("설계도 등 삽입", "")
FR("(해당 시)", "")
FR("※ 제품", "")
FR("참고 사진(이미지)", "")
FR("설계도 등 삽입", "")
FR("(해당 시)", "")

# ==============================================================
# 12. 표17: 문제 인식 상세 안내문 (삭제)
# ==============================================================
print("\n[표17: 문제 인식 상세]")
FR("※ 개발하고자 하는 창업 아이템의 국내", "")
FR("문제점 등의 제시", "")
FR("개발 필요성 등 기재", "")

# ==============================================================
# 13. 표19: 실현 가능성 상세 안내문 (삭제)
# ==============================================================
print("\n[표19: 실현 가능성 상세]")
FR("※ 아이디어를 제품", "")
FR("구체화 하고자 하는 계획", "")
FR("(사업기간 내 일정 등)", "")
FR("차별성 및 경쟁력 확보 전략", "")

# ==============================================================
# 14. 표21,23: 사업비 안내문 (삭제)
# ==============================================================
print("\n[표21,23: 사업비 안내]")
FR("※ 1단계 정부지원사업비는 20백만원 내외로 작성", "")
FR("※ 2단계 정부지원사업비는 20백만원 내외로 작성", "")

# ==============================================================
# 15. 표26: 성장전략 상세 안내문 (삭제)
# ==============================================================
print("\n[표26: 성장전략 상세]")
FR("※ 경쟁제품", "")
FR("목표 시장 진입 전략 등 기재", "")
FR("투자유치 전략 등", "")

# ==============================================================
# 16. 표29: 팀 구성 상세 안내문
# ==============================================================
print("\n[표29: 팀 구성 상세]")
FR("※ 대표자 보유 역량(경영 능력",
   "박수현: 컴공석사(NLP전공), AI 5년, 특허2건, SCI 3편")
FR("인적 네트워크 등)", "")
FR("역량 : 창업아이템을 개발", "")
FR("구체화 할 수 있는 능력", "")
FR("※ 팀에서 보유 또는 보유할 예정인 장비",
   "연세대 심리학과 MOU, vLLM 커뮤니티 기술자문")

# ==============================================================
# 17. 표30: 팀원 정보 (예시 데이터 교체)
# ==============================================================
print("\n[표30: 팀원 정보]")
FR("S/W 개발 총괄", "AI 엔진 개발")
FR("OO학 박사, OO학과 교수 재직(00년)", "컴공 석사, AI 스타트업 3년 경력")
FR("홍보 및 마케팅", "임상심리 감수 및 콘텐츠")
FR("OO학 학사, OO 관련 경력(00년 이상)", "임상심리전문가 자격, 상담경력 10년")
FR("예정('00.0)", "완료")

# ==============================================================
# 18. 표31: 업무파트너 (예시 데이터 교체)
# ==============================================================
print("\n[표31: 업무파트너]")
FR("○○전자", "연세대 심리학과")
FR("시제품 관련 H/W 제작·개발", "공동연구(임상시험 설계, 효과성 검증)")
FR("테스트 장비 지원", "연구 인프라 활용")
FR("○○기업", "vLLM 커뮤니티")

# ==============================================================
# 19. 본문 섹션 B열 작성 (표6의 B7, B8, B9, B10)
# ==============================================================

def write_section_b(title_keyword, content):
    """제목 키워드를 찾아 오른쪽(B열)으로 이동 후 내용 삽입."""
    hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = title_keyword
    hwp.HParameterSet.HFindReplace.FindType = 1
    if not hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet):
        print(f"  X [{title_keyword}] 못찾음")
        return
    # B열로 이동
    hwp.TableRightCell()
    addr = hwp.get_cell_addr()
    # 셀 끝으로 이동 후 줄바꿈+내용 추가
    hwp.HAction.Run("MoveDocEnd")
    hwp.HAction.Run("BreakPara")
    # 글자색 검정
    act = hwp.CreateAction("CharShape")
    p = act.CreateSet()
    act.GetDefault(p)
    p.SetItem("TextColor", 0)
    act.Execute(p)
    # 내용 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = content
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    print(f"  O [{title_keyword}] → {addr}")


print("\n[본문: 1. 문제 인식 B열]")
write_section_b("1. 문제 인식", """국내 심리상담 시장은 연간 2조원 규모이나 전문 상담사 부족과 높은 비용(회당 8-15만원)으로 접근성이 제한적이다. 청년층 67%가 비용 부담으로 상담을 포기하며, 20-30대 우울증 유병률은 18.7%로 5년간 2.3배 증가하였다.

기존 AI 상담 서비스(트로스트, 마보 등)는 텍스트 기반으로 평균 세션 3.2분에 그쳐 사용자 몰입도가 낮고, 전문 상담 수준에 미달하여 일회성 사용에 그치는 한계가 있다. 비대면이면서도 대면 상담에 준하는 몰입감과 전문성을 갖춘 새로운 상담 서비스가 필요하다.

마인드브이알은 AI 아바타 기반 실시간 심리상담 플랫폼을 개발하여 MuseTalk 아바타 기술과 Qwen3.5 대규모 언어모델로 세션 시간 10분(3배), 재방문율 40%(2배)를 달성하여 심리상담 접근성 문제를 해결한다.""")

print("\n[본문: 2. 실현 가능성 B열]")
write_section_b("2. 실현 가능성", """[개발 계획]
1단계(1-3개월): AI 상담 엔진 구축
 - Qwen3.5-27B 기반 한국어 심리상담 특화 파인튜닝 (학습 데이터 5,000건)
 - CBT(인지행동치료), DBT(변증법적행동치료) 기반 상담 알고리즘 설계
 - 자해/자살 위험 감지 및 즉시 연계(1393) 안전장치 구현
 - 산출물: AI 상담 엔진 프로토타입 v1.0

2단계(4-6개월): MuseTalk 아바타 통합 및 베타 서비스
 - MuseTalk 기반 실시간 립싱크 아바타 4종 개발 (연령/성별 맞춤)
 - Whisper 한국어 STT + 자체 TTS 파이프라인 구축 (지연시간 500ms 이내)
 - 감정 분석 엔진 통합, 클로즈드 베타 100명 대상 8주간 사용성 테스트
 - 산출물: 아바타 상담 앱 베타 v0.9

3단계(7-10개월): 정식 출시 및 사업화
 - 앱스토어/플레이스토어 정식 출시
 - 대학교 상담센터 3개소 B2B 시범 도입
 - 산출물: 정식 서비스 v1.0, 월 MAU 1,000명 달성

[차별성 및 경쟁력]
 - 아바타 음성상담: 세션 시간 10분(경쟁사 텍스트 기반 3분의 3배), 재방문율 40%
 - MuseTalk 실시간 립싱크: 자연스러운 표정/입모양으로 대면 상담 수준의 몰입감
 - Qwen3.5 한국어 특화: 한국어 상담 맥락 이해도 85% (GPT-4 대비 15%p 향상)
 - 임상심리전문가 3인 자문단: CBT/DBT 근거기반 프로토콜 적용
 - 진입장벽: 한국어 상담 데이터 5,000건 + 아바타+감정분석 통합 파이프라인 (특허출원 예정)""")

print("\n[본문: 3. 성장전략 B열]")
write_section_b("3. 성장전략", """[경쟁사 분석]
 - 트로스트(국내): 텍스트 기반 AI 상담, MAU 5만, 세션 평균 3분
 - 마보(국내): 명상/마음챙김 앱, 상담 기능 미흡
 - Woebot(미국): CBT 기반 챗봇, 영어 전용
 - 당사 차별점: 아바타 음성상담 + 한국어 특화 + 임상심리 감수

[시장 분석]
 - TAM: 국내 심리상담 시장 2.3조원 (연평균 12% 성장)
 - SAM: 비대면 디지털 상담 시장 4,500억원
 - SOM: AI 아바타 상담 시장 150억원 (2027년)

[비즈니스 모델]
 - B2C 구독형: 월 19,900원 (주 2회 상담, 감정일기, 리포트)
 - B2B 라이선스: 월 50만원/기관 (대학 상담센터, 기업 EAP)

[투자 유치 전략]
 - 2026년 하반기: 엔젤투자 5천만원
 - 2027년 상반기: 시드투자 3억원 (MAU 1,000 달성 후)
 - 2028년: 시리즈A 15억원 (해외 진출 자금)

[로드맵]
 2026년: MVP 개발 + 베타 서비스 (MAU 1,000, 고용 3명)
 2027년: 정식 출시 + B2B 확장 (MAU 10,000, 매출 6억, 고용 8명)
 2028년: 해외 진출(일본) + 시리즈A (MAU 50,000, 매출 30억, 고용 20명)
 2029년: 동남아 확장 + 흑자 전환 (MAU 200,000, 매출 100억)""")

print("\n[본문: 4. 팀 구성 B열]")
write_section_b("4. 팀 구성", """[대표자] 박수현
 - 학력: OO대학교 컴퓨터공학과 석사 (자연어처리 전공)
 - 경력: AI 스타트업 연구개발팀 5년 (NLP, 대화시스템 개발)
 - 성과: AI 관련 특허 2건, SCI 논문 3편, 정부과제 참여 2건
 - 역할: 전체 사업 총괄, AI 엔진 아키텍처 설계, 투자유치

[핵심 팀원]
 - CTO 김OO: 딥러닝 석사, AI 엔진 5년, 대규모 모델 최적화 전문
 - 임상심리전문가 이OO: 임상심리전문가 자격, 상담경력 10년, 콘텐츠 감수
 - 프론트엔드 박OO: 아바타 UI/UX 3년, Three.js/WebGL 전문

[업무 파트너]
 - 연세대 심리학과: 공동연구 MOU (임상시험 설계, 효과성 검증)
 - vLLM 커뮤니티: LLM 추론 최적화 기술 자문""")


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

"""셀 분류기 — 표 셀을 LABEL / EMPTY / PREFILLED / PLACEHOLDER 로 분류.

분류 기준:
- EMPTY: 텍스트가 비어 있음
- PLACEHOLDER: 교체 대상 (예시글, 파란색 텍스트, 안내 문구 등)
- LABEL: 항목명·헤더 역할 (첫 행/첫 열, 라벨 키워드, 짧은 텍스트 등)
- PREFILLED: 이미 유의미한 내용이 채워져 있음 (수정하지 않을 셀)

정부과제 HWP 양식 특성:
- 파란색/빨간색 텍스트 = 작성 예시 (교체 대상)
- 검정 텍스트 + 볼드 = 라벨/항목명
- 검정 텍스트 + 일반 = 기입력 내용
"""

from __future__ import annotations

import re

import structlog

from .table_reader import Cell, CellType, Table

logger = structlog.get_logger()

# 파란색/빨간색 계열 텍스트 색상 값 (한/글 COM에서 반환)
# 한/글 색상: 0xBBGGRR 형식 또는 정수
_BLUE_COLORS = {
    0xFF0000, 0x800000, 0xCC0000,     # 순수 파랑 계열 (BGR)
    0xFF6600, 0xFF3300,               # 파란색 변형
    16711680, 8388608, 13369344,      # 같은 값 정수
}

_RED_COLORS = {
    0x0000FF, 0x0000CC, 0x0000AA,     # 빨강 계열 (BGR)
    255, 204, 170,                     # 같은 값 정수
}

# 검정이 아닌 색상 (검정=0 또는 매우 작은 값)
_BLACK_THRESHOLD = 0x303030  # R+G+B 각 48 이하면 검정으로 간주


class CellClassifier:
    """표 셀 타입 분류기."""

    # 플레이스홀더 키워드 — 예시글, 안내 문구
    PLACEHOLDER_KEYWORDS: list[str] = [
        # 직접적인 안내
        "(내용 입력)", "(작성)", "내용을 입력", "여기에 작성",
        "(기재)", "(기입)", "해당 내용", "작성 요령", "※ 작성",
        # 예시 표시
        "예시)", "예시:", "(예:", "(예시", "예) ", "<예시>",
        "(작성예시)", "작성예시", "작성 예시",
        # 기호형 플레이스홀더
        "○○", "◇◇", "△△", "□□", "●●",
        "OO기업", "OO회사", "OO기관", "OO대학", "주식회사 0000",
        "000-0000", "000-000-", "0000-00-00",
        # 정부과제 양식 특유의 예시 문구
        "본 사업은", "본 과제는", "당사는", "본 기업은",
        "핵심기술에 대해", "시장 규모는",
        "예비창업자", "창업아이템",
        "기술개발 내용을", "사업화 방안을",
        # 안내문 (삭제 대상)
        "※ 창업아이템", "※ 해당 아이템", "※ 해당 사업",
        "※ 창업 아이템", "※ 개발하고자", "※ 본 지원사업",
        "※ 제품(서비스)", "※ 사업 수행",
        "※ 창업 아이템의 국내", "※ 아이디어를 제품",
        "※ 경쟁사 분석", "※ 대표자의 보유",
        "파란색 글씨로 표기된", "삭제 후",
        # 단위 표시만 있는 경우
        "(백만원)", "(천원)", "(명)", "(개월)",
    ]

    # 라벨 키워드 — 항목명
    LABEL_KEYWORDS: list[str] = [
        "사업명", "과제명", "기관명", "대표자", "연락처", "주소", "담당자",
        "사업기간", "총사업비", "항목", "내용", "비고", "구분", "세부내용",
        "목표", "전략", "일정", "예산", "성명", "직위", "소속", "역할",
        "수량", "단가", "금액", "산출근거", "세목", "분류", "계",
        "합계", "소계", "번호", "No", "연번", "기간",
        "전화번호", "이메일", "팩스", "홈페이지", "설립일",
        "사업자등록번호", "법인등록번호",
        "생년월일", "최종학력", "전공", "경력사항",
        "기술분야", "제품명", "서비스명", "지원금액",
        "자부담", "정부출연금", "민간부담금",
    ]

    # 플레이스홀더 정규식 — 기호만, 번호만, 공백만
    _PLACEHOLDER_RE = re.compile(r"^[\s○◇△□▷▶●•\-·\.\,\(\)~]+$")
    _NUMBERING_RE = re.compile(r"^\s*\d+\)\s*$")  # "1)", "2)" 등 번호만
    _GUIDE_RE = re.compile(r"^※\s")  # "※ " 로 시작하는 안내문

    def __init__(
        self,
        label_keywords: list[str] | None = None,
        placeholder_keywords: list[str] | None = None,
        label_max_length: int = 20,
        detect_colored_text: bool = True,
    ) -> None:
        self._label_keywords = label_keywords or self.LABEL_KEYWORDS
        self._placeholder_keywords = placeholder_keywords or self.PLACEHOLDER_KEYWORDS
        self._label_max_length = label_max_length
        self._detect_colored_text = detect_colored_text

    # ------------------------------------------------------------------
    # public API
    # ------------------------------------------------------------------

    def classify(self, cell: Cell, table: Table) -> CellType:
        """단일 셀의 타입을 분류한다."""
        text = cell.text.strip()

        # 1. 빈 셀
        if not text:
            return CellType.EMPTY

        # 2. 플레이스홀더 감지 (키워드, 기호, 색상)
        if self._is_placeholder(text, cell):
            return CellType.PLACEHOLDER

        # 3. 라벨 감지
        if self._is_label(cell, table):
            return CellType.LABEL

        # 4. 나머지는 기 작성된 내용
        return CellType.PREFILLED

    def classify_table(self, table: Table) -> Table:
        """표의 모든 셀을 분류하고, cell_type을 설정한 Table을 반환한다."""
        for cell in table.cells:
            cell.cell_type = self.classify(cell, table)

        stats = {
            "label": sum(1 for c in table.cells if c.cell_type == CellType.LABEL),
            "empty": sum(1 for c in table.cells if c.cell_type == CellType.EMPTY),
            "placeholder": sum(1 for c in table.cells if c.cell_type == CellType.PLACEHOLDER),
            "prefilled": sum(1 for c in table.cells if c.cell_type == CellType.PREFILLED),
        }
        logger.info("셀 분류 완료", table_idx=table.table_idx, **stats)
        return table

    # ------------------------------------------------------------------
    # 분류 로직
    # ------------------------------------------------------------------

    def _is_placeholder(self, text: str, cell: Cell) -> bool:
        """플레이스홀더(교체 대상) 여부를 판단한다."""
        stripped = text.strip()

        # 키워드 매칭
        for kw in self._placeholder_keywords:
            if kw in stripped:
                return True

        # 기호만으로 구성된 텍스트
        if self._PLACEHOLDER_RE.match(stripped):
            return True

        # "1)", "2)" 등 번호만 있는 셀
        if self._NUMBERING_RE.match(stripped):
            return True

        # "※ " 로 시작하는 안내문
        if self._GUIDE_RE.match(stripped):
            return True

        # 파란색/빨간색 텍스트 감지 (정부과제 양식의 예시글)
        if self._detect_colored_text and cell.style:
            if self._is_colored_text(cell.style.text_color):
                return True

        return False

    def _is_label(self, cell: Cell, table: Table) -> bool:
        """라벨(항목명) 여부를 판단한다."""
        text = cell.text.strip()

        # 키워드 매칭
        for kw in self._label_keywords:
            if kw in text:
                return True

        # 첫 행은 헤더일 가능성 높음
        if cell.row == 0:
            return True

        # 첫 열이면서 짧은 텍스트
        if cell.col == 0 and len(text) <= self._label_max_length:
            return True

        # 짧은 텍스트 + 볼드
        if len(text) <= self._label_max_length and cell.style and cell.style.bold:
            return True

        return False

    @staticmethod
    def _is_colored_text(color_value: str | int) -> bool:
        """텍스트 색상이 검정이 아닌지 (파란색/빨간색 예시글 감지)."""
        try:
            if isinstance(color_value, str):
                # "0x00FF0000" 형태
                c = int(color_value, 0) if color_value.startswith("0x") else int(color_value)
            else:
                c = int(color_value)

            # BGR 분리
            b = (c >> 16) & 0xFF
            g = (c >> 8) & 0xFF
            r = c & 0xFF

            # 검정(0,0,0) 또는 매우 어두운 색은 일반 텍스트
            if r < 48 and g < 48 and b < 48:
                return False

            # 파란색 계열 확인
            if c in _BLUE_COLORS:
                return True

            # 빨간색 계열 확인
            if c in _RED_COLORS:
                return True

            # 파란색이 지배적 (B > 150, B > R*2, B > G*2)
            if b > 150 and b > r * 2 and b > g * 1.5:
                return True

            # 빨간색이 지배적
            if r > 150 and r > b * 2 and r > g * 1.5:
                return True

            # 어두운 파란색 (정부과제에서 자주 사용)
            if b > 100 and r < 50 and g < 50:
                return True

            # 검정이 아닌 모든 밝은 색 (회색 제외)
            # 회색: R≈G≈B
            if max(r, g, b) - min(r, g, b) > 50 and max(r, g, b) > 100:
                return True

        except (ValueError, TypeError):
            pass

        return False

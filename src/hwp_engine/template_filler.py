"""템플릿 채우기 — 정부과제 양식의 구조를 인식하고 적절한 위치에 내용을 삽입.

정부과제 양식 구조:
  - 본문 표의 B열에 모든 섹션이 하나의 셀로 들어있음
  - ※ 안내문 (파란색) → 삭제 대상
  - ◦, - 구조 기호 → 보존하고 뒤에 내용 추가
  - 섹션 제목 (1. 문제 인식, 2. 실현 가능성...) → 보존

사용법::

    filler = TemplateFiller(hwp)
    structure = filler.analyze_template()
    summary = filler.get_fillable_summary(structure)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import structlog

logger = structlog.get_logger()


@dataclass
class InfoField:
    """기업정보 필드 (패턴 A)."""

    table_idx: int
    cell_addr: str  # "D1" 등
    label: str      # "기업명", "대표자명" 등
    example: str    # "주식회사 0000" 등 현재 값
    moves: int      # A1에서 이 셀까지 TableRightCell 횟수


@dataclass
class BodySection:
    """본문 섹션 — 2열 표의 B열."""

    table_idx: int          # 이 섹션이 속한 표 인덱스
    section_num: int        # 1, 2, 3, 4
    title: str              # "1. 문제 인식(Problem)_창업 아이템의 필요성"
    guide_text: str = ""    # "※ 안내문..."
    markers: list[dict[str, Any]] = field(default_factory=list)
    # markers: [{"type": "◦" 또는 "-", "para_index": int, "text": "◦" 등}]


@dataclass
class NarrativeSection:
    """서술형 섹션 (패턴 B) — 하위 호환용."""

    table_idx: int
    title: str
    sub_items: list[str] = field(default_factory=list)
    guide_text: str = ""
    full_text: str = ""


@dataclass
class DataTable:
    """데이터 표 (예산, 일정 등 다열 구조)."""

    table_idx: int
    title: str  # 표 제목 (첫 행 또는 헤더에서 추출)
    headers: list[str] = field(default_factory=list)
    rows: int = 0
    cols: int = 0
    empty_cells: list[dict[str, Any]] = field(default_factory=list)
    # empty_cells: [{"row": int, "col": int, "header": str}]


@dataclass
class TemplateStructure:
    """양식 전체 구조."""

    info_tables: list[dict[str, Any]] = field(default_factory=list)
    narrative_sections: list[NarrativeSection] = field(default_factory=list)
    other_tables: list[int] = field(default_factory=list)
    body_sections: list[BodySection] = field(default_factory=list)
    data_tables: list[DataTable] = field(default_factory=list)


# 예시 데이터 패턴 (교체 대상)
_EXAMPLE_PATTERNS = re.compile(
    r"^("
    r"주식회사\s*0+|"
    r"0{3,}[-.]?0{2,}[-.]?0{3,}|"
    r"0{3,}[-.]0{4,}[-.]0{4,}|"
    r"20\d{2}[.\s]*0{2}[.\s]*0{2}|"
    r"0{2,}|"
    r"\(주\)\s*0+|"
    r".{0,5}@.{0,5}\..{0,3}"
    r")$"
)

# 섹션 제목 패턴
_SECTION_TITLE = re.compile(
    r"^\s*(\d+)\.\s*(문제\s*인식|실현\s*가능성|성장\s*전략|팀\s*구성)"
)


class TemplateFiller:
    """정부과제 양식 구조를 분석하고 적절한 위치에 내용을 삽입한다."""

    def __init__(self, hwp: Any) -> None:
        self._hwp = hwp

    # ------------------------------------------------------------------
    # 분석
    # ------------------------------------------------------------------

    def analyze_template(self) -> TemplateStructure:
        """양식 전체를 분석하여 구조를 반환한다."""
        hwp = self._hwp
        structure = TemplateStructure()

        # 표 개수
        ctrl = hwp.HeadCtrl
        table_count = 0
        while ctrl:
            if ctrl.CtrlID == "tbl":
                table_count += 1
            ctrl = ctrl.Next

        for ti in range(table_count):
            try:
                hwp.get_into_nth_table(ti)
                df = hwp.table_to_df()
                rows, cols_count = df.shape

                if cols_count >= 3 and rows >= 1:
                    # 패턴 A: 다열 기업정보 표
                    info = self._analyze_info_table(ti, df)
                    if info["fields"]:
                        structure.info_tables.append(info)
                    else:
                        # 기업정보가 아닌 다열 표 (예산, 일정 등)
                        data_table = self._analyze_data_table(ti, df)
                        if data_table:
                            structure.data_tables.append(data_table)

                elif cols_count == 2:
                    # 2열 표 — 본문 표일 수 있음 (A열=제목, B열=내용)
                    a_text = str(df.columns[0]).strip()
                    if _SECTION_TITLE.match(a_text):
                        self._analyze_body_cell(ti, a_text, structure)

                elif cols_count == 1:
                    # 1열 표 — 서술형 또는 안내문
                    section = self._analyze_narrative(ti, df)
                    if section:
                        structure.narrative_sections.append(section)

                else:
                    structure.other_tables.append(ti)

            except Exception:
                pass

        logger.info(
            "양식 분석 완료",
            info_tables=len(structure.info_tables),
            data_tables=len(structure.data_tables),
            body_sections=len(structure.body_sections),
            narratives=len(structure.narrative_sections),
        )
        return structure

    def _analyze_body_cell(self, table_idx: int, title_text: str, structure: TemplateStructure) -> None:
        """본문 표(2열)의 B1 셀을 단락 단위로 분석한다.

        각 2열 표는 하나의 섹션에 대응:
          A열 = 제목 ("1. 문제 인식...")
          B열 = 안내문(※) + 구조 기호(◦, -)
        """
        hwp = self._hwp
        hwp.get_into_nth_table(table_idx)
        hwp.TableRightCell()  # B1으로 이동

        # 섹션 번호 추출
        m = _SECTION_TITLE.match(title_text)
        section_num = int(m.group(1)) if m else 0

        section = BodySection(
            table_idx=table_idx,
            section_num=section_num,
            title=title_text.replace("\r\n", "").replace("\n", ""),
        )

        # 단락 텍스트 수집
        paragraphs = self._read_cell_paragraphs()

        for para_idx, text in enumerate(paragraphs):
            stripped = text.strip()
            if not stripped:
                continue

            # 다른 섹션 제목이 나타나면 이 셀의 분석 영역 끝
            # (병합 셀의 경우 여러 섹션이 연속으로 들어있을 수 있음)
            m2 = _SECTION_TITLE.match(stripped)
            if m2 and int(m2.group(1)) != section_num:
                break

            # ※ 안내문 (첫 번째만)
            if stripped.startswith("※") and not section.guide_text:
                section.guide_text = stripped
                continue

            # ◦/○ 소제목 (빈 마커 또는 짧은 텍스트 포함)
            if stripped in ("◦", "○") or (
                (stripped.startswith("◦") or stripped.startswith("○"))
                and len(stripped) <= 5
            ):
                section.markers.append({
                    "type": "◦",
                    "para_index": para_idx,
                    "text": stripped,
                })
                continue

            # - 하위항목
            if stripped == "-" or (stripped.startswith("-") and len(stripped) <= 5):
                section.markers.append({
                    "type": "-",
                    "para_index": para_idx,
                    "text": stripped,
                })
                continue

        # 마커 유무와 관계없이 본문 섹션으로 추가
        structure.body_sections.append(section)

    def _read_cell_paragraphs(self) -> list[str]:
        """현재 셀의 모든 단락 텍스트를 반환한다."""
        hwp = self._hwp

        try:
            import win32clipboard
        except ImportError:
            return []

        hwp.HAction.Run("MoveColBegin")
        paragraphs: list[str] = []
        seen_positions: set[tuple[int, ...]] = set()

        for _ in range(500):  # 안전 상한
            cur_pos = hwp.GetPos()
            pos_key = tuple(cur_pos) if isinstance(cur_pos, (list, tuple)) else (cur_pos,)
            if pos_key in seen_positions:
                break
            seen_positions.add(pos_key)

            # 단락 끝까지 선택 후 클립보드로 복사
            hwp.HAction.Run("MoveSelParaEnd")
            hwp.HAction.Run("Copy")

            try:
                win32clipboard.OpenClipboard()
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
            except Exception:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
                text = ""

            paragraphs.append(text.strip())

            # 다음 단락으로
            hwp.HAction.Run("MoveParaEnd")
            hwp.HAction.Run("MoveRight")

        return paragraphs

    # ------------------------------------------------------------------
    # 채우기
    # ------------------------------------------------------------------

    def _replace_text(self, old: str, new: str) -> bool:
        """문서 전체에서 old를 new로 찾아바꾸기한다.

        find_replace 방식은 원본의 글머리표, 들여쓰기, 문단 서식을
        100% 보존하면서 텍스트만 교체한다.
        """
        hwp = self._hwp
        pset = hwp.HParameterSet.HFindReplace
        hwp.HAction.GetDefault("AllReplace", pset.HSet)
        pset.FindString = old
        pset.ReplaceString = new
        pset.IgnoreMessage = 1
        pset.FindType = 1  # 전체 범위
        result = hwp.HAction.Execute("AllReplace", pset.HSet)
        return bool(result)

    def _enter_table(self, table_idx: int) -> bool:
        """표에 진입한다. 실패하면 False를 반환한다."""
        result = self._hwp.get_into_nth_table(table_idx)
        if result is False:
            logger.warning("표 진입 실패, 쓰기 건너뜀", table_idx=table_idx)
            return False
        return True

    def fill_info_field(self, table_idx: int, moves: int, text: str, example: str = "") -> None:
        """기업정보 표의 특정 셀에 값을 채운다.

        하이브리드: 예시값이 있으면 find_replace로 교체 (서식 보존 + 셀 위치 무관).
        예시값이 없으면 SelectAll 폴백.
        """
        hwp = self._hwp

        # 1차: 예시값으로 find_replace (병합 셀에서도 정확하게 작동)
        if example and len(example) > 1:
            result = self._replace_text(example, text)
            if result:
                return

        # 2차: SelectAll 폴백
        if not self._enter_table(table_idx):
            return
        for _ in range(moves):
            hwp.TableRightCell()

        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Delete")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    def fill_body_section(self, table_idx: int, section: BodySection, contents: list[str]) -> None:
        """본문 섹션의 내용을 채운다.

        이 양식은 2열 헤더 표(A=제목, B=빈칸) + 1열 안내문 표(※)로 구성.
        내용은 ※ 안내문이 있는 다음 표(table_idx+1)에 넣어야 한다.
        find_replace로 ※ 안내문을 교체하여 서식을 보존한다.
        """
        hwp = self._hwp

        # ※ 안내문을 find_replace로 교체 (서식 보존)
        if section.guide_text:
            guide = section.guide_text.strip()
            find_key = guide[:40] if len(guide) > 40 else guide
            full_text = "\n".join(contents)
            replace_text = full_text[:40] if len(full_text) > 40 else full_text
            result = self._replace_text(find_key, replace_text)
            if result:
                logger.info("섹션 채우기 완료 (find_replace)",
                            section=section.section_num, title=section.title[:30],
                            filled=len(contents), total_markers=len(section.markers))
                return

        # find_replace 실패 시: 다음 표(안내문 표)에 SelectAll로 넣기
        next_table = table_idx + 1
        if not self._enter_table(next_table):
            # 다음 표도 없으면 현재 표 B1에 시도
            if not self._enter_table(table_idx):
                return
            hwp.TableRightCell()

        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Delete")

        full_text = "\n".join(contents)
        lines = full_text.split("\n")
        for i, line in enumerate(lines):
            if line.strip():
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = line
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            if i < len(lines) - 1:
                hwp.HAction.Run("BreakPara")

        logger.info("섹션 채우기 완료",
                     section=section.section_num,
                     title=section.title[:30],
                     filled=len(contents),
                     total_markers=len(section.markers))

    def _delete_guide_paragraphs(self, section: BodySection) -> None:
        """현재 B1 셀에서 해당 섹션의 ※ 안내문 단락을 찾아 삭제한다."""
        hwp = self._hwp

        if not section.guide_text:
            return

        # B1 처음부터 순회하며 ※ 안내문 찾기
        hwp.HAction.Run("MoveColBegin")
        guide_keyword = section.guide_text[:20]  # 안내문 식별용 앞부분

        try:
            import win32clipboard
        except ImportError:
            return

        seen: set[tuple[int, ...]] = set()
        for _ in range(500):
            cur_pos = hwp.GetPos()
            pos_key = tuple(cur_pos) if isinstance(cur_pos, (list, tuple)) else (cur_pos,)
            if pos_key in seen:
                break
            seen.add(pos_key)

            # 단락 텍스트 읽기
            hwp.HAction.Run("MoveSelParaEnd")
            hwp.HAction.Run("Copy")

            try:
                win32clipboard.OpenClipboard()
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
            except Exception:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
                text = ""

            stripped = text.strip()

            if stripped.startswith("※") and guide_keyword[:15] in stripped:
                # 이 단락 전체를 선택하여 삭제
                hwp.HAction.Run("MoveParaBegin")
                hwp.HAction.Run("MoveSelParaEnd")
                hwp.HAction.Run("MoveSelRight")  # 줄바꿈 포함
                hwp.HAction.Run("Delete")
                logger.debug("안내문 삭제", text=stripped[:40])
                # 삭제 후 위치가 바뀌므로 처음부터 다시 시작
                hwp.HAction.Run("MoveColBegin")
                seen.clear()
                continue

            # 다음 단락으로
            hwp.HAction.Run("MoveParaEnd")
            hwp.HAction.Run("MoveRight")

    def _fill_markers(self, section: BodySection, contents: list[str]) -> None:
        """현재 B1 셀의 ◦/- 마커를 찾아 내용을 채운다.

        각 2열 표의 B1 셀은 해당 섹션 전용이므로
        바로 ◦/- 마커를 순회한다.
        """
        hwp = self._hwp

        if not section.markers or not contents:
            return

        try:
            import win32clipboard
        except ImportError:
            return

        hwp.HAction.Run("MoveColBegin")
        content_idx = 0
        seen: set[tuple[int, ...]] = set()

        for _ in range(500):
            if content_idx >= len(contents):
                break

            cur_pos = hwp.GetPos()
            pos_key = tuple(cur_pos) if isinstance(cur_pos, (list, tuple)) else (cur_pos,)
            if pos_key in seen:
                break
            seen.add(pos_key)

            # 선택 해제 후 단락 텍스트 읽기
            hwp.HAction.Run("Cancel")
            hwp.HAction.Run("MoveSelParaEnd")
            hwp.HAction.Run("Copy")

            try:
                win32clipboard.OpenClipboard()
                text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
            except Exception:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
                text = ""

            stripped = text.strip()

            # ◦/○ 또는 - 마커 찾기 (길이 5 이하)
            is_circle = stripped in ("◦", "○") or (
                (stripped.startswith("◦") or stripped.startswith("○"))
                and len(stripped) <= 5
            )
            is_dash = stripped == "-" or (stripped.startswith("-") and len(stripped) <= 5)

            if is_circle or is_dash:
                # 단락 끝으로 이동
                hwp.HAction.Run("MoveParaEnd")

                # 밑줄/취소선 서식 제거
                self._clear_formatting(hwp)

                # 마커 뒤에 내용 추가
                content_text = " " + contents[content_idx]
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = content_text
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                marker_type = "◦" if is_circle else "-"
                logger.debug("마커에 내용 추가", marker=marker_type, content=content_text[:30])
                content_idx += 1

            # 다음 단락
            hwp.HAction.Run("MoveParaEnd")
            hwp.HAction.Run("MoveRight")

        logger.info(
            "섹션 채우기 완료",
            section=section.section_num,
            title=section.title[:30],
            filled=content_idx,
            total_markers=len(section.markers),
        )

    @staticmethod
    def _clear_formatting(hwp: Any) -> None:
        """현재 위치의 밑줄/취소선/문단 테두리 서식을 제거한다."""
        # 1. 글자 서식: 밑줄, 취소선 제거
        try:
            act = hwp.HParameterSet.HCharShape
            hwp.HAction.GetDefault("CharShape", act.HSet)
            act.UnderlineType = 0   # 밑줄 없음
            act.StrikeOutType = 0   # 취소선 없음
            hwp.HAction.Execute("CharShape", act.HSet)
        except Exception:
            pass

        # 2. 문단 테두리 제거 (※ 안내문 서식 상속 방지)
        try:
            pset = hwp.get_parashape()
            hwp.HAction.GetDefault("ParaShape", pset.HSet)
            # 문단 테두리 속성 초기화
            pset.LeftBorder = 0
            pset.RightBorder = 0
            pset.TopBorder = 0
            pset.BottomBorder = 0
            pset.ConnectBorder = 0
            hwp.HAction.Execute("ParaShape", pset.HSet)
        except Exception:
            pass

        # 3. 셀 대각선 테두리 제거 (CellBorderFill 사용)
        try:
            cbf = hwp.HParameterSet.HCellBorderFill
            hwp.HAction.GetDefault("CellBorderFill", cbf.HSet)
            # Diagonal 속성이 있으면 0으로 설정
            try:
                cbf.DiagonalType = 0
            except AttributeError:
                pass
            try:
                cbf.Diagonal = 0
            except AttributeError:
                pass
            hwp.HAction.Execute("CellBorderFill", cbf.HSet)
        except Exception:
            pass

    def fill_body_narrative(self, table_idx: int, section: BodySection, content: str) -> None:
        """마커 없는 본문 섹션에 내용을 작성한다.

        find_replace로 ※ 안내문을 교체하여 서식 보존.
        실패 시 다음 표(안내문 표)에 SelectAll 폴백.
        """
        hwp = self._hwp

        # find_replace로 ※ 안내문 교체 (서식 보존)
        replaced = False
        if section.guide_text:
            guide = section.guide_text.strip()
            if len(guide) > 5:
                find_key = guide[:40] if len(guide) > 40 else guide
                replace_val = content[:40] if len(content) > 40 else content
                result = self._replace_text(find_key, replace_val)
                if result:
                    replaced = True
                # 나머지 안내문 조각 제거
                for part in section.guide_text.split("\n"):
                    part = part.strip()
                    if part and len(part) > 5:
                        self._replace_text(part, "")

        if not replaced:
            # 폴백: 다음 표(안내문 표)에 넣기
            next_table = table_idx + 1
            if self._enter_table(next_table):
                hwp.HAction.Run("SelectAll")
                hwp.HAction.Run("Delete")
                lines = content.split("\n")
                for i, line in enumerate(lines):
                    if line.strip():
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.HParameterSet.HInsertText.Text = line
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    if i < len(lines) - 1:
                        hwp.HAction.Run("BreakPara")

        logger.info("본문 서술형 채우기 완료", section=section.section_num, title=section.title[:30])

    def fill_narrative(self, table_idx: int, content: str) -> None:
        """서술형 셀의 안내문을 교체하고 내용을 작성한다.

        하이브리드: find_replace로 ※ 안내문을 제거/교체하여 서식 보존.
        find_replace 실패 시에만 SelectAll 폴백.
        """
        hwp = self._hwp

        # 표의 원본 텍스트에서 ※ 안내문 부분을 find_replace로 교체
        # 이렇게 하면 제목(①, ②), 번호(1), 2)) 등 서식이 보존됨
        replaced = False

        # 1차: ※ 안내문을 찾아서 내용으로 교체
        if not self._enter_table(table_idx):
            return

        # 현재 셀의 전체 텍스트 읽기 (표 내용 파악)
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Copy")
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            current = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except Exception:
            current = ""
        hwp.Cancel()

        current = (current or "").strip()

        # ※ 안내문 부분을 find_replace로 교체
        if current:
            lines = current.split("\n")
            for line in lines:
                stripped = line.strip()
                if stripped.startswith("※") and len(stripped) > 5:
                    # ※ 안내문을 내용으로 교체 (첫 번째만)
                    if not replaced:
                        self._replace_text(stripped[:40], content[:40] if len(content) > 40 else content)
                        replaced = True
                    else:
                        # 나머지 ※ 안내문은 삭제
                        self._replace_text(stripped[:40], "")

        if not replaced:
            # find_replace 실패 시 SelectAll 폴백
            if self._enter_table(table_idx):
                hwp.HAction.Run("SelectAll")
                hwp.HAction.Run("Delete")
                lines = content.split("\n")
                for i, line in enumerate(lines):
                    if line.strip():
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.HParameterSet.HInsertText.Text = line
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    if i < len(lines) - 1:
                        hwp.HAction.Run("BreakPara")

    # ------------------------------------------------------------------
    # 내부 분석 로직
    # ------------------------------------------------------------------

    def _analyze_info_table(self, table_idx: int, df: Any) -> dict[str, Any]:
        """기업정보 표를 분석한다."""
        fields: list[dict[str, str]] = []
        cols = list(df.columns)
        all_vals = list(cols)

        for ri in range(df.shape[0]):
            all_vals.extend(list(df.iloc[ri].values))

        for i, val in enumerate(all_vals):
            s = str(val).strip()
            if _EXAMPLE_PATTERNS.match(s):
                label = ""
                for j in range(i - 1, -1, -1):
                    candidate = str(all_vals[j]).strip()
                    if candidate and not _EXAMPLE_PATTERNS.match(candidate) and len(candidate) < 20:
                        label = candidate
                        break
                fields.append({
                    "table_idx": table_idx,
                    "index": i,
                    "label": label,
                    "example": s,
                })

        return {"table_idx": table_idx, "fields": fields}

    def _analyze_narrative(self, table_idx: int, df: Any) -> NarrativeSection | None:
        """서술형 셀(1열 표)을 분석한다."""
        try:
            full = str(df.columns[0])
            if df.shape[0] > 0:
                for ri in range(df.shape[0]):
                    full += "\n" + str(df.iloc[ri, 0])
        except Exception:
            return None

        full = full.strip()
        if not full or len(full) < 5:
            return None

        # ※ 안내문 전용 표는 채우기 대상이 아님 (삭제 대상)
        if full.startswith("※"):
            return None

        lines = full.split("\n")
        title = lines[0].strip() if lines else ""

        # 제목만 있고 내용이 없는 표 (섹션 헤더 등)도 제외
        if len(lines) <= 1 and len(title) < 10:
            return None

        sub_items = []
        for line in lines:
            stripped = line.strip()
            if re.match(r"^\d+\)\s*$", stripped) or re.match(r"^\d+\)\s+\S", stripped):
                sub_items.append(stripped)

        guide = ""
        for line in lines:
            if line.strip().startswith("※"):
                guide = line.strip()
                break

        if not title:
            return None

        return NarrativeSection(
            table_idx=table_idx,
            title=title,
            sub_items=sub_items,
            guide_text=guide,
            full_text=full,
        )

    def _analyze_data_table(self, table_idx: int, df: Any) -> DataTable | None:
        """다열 데이터 표(예산, 일정 등)를 분석한다.

        예시 데이터(파란색 텍스트 등)가 들어있는 셀과 빈 셀을 모두 감지한다.
        헤더 행과 합계 행은 제외한다.
        """
        rows, cols = df.shape
        if rows < 1 or cols < 2:
            return None

        headers = [str(c).strip() for c in df.columns]

        # 표 제목 추정
        title = headers[0]
        for candidate in headers:
            if len(candidate) > 3 and not _EXAMPLE_PATTERNS.match(candidate):
                title = candidate
                break

        # 전체 표 데이터를 예시로 수집 (헤더 제외, 합계 제외)
        example_rows: list[list[str]] = []
        for ri in range(rows):
            row_vals = [str(df.iloc[ri, ci]).strip() for ci in range(cols)]
            first_val = row_vals[0] if row_vals else ""
            # 합계 행 또는 빈 행("…", "...") 건너뛰기
            if first_val in ("합  계", "합계", "…", "...", ""):
                continue
            example_rows.append(row_vals)

        if not example_rows:
            return None

        # 모든 데이터 행을 교체 대상으로 설정
        empty_cells: list[dict[str, Any]] = []
        for ri_orig in range(rows):
            row_vals = [str(df.iloc[ri_orig, ci]).strip() for ci in range(cols)]
            first_val = row_vals[0] if row_vals else ""
            if first_val in ("합  계", "합계", "…", "...", ""):
                continue
            for ci in range(cols):
                header = headers[ci] if ci < len(headers) else f"col{ci}"
                empty_cells.append({
                    "row": ri_orig + 1,  # 1-based (헤더 제외)
                    "col": ci,
                    "header": header,
                    "example": row_vals[ci],
                })

        return DataTable(
            table_idx=table_idx,
            title=title,
            headers=headers,
            rows=rows,
            cols=cols,
            empty_cells=empty_cells,
        )

    def fill_data_cell(self, table_idx: int, row: int, col: int, text: str, example: str = "") -> None:
        """데이터 표의 특정 셀에 값을 채운다.

        하이브리드: 예시값이 있으면 find_replace (서식 보존 + 병합 셀 안전).
        없으면 SelectAll 폴백.
        """
        hwp = self._hwp

        # 1차: 예시값으로 find_replace
        if example and len(example) > 1:
            result = self._replace_text(example, text)
            if result:
                return

        # 2차: SelectAll 폴백
        if not self._enter_table(table_idx):
            return

        total_moves = row * self._get_col_count(table_idx) + col
        for _ in range(total_moves):
            hwp.TableRightCell()

        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Delete")

        # 텍스트 색상을 검정으로
        try:
            act = hwp.HParameterSet.HCharShape
            hwp.HAction.GetDefault("CharShape", act.HSet)
            act.TextColor = 0x00000000
            hwp.HAction.Execute("CharShape", act.HSet)
        except Exception:
            pass

        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    def _get_col_count(self, table_idx: int) -> int:
        """표의 열 수를 반환한다."""
        hwp = self._hwp
        hwp.get_into_nth_table(table_idx)
        df = hwp.table_to_df()
        return df.shape[1]

    def get_fillable_summary(self, structure: TemplateStructure) -> list[dict[str, Any]]:
        """채울 수 있는 항목 요약을 반환한다."""
        items: list[dict[str, Any]] = []

        for info in structure.info_tables:
            for f in info["fields"]:
                items.append({
                    "type": "info",
                    "table_idx": f["table_idx"],
                    "label": f["label"],
                    "example": f["example"],
                    "index": f["index"],
                })

        # 본문 섹션 (새로운 방식)
        for bs in structure.body_sections:
            marker_count = len(bs.markers)
            circle_count = len([m for m in bs.markers if m["type"] == "◦"])
            dash_count = len([m for m in bs.markers if m["type"] == "-"])
            items.append({
                "type": "body_section",
                "table_idx": bs.table_idx,
                "section_num": bs.section_num,
                "title": bs.title,
                "guide": bs.guide_text[:80] if bs.guide_text else "",
                "markers": bs.markers,
                "marker_count": marker_count,
                "circle_count": circle_count,
                "dash_count": dash_count,
            })

        # 하위 호환: narrative_sections
        for ns in structure.narrative_sections:
            items.append({
                "type": "narrative",
                "table_idx": ns.table_idx,
                "title": ns.title,
                "sub_items": ns.sub_items,
                "guide": ns.guide_text[:50],
            })

        # 데이터 표 (예산, 일정 등)
        for dt in structure.data_tables:
            items.append({
                "type": "data_table",
                "table_idx": dt.table_idx,
                "title": dt.title,
                "headers": dt.headers,
                "rows": dt.rows,
                "cols": dt.cols,
                "empty_cells": dt.empty_cells,
                "empty_count": len(dt.empty_cells),
            })

        return items

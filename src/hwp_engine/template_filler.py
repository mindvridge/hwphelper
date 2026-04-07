"""템플릿 채우기 — 정부과제 양식의 구조를 인식하고 적절한 위치에 내용을 삽입.

정부과제 양식의 두 가지 패턴:

패턴 A (기업정보 표): 다열 표, "주식회사 0000" 같은 예시값 교체
패턴 B (서술형 셀): 1셀 표, 내부에 "1)", "2)", "※" 구조 → 안내문 삭제 후 내용 작성

사용법::

    filler = TemplateFiller(hwp)
    sections = filler.analyze_template()
    # → [{"type": "info", "table_idx": 1, "fields": [...]},
    #    {"type": "narrative", "table_idx": 6, "title": "① 창업동기", ...}]
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
class NarrativeSection:
    """서술형 섹션 (패턴 B)."""

    table_idx: int
    title: str       # "① 창업동기" 등
    sub_items: list[str] = field(default_factory=list)  # ["1)", "2)"]
    guide_text: str = ""  # "※ 안내문..."
    full_text: str = ""   # 셀 전체 텍스트


@dataclass
class TemplateStructure:
    """양식 전체 구조."""

    info_tables: list[dict[str, Any]] = field(default_factory=list)  # 기업정보 표
    narrative_sections: list[NarrativeSection] = field(default_factory=list)  # 서술형
    other_tables: list[int] = field(default_factory=list)  # 기타 표


# 예시 데이터 패턴 (교체 대상)
_EXAMPLE_PATTERNS = re.compile(
    r"^("
    r"주식회사\s*0+|"
    r"0{3,}[-.]?0{2,}[-.]?0{3,}|"  # 000-00-00000
    r"0{3,}[-.]0{4,}[-.]0{4,}|"    # 010-0000-0000
    r"20\d{2}[.\s]*0{2}[.\s]*0{2}|"  # 2026.02.12
    r"0{2,}|"
    r"\(주\)\s*0+|"
    r".{0,5}@.{0,5}\..{0,3}"  # 이메일
    r")$"
)


class TemplateFiller:
    """정부과제 양식 구조를 분석하고 적절한 위치에 내용을 삽입한다."""

    def __init__(self, hwp: Any) -> None:
        self._hwp = hwp

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
                elif cols_count <= 2:
                    # 패턴 B: 서술형 셀 (1~2열)
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
            narratives=len(structure.narrative_sections),
        )
        return structure

    def fill_info_field(self, table_idx: int, moves: int, text: str) -> None:
        """기업정보 표의 특정 셀에 값을 채운다."""
        hwp = self._hwp
        hwp.get_into_nth_table(table_idx)
        for _ in range(moves):
            hwp.TableRightCell()

        # 셀 내용 교체
        hwp.HAction.Run("MoveColBegin")
        hwp.HAction.Run("MoveSelColEnd")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    def fill_narrative(self, table_idx: int, content: str) -> None:
        """서술형 셀의 안내문을 삭제하고 내용을 작성한다.

        기존 구조:
            ① 창업동기
              1)
              2)
            ※ 안내문...

        작성 후:
            ① 창업동기
              1) [LLM이 생성한 내용]
              2) [LLM이 생성한 내용]
        """
        hwp = self._hwp
        hwp.get_into_nth_table(table_idx)

        # 셀 전체 텍스트를 교체
        hwp.HAction.Run("MoveColBegin")
        hwp.HAction.Run("MoveSelColEnd")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = content
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

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

        # 라벨-값 쌍 찾기
        for i, val in enumerate(all_vals):
            s = str(val).strip()
            if _EXAMPLE_PATTERNS.match(s):
                # 왼쪽에서 라벨 찾기
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
        """서술형 셀을 분석한다."""
        # 첫 번째 열의 전체 텍스트
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

        # 제목 추출 (첫 줄에서 ①, ②, 1., 2. 등)
        lines = full.split("\n")
        title = lines[0].strip() if lines else ""

        # 번호 항목 찾기
        sub_items = []
        for line in lines:
            stripped = line.strip()
            if re.match(r"^\d+\)\s*$", stripped) or re.match(r"^\d+\)\s+\S", stripped):
                sub_items.append(stripped)

        # 안내문 찾기
        guide = ""
        for line in lines:
            if line.strip().startswith("※"):
                guide = line.strip()
                break

        # 서술형 섹션으로 인식하려면 제목이 있어야 함
        if not title:
            return None

        return NarrativeSection(
            table_idx=table_idx,
            title=title,
            sub_items=sub_items,
            guide_text=guide,
            full_text=full,
        )

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

        for ns in structure.narrative_sections:
            items.append({
                "type": "narrative",
                "table_idx": ns.table_idx,
                "title": ns.title,
                "sub_items": ns.sub_items,
                "guide": ns.guide_text[:50],
            })

        return items

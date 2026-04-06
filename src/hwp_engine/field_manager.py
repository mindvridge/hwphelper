"""누름틀 필드 관리 — HWP 문서의 누름틀(필드)을 조회·채우기·매핑.

한/글의 누름틀(FormField)은 ``{{이름}}`` 형태로 문서에 삽입되며,
``GetFieldList`` / ``PutFieldText`` API로 프로그래밍 방식 접근이 가능하다.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

import structlog

if TYPE_CHECKING:
    from .com_controller import HwpController

logger = structlog.get_logger()


@dataclass
class FieldInfo:
    """누름틀 필드 정보."""

    name: str
    direction: str = ""     # 안내문 (툴팁)
    memo: str = ""          # 메모
    value: str = ""         # 현재 값
    field_type: str = ""    # 필드 종류


class FieldManager:
    """HWP 누름틀 필드를 조회하고 값을 채운다."""

    def __init__(self, hwp_ctrl: HwpController) -> None:
        self._ctrl = hwp_ctrl

    # ------------------------------------------------------------------
    # public API
    # ------------------------------------------------------------------

    def list_fields(self) -> list[FieldInfo]:
        """문서 내 모든 누름틀 필드를 조회한다."""
        hwp = self._ctrl.hwp
        fields: list[FieldInfo] = []

        try:
            # GetFieldList → 줄바꿈으로 구분된 필드명 문자열
            field_list_str = hwp.GetFieldList(0, 0)  # 옵션: 0=전체
            if not field_list_str:
                logger.info("필드 없음")
                return fields

            # 여러 구분자 처리 (\r\n, \n, 탭 등)
            field_names: list[str] = []
            for sep in ["\r\n", "\n", "\x02"]:
                if sep in field_list_str:
                    field_names = [n.strip() for n in field_list_str.split(sep) if n.strip()]
                    break
            if not field_names:
                field_names = [field_list_str.strip()] if field_list_str.strip() else []

            for name in field_names:
                value = ""
                try:
                    value = hwp.GetFieldText(name) or ""
                except Exception:
                    pass

                fields.append(FieldInfo(name=name, value=value))

        except Exception:
            logger.warning("필드 목록 조회 실패")

        logger.info("필드 조회 완료", count=len(fields))
        return fields

    def fill_field(self, field_name: str, text: str) -> bool:
        """누름틀 필드에 값을 채운다."""
        hwp = self._ctrl.hwp
        try:
            hwp.PutFieldText(field_name, text)
            logger.info("필드 채우기 완료", name=field_name, length=len(text))
            return True
        except Exception:
            logger.warning("필드 채우기 실패", name=field_name)
            return False

    def fill_fields_batch(self, fills: dict[str, str]) -> dict[str, bool]:
        """여러 필드에 값을 일괄 채운다.

        Returns
        -------
        dict[str, bool]
            필드명 → 성공 여부 매핑.
        """
        results: dict[str, bool] = {}
        for name, text in fills.items():
            results[name] = self.fill_field(name, text)

        success = sum(results.values())
        logger.info("일괄 필드 채우기 완료", total=len(fills), success=success)
        return results

    def create_field_template(
        self,
        table_idx: int,
        field_mapping: dict[str, tuple[int, int]],
    ) -> None:
        """표의 특정 셀에 누름틀 필드를 생성한다.

        Parameters
        ----------
        table_idx : int
            대상 표 인덱스.
        field_mapping : dict[str, tuple[int, int]]
            필드명 → (row, col) 매핑.
            예: {"사업명": (0, 1), "기관명": (1, 1)}
        """
        hwp = self._ctrl.hwp

        for field_name, (row, col) in field_mapping.items():
            try:
                # 셀로 이동
                self._move_to_cell(table_idx, row, col)

                # 셀 내용 전체 선택 후 필드 삽입
                hwp.HAction.Run("SelectAll")

                # 누름틀 삽입
                act = hwp.CreateAction("InsertFieldRevision")
                pset = act.CreateSet()
                act.GetDefault(pset)
                pset.SetItem("FieldName", field_name)
                pset.SetItem("Direction", field_name)  # 안내문
                act.Execute(pset)

                logger.info("필드 생성", name=field_name, row=row, col=col)
            except Exception:
                logger.warning("필드 생성 실패", name=field_name, row=row, col=col)

    # ------------------------------------------------------------------
    # 내부 유틸리티
    # ------------------------------------------------------------------

    def _move_to_cell(self, table_idx: int, row: int, col: int) -> None:
        """표 내 특정 셀로 커서를 이동한다."""
        hwp = self._ctrl.hwp
        try:
            if hasattr(hwp, "ShapeObjTableSelCell"):
                hwp.ShapeObjTableSelCell(table_idx, row, col)
                return
        except Exception:
            pass

        # 폴백
        hwp.MovePos(2)
        ctrl = hwp.HeadCtrl
        idx = 0
        while ctrl:
            if ctrl.CtrlID == "tbl":
                if idx == table_idx:
                    try:
                        hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    except Exception:
                        pass
                    break
                idx += 1
            ctrl = ctrl.Next

        cols = 1
        if ctrl:
            try:
                cols = ctrl.ColCount
            except AttributeError:
                pass
        target = row * cols + col
        for _ in range(target):
            hwp.HAction.Run("TableRightCell")

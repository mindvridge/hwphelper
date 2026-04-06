"""서식 검증 — 정부과제별 서식 규정(폰트, 크기, 자간, 행간, 여백) 검사 및 자동 수정.

정부지원사업 계획서는 과제별로 지정된 서식 규정이 있다.
이 모듈은 HWP COM API를 통해 문서의 실제 서식을 읽어서 규정 위반 여부를 검사하고,
위반 항목을 자동으로 수정하는 기능을 제공한다.

사용법::

    checker = FormatChecker("config/format_rules.yaml")
    report = checker.check_document(hwp_ctrl, "예비창업패키지")
    if not report.passed:
        fixes = checker.auto_fix(hwp_ctrl, "예비창업패키지")
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING, Any

import structlog
import yaml

if TYPE_CHECKING:
    from src.hwp_engine.com_controller import HwpController

logger = structlog.get_logger()


# ------------------------------------------------------------------
# 데이터 클래스
# ------------------------------------------------------------------


class Severity(str, Enum):
    WARNING = "warning"
    ERROR = "error"


@dataclass
class FormatWarning:
    """서식 경고 — 규정 위반이지만 자동 수정 가능."""

    location: str           # "표 1, (2,3)" 또는 "문단 5" 등
    rule: str               # "font_name", "font_size", "char_spacing" 등
    current_value: str      # 현재 값
    expected: str           # 규정 값
    auto_fixable: bool = True
    severity: Severity = Severity.WARNING


@dataclass
class FormatError:
    """서식 에러 — 심각한 위반 (자동 수정 어려움)."""

    location: str
    rule: str
    current_value: str
    expected: str
    message: str = ""


@dataclass
class FixAction:
    """자동 수정 액션 로그."""

    location: str
    rule: str
    old_value: str
    new_value: str
    success: bool = True


@dataclass
class FormatReport:
    """서식 검증 보고서."""

    program_name: str
    passed: bool = True
    total_checks: int = 0
    passed_checks: int = 0
    warnings: list[FormatWarning] = field(default_factory=list)
    errors: list[FormatError] = field(default_factory=list)

    @property
    def failed_checks(self) -> int:
        return self.total_checks - self.passed_checks

    def summary(self) -> str:
        status = "PASS" if self.passed else "FAIL"
        return (
            f"[{status}] {self.program_name}: "
            f"{self.passed_checks}/{self.total_checks} 통과, "
            f"경고 {len(self.warnings)}건, 에러 {len(self.errors)}건"
        )


@dataclass
class ProgramRules:
    """정부과제별 서식 규정."""

    name: str
    fonts_allowed: list[str] = field(default_factory=lambda: ["맑은 고딕"])
    body_size_range: tuple[float, float] = (10.0, 12.0)
    title_size_range: tuple[float, float] = (13.0, 16.0)
    char_spacing_range: tuple[float, float] = (-5.0, 0.0)
    line_spacing_range: tuple[float, float] = (160.0, 180.0)
    margins: dict[str, float] = field(default_factory=dict)


# ------------------------------------------------------------------
# FormatChecker
# ------------------------------------------------------------------


class FormatChecker:
    """정부과제별 서식 규정을 검사하고 자동 수정한다."""

    def __init__(self, rules_path: str = "config/format_rules.yaml") -> None:
        self._rules_path = rules_path
        self._programs: dict[str, ProgramRules] = {}
        self._load_rules()

    # ------------------------------------------------------------------
    # 규정 로드
    # ------------------------------------------------------------------

    def _load_rules(self) -> None:
        """YAML에서 과제별 규정을 로드한다."""
        path = Path(self._rules_path)
        if not path.exists():
            logger.warning("서식 규정 파일 없음, 기본 규정 사용", path=self._rules_path)
            self._programs["기본"] = ProgramRules(name="기본")
            return

        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f)

        for name, cfg in data.get("programs", {}).items():
            body = cfg.get("body_size_range", [10, 12])
            title = cfg.get("title_size_range", [13, 16])
            cs = cfg.get("char_spacing_range", [-5, 0])
            ls = cfg.get("line_spacing_range", [160, 180])

            self._programs[name] = ProgramRules(
                name=name,
                fonts_allowed=cfg.get("fonts_allowed", ["맑은 고딕"]),
                body_size_range=(float(body[0]), float(body[1])),
                title_size_range=(float(title[0]), float(title[1])),
                char_spacing_range=(float(cs[0]), float(cs[1])),
                line_spacing_range=(float(ls[0]), float(ls[1])),
                margins=cfg.get("margins", {}),
            )

        logger.info("서식 규정 로드 완료", programs=list(self._programs.keys()))

    def get_rules(self, program_name: str) -> ProgramRules:
        """과제명에 해당하는 규정을 반환한다. 없으면 '기본' 규정."""
        return self._programs.get(program_name, self._programs.get("기본", ProgramRules(name="기본")))

    @property
    def available_programs(self) -> list[str]:
        return list(self._programs.keys())

    # ------------------------------------------------------------------
    # 문서 전체 검증
    # ------------------------------------------------------------------

    def check_document(
        self,
        hwp_ctrl: HwpController,
        program_name: str,
    ) -> FormatReport:
        """COM으로 문서 전체 서식을 검증한다.

        검사 항목: 폰트, 크기, 자간, 행간, 여백.
        """
        rules = self.get_rules(program_name)
        report = FormatReport(program_name=program_name)

        # 1. 문단별 글자/문단 모양 검사
        para_results = self._check_paragraphs(hwp_ctrl, rules)
        report.warnings.extend(para_results["warnings"])
        report.errors.extend(para_results["errors"])
        report.total_checks += para_results["total"]
        report.passed_checks += para_results["passed"]

        # 2. 여백 검사
        if rules.margins:
            margin_results = self._check_margins(hwp_ctrl, rules)
            report.warnings.extend(margin_results["warnings"])
            report.errors.extend(margin_results["errors"])
            report.total_checks += margin_results["total"]
            report.passed_checks += margin_results["passed"]

        report.passed = len(report.errors) == 0 and len(report.warnings) == 0
        logger.info("문서 검증 완료", summary=report.summary())
        return report

    def check_table_cells(
        self,
        hwp_ctrl: HwpController,
        table_idx: int,
        program_name: str,
    ) -> list[FormatWarning]:
        """특정 표의 셀별 서식을 검사한다."""
        from src.hwp_engine.table_reader import TableReader

        rules = self.get_rules(program_name)
        warnings: list[FormatWarning] = []

        reader = TableReader(hwp_ctrl)
        table = reader.read_table(table_idx)

        for cell in table.cells:
            if cell.style is None:
                continue

            loc = f"표 {table_idx}, ({cell.row},{cell.col})"
            warnings.extend(self._check_style(loc, cell.style, rules))

        logger.info("표 셀 검증 완료", table_idx=table_idx, warnings=len(warnings))
        return warnings

    # ------------------------------------------------------------------
    # 자동 수정
    # ------------------------------------------------------------------

    def auto_fix(
        self,
        hwp_ctrl: HwpController,
        program_name: str,
    ) -> list[FixAction]:
        """서식 위반 항목을 자동으로 수정한다.

        수정 항목:
        - 비허용 폰트 → 규정 첫 번째 폰트
        - 자간/행간 범위 초과 → 규정 범위 내로 조정
        """
        rules = self.get_rules(program_name)
        fixes: list[FixAction] = []

        hwp = hwp_ctrl.hwp

        # 문서 전체 선택
        hwp.MovePos(2)   # 문서 처음
        hwp.HAction.Run("SelectAll")

        # 현재 글자 모양 가져오기
        try:
            char_shape = hwp_ctrl.get_char_shape()
        except Exception:
            logger.warning("글자 모양 읽기 실패 — 자동 수정 건너뜀")
            return fixes

        new_shape: dict[str, Any] = {}

        # 폰트 수정
        current_font = str(char_shape.get("font_name", ""))
        if current_font and current_font not in rules.fonts_allowed:
            target_font = rules.fonts_allowed[0]
            new_shape["font_name"] = target_font
            fixes.append(FixAction(
                location="문서 전체",
                rule="font_name",
                old_value=current_font,
                new_value=target_font,
            ))

        # 자간 수정
        current_cs = float(char_shape.get("char_spacing", 0))
        cs_min, cs_max = rules.char_spacing_range
        if current_cs < cs_min or current_cs > cs_max:
            target_cs = max(cs_min, min(current_cs, cs_max))
            new_shape["char_spacing"] = target_cs
            fixes.append(FixAction(
                location="문서 전체",
                rule="char_spacing",
                old_value=str(current_cs),
                new_value=str(target_cs),
            ))

        # 적용
        if new_shape:
            try:
                hwp_ctrl.set_char_shape(new_shape)
                logger.info("글자 모양 자동 수정 완료", fixes=len(fixes))
            except Exception:
                for f in fixes:
                    f.success = False
                logger.warning("글자 모양 자동 수정 실패")

        # 행간 수정 (문단 모양)
        try:
            para_shape = hwp_ctrl.get_para_shape()
            current_ls = float(para_shape.get("line_spacing", 160))
            ls_min, ls_max = rules.line_spacing_range
            if current_ls < ls_min or current_ls > ls_max:
                target_ls = max(ls_min, min(current_ls, ls_max))
                self._fix_line_spacing(hwp_ctrl, target_ls)
                fixes.append(FixAction(
                    location="문서 전체",
                    rule="line_spacing",
                    old_value=str(current_ls),
                    new_value=str(target_ls),
                ))
        except Exception:
            logger.debug("행간 수정 건너뜀")

        # 여백 수정
        if rules.margins:
            margin_fixes = self._fix_margins(hwp_ctrl, rules)
            fixes.extend(margin_fixes)

        logger.info("자동 수정 완료", total_fixes=len(fixes), success=sum(1 for f in fixes if f.success))
        return fixes

    # ------------------------------------------------------------------
    # 내부 검사 로직
    # ------------------------------------------------------------------

    def _check_paragraphs(
        self,
        hwp_ctrl: HwpController,
        rules: ProgramRules,
    ) -> dict[str, Any]:
        """문단별 글자/문단 모양을 검사한다."""
        warnings: list[FormatWarning] = []
        errors: list[FormatError] = []
        total = 0
        passed = 0

        hwp = hwp_ctrl.hwp

        # 문서 처음으로 이동
        hwp.MovePos(2)

        # 문단 순회 (최대 200 문단)
        for para_idx in range(200):
            try:
                char_shape = hwp_ctrl.get_char_shape()
                para_shape = hwp_ctrl.get_para_shape()
            except Exception:
                break

            loc = f"문단 {para_idx + 1}"

            # 글자 모양 검사
            style_warnings = self._check_char_shape(loc, char_shape, rules)
            para_warnings = self._check_para_shape(loc, para_shape, rules)

            checks_here = 4  # font, size, char_spacing, line_spacing
            violations = len(style_warnings) + len(para_warnings)
            total += checks_here
            passed += checks_here - violations

            warnings.extend(style_warnings)
            warnings.extend(para_warnings)

            # 다음 문단으로 이동
            try:
                hwp.HAction.Run("MoveNextParaBegin")
                # 위치가 변하지 않으면 문서 끝
                new_pos = hwp_ctrl.get_pos()
                if para_idx > 0:
                    # 간단한 종료 조건 (실제로는 더 정교해야 함)
                    pass
            except Exception:
                break

        return {"warnings": warnings, "errors": errors, "total": total, "passed": passed}

    def _check_char_shape(
        self,
        location: str,
        char_shape: dict[str, Any],
        rules: ProgramRules,
    ) -> list[FormatWarning]:
        """글자 모양을 규정과 비교한다."""
        warnings: list[FormatWarning] = []

        # 폰트
        font = str(char_shape.get("font_name", ""))
        if font and font not in rules.fonts_allowed:
            warnings.append(FormatWarning(
                location=location,
                rule="font_name",
                current_value=font,
                expected=f"허용 폰트: {', '.join(rules.fonts_allowed)}",
                auto_fixable=True,
            ))

        # 글자 크기
        size = float(char_shape.get("font_size", 10.0))
        body_min, body_max = rules.body_size_range
        title_min, title_max = rules.title_size_range
        # 본문 또는 제목 범위에 들어가면 OK
        if not (body_min <= size <= body_max or title_min <= size <= title_max):
            warnings.append(FormatWarning(
                location=location,
                rule="font_size",
                current_value=f"{size}pt",
                expected=f"본문 {body_min}-{body_max}pt 또는 제목 {title_min}-{title_max}pt",
                auto_fixable=False,  # 크기 자동 수정은 위험
            ))

        # 자간
        cs = float(char_shape.get("char_spacing", 0))
        cs_min, cs_max = rules.char_spacing_range
        if cs < cs_min or cs > cs_max:
            warnings.append(FormatWarning(
                location=location,
                rule="char_spacing",
                current_value=f"{cs}%",
                expected=f"{cs_min}% ~ {cs_max}%",
                auto_fixable=True,
            ))

        return warnings

    def _check_para_shape(
        self,
        location: str,
        para_shape: dict[str, Any],
        rules: ProgramRules,
    ) -> list[FormatWarning]:
        """문단 모양을 규정과 비교한다."""
        warnings: list[FormatWarning] = []

        ls = float(para_shape.get("line_spacing", 160))
        ls_min, ls_max = rules.line_spacing_range
        if ls < ls_min or ls > ls_max:
            warnings.append(FormatWarning(
                location=location,
                rule="line_spacing",
                current_value=f"{ls}%",
                expected=f"{ls_min}% ~ {ls_max}%",
                auto_fixable=True,
            ))

        return warnings

    def _check_style(
        self,
        location: str,
        style: Any,
        rules: ProgramRules,
    ) -> list[FormatWarning]:
        """CellStyle 객체로 서식을 검사한다 (표 셀용)."""
        warnings: list[FormatWarning] = []

        # 폰트
        if style.font_name and style.font_name not in rules.fonts_allowed:
            warnings.append(FormatWarning(
                location=location,
                rule="font_name",
                current_value=style.font_name,
                expected=f"허용: {', '.join(rules.fonts_allowed)}",
                auto_fixable=True,
            ))

        # 크기
        body_min, body_max = rules.body_size_range
        title_min, title_max = rules.title_size_range
        if not (body_min <= style.font_size <= body_max or title_min <= style.font_size <= title_max):
            warnings.append(FormatWarning(
                location=location,
                rule="font_size",
                current_value=f"{style.font_size}pt",
                expected=f"본문 {body_min}-{body_max}pt / 제목 {title_min}-{title_max}pt",
                auto_fixable=False,
            ))

        # 자간
        cs_min, cs_max = rules.char_spacing_range
        if style.char_spacing < cs_min or style.char_spacing > cs_max:
            warnings.append(FormatWarning(
                location=location,
                rule="char_spacing",
                current_value=f"{style.char_spacing}%",
                expected=f"{cs_min}% ~ {cs_max}%",
                auto_fixable=True,
            ))

        # 행간
        ls_min, ls_max = rules.line_spacing_range
        if style.line_spacing < ls_min or style.line_spacing > ls_max:
            warnings.append(FormatWarning(
                location=location,
                rule="line_spacing",
                current_value=f"{style.line_spacing}%",
                expected=f"{ls_min}% ~ {ls_max}%",
                auto_fixable=True,
            ))

        return warnings

    def _check_margins(
        self,
        hwp_ctrl: HwpController,
        rules: ProgramRules,
    ) -> dict[str, Any]:
        """페이지 여백을 검사한다."""
        warnings: list[FormatWarning] = []
        errors: list[FormatError] = []
        total = 0
        passed = 0

        try:
            hwp = hwp_ctrl.hwp
            act = hwp.CreateAction("PageSetup")
            param = act.CreateSet()
            act.GetDefault(param)

            # 여백 읽기 (1/100mm → mm)
            current_margins = {
                "top": param.Item("TopMargin") / 100,
                "bottom": param.Item("BottomMargin") / 100,
                "left": param.Item("LeftMargin") / 100,
                "right": param.Item("RightMargin") / 100,
            }

            for side, expected_mm in rules.margins.items():
                total += 1
                current_mm = current_margins.get(side, 0)
                # 1mm 오차 허용
                if abs(current_mm - expected_mm) > 1.0:
                    warnings.append(FormatWarning(
                        location="페이지 여백",
                        rule=f"margin_{side}",
                        current_value=f"{current_mm:.1f}mm",
                        expected=f"{expected_mm}mm",
                        auto_fixable=True,
                    ))
                else:
                    passed += 1

        except Exception:
            logger.debug("여백 검사 건너뜀 (COM API 미지원)")

        return {"warnings": warnings, "errors": errors, "total": total, "passed": passed}

    # ------------------------------------------------------------------
    # 내부 수정 로직
    # ------------------------------------------------------------------

    def _fix_line_spacing(self, hwp_ctrl: HwpController, target: float) -> None:
        """문서 전체 행간을 수정한다."""
        hwp = hwp_ctrl.hwp
        hwp.MovePos(2)
        hwp.HAction.Run("SelectAll")

        act = hwp.CreateAction("ParaShape")
        param = act.CreateSet()
        act.GetDefault(param)
        param.SetItem("LineSpacing", int(target))
        param.SetItem("LineSpacingType", 0)  # 비율
        act.Execute(param)

    def _fix_margins(self, hwp_ctrl: HwpController, rules: ProgramRules) -> list[FixAction]:
        """페이지 여백을 수정한다."""
        fixes: list[FixAction] = []

        try:
            hwp = hwp_ctrl.hwp
            act = hwp.CreateAction("PageSetup")
            param = act.CreateSet()
            act.GetDefault(param)

            margin_map = {
                "top": "TopMargin",
                "bottom": "BottomMargin",
                "left": "LeftMargin",
                "right": "RightMargin",
            }

            changed = False
            for side, expected_mm in rules.margins.items():
                key = margin_map.get(side)
                if not key:
                    continue
                current_mm = param.Item(key) / 100
                if abs(current_mm - expected_mm) > 1.0:
                    param.SetItem(key, int(expected_mm * 100))
                    fixes.append(FixAction(
                        location="페이지 여백",
                        rule=f"margin_{side}",
                        old_value=f"{current_mm:.1f}mm",
                        new_value=f"{expected_mm}mm",
                    ))
                    changed = True

            if changed:
                act.Execute(param)

        except Exception:
            for f in fixes:
                f.success = False
            logger.debug("여백 수정 건너뜀")

        return fixes

"""서식 검증기 테스트.

COM 연동 테스트는 skip, 데이터 클래스/규정 로드/스타일 검사 로직은 단위 테스트.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest
import yaml

from src.hwp_engine.com_controller import HAS_HWP
from src.hwp_engine.table_reader import CellStyle
from src.validator.format_checker import (
    FixAction,
    FormatChecker,
    FormatError,
    FormatReport,
    FormatWarning,
    ProgramRules,
    Severity,
)


# ------------------------------------------------------------------
# 픽스처
# ------------------------------------------------------------------


@pytest.fixture
def rules_file(tmp_path: Path) -> str:
    """테스트용 규정 파일."""
    data = {
        "programs": {
            "테스트과제": {
                "fonts_allowed": ["맑은 고딕", "함초롬바탕"],
                "body_size_range": [10, 12],
                "title_size_range": [13, 16],
                "char_spacing_range": [-5, 0],
                "line_spacing_range": [160, 180],
                "margins": {"top": 20, "bottom": 15, "left": 20, "right": 20},
            },
            "엄격과제": {
                "fonts_allowed": ["맑은 고딕"],
                "body_size_range": [10, 10],
                "title_size_range": [14, 14],
                "char_spacing_range": [0, 0],
                "line_spacing_range": [160, 160],
            },
            "기본": {
                "fonts_allowed": ["맑은 고딕", "돋움", "나눔고딕"],
                "body_size_range": [9, 12],
                "title_size_range": [12, 18],
                "char_spacing_range": [-10, 5],
                "line_spacing_range": [130, 200],
            },
        }
    }
    path = tmp_path / "format_rules.yaml"
    path.write_text(yaml.dump(data, allow_unicode=True), encoding="utf-8")
    return str(path)


@pytest.fixture
def checker(rules_file: str) -> FormatChecker:
    return FormatChecker(rules_path=rules_file)


# ------------------------------------------------------------------
# 데이터 클래스 테스트
# ------------------------------------------------------------------


class TestDataClasses:
    """데이터 클래스 기본 동작."""

    def test_format_warning(self) -> None:
        w = FormatWarning(
            location="문단 1",
            rule="font_name",
            current_value="굴림",
            expected="맑은 고딕",
        )
        assert w.auto_fixable is True
        assert w.severity == Severity.WARNING

    def test_format_error(self) -> None:
        e = FormatError(
            location="표 1",
            rule="font_size",
            current_value="8pt",
            expected="10-12pt",
            message="크기가 너무 작습니다",
        )
        assert e.message == "크기가 너무 작습니다"

    def test_fix_action(self) -> None:
        fa = FixAction(
            location="문서 전체",
            rule="font_name",
            old_value="굴림",
            new_value="맑은 고딕",
        )
        assert fa.success is True

    def test_format_report_pass(self) -> None:
        r = FormatReport(program_name="테스트", total_checks=10, passed_checks=10)
        assert r.passed is True
        assert r.failed_checks == 0
        assert "PASS" in r.summary()

    def test_format_report_fail(self) -> None:
        r = FormatReport(
            program_name="테스트",
            passed=False,
            total_checks=10,
            passed_checks=7,
            warnings=[FormatWarning("loc", "rule", "cur", "exp")],
        )
        assert r.failed_checks == 3
        assert "FAIL" in r.summary()
        assert "경고 1건" in r.summary()

    def test_program_rules_defaults(self) -> None:
        rules = ProgramRules(name="기본")
        assert rules.fonts_allowed == ["맑은 고딕"]
        assert rules.body_size_range == (10.0, 12.0)


# ------------------------------------------------------------------
# 규정 로드 테스트
# ------------------------------------------------------------------


class TestRulesLoading:
    """규정 파일 로드."""

    def test_load_programs(self, checker: FormatChecker) -> None:
        programs = checker.available_programs
        assert "테스트과제" in programs
        assert "엄격과제" in programs
        assert "기본" in programs

    def test_get_rules(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("테스트과제")
        assert rules.name == "테스트과제"
        assert "맑은 고딕" in rules.fonts_allowed
        assert "함초롬바탕" in rules.fonts_allowed
        assert rules.body_size_range == (10.0, 12.0)
        assert rules.margins["top"] == 20

    def test_get_rules_fallback(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("없는과제")
        assert rules.name == "기본"

    def test_missing_rules_file(self, tmp_path: Path) -> None:
        c = FormatChecker(rules_path=str(tmp_path / "nope.yaml"))
        rules = c.get_rules("아무거나")
        assert rules.name == "기본"

    def test_strict_rules(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("엄격과제")
        assert rules.fonts_allowed == ["맑은 고딕"]
        assert rules.body_size_range == (10.0, 10.0)
        assert rules.char_spacing_range == (0.0, 0.0)


# ------------------------------------------------------------------
# 스타일 검사 로직 (COM 없이 검증)
# ------------------------------------------------------------------


class TestStyleChecks:
    """_check_style 내부 로직 테스트 (CellStyle 객체 사용)."""

    def test_valid_style(self, checker: FormatChecker) -> None:
        style = CellStyle(font_name="맑은 고딕", font_size=10.0, char_spacing=-2.0, line_spacing=170.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀 (0,0)", style, rules)
        assert warnings == []

    def test_invalid_font(self, checker: FormatChecker) -> None:
        style = CellStyle(font_name="굴림", font_size=10.0, char_spacing=0.0, line_spacing=160.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀 (0,0)", style, rules)
        font_warnings = [w for w in warnings if w.rule == "font_name"]
        assert len(font_warnings) == 1
        assert font_warnings[0].auto_fixable is True
        assert "굴림" in font_warnings[0].current_value

    def test_invalid_font_size(self, checker: FormatChecker) -> None:
        style = CellStyle(font_name="맑은 고딕", font_size=8.0, char_spacing=0.0, line_spacing=160.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀", style, rules)
        size_warnings = [w for w in warnings if w.rule == "font_size"]
        assert len(size_warnings) == 1
        assert size_warnings[0].auto_fixable is False  # 크기 자동 수정 위험

    def test_title_size_ok(self, checker: FormatChecker) -> None:
        """제목 크기 범위에 있으면 OK."""
        style = CellStyle(font_name="맑은 고딕", font_size=14.0, char_spacing=0.0, line_spacing=160.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀", style, rules)
        size_warnings = [w for w in warnings if w.rule == "font_size"]
        assert len(size_warnings) == 0

    def test_invalid_char_spacing(self, checker: FormatChecker) -> None:
        style = CellStyle(font_name="맑은 고딕", font_size=10.0, char_spacing=-10.0, line_spacing=160.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀", style, rules)
        cs_warnings = [w for w in warnings if w.rule == "char_spacing"]
        assert len(cs_warnings) == 1
        assert cs_warnings[0].auto_fixable is True

    def test_invalid_line_spacing(self, checker: FormatChecker) -> None:
        style = CellStyle(font_name="맑은 고딕", font_size=10.0, char_spacing=0.0, line_spacing=130.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀", style, rules)
        ls_warnings = [w for w in warnings if w.rule == "line_spacing"]
        assert len(ls_warnings) == 1

    def test_multiple_violations(self, checker: FormatChecker) -> None:
        """여러 규정 동시 위반."""
        style = CellStyle(font_name="굴림", font_size=8.0, char_spacing=-10.0, line_spacing=100.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀", style, rules)
        assert len(warnings) == 4  # font, size, char_spacing, line_spacing

    def test_strict_rules_exact_match(self, checker: FormatChecker) -> None:
        """엄격과제: 정확한 값만 통과."""
        rules = checker.get_rules("엄격과제")
        # 정확히 맞는 경우
        ok_style = CellStyle(font_name="맑은 고딕", font_size=10.0, char_spacing=0.0, line_spacing=160.0)
        assert checker._check_style("셀", ok_style, rules) == []

        # 자간 -1이면 위반 (범위 0~0)
        bad_style = CellStyle(font_name="맑은 고딕", font_size=10.0, char_spacing=-1.0, line_spacing=160.0)
        warnings = checker._check_style("셀", bad_style, rules)
        assert any(w.rule == "char_spacing" for w in warnings)

    def test_empty_font_name_no_warning(self, checker: FormatChecker) -> None:
        """폰트명이 비어있으면 검사하지 않음."""
        style = CellStyle(font_name="", font_size=10.0, char_spacing=0.0, line_spacing=160.0)
        rules = checker.get_rules("테스트과제")
        warnings = checker._check_style("셀", style, rules)
        font_warnings = [w for w in warnings if w.rule == "font_name"]
        assert len(font_warnings) == 0


# ------------------------------------------------------------------
# CharShape / ParaShape 검사 (딕셔너리 기반)
# ------------------------------------------------------------------


class TestCharParaShapeChecks:
    """_check_char_shape / _check_para_shape 딕셔너리 검사."""

    def test_char_shape_valid(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("테스트과제")
        shape = {"font_name": "맑은 고딕", "font_size": 11.0, "char_spacing": -3.0}
        warnings = checker._check_char_shape("문단 1", shape, rules)
        assert warnings == []

    def test_char_shape_bad_font(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("테스트과제")
        shape = {"font_name": "Arial", "font_size": 10.0, "char_spacing": 0.0}
        warnings = checker._check_char_shape("문단 1", shape, rules)
        assert len(warnings) == 1
        assert warnings[0].rule == "font_name"

    def test_para_shape_valid(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("테스트과제")
        shape = {"line_spacing": 170.0}
        warnings = checker._check_para_shape("문단 1", shape, rules)
        assert warnings == []

    def test_para_shape_bad_line_spacing(self, checker: FormatChecker) -> None:
        rules = checker.get_rules("테스트과제")
        shape = {"line_spacing": 200.0}
        warnings = checker._check_para_shape("문단 1", shape, rules)
        assert len(warnings) == 1
        assert warnings[0].rule == "line_spacing"


# ------------------------------------------------------------------
# COM 연동 테스트
# ------------------------------------------------------------------


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestFormatCheckerCOM:
    """실제 COM 연동 검증 테스트."""

    def test_check_document(self) -> None:
        pass

    def test_auto_fix(self) -> None:
        pass

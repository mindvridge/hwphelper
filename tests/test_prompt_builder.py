"""프롬프트 빌더 테스트."""

from __future__ import annotations

from src.ai.prompt_builder import PromptBuilder


class TestPromptBuilder:
    """PromptBuilder 테스트."""

    def test_system_prompt_exists(self) -> None:
        pb = PromptBuilder()
        assert len(pb.system_prompt) > 0
        assert "계획서" in pb.system_prompt

    def test_build_cell_prompt_basic(self) -> None:
        pb = PromptBuilder()
        prompt = pb.build_cell_prompt(
            cell_schema={"row": 0, "col": 1, "context": {}},
        )
        assert "(0행, 1열)" in prompt

    def test_build_cell_prompt_with_context(self) -> None:
        pb = PromptBuilder()
        prompt = pb.build_cell_prompt(
            cell_schema={
                "row": 1,
                "col": 2,
                "context": {
                    "row_label": "사업명",
                    "col_header": "내용",
                },
            },
            program_name="예비창업패키지",
            company_info="AI 스타트업",
        )
        assert "예비창업패키지" in prompt
        assert "AI 스타트업" in prompt
        assert "사업명" in prompt
        assert "내용" in prompt

    def test_build_cell_prompt_with_rag(self) -> None:
        pb = PromptBuilder()
        prompt = pb.build_cell_prompt(
            cell_schema={"row": 0, "col": 0, "context": {}},
            rag_context="과거 계획서 참고 내용",
        )
        assert "과거 계획서" in prompt

    def test_build_batch_prompt(self) -> None:
        pb = PromptBuilder()
        cells = [
            {"row": 0, "col": 1, "context": {"row_label": "사업명"}},
            {"row": 1, "col": 1, "context": {"row_label": "기관명"}},
        ]
        prompt = pb.build_batch_prompt(
            cells=cells,
            program_name="TIPS",
        )
        assert "TIPS" in prompt
        assert "2개" in prompt
        assert "사업명" in prompt
        assert "기관명" in prompt
        assert "JSON" in prompt

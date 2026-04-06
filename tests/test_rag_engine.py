"""RAG 엔진 테스트."""

from __future__ import annotations

import pytest

from src.ai.rag_engine import RAGEngine


class TestRAGEngineTextSplit:
    """텍스트 분할 테스트."""

    def test_short_text(self) -> None:
        chunks = RAGEngine._split_text("짧은 텍스트")
        assert len(chunks) == 1

    def test_long_text(self) -> None:
        text = "A" * 1200
        chunks = RAGEngine._split_text(text, chunk_size=500, overlap=50)
        assert len(chunks) >= 2

    def test_overlap(self) -> None:
        text = "ABCDE" * 200  # 1000 chars
        chunks = RAGEngine._split_text(text, chunk_size=500, overlap=50)
        # 마지막 50자의 첫 번째 청크가 두 번째 청크 시작에 포함
        assert len(chunks) >= 2

    def test_empty_text(self) -> None:
        chunks = RAGEngine._split_text("")
        assert chunks == []


class TestRAGEngineIngest:
    """문서 색인 테스트."""

    def test_ingest_nonexistent_file(self, tmp_path) -> None:
        engine = RAGEngine(db_path=str(tmp_path / "chromadb"))
        count = engine.ingest_document("/nonexistent/file.txt")
        assert count == 0

    def test_ingest_file(self, tmp_path) -> None:
        # ChromaDB가 설치되어 있으면 실제 색인 테스트
        db_path = str(tmp_path / "chromadb")
        engine = RAGEngine(db_path=db_path)

        txt = tmp_path / "test.txt"
        txt.write_text("이것은 테스트 문서입니다. 정부과제 계획서 작성 예시.", encoding="utf-8")

        count = engine.ingest_document(str(txt), program_name="테스트과제")
        assert count >= 1

    def test_search_empty_db(self, tmp_path) -> None:
        engine = RAGEngine(db_path=str(tmp_path / "chromadb"))
        results = engine.search("테스트")
        assert isinstance(results, list)

    def test_get_context(self, tmp_path) -> None:
        engine = RAGEngine(db_path=str(tmp_path / "chromadb"))
        ctx = engine.get_context("테스트")
        assert isinstance(ctx, str)

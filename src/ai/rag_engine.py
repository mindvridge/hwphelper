"""RAG 파이프라인 — 과거 계획서를 벡터 DB에 저장하고 검색.

ChromaDB를 사용하여 과거 성공 계획서의 유사 내용을 검색,
새 계획서 작성 시 참고 자료로 활용한다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import structlog

logger = structlog.get_logger()


class RAGEngine:
    """과거 계획서 데이터를 활용한 RAG 파이프라인."""

    def __init__(self, db_path: str = "./data/chroma_db") -> None:
        self._db_path = db_path
        self._client: Any = None
        self._collection: Any = None

    def _ensure_client(self) -> None:
        """ChromaDB 클라이언트를 초기화한다."""
        if self._client is not None:
            return

        try:
            import chromadb

            self._client = chromadb.PersistentClient(path=self._db_path)
            self._collection = self._client.get_or_create_collection(
                name="proposals",
                metadata={"hnsw:space": "cosine"},
            )
            logger.info("ChromaDB 초기화 완료", path=self._db_path)
        except Exception:
            logger.warning("ChromaDB 초기화 실패 — RAG 비활성")

    def ingest_document(self, file_path: str, program_name: str = "") -> int:
        """텍스트 파일을 벡터 DB에 색인한다.

        Parameters
        ----------
        file_path : str
            텍스트 파일 경로 (.txt).
        program_name : str
            과제명 메타데이터.

        Returns
        -------
        int
            색인된 청크 수.
        """
        self._ensure_client()
        if self._collection is None:
            return 0

        path = Path(file_path)
        if not path.exists():
            logger.warning("파일 없음", path=file_path)
            return 0

        text = path.read_text(encoding="utf-8")
        chunks = self._split_text(text)

        ids = [f"{path.stem}_{i}" for i in range(len(chunks))]
        metadatas = [{"source": path.name, "program": program_name}] * len(chunks)

        self._collection.upsert(
            ids=ids,
            documents=chunks,
            metadatas=metadatas,
        )

        logger.info("문서 색인 완료", file=path.name, chunks=len(chunks))
        return len(chunks)

    def search(self, query: str, program_name: str | None = None, top_k: int = 3) -> list[str]:
        """쿼리와 관련된 문서 청크를 검색한다."""
        self._ensure_client()
        if self._collection is None:
            return []

        where_filter = None
        if program_name:
            where_filter = {"program": program_name}

        try:
            results = self._collection.query(
                query_texts=[query],
                n_results=top_k,
                where=where_filter,
            )
            documents = results.get("documents", [[]])[0]
            logger.info("RAG 검색 완료", query=query[:50], results=len(documents))
            return documents
        except Exception:
            logger.warning("RAG 검색 실패", query=query[:50])
            return []

    def get_context(self, query: str, program_name: str | None = None, top_k: int = 3) -> str:
        """검색 결과를 하나의 컨텍스트 문자열로 반환한다."""
        docs = self.search(query, program_name, top_k)
        if not docs:
            return ""
        return "\n\n---\n\n".join(docs)

    @staticmethod
    def _split_text(text: str, chunk_size: int = 500, overlap: int = 50) -> list[str]:
        """텍스트를 청크로 분할한다."""
        chunks: list[str] = []
        start = 0
        while start < len(text):
            end = start + chunk_size
            chunk = text[start:end].strip()
            if chunk:
                chunks.append(chunk)
            start = end - overlap
        return chunks

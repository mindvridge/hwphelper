"""이미지 생성기 — Gemini 프록시 API를 사용한 이미지 생성.

사용법::

    gen = ImageGenerator()
    result = await gen.generate("사이버펑크 고양이")
    # result.url → 이미지 URL
    # result.base64 → base64 데이터
"""

from __future__ import annotations

import base64
import os
from dataclasses import dataclass
from typing import Any

import httpx
import structlog
from dotenv import load_dotenv

load_dotenv()
logger = structlog.get_logger()

IMAGE_API_URL = "https://mlapi.run/820ebe88-0383-4fa4-b5e9-06fcf26b3420"


@dataclass
class ImageResult:
    """이미지 생성 결과."""

    url: str = ""
    base64_data: str = ""
    prompt: str = ""
    error: str = ""


class ImageGenerator:
    """Gemini 프록시 API를 사용한 이미지 생성."""

    def __init__(self) -> None:
        self._api_url = os.environ.get("IMAGE_API_URL", IMAGE_API_URL)
        self._api_key = os.environ.get(
            "IMAGE_API_KEY",
            "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3NzMwMjQ0NDYsIm5iZiI6MTc3MzAyNDQ0Niwia2V5X2lkIjoiNWQ5NmVkOTEtNzZlMy00MmY0LTllZGItOWM0YzQ4NjQ5MjFmIn0.D8sK3v9F4cVXA-LAXHi1RbOnuq5VQb42DzaKJLUL0Xw",
        )
        self._client = httpx.AsyncClient(timeout=120)

    async def generate(
        self,
        prompt: str,
        size: str = "1024x1024",
        aspect_ratio: str = "1:1",
    ) -> ImageResult:
        """텍스트 프롬프트로 이미지를 생성한다."""
        try:
            payload: dict[str, Any] = {
                "model": "google/gemini-3-pro-image-preview",
                "prompt": prompt,
                "n": 1,
                "size": size,
            }
            if aspect_ratio != "1:1":
                payload["aspect_ratio"] = aspect_ratio

            resp = await self._client.post(
                f"{self._api_url}/v1/images/generations",
                json=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self._api_key}",
                },
            )
            resp.raise_for_status()
            data = resp.json()

            url = data.get("data", [{}])[0].get("url", "")
            b64 = data.get("data", [{}])[0].get("b64_json", "")

            logger.info("이미지 생성 완료", prompt=prompt[:50])
            return ImageResult(url=url, base64_data=b64, prompt=prompt)

        except Exception as e:
            logger.warning("이미지 생성 실패", error=str(e))
            return ImageResult(prompt=prompt, error=str(e))

    async def edit(
        self,
        prompt: str,
        image_base64: str,
        mime_type: str = "image/png",
    ) -> ImageResult:
        """기존 이미지를 편집한다."""
        try:
            payload = {
                "model": "google/gemini-3-pro-image-preview",
                "prompt": prompt,
                "n": 1,
                "image": image_base64,
            }

            resp = await self._client.post(
                f"{self._api_url}/v1/images/generations",
                json=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self._api_key}",
                },
            )
            resp.raise_for_status()
            data = resp.json()

            url = data.get("data", [{}])[0].get("url", "")
            b64 = data.get("data", [{}])[0].get("b64_json", "")

            logger.info("이미지 편집 완료", prompt=prompt[:50])
            return ImageResult(url=url, base64_data=b64, prompt=prompt)

        except Exception as e:
            logger.warning("이미지 편집 실패", error=str(e))
            return ImageResult(prompt=prompt, error=str(e))

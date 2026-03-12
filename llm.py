from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional

from openai import AsyncOpenAI, OpenAI

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _load_env_file(env_path: Optional[Path] = None) -> None:
    """Load key=value pairs from .env into process env if missing."""
    if env_path is None:
        env_path = Path(__file__).resolve().parent / ".env"

    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


_load_env_file()


class LLMAPIClient:
    def __init__(self) -> None:
        self.api_key = os.getenv("XHANG_API_KEY") or os.getenv("OPENAI_API_KEY")
        self.base_url = os.getenv("XHANG_BASE_URL", "https://xhang.buaa.edu.cn/xhang/v1")
        self.model_name = os.getenv("XHANG_MODEL", "xhang")

        if not self.api_key:
            raise ValueError("Missing API key. Set XHANG_API_KEY or OPENAI_API_KEY in .env")

        self.client = OpenAI(api_key=self.api_key, base_url=self.base_url)
        self.async_client = AsyncOpenAI(api_key=self.api_key, base_url=self.base_url)

    def chat_completion(
        self,
        messages: List[Dict[str, str]],
        stream: bool = True,
        temperature: float = 0.7,
        max_tokens: int = 300,
    ) -> Any:
        try:
            return self.client.chat.completions.create(
                model=self.model_name,
                messages=messages,
                stream=stream,
                temperature=temperature,
                max_tokens=max_tokens,
            )
        except Exception as e:
            logger.error("LLM API call failed: %s", e)
            raise

    def stream_chat(self, question: str, system_prompt: Optional[str] = None) -> Generator[str, None, None]:
        try:
            messages: List[Dict[str, str]] = []
            if system_prompt:
                messages.append({"role": "system", "content": system_prompt})
            messages.append({"role": "user", "content": question})

            response = self.chat_completion(messages=messages, stream=True)
            for chunk in response:
                delta = chunk.choices[0].delta.content
                if delta:
                    yield delta
        except Exception as e:
            logger.error("Streaming chat failed: %s", e)
            yield f"Chat failed: {e}"

    async def async_chat_completion(
        self, messages: List[Dict[str, str]], temperature: float = 0.7, max_tokens: int = 1000
    ) -> str:
        try:
            response = await self.async_client.chat.completions.create(
                model=self.model_name,
                messages=messages,
                stream=False,
                temperature=temperature,
                max_tokens=max_tokens,
            )
            return response.choices[0].message.content or ""
        except Exception as e:
            logger.error("Async LLM API call failed: %s", e)
            return f"Call failed: {e}"

    async def async_chat(self, question: str, system_prompt: Optional[str] = None) -> str:
        try:
            messages: List[Dict[str, str]] = []
            if system_prompt:
                messages.append({"role": "system", "content": system_prompt})
            messages.append({"role": "user", "content": question})
            return await self.async_chat_completion(messages=messages)
        except Exception as e:
            logger.error("Async chat failed: %s", e)
            return f"Chat failed: {e}"


if __name__ == "__main__":
    client = LLMAPIClient()
    print("Answer: ", end="")
    for chunk in client.stream_chat("What is time complexity?"):
        print(chunk, end="", flush=True)
    print()

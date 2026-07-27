"""Groq generation client (OpenAI-compatible, generous free tier).

Groq only. The digest runs on a weekly schedule with no reply deadline, so there
is no need for the multi-provider fallback the RAG project carries — one provider
means one dependency and one failure mode.
"""
from __future__ import annotations

import time

from groq import Groq
from groq import APIStatusError, APIConnectionError

from .config import Config

_RETRYABLE = {429, 500, 502, 503}


class GroqClient:
    def __init__(self, cfg: Config):
        cfg.require_groq()  # validated here, not at config load
        self.cfg = cfg
        self.client = Groq(api_key=cfg.groq_api_key)

    def generate(
        self, prompt: str, *, max_output_tokens: int | None = None, attempts: int = 5
    ) -> str:
        delay = 3.0
        for i in range(attempts):
            try:
                resp = self.client.chat.completions.create(
                    model=self.cfg.groq_model,
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=max_output_tokens or 1024,
                    # Low but non-zero: the digest should read naturally while
                    # staying tightly anchored to the brief.
                    temperature=0.2,
                )
                return (resp.choices[0].message.content or "").strip()
            except APIStatusError as e:
                if e.status_code in _RETRYABLE and i < attempts - 1:
                    print(f"  groq {e.status_code}; waiting {delay:.0f}s...")
                    time.sleep(delay)
                    delay = min(delay * 2, 30.0)
                    continue
                raise
            except APIConnectionError:
                if i < attempts - 1:
                    time.sleep(delay)
                    delay = min(delay * 2, 30.0)
                    continue
                raise
        raise RuntimeError("unreachable: retry loop exhausted without raising")

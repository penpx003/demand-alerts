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

# Upper bound when growing the budget after a truncated answer. Well under the
# model's completion limit, and far above anything a weekly digest should need.
MAX_TOKEN_CEILING = 16000


class GroqClient:
    def __init__(self, cfg: Config):
        cfg.require_groq()  # validated here, not at config load
        self.cfg = cfg
        self.client = Groq(api_key=cfg.groq_api_key)

    def generate(
        self, prompt: str, *, max_output_tokens: int | None = None, attempts: int = 5
    ) -> str:
        """Generate, enlarging the token budget once if the answer was cut off.

        A truncated completion is the nastiest failure mode here: the API returns
        200, the text reads normally, and it simply stops mid-sentence with later
        sections missing. Nothing downstream can tell that apart from a genuinely
        short digest, so it is caught here and retried with more room.

        Reasoning models make this easy to hit — their internal reasoning counts
        against max_tokens, so a budget that looks generous can leave very little
        for the visible answer.
        """
        budget = max_output_tokens or 1024

        text, truncated = self._complete(prompt, budget, attempts)
        if not truncated:
            return text

        bigger = min(budget * 3, MAX_TOKEN_CEILING)
        if bigger > budget:
            print(
                f"  groq stopped at the {budget}-token limit; "
                f"retrying with {bigger}"
            )
            text, truncated = self._complete(prompt, bigger, attempts)

        if truncated:
            # Return what we have — a partial digest beats none — but say so
            # loudly, because it will not look wrong on its own.
            print(
                "WARNING: the narrative was still truncated at the token limit. "
                "The digest is INCOMPLETE — later sections may be missing. "
                "Raise DEMAND_MAX_OUTPUT_TOKENS, or shorten the brief with a "
                "lower DEMAND_TOP_N."
            )
        return text

    def _complete(self, prompt: str, budget: int, attempts: int) -> tuple[str, bool]:
        """One completion with transient-error retries. Returns (text, truncated)."""
        delay = 3.0
        for i in range(attempts):
            try:
                resp = self.client.chat.completions.create(
                    model=self.cfg.groq_model,
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=budget,
                    # Low but non-zero: the digest should read naturally while
                    # staying tightly anchored to the brief.
                    temperature=0.2,
                )
                choice = resp.choices[0]
                text = (choice.message.content or "").strip()
                truncated = getattr(choice, "finish_reason", None) == "length"
                return text, truncated
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

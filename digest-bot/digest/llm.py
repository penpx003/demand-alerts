"""Groq generation client (OpenAI-compatible, generous free tier).

Groq only. The digest runs on a weekly schedule with no reply deadline, so there
is no need for the multi-provider fallback the RAG project carries — one provider
means one dependency and one failure mode.
"""
from __future__ import annotations

import os
import time

from groq import Groq
from groq import APIStatusError, APIConnectionError

from .config import Config

_RETRYABLE = {429, 500, 502, 503}

# Upper bound when growing the budget after a truncated answer. Well under the
# model's completion limit, and far above anything a weekly digest should need.
MAX_TOKEN_CEILING = 16000

# Groq charges the tokens-per-minute allowance for the PROMPT PLUS the whole
# max_tokens reservation, not the tokens actually generated. So a large brief
# and a generous output budget can be rejected outright with
#   413 ... on tokens per minute (TPM): Limit 8000, Requested 8480
# even though the answer would have been far shorter. The budget is therefore
# capped to fit, rather than discovered by failing.
#
# Free tier is 8000 TPM for the standard chat models (verified Aug 2026 from the
# x-ratelimit-limit-tokens header). Raise this if the account is upgraded.
TPM_LIMIT = int(os.getenv("GROQ_TPM_LIMIT", "8000"))
# Slack for the difference between the character-based estimate and Groq's count.
TPM_SAFETY_MARGIN = 400
# Below this an answer is not worth attempting; the brief needs to shrink instead.
MIN_OUTPUT_BUDGET = 600


def estimate_tokens(text: str) -> int:
    """Rough prompt size. ~4 characters per token is close enough to budget by."""
    return len(text) // 4 + 1


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
        requested = max_output_tokens or 1024
        headroom = self._headroom(prompt)
        budget = max(MIN_OUTPUT_BUDGET, min(requested, headroom))

        if budget < requested:
            print(
                f"  output budget capped to {budget} (asked {requested}) to stay "
                f"within the {TPM_LIMIT} tokens-per-minute limit"
            )

        text, truncated = self._complete(prompt, budget, attempts)
        if not truncated:
            return text

        bigger = min(budget * 3, MAX_TOKEN_CEILING, headroom)
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
                "The brief is leaving too little room under the "
                f"{TPM_LIMIT} tokens-per-minute cap: shorten it with a lower "
                "DEMAND_TOP_N."
            )
        return text

    def _headroom(self, prompt: str) -> int:
        """Output tokens that still fit under the per-minute allowance."""
        prompt_tokens = estimate_tokens(prompt)
        room = TPM_LIMIT - prompt_tokens - TPM_SAFETY_MARGIN
        if room < MIN_OUTPUT_BUDGET:
            print(
                f"WARNING: the brief is ~{prompt_tokens} tokens and nearly fills "
                f"the {TPM_LIMIT} tokens-per-minute allowance on its own. Lower "
                "DEMAND_TOP_N so the model has room to answer."
            )
        return room

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
                # 413 means prompt + reservation exceeded the per-minute
                # allowance. Waiting cannot help — the request is too big by
                # construction — so shrink the reservation and try again.
                if e.status_code == 413 and budget > MIN_OUTPUT_BUDGET and i < attempts - 1:
                    budget = max(MIN_OUTPUT_BUDGET, budget // 2)
                    print(
                        f"  groq 413 (request larger than the per-minute limit); "
                        f"retrying with a {budget}-token reservation"
                    )
                    continue
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

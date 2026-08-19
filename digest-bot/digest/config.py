"""Central configuration, loaded from environment / .env.

Nothing is required at load time. Each credential is validated where it is
actually used, so a partially configured service degrades instead of failing
whole: no Supabase key means no trend history but a digest still generates, and
a week with no alerts needs no Groq key at all because the LLM is never called.
"""
from __future__ import annotations

import os
from dataclasses import dataclass

from dotenv import load_dotenv

load_dotenv()


class MissingConfig(RuntimeError):
    """Raised at the point a missing credential is actually needed."""


@dataclass(frozen=True)
class Config:
    # --- generation -------------------------------------------------------
    groq_api_key: str
    groq_model: str

    # --- storage ----------------------------------------------------------
    supabase_url: str
    supabase_key: str

    # --- presentation -----------------------------------------------------
    # Link to the alert workbook, appended to the digest so a planner can open
    # the detail behind the narrative. Blank omits the link entirely.
    workbook_url: str

    # The digest tuning knobs (DEMAND_TREND_WEEKS, DEMAND_SNAPSHOT_PC_LIMIT,
    # DEMAND_TOP_N) are read as module constants in digest.py, where they are
    # used. Duplicating them here would give two sources of truth.

    @classmethod
    def load(cls) -> "Config":
        return cls(
            groq_api_key=_clean(os.getenv("GROQ_API_KEY")),
            # Groq retires model ids without notice — the Llama 3.x line vanished
            # in Aug 2026, giving `model_not_found` 404s on a previously working
            # deployment. If that happens, list the current ids with:
            #   GET https://api.groq.com/openai/v1/models  (Bearer GROQ_API_KEY)
            # and set GROQ_CHAT_MODEL to one of them.
            groq_model=os.getenv("GROQ_CHAT_MODEL", "openai/gpt-oss-120b"),
            supabase_url=_clean(os.getenv("SUPABASE_URL")),
            supabase_key=_clean(os.getenv("SUPABASE_KEY")),
            workbook_url=_clean(os.getenv("WORKBOOK_URL")),
        )

    def require_groq(self) -> None:
        if not self.groq_api_key:
            raise MissingConfig(
                "GROQ_API_KEY is not set, so the narrative cannot be generated. "
                "Copy .env.example to .env and fill it in, or set it in the host "
                "environment."
            )

    def require_supabase(self) -> None:
        missing = [
            name
            for name, value in (("SUPABASE_URL", self.supabase_url), ("SUPABASE_KEY", self.supabase_key))
            if not value
        ]
        if missing:
            raise MissingConfig(
                f"{' and '.join(missing)} not set, so alert snapshots and trend "
                f"history are unavailable. The digest still generates, but nothing "
                f"can be reported as recurring."
            )

    def missing(self) -> list[str]:
        """Everything unset, for a startup warning."""
        return [
            name
            for name, value in (
                ("GROQ_API_KEY", self.groq_api_key),
                ("SUPABASE_URL", self.supabase_url),
                ("SUPABASE_KEY", self.supabase_key),
            )
            if not value
        ]


def _clean(value: str | None) -> str:
    """Treat the .env.example placeholders as unset."""
    val = (value or "").strip()
    return "" if val.startswith("your-") else val

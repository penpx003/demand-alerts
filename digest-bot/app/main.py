"""FastAPI service for the weekly Demand Planning digest.

Endpoints:
  GET  /health                     -> liveness check (also the Render health path)
  POST /api/demand-digest          -> alert tables in, narrative out
  GET  /api/demand-digest/latest   -> most recently generated digest

Deliberately standalone: this service shares no code, no Supabase tables and no
Teams channel with the RAG bot. The only thing they have in common is the
provider stack (Groq + Supabase + Power Automate).
"""
from __future__ import annotations

import hmac
import os

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

# Importing the package injects truststore before any HTTPS client is built.
from digest.config import Config
from digest.digest import DemandStore, run_digest

app = FastAPI(title="Demand Planning digest")


@app.get("/health")
def health():
    return {"status": "ok"}


def _authorised(request: Request) -> bool:
    """Shared-token check. The service is publicly reachable, so a deployed
    environment must set DEMAND_DIGEST_TOKEN; blank disables the check and is
    intended for local testing only."""
    token = os.getenv("DEMAND_DIGEST_TOKEN", "")
    if not token:
        return True
    return hmac.compare_digest(request.headers.get("X-Digest-Token", ""), token)


@app.post("/api/demand-digest")
async def demand_digest(request: Request):
    """Weekly Demand Planning digest.

    Body: {"week_of": "2026-07-27", "alerts": {"alert1": [...], ..., "alert4": [...]}}
    where each list holds the rows of the matching output table from the
    `Demand Alerts` Office Script, as returned by Power Automate's
    "List rows present in a table".
    """
    if not _authorised(request):
        return JSONResponse({"error": "Unauthorized."}, status_code=401)

    try:
        payload = await request.json()
    except Exception:
        return JSONResponse({"error": "Body must be JSON."}, status_code=400)
    if not isinstance(payload, dict):
        return JSONResponse({"error": "Body must be a JSON object."}, status_code=400)

    try:
        result = run_digest(payload)
    except Exception as e:
        print(f"demand digest failed: {e}")
        return JSONResponse({"error": f"Digest failed: {e}"}, status_code=500)

    # The brief is verbose and only useful when debugging a bad narrative.
    if not payload.get("include_brief"):
        result.pop("brief", None)
    return result


@app.get("/api/demand-digest/latest")
def demand_digest_latest(request: Request):
    """The most recent stored digest, for re-posting or a quick check."""
    if not _authorised(request):
        return JSONResponse({"error": "Unauthorized."}, status_code=401)

    latest = DemandStore(Config.load()).latest_digest()
    if not latest:
        return {"week_of": None, "narrative": None}
    return latest

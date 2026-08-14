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
from digest.digest import DemandStore, normalise_country, run_digest

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

    Body: {"week_of": "2026-07-27", "country": "ES", "workbook_url": "https://...",
           "alerts": {"alert1": [...], ..., "alert4": [...]}}
    where each list holds the rows of the matching output table from the
    `Demand Alerts` Office Script, as returned by Power Automate's
    "List rows present in a table".

    `country` scopes storage and trend history, so one flow per country can share
    this service without overwriting each other. `workbook_url` overrides the
    configured default, so each country's digest links to its own workbook. Both
    are optional; omitted, the service behaves as a single-scope installation.
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
def demand_digest_latest(request: Request, country: str | None = None):
    """The most recent stored digest, for re-posting or a quick check.

    ?country=ES scopes it to one country. Omitted, it returns the most recent
    digest of any country — which with several country flows running is whichever
    happened to finish last, so pass the country when you mean a specific one.
    """
    if not _authorised(request):
        return JSONResponse({"error": "Unauthorized."}, status_code=401)

    scope = normalise_country(country) if country is not None else None
    latest = DemandStore(Config.load()).latest_digest(scope)
    if not latest:
        return {"week_of": None, "country": scope, "narrative": None}
    return latest

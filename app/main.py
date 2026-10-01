"""Application FastAPI pour le complément Triage_Outlook."""
from __future__ import annotations

import hmac
import logging
import os
import uuid
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, ConfigDict, Field, model_validator

from . import __version__
from .ai import (
    AIError,
    ClassificationResult,
    EmailToClassify,
    TargetFolderRule,
    classify_batch,
    provider_status,
)

load_dotenv(override=False)

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(name)s — %(message)s",
)
logger = logging.getLogger("triage_outlook")

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

ADDIN_ID = os.getenv("ADDIN_ID", "e78b2a1c-9012-4fe3-8899-1a2b3c4d5e6f")
ADDIN_VERSION = os.getenv("ADDIN_VERSION", "1.0.0.0")
PROVIDER_NAME = os.getenv("ADDIN_PROVIDER_NAME", "Triage_Outlook")
AUTH_TOKEN = os.getenv("AUTH_TOKEN", "").strip()

app = FastAPI(
    title="Triage_Outlook",
    description="Complément Outlook de tri automatique d'e-mails par lot par IA.",
    version=__version__,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=os.getenv("CORS_ORIGINS", "*").split(","),
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


def public_base_url(request: Request) -> str:
    explicit = (os.getenv("PUBLIC_BASE_URL") or "").strip().rstrip("/")
    if explicit:
        return explicit
    railway_domain = (os.getenv("RAILWAY_PUBLIC_DOMAIN") or "").strip().rstrip("/")
    if railway_domain:
        return f"https://{railway_domain}"
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or request.url.netloc
    scheme = request.headers.get("x-forwarded-proto", request.url.scheme).split(",")[0]
    if not host.split(":")[0] in {"localhost", "127.0.0.1"}:
        scheme = "https"
    return f"{scheme}://{host}".rstrip("/")


class EmailIn(BaseModel):
    id: str
    subject: str = ""
    sender: str = ""
    body: str = ""


def _colonnes_possibles(valeur: str, *candidats: str) -> str:
    """Retourne la premiere valeur non vide parmi les candidats."""
    for candidat in candidats:
        if isinstance(candidat, str) and candidat.strip():
            return candidat.strip()
    return (valeur or "").strip()


def _est_actif(valeur: Any) -> bool:
    """Interprete la colonne Actif d'une liste SharePoint.

    Absente : on considere le dossier actif, pour rester compatible avec les
    appels historiques qui ne l'envoient pas.
    """
    if valeur is None:
        return True
    if isinstance(valeur, bool):
        return valeur
    if isinstance(valeur, (int, float)):
        return valeur != 0
    texte = str(valeur).strip().lower()
    return texte not in {"non", "no", "false", "0", "inactif", "desactive", ""}


class FolderIn(BaseModel):
    """Dossier cible.

    Accepte les noms du contrat d'origine (name, folder_id, rule) comme ceux des
    colonnes d'une liste SharePoint (Titre ou Title, Identifiant, Consigne,
    Actif). Le flux peut ainsi envoyer directement les lignes de la liste, sans
    transformation intermediaire.
    """

    model_config = ConfigDict(extra="ignore")

    name: str = ""
    folder_id: str = ""
    rule: str = ""

    # Colonnes d'une liste SharePoint
    Titre: str = ""
    Title: str = ""
    Identifiant: str = ""
    Consigne: str = ""
    Actif: Any = None

    @model_validator(mode="after")
    def _resoudre(self) -> "FolderIn":
        """Ramene les deux nommages vers name / folder_id / rule."""
        self.name = _colonnes_possibles(self.name, self.name, self.Titre, self.Title)
        self.folder_id = _colonnes_possibles(self.folder_id, self.folder_id, self.Identifiant)
        self.rule = _colonnes_possibles(self.rule, self.rule, self.Consigne)
        return self

    def utilisable(self) -> bool:
        """Un dossier sans identifiant ne peut pas recevoir de deplacement."""
        return bool(self.name and self.folder_id)


class TriageRequest(BaseModel):
    emails: list[EmailIn]
    folders: list[FolderIn]
    auth_token: str = Field("", description="Code secret d'accès")


class VerifyTokenRequest(BaseModel):
    token: str = ""


@app.get("/", response_class=HTMLResponse)
async def home(request: Request) -> HTMLResponse:
    base = public_base_url(request)
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "base_url": base,
            "manifest_url": f"{base}/manifest.xml",
            "taskpane_url": f"{base}/taskpane.html",
            "status": provider_status(),
            "auth_required": bool(AUTH_TOKEN),
            "version": __version__,
        },
    )


@app.get("/taskpane.html", response_class=HTMLResponse)
async def taskpane(request: Request) -> HTMLResponse:
    base = public_base_url(request)
    resp = templates.TemplateResponse(
        request,
        "taskpane.html",
        {
            "api_url": f"{base}/api/triage",
            "verify_url": f"{base}/api/verify-token",
            "status": provider_status(),
            "auth_required": bool(AUTH_TOKEN),
            "version": __version__,
        },
    )
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.get("/commands.html", response_class=HTMLResponse, include_in_schema=False)
async def commands() -> HTMLResponse:
    return HTMLResponse(
        '<!DOCTYPE html><html><head><meta charset="utf-8"><script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script></head><body></body></html>'
    )


@app.get("/manifest.xml")
async def manifest(request: Request, download: bool = False) -> Response:
    base = public_base_url(request)
    xml = templates.get_template("manifest.xml.j2").render(
        base_url=base,
        addin_id=ADDIN_ID,
        version=ADDIN_VERSION,
        provider_name=PROVIDER_NAME,
    )
    headers = {"Cache-Control": "no-store"}
    if download:
        headers["Content-Disposition"] = 'attachment; filename="manifest.xml"'
    return Response(content=xml, media_type="text/xml; charset=utf-8", headers=headers)


@app.post("/api/verify-token")
async def api_verify_token(payload: VerifyTokenRequest) -> dict:
    if not AUTH_TOKEN:
        return {"valid": True, "auth_required": False}
    is_valid = bool(payload.token and hmac.compare_digest(payload.token.strip(), AUTH_TOKEN))
    return {"valid": is_valid, "auth_required": True}


@app.post("/api/triage")
async def api_triage(request: Request, payload: TriageRequest) -> dict:
    if AUTH_TOKEN:
        token = request.headers.get("x-auth-token") or payload.auth_token
        if not token or not hmac.compare_digest(token.strip(), AUTH_TOKEN):
            raise HTTPException(status_code=401, detail="Code secret invalide ou manquant.")

    if not payload.emails:
        return {"classifications": []}

    email_objs = [
        EmailToClassify(
            item_id=e.id,
            subject=e.subject,
            sender=e.sender,
            body_snippet=e.body,
        )
        for e in payload.emails
    ]

    ignores = [
        f.name or f.Titre or f.Title or "(sans nom)"
        for f in payload.folders
        if not _est_actif(f.Actif) or not f.utilisable()
    ]
    if ignores:
        logger.info("Dossiers ignores : %s", ", ".join(ignores))

    folder_objs = [
        TargetFolderRule(
            name=f.name,
            folder_id=f.folder_id,
            rule=f.rule,
        )
        for f in payload.folders
        if _est_actif(f.Actif) and f.utilisable()
    ]

    try:
        results = await classify_batch(email_objs, folder_objs)
    except AIError as exc:
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return {
        "classifications": [
            {
                "item_id": r.item_id,
                "target_folder_name": r.target_folder_name,
                "target_folder_id": r.target_folder_id,
                "confidence": r.confidence,
                "justification": r.justification,
            }
            for r in results
        ]
    }


@app.get("/healthz")
async def healthz() -> JSONResponse:
    status = provider_status()
    return JSONResponse(
        {
            "status": "ok",
            "version": __version__,
            "ai_provider": status["provider"],
            "demo_mode": status["demo"],
        }
    )


@app.get("/robots.txt", response_class=PlainTextResponse, include_in_schema=False)
async def robots() -> str:
    return "User-agent: *\nDisallow: /\n"

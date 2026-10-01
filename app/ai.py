"""Module IA pour le tri automatique d\'e-mails par lot.
Supporte Gemini (Google), Claude (Anthropic) et mode démo sans clé.
"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import re
from dataclasses import dataclass
from typing import Any, Final

import httpx

logger = logging.getLogger("triage_outlook.ai")

DEFAULT_GEMINI_MODEL: Final = "gemini-3.6-flash"
DEFAULT_ANTHROPIC_MODEL: Final = "claude-sonnet-5"

ANTHROPIC_URL: Final = "https://api.anthropic.com/v1/messages"
DEEPSEEK_URL: Final = "https://api.deepseek.com/chat/completions"
DEFAULT_DEEPSEEK_MODEL: Final = "deepseek-chat"
ANTHROPIC_VERSION: Final = "2023-06-01"
GEMINI_URL_TEMPLATE: Final = (
    "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
)

REQUEST_TIMEOUT: Final = 60.0

# Statuts qui signalent une surtension passagere du cote du fournisseur : cela
# vaut un nouvel essai, pas un abandon. Le 503 « high demand » de Google en fait
# partie, il est intermittent.
STATUTS_REESSAYABLES: Final = frozenset({429, 500, 502, 503, 504})

# Trois tentatives par modele, avec des attentes croissantes (en secondes).
TENTATIVES_PAR_MODELE: Final = 3
ATTENTES: Final = (0.0, 1.5, 4.0)


class AIError(RuntimeError):
    """Erreur remontée à l\'API."""

    def __init__(self, message: str, statut: int | None = None) -> None:
        super().__init__(message)
        self.statut = statut

    @property
    def passager(self) -> bool:
        """Vrai si l'erreur tient a une surtension, donc a un autre essai."""
        return self.statut in STATUTS_REESSAYABLES


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def _is_usable_key(val: str) -> bool:
    v = (val or "").strip()
    if not v or len(v) < 15:
        return False
    lower = v.lower()
    placeholders = ["sk-ant-api03-...", "aizasy...", "votre_", "your_", "todo", "xxx", "placeholder"]
    return not any(p in lower for p in placeholders)


def _fournisseurs_disponibles() -> list[str]:
    """Les fournisseurs utilisables d'apres les cles presentes."""
    dispo = []
    if _is_usable_key(_env("GEMINI_API_KEY")):
        dispo.append("gemini")
    if _is_usable_key(_env("DEEPSEEK_API_KEY")):
        dispo.append("deepseek")
    if _is_usable_key(_env("ANTHROPIC_API_KEY")):
        dispo.append("anthropic")
    return dispo


def resolve_provider() -> str:
    forced = _env("AI_PROVIDER", "auto").lower()
    has_anthropic = _is_usable_key(_env("ANTHROPIC_API_KEY"))
    has_gemini = _is_usable_key(_env("GEMINI_API_KEY"))
    has_deepseek = _is_usable_key(_env("DEEPSEEK_API_KEY"))

    if forced in {"anthropic", "claude"}:
        if not has_anthropic:
            raise AIError("AI_PROVIDER=anthropic mais ANTHROPIC_API_KEY est absente.")
        return "anthropic"
    if forced in {"deepseek"}:
        if not has_deepseek:
            raise AIError("AI_PROVIDER=deepseek mais DEEPSEEK_API_KEY est absente.")
        return "deepseek"
    if forced in {"gemini", "google"}:
        if not has_gemini:
            raise AIError("AI_PROVIDER=gemini mais GEMINI_API_KEY est absente.")
        return "gemini"
    if forced in {"demo", "mock", "test"}:
        return "demo"

    disponibles = _fournisseurs_disponibles()
    return disponibles[0] if disponibles else "demo"


def provider_status() -> dict[str, Any]:
    try:
        provider = resolve_provider()
    except AIError:
        provider = "demo"
    return {
        "provider": provider,
        "demo": provider == "demo",
        "gemini_key": _is_usable_key(_env("GEMINI_API_KEY")),
        "deepseek_key": _is_usable_key(_env("DEEPSEEK_API_KEY")),
        "anthropic_key": _is_usable_key(_env("ANTHROPIC_API_KEY")),
        "repli": _fournisseurs_disponibles()[1:] or [],
        "model": {
            "gemini": _env("GEMINI_MODEL", DEFAULT_GEMINI_MODEL),
            "deepseek": _env("DEEPSEEK_MODEL", DEFAULT_DEEPSEEK_MODEL),
            "anthropic": _env("ANTHROPIC_MODEL", DEFAULT_ANTHROPIC_MODEL),
        }.get(provider, DEFAULT_GEMINI_MODEL),
    }


@dataclass
class EmailToClassify:
    item_id: str
    subject: str
    sender: str
    body_snippet: str


@dataclass
class TargetFolderRule:
    name: str
    folder_id: str
    rule: str


@dataclass
class ClassificationResult:
    item_id: str
    target_folder_name: str
    target_folder_id: str
    confidence: str  # high, medium, low
    justification: str


async def classify_batch(
    emails: list[EmailToClassify],
    folders: list[TargetFolderRule],
) -> list[ClassificationResult]:
    if not emails:
        return []
    if not folders:
        raise AIError("Aucun dossier cible actif. Activez au moins un dossier dans les règles.")

    provider = resolve_provider()
    logger.info("Tri de %d mails vers %d dossiers cibles via %s", len(emails), len(folders), provider)

    if provider == "demo":
        return _demo_classify(emails, folders)

    system_prompt = (
        "Tu es un assistant IA spécialisé dans le tri et la catégorisation de courriels d\'entreprise.\n"
        "Ta mission : analyser chaque e-mail d\'une liste et déterminer vers quel dossier le déplacer, "
        "en respectant STRICTEMENT les règles sémantiques définies pour chaque dossier.\n"
        "Tu dois répondre UNIQUEMENT sous forme d\'un tableau JSON valide, sans texte additionnel."
    )

    folders_desc = "\n".join(
        f"- Dossier: \"{f.name}\" (ID: {f.folder_id}) | Règle d\'attribution: {f.rule or 'Général'}"
        for f in folders
    )

    emails_desc = "\n---\n".join(
        f"Item ID: {e.item_id}\nDe: {e.sender}\nObjet: {e.subject}\nCorps: {e.body_snippet[:600]}"
        for e in emails
    )

    user_prompt = f"""### DOSSIERS CIBLES ET LEURS RÈGLES :
{folders_desc}

### LISTE DES E-MAILS À CLASSER :
{emails_desc}

### INSTRUCTIONS DE RÉPONSE :
Attribue à chaque e-mail le dossier cible le plus pertinent d\'après les règles ci-dessus.
Réponds UNIQUEMENT avec un tableau JSON au format exact suivant :
[
  {{
    "item_id": "...",
    "target_folder_name": "nom exact du dossier choisi",
    "target_folder_id": "ID du dossier choisi",
    "confidence": "high" ou "medium" ou "low",
    "justification": "Explication courte en français (5 à 10 mots) du motif de classement"
  }}
]
"""

    raw_json = await _appeler_modele(system_prompt, user_prompt, provider)

    return _parse_classifications(raw_json, emails, folders)


async def _appeler_modele(system: str, user: str, principal: str) -> str:
    """Interroge le fournisseur principal, puis les autres si celui-ci echoue.

    Une surtension n'est jamais simultanee chez deux fournisseurs differents :
    c'est la raison d'etre de ce repli. On essaie donc chacun dans l'ordre, et on
    ne remonte l'erreur que si aucun n'a repondu.
    """
    appels = {
        "gemini": _call_gemini,
        "deepseek": _call_deepseek,
        "anthropic": _call_anthropic,
    }
    ordre = [principal] + [f for f in _fournisseurs_disponibles() if f != principal]
    derniere: AIError | None = None

    for nom in ordre:
        appel = appels.get(nom)
        if appel is None:
            continue
        try:
            texte = await appel(system, user)
            if texte:
                if nom != principal:
                    logger.warning("Tri effectue par le fournisseur de repli : %s", nom)
                return texte
            derniere = AIError(f"{nom} : reponse vide.")
        except AIError as exc:
            derniere = exc
            logger.warning("Fournisseur %s indisponible (%s)", nom, exc)

    if derniere:
        raise derniere
    raise AIError("Aucun fournisseur d'IA disponible.")


def _parse_classifications(
    raw_text: str,
    emails: list[EmailToClassify],
    folders: list[TargetFolderRule],
) -> list[ClassificationResult]:
    # Nettoyage JSON
    text = raw_text.strip()
    match = re.search(r"\[.*?\]", text, re.DOTALL)
    if match:
        text = match.group(0)

    folder_map = {f.name.lower(): f for f in folders}
    folder_id_map = {f.folder_id: f for f in folders}
    fallback_folder = folders[0]

    results: list[ClassificationResult] = []
    try:
        data = json.loads(text)
        if isinstance(data, list):
            items_by_id = {item.get("item_id"): item for item in data if isinstance(item, dict)}
            for e in emails:
                item = items_by_id.get(e.item_id)
                if item:
                    f_name = item.get("target_folder_name", "")
                    f_id = item.get("target_folder_id", "")
                    target = folder_id_map.get(f_id) or folder_map.get(f_name.lower()) or fallback_folder
                    results.append(
                        ClassificationResult(
                            item_id=e.item_id,
                            target_folder_name=target.name,
                            target_folder_id=target.folder_id,
                            confidence=item.get("confidence", "high"),
                            justification=item.get("justification", "Classement basé sur les règles"),
                        )
                    )
                else:
                    results.append(
                        ClassificationResult(
                            item_id=e.item_id,
                            target_folder_name=fallback_folder.name,
                            target_folder_id=fallback_folder.folder_id,
                            confidence="low",
                            justification="Dossier par défaut",
                        )
                    )
            return results
    except Exception as exc:
        logger.warning("Erreur décodage JSON IA (%s), repli démo", exc)

    return _demo_classify(emails, folders)


def _demo_classify(
    emails: list[EmailToClassify],
    folders: list[TargetFolderRule],
) -> list[ClassificationResult]:
    results = []
    for i, e in enumerate(emails):
        # Heuristique simple par mots-clés pour le mode démo
        text = f"{e.subject} {e.body_snippet}".lower()
        chosen = folders[0]
        reason = "Attribution par défaut"
        confidence = "medium"

        for f in folders:
            f_rule_lower = f.rule.lower()
            keywords = [w for w in re.findall(r"\w{4,}", f_rule_lower) if w not in {"dans", "cette", "pour", "avec", "sont", "doit", "client", "mails", "dossier"}]
            matches = sum(1 for kw in keywords if kw in text)
            if matches > 0:
                chosen = f
                reason = f"Correspond à la règle : {f.name}"
                confidence = "high"
                break

        results.append(
            ClassificationResult(
                item_id=e.item_id,
                target_folder_name=chosen.name,
                target_folder_id=chosen.folder_id,
                confidence=confidence,
                justification=reason,
            )
        )
    return results


async def _call_gemini(system: str, user: str) -> str:
    api_key = _env("GEMINI_API_KEY")
    primary_model = _env("GEMINI_MODEL", DEFAULT_GEMINI_MODEL)
    # Repli sur des modeles verifies comme disponibles avec la cle du service.
    # « gemini-flash-latest » suit les mises a jour de Google et sert de filet.
    candidate_models = [primary_model]
    for m in ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-flash-latest"]:
        if m not in candidate_models:
            candidate_models.append(m)

    payload = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [{"role": "user", "parts": [{"text": user}]}],
        "generationConfig": {
            "temperature": 0.2,
            "responseMimeType": "application/json",
        },
    }
    headers = {"x-goog-api-key": api_key, "content-type": "application/json"}

    last_error = None
    for model in candidate_models:
        try:
            url = GEMINI_URL_TEMPLATE.format(model=model)
            data = await _post_json(url, payload, headers=headers, label=f"Gemini ({model})")
            parts = ((data.get("candidates", [{}])[0].get("content", {})).get("parts", [{}]))
            text = parts[0].get("text", "").strip() if parts else ""
            if text:
                return text
        except AIError as exc:
            last_error = exc
            err_str = str(exc).lower()
            # On passe au modele suivant si celui-ci n'existe pas, ou s'il est en
            # surtension : un autre modele a des chances de repondre.
            if exc.passager or "404" in err_str or "not found" in err_str:
                logger.warning("Modèle %s indisponible (%s), repli", model, exc)
                continue
            raise
    if last_error:
        raise last_error
    raise AIError("Aucun modèle Gemini disponible.")


async def _call_anthropic(system: str, user: str) -> str:
    api_key = _env("ANTHROPIC_API_KEY")
    model = _env("ANTHROPIC_MODEL", DEFAULT_ANTHROPIC_MODEL)
    payload = {
        "model": model,
        "max_tokens": 2048,
        "temperature": 0.2,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
        "content-type": "application/json",
    }
    data = await _post_json(ANTHROPIC_URL, payload, headers=headers, label="Anthropic")
    blocks = data.get("content") or []
    text = "".join(b.get("text", "") for b in blocks if isinstance(b, dict)).strip()
    return text


async def _call_deepseek(system: str, user: str) -> str:
    """Appelle DeepSeek, qui parle le dialecte d'OpenAI.

    deepseek-chat rend un JSON propre en quelques dixiemes de seconde. On ecarte
    deepseek-reasoner : avec response_format il consomme ses jetons en
    raisonnement et rend un contenu vide.
    """
    api_key = _env("DEEPSEEK_API_KEY")
    model = _env("DEEPSEEK_MODEL", DEFAULT_DEEPSEEK_MODEL)
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": 0.2,
        "max_tokens": 2048,
        "response_format": {"type": "json_object"},
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "content-type": "application/json",
    }
    data = await _post_json(DEEPSEEK_URL, payload, headers=headers, label=f"DeepSeek ({model})")
    choices = data.get("choices") or []
    if not choices:
        raise AIError(f"DeepSeek ({model}) : reponse sans choix.")
    return (choices[0].get("message", {}) or {}).get("content", "").strip()


async def _post_json(url: str, payload: dict[str, Any], headers: dict[str, str], label: str) -> dict[str, Any]:
    """Appelle l'API en reessayant les erreurs passageres.

    Une surtension du fournisseur (429, 503...) ne doit pas faire echouer un tri :
    on reessaie jusqu'a TENTATIVES_PAR_MODELE fois, avec une attente croissante.
    Les autres erreurs (400, 401, 404) sont definitives et remontees aussitot.
    """
    derniere: AIError | None = None

    for essai in range(TENTATIVES_PAR_MODELE):
        if essai:
            attente = ATTENTES[min(essai, len(ATTENTES) - 1)]
            logger.info("%s : surtension, nouvel essai %d/%d dans %.1f s",
                        label, essai + 1, TENTATIVES_PAR_MODELE, attente)
            await asyncio.sleep(attente)

        try:
            async with httpx.AsyncClient(timeout=REQUEST_TIMEOUT) as client:
                resp = await client.post(url, json=payload, headers=headers)
        except Exception as exc:
            derniere = AIError(f"{label} : erreur réseau ({exc.__class__.__name__}).", statut=None)
            continue

        if resp.status_code < 400:
            return resp.json()

        message = f"{label} : erreur {resp.status_code} — {resp.text[:200]}"
        if resp.status_code not in STATUTS_REESSAYABLES:
            raise AIError(message, statut=resp.status_code)

        derniere = AIError(message, statut=resp.status_code)

    assert derniere is not None
    raise derniere

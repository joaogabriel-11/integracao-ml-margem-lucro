import json
import os
import time

import requests
from dotenv import load_dotenv

load_dotenv()

API = "https://api.mercadolibre.com"
CLIENT_ID = os.getenv("ML_CLIENT_ID")
CLIENT_SECRET = os.getenv("ML_CLIENT_SECRET")
REDIRECT_URI = os.getenv("ML_REDIRECT_URI")
TOKEN_FILE = "tokens.json"

HEADERS = {
    "accept": "application/json",
    "content-type": "application/x-www-form-urlencoded",
}


def _salvar_tokens(data: dict) -> dict:
    data["obtido_em"] = int(time.time())
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    return data


def _carregar_tokens() -> dict:
    with open(TOKEN_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def trocar_code_por_token(code: str) -> dict:
    """Primeira autenticação: troca o código TG- pelo access_token."""
    payload = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
    }
    r = requests.post(f"{API}/oauth/token", headers=HEADERS, data=payload, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"Erro {r.status_code}: {r.text}")
    return _salvar_tokens(r.json())


def renovar_token() -> dict:
    """Renova usando o refresh_token (o antigo é descartado, salvamos o novo)."""
    tokens = _carregar_tokens()
    payload = {
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": tokens["refresh_token"],
    }
    r = requests.post(f"{API}/oauth/token", headers=HEADERS, data=payload, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"Erro ao renovar {r.status_code}: {r.text}")
    return _salvar_tokens(r.json())


def get_access_token() -> str:
    """Retorna um token válido, renovando automaticamente se estiver perto de expirar."""
    tokens = _carregar_tokens()
    expira_em = tokens["obtido_em"] + tokens["expires_in"] - 300  # margem de 5 min
    if time.time() >= expira_em:
        tokens = renovar_token()
    return tokens["access_token"]


def auth_headers() -> dict:
    return {"Authorization": f"Bearer {get_access_token()}"}
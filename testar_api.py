import requests
from ml_auth import API, auth_headers

# 1) Quem sou eu?
r = requests.get(f"{API}/users/me", headers=auth_headers(), timeout=30)
r.raise_for_status()
me = r.json()
print("Conta:", me["nickname"], "| ID:", me["id"])

# 2) Consegue listar anúncios ativos?
seller_id = me["id"]
r = requests.get(
    f"{API}/users/{seller_id}/items/search",
    headers=auth_headers(),
    params={"status": "active", "limit": 5},
    timeout=30,
)
r.raise_for_status()
dados = r.json()
print("Total de anúncios ativos:", dados["paging"]["total"])
print("Primeiros IDs:", dados["results"])
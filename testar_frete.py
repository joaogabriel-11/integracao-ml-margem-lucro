import json

import pandas as pd
import requests

from ml_auth import API, auth_headers

seller_id = requests.get(f"{API}/users/me", headers=auth_headers(), timeout=30).json()["id"]

df = pd.read_csv("anuncios_ml.csv")
amostra = df[df["frete_gratis"] == True].head(5)

for i, row in enumerate(amostra.itertuples()):
    params = {
        "item_id": row.item_id,
        "verbose": "true",
        "free_shipping": "true",
        "mode": "me2",
        "logistic_type": "xd_drop_off",
    }
    r = requests.get(
        f"{API}/users/{seller_id}/shipping_options/free",
        headers=auth_headers(),
        params=params,
        timeout=30,
    )
    print(f"\n{row.item_id} | preço R$ {row.preco} | HTTP {r.status_code}")
    if r.status_code != 200:
        print("  erro:", r.text)
        continue
    d = r.json()
    custo = d.get("coverage", {}).get("all_country", {}).get("list_cost")
    print("  list_cost (vendedor):", custo)
    if i == 0:
        print("  resposta bruta:")
        print(json.dumps(d, indent=2, ensure_ascii=False)[:2500])
import time

import pandas as pd
import requests

from ml_auth import API, auth_headers


def get(path, **params):
    r = requests.get(f"{API}{path}", headers=auth_headers(), params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def listar_ids(seller_id):
    ids, offset = [], 0
    while True:
        d = get(f"/users/{seller_id}/items/search", status="active", limit=100, offset=offset)
        ids += d["results"]
        offset += 100
        if offset >= d["paging"]["total"]:
            return ids


def buscar_detalhes(ids):
    itens = []
    for i in range(0, len(ids), 20):  # multiget aceita até 20 IDs
        for x in get("/items", ids=",".join(ids[i:i + 20])):
            if x["code"] == 200:
                itens.append(x["body"])
            else:
                print("Falha em", x)
    return itens


def extrair_sku(item):
    if item.get("seller_custom_field"):
        return item["seller_custom_field"]
    for a in item.get("attributes", []):
        if a["id"] == "SELLER_SKU" and a.get("value_name"):
            return a["value_name"]
    for v in item.get("variations", []):
        if v.get("seller_custom_field"):
            return v["seller_custom_field"]
    return None


def buscar_comissao(item):
    d = get(
        "/sites/MLB/listing_prices",
        price=item["price"],
        listing_type_id=item["listing_type_id"],
        category_id=item["category_id"],
    )
    if isinstance(d, list):
        d = next(x for x in d if x["listing_type_id"] == item["listing_type_id"])
    return d["sale_fee_amount"]


if __name__ == "__main__":
    seller_id = get("/users/me")["id"]
    ids = listar_ids(seller_id)
    print(f"{len(ids)} anúncios ativos encontrados")

    itens = buscar_detalhes(ids)

    linhas = []
    for n, it in enumerate(itens, 1):
        ship = it.get("shipping", {})
        linhas.append({
            "item_id": it["id"],
            "titulo": it["title"],
            "sku": extrair_sku(it),
            "preco": it["price"],
            "categoria": it["category_id"],
            "tipo_anuncio": it["listing_type_id"],
            "comissao": buscar_comissao(it),
            "frete_gratis": ship.get("free_shipping"),
            "logistica": ship.get("logistic_type"),
            "modo_envio": ship.get("mode"),
        })
        if n % 50 == 0:
            print(f"  {n}/{len(itens)} processados")
        time.sleep(0.1)

    df = pd.DataFrame(linhas)
    df.to_csv("anuncios_ml.csv", index=False, encoding="utf-8-sig")

    print("\n--- Resumo ---")
    print("Anúncios com SKU:", df["sku"].notna().sum(), "de", len(df))
    print("Com frete grátis:", df["frete_gratis"].sum())
    print(df["logistica"].value_counts(dropna=False))
    print("\nExemplo de shipping (bruto):")
    print(itens[0].get("shipping"))
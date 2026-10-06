"""Auditor de Margem de Lucro - Mercado Livre.

Uso:
    python auditor_margem.py          # modo real: consulta a API do Mercado Livre
    python auditor_margem.py --demo   # modo demo: lê anuncios_ml.csv e custos.xlsx locais
"""
import argparse
import re
import time

import pandas as pd

# ================= CONFIGURAÇÃO =================
ARQUIVO_CUSTOS = "custos.xlsx"
COL_SKU = "SKU"
COL_CUSTO = "Custo do Fornecedor"
MARGEM_MIN_PCT = 10.0    # alerta se a margem (% sobre o preço) ficar abaixo disso
LUCRO_MIN_REAIS = 0.0    # alerta se o lucro em R$ ficar abaixo disso
ARQUIVO_SAIDA = "alertas_revisao_precos.xlsx"
LIMITE_TESTE = None      # modo real: processa só N anúncios (ex.: 20); None = todos
# ================================================


def normalizar_sku(valor):
    return str(valor).strip().upper()


def para_float(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = re.sub(r"[^\d,.\-]", "", str(valor))  # tira "R$", espaços etc.
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def carregar_custos(arquivo):
    df = pd.read_excel(arquivo, dtype={COL_SKU: str})
    for col in (COL_SKU, COL_CUSTO):
        if col not in df.columns:
            raise SystemExit(f"❌ Coluna '{col}' não encontrada. Colunas do arquivo: {list(df.columns)}")
    df = df[[COL_SKU, COL_CUSTO]].dropna(subset=[COL_SKU])
    df["sku_norm"] = df[COL_SKU].map(normalizar_sku)
    df["custo"] = df[COL_CUSTO].map(para_float)

    dup = df[df.duplicated("sku_norm", keep=False)]
    if not dup.empty:
        print(f"⚠️  {dup['sku_norm'].nunique()} SKUs repetidos no custos.xlsx (usando a 1ª ocorrência):")
        print("   ", ", ".join(sorted(dup["sku_norm"].unique())[:10]))
    return df.drop_duplicates("sku_norm").set_index("sku_norm")["custo"]


def montar_linha(base, comissao, frete, custo):
    """Calcula lucro e margem. Usado tanto no modo real quanto no demo."""
    linha = {**base, "comissao": comissao, "frete": frete, "custo": custo}
    if custo is None or pd.isna(custo):
        linha["status"] = "sem_custo"
        return linha

    preco = base["preco"]
    lucro = preco - comissao - frete - custo
    linha.update(
        status="ok",
        lucro=round(lucro, 2),
        margem_pct=round(lucro / preco * 100, 2),
    )
    # Preço estimado para atingir a margem mínima (aproximação: mantém a
    # taxa de comissão atual e o frete atual)
    denom = 1 - comissao / preco - MARGEM_MIN_PCT / 100
    if denom > 0:
        linha["preco_sugerido_aprox"] = round((custo + frete) / denom, 2)
    return linha


# ---------------------- MODO REAL (API) ----------------------

def buscar_frete_vendedor(seller_id, item):
    import requests
    from ml_auth import API, auth_headers

    ship = item.get("shipping", {})
    if not ship.get("free_shipping") or ship.get("logistic_type") in (None, "not_specified"):
        return 0.0  # comprador paga o frete, ou envio fora do Mercado Envios

    params = {
        "item_id": item["id"],
        "verbose": "true",
        "free_shipping": "true",
        "mode": ship.get("mode", "me2"),
        "logistic_type": ship["logistic_type"],
    }
    for tentativa in range(3):
        r = requests.get(
            f"{API}/users/{seller_id}/shipping_options/free",
            headers=auth_headers(), params=params, timeout=30,
        )
        if r.status_code == 429:
            time.sleep(2 * (tentativa + 1))
            continue
        r.raise_for_status()
        return r.json()["coverage"]["all_country"]["list_cost"]
    raise RuntimeError("limite de requisições (429)")


def coletar_via_api(custos):
    from coletar_anuncios import buscar_comissao, buscar_detalhes, extrair_sku, get, listar_ids

    seller_id = get("/users/me")["id"]
    ids = listar_ids(seller_id)
    itens = buscar_detalhes(ids)
    if LIMITE_TESTE:
        itens = itens[:LIMITE_TESTE]
    print(f"Auditando {len(itens)} anúncios na API...")

    linhas = []
    for n, it in enumerate(itens, 1):
        sku = extrair_sku(it)
        base = {
            "item_id": it["id"], "titulo": it["title"], "sku": sku,
            "preco": it["price"], "link": it.get("permalink"),
        }
        try:
            comissao = buscar_comissao(it)
            frete = buscar_frete_vendedor(seller_id, it)
        except Exception as e:
            linhas.append({**base, "status": f"erro_api: {e}"})
            continue

        custo = custos.get(normalizar_sku(sku)) if sku else None
        linhas.append(montar_linha(base, comissao, frete, custo))

        if n % 25 == 0:
            print(f"  {n}/{len(itens)}")
        time.sleep(0.1)
    return linhas


# ---------------------- MODO DEMO (CSV local) ----------------------

def coletar_via_csv(custos, caminho_csv):
    df = pd.read_csv(caminho_csv, dtype={"sku": str})
    obrigatorias = ["item_id", "titulo", "sku", "preco", "comissao", "frete"]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise SystemExit(
            f"❌ O CSV '{caminho_csv}' não tem a(s) coluna(s): {faltando}. "
            "O modo demo precisa da coluna 'frete' (o CSV gerado pela coleta real não tem)."
        )
    print(f"Modo demo: {len(df)} anúncios lidos de {caminho_csv} (sem chamar a API).")

    linhas = []
    for row in df.itertuples():
        base = {
            "item_id": row.item_id, "titulo": row.titulo, "sku": row.sku,
            "preco": float(row.preco), "link": None,
        }
        custo = custos.get(normalizar_sku(row.sku)) if pd.notna(row.sku) else None
        linhas.append(montar_linha(base, float(row.comissao), float(row.frete), custo))
    return linhas


# ---------------------- RELATÓRIO ----------------------

def gerar_relatorio(linhas, arquivo_saida):
    df = pd.DataFrame(linhas)
    ok = df[df["status"] == "ok"]
    alertas = ok[(ok["lucro"] < LUCRO_MIN_REAIS) | (ok["margem_pct"] < MARGEM_MIN_PCT)]
    alertas = alertas.sort_values("margem_pct")
    pendencias = df[df["status"] != "ok"]

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as xw:
        alertas.to_excel(xw, sheet_name="Alertas", index=False)
        pendencias.to_excel(xw, sheet_name="Sem custo ou erro", index=False)
        df.to_excel(xw, sheet_name="Auditoria completa", index=False)

    print("\n--- Resumo ---")
    print(f"Com margem calculada : {len(ok)}")
    print(f"Em alerta            : {len(alertas)}")
    print(f"Sem custo/erro       : {len(pendencias)}")
    print(f"📄 Arquivo gerado: {arquivo_saida}")


def main():
    parser = argparse.ArgumentParser(description="Auditor de margem de lucro - Mercado Livre")
    parser.add_argument("--demo", action="store_true",
                        help="usa o CSV local em vez da API (não precisa de credenciais)")
    parser.add_argument("--csv", default="anuncios_ml.csv",
                        help="CSV de anúncios usado no modo demo (padrão: anuncios_ml.csv)")
    parser.add_argument("--custos", default=ARQUIVO_CUSTOS,
                        help=f"planilha de custos (padrão: {ARQUIVO_CUSTOS})")
    args = parser.parse_args()

    custos = carregar_custos(args.custos)

    if args.demo:
        linhas = coletar_via_csv(custos, args.csv)
        saida = ARQUIVO_SAIDA.replace(".xlsx", "_demo.xlsx")  # não sobrescreve o relatório real
    else:
        linhas = coletar_via_api(custos)
        saida = ARQUIVO_SAIDA

    gerar_relatorio(linhas, saida)


if __name__ == "__main__":
    main()
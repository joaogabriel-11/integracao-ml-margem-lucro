import glob

import pandas as pd

arquivos = sorted(glob.glob("historico/auditoria_*.csv"))
if len(arquivos) < 2:
    raise SystemExit("Preciso de pelo menos 2 execuções em historico/ para comparar.")

antes = pd.read_csv(arquivos[-2]).set_index("item_id")
depois = pd.read_csv(arquivos[-1]).set_index("item_id")
comum = antes.index.intersection(depois.index)
a, d = antes.loc[comum], depois.loc[comum]

df = pd.DataFrame({
    "titulo": d["titulo"], "sku": d["sku"],
    "preco_antes": a["preco"], "preco_depois": d["preco"],
    "comissao_antes": a["comissao"], "comissao_depois": d["comissao"],
    "frete_antes": a["frete"], "frete_depois": d["frete"],
    "margem_antes": a["margem_pct"], "margem_depois": d["margem_pct"],
})
df["delta_margem_pp"] = (df["margem_depois"] - df["margem_antes"]).round(2)

mudou = df[
    ((df["comissao_antes"] - df["comissao_depois"]).abs() > 0.01)
    | ((df["frete_antes"] - df["frete_depois"]).abs() > 0.01)
].sort_values("delta_margem_pp")

mudou.to_excel("mudancas_tarifas.xlsx")
print(f"Comparando {arquivos[-2]} → {arquivos[-1]}")
print(f"{len(mudou)} anúncios com mudança de comissão ou frete")
"""Atualização de preços dos anúncios em alerta, via API do Mercado Livre.

Lê o alertas_revisao_precos.xlsx (gerado pelo auditor_margem.py) e leva cada anúncio
para o preço sugerido (coluna preco_sugerido_aprox), no modo automático ou assistido.

SEGURANÇA: por padrão é uma SIMULAÇÃO, nada é enviado ao Mercado Livre.
Só altera preços reais com --aplicar (e o app precisa de permissão de escrita).

Uso:
    python src/atualizar_precos.py             # simulação (lê saida/alertas_revisao_precos.xlsx)
    python src/atualizar_precos.py --aplicar   # altera preços de verdade (pede confirmação)
    python src/atualizar_precos.py --demo      # usa demo/saida/alertas_revisao_precos_demo.xlsx (sempre simulação)
    python src/atualizar_precos.py --max 1     # processa só 1 anúncio (bom para o primeiro teste real)
    python src/atualizar_precos.py --sim       # responde S à pergunta inicial (modo automático, sem perguntar)

Modo automático: aplica o preço sugerido em todos, sem perguntar anúncio a anúncio.
Modo assistido : mostra um anúncio por vez; você aceita o sugerido (ENTER) ou digita outro preço.

Regras de proteção:
    - sem preço sugerido, ou sugerido que não aumenta o preço .... IGNORADO (no automático)
    - aumento maior que LIMITE_AUMENTO_PCT ........................ CONFERIR (não é aplicado no automático)
    - anúncios com variações precisam de ajuste por variação (a API recusa) .. ERRO, e o robô segue
"""
import argparse
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

# ================= CONFIGURAÇÃO =================
ARQUIVO_ALERTAS = "saida/alertas_revisao_precos.xlsx"
ARQUIVO_RESULTADO = "saida/resultado_precos.xlsx"
PASTA_DEMO = Path("demo")
PASTA_LOGS = Path("saida/logs")
LIMITE_AUMENTO_PCT = 50.0       # no modo automático, aumentos acima disso ficam para conferência
COLUNAS_OBRIGATORIAS = ["item_id", "sku", "titulo", "preco", "preco_sugerido_aprox", "margem_pct"]
# ================================================

logger = logging.getLogger("atualizar_precos")


# ---------------------- utilitários ----------------------

def parse_brl(texto):
    """'R$ 1.234,56' -> 1234.56 | '849,90' -> 849.9 | '849.90' -> 849.9 | '1.234' -> 1234.0"""
    s = re.sub(r"[^\d,.\-]", "", str(texto))
    if not re.search(r"\d", s):
        raise ValueError(f"valor monetário inválido: {texto!r}")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        if s.count(".") > 1 or len(s.split(".")[-1]) == 3:
            s = s.replace(".", "")
    return float(s)


def brl(valor):
    texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def pct(valor):
    return "-" if valor is None or pd.isna(valor) else f"{valor:.1f}%".replace(".", ",")


def configurar_log():
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    arquivo = PASTA_LOGS / f"precos_{datetime.now():%Y-%m-%d}.log"
    if not logger.handlers:
        handler = logging.FileHandler(arquivo, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S"))
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    return arquivo


def falhar(mensagem):
    logger.error(mensagem)
    sys.exit(f"\n❌ {mensagem}")


def perguntar(texto, validas=None):
    """input() que repete até receber uma opção válida (ex.: S/N)."""
    while True:
        resposta = input(texto).strip().upper()
        if validas is None or resposta in validas:
            return resposta
        print(f"  Opção inválida. Digite uma destas: {', '.join(validas)}")


# ---------------------- leitura dos alertas ----------------------

def carregar_alertas(caminho):
    if not Path(caminho).is_file():
        falhar(f"Arquivo não encontrado: {caminho}. Rode o auditor_margem.py antes.")
    try:
        df = pd.read_excel(caminho, sheet_name="Alertas")
    except ValueError:
        falhar(f"{caminho} não tem a aba 'Alertas'. Foi gerado pelo auditor_margem.py?")
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltando:
        falhar(f"{caminho} não tem a(s) coluna(s): {faltando}")
    for coluna in ("preco", "preco_sugerido_aprox", "margem_pct"):
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce")
    return df


# ---------------------- envio do preço ----------------------

def enviar_preco(item_id, novo_preco, real):
    """Devolve (status, mensagem). Na simulação, nada é enviado."""
    if not real:
        return "SIMULADO", "simulação: nenhuma chamada enviada ao Mercado Livre"

    import requests
    from ml_auth import API, auth_headers

    try:
        r = requests.put(f"{API}/items/{item_id}", headers=auth_headers(),
                         json={"price": novo_preco}, timeout=30)
        if r.status_code == 403:
            return "ERRO", ("sem permissão de escrita: habilite 'Publicação e sincronização' como "
                            "Leitura e escrita no app e reautorize (gerar_token.py)")
        if r.status_code != 200:
            return "ERRO", f"HTTP {r.status_code}: {r.text[:200]}"

        # confirma lendo o anúncio de volta
        conferencia = requests.get(f"{API}/items/{item_id}", headers=auth_headers(), timeout=30)
        if conferencia.ok and abs(conferencia.json().get("price", -1) - novo_preco) < 0.005:
            return "ALTERADO", "preço confirmado na API"
        return "ERRO", "a API aceitou a alteração, mas o preço lido em seguida não confere"
    except requests.RequestException as e:
        return "ERRO", f"falha de rede: {e}"


def registrar(resultados, linha, aplicado, status, mensagem):
    resultados.append({
        "sku": linha.sku, "item_id": linha.item_id, "produto": linha.titulo,
        "preco_anterior": linha.preco, "preco_sugerido": linha.preco_sugerido_aprox,
        "preco_aplicado": aplicado, "status": status, "mensagem": mensagem,
        "data_hora": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
    })
    logger.info("SKU=%s | ITEM=%s | ANTES=%s | SUGERIDO=%s | NOVO=%s | STATUS=%s%s", linha.sku, linha.item_id,
                linha.preco, linha.preco_sugerido_aprox, aplicado if aplicado is not None else "-", status,
                f" | {mensagem}" if mensagem else "")


# ---------------------- modos de execução ----------------------

def modo_automatico(linhas, real, resultados):
    for n, linha in enumerate(linhas, start=1):
        sugerido, atual = linha.preco_sugerido_aprox, linha.preco
        if pd.isna(sugerido) or pd.isna(atual):
            status, msg, aplicado = "IGNORADO", "sem preço sugerido", None
        elif sugerido <= atual:
            status, msg, aplicado = "IGNORADO", "o preço sugerido não aumenta o preço atual", None
        elif (sugerido - atual) / atual * 100 > LIMITE_AUMENTO_PCT:
            status, msg, aplicado = "CONFERIR", f"aumento acima de {LIMITE_AUMENTO_PCT:g}%: confira antes de aplicar", None
        else:
            aplicado = round(float(sugerido), 2)
            status, msg = enviar_preco(linha.item_id, aplicado, real)
            time.sleep(0.2)

        registrar(resultados, linha, aplicado, status, msg)
        simbolo = {"ALTERADO": "✅", "SIMULADO": "🧪", "IGNORADO": "➖", "CONFERIR": "⚠️", "ERRO": "❌"}[status]
        detalhe = f"{brl(atual)} -> {brl(aplicado)}" if aplicado is not None else msg
        print(f"  [{n}/{len(linhas)}] {simbolo} {linha.sku:<10} {status:<9} {detalhe}")
    return False  # não foi cancelado


def modo_assistido(linhas, real, resultados):
    total = len(linhas)
    for n, linha in enumerate(linhas, start=1):
        sugerido = linha.preco_sugerido_aprox
        print("\n" + "=" * 40)
        print(f" ANÚNCIO {n} DE {total}")
        print("=" * 40 + "\n")
        print(f"SKU: {linha.sku}")
        print(f"Produto: {linha.titulo}")
        print(f"Preço atual: {brl(linha.preco)}")
        print(f"Preço sugerido: {brl(sugerido) if not pd.isna(sugerido) else '(sem sugestão)'}")
        print(f"Margem atual: {pct(linha.margem_pct)}")
        link = getattr(linha, "link", None)
        print(f"\nLink:\n{link if isinstance(link, str) and link else linha.item_id}\n")

        # preço escolhido
        while True:
            entrada = input("Digite o novo preço ou pressione ENTER para utilizar o preço sugerido:\n\nNovo preço:\n> ").strip()
            if not entrada:
                if pd.isna(sugerido):
                    print("  Não há preço sugerido: digite um valor.")
                    continue
                escolhido = round(float(sugerido), 2)
            else:
                try:
                    escolhido = round(parse_brl(entrada), 2)
                except ValueError:
                    print("  Valor inválido. Exemplos: 829.90 ou 829,90")
                    continue
            if escolhido <= 0:
                print("  O preço precisa ser maior que zero.")
                continue
            break
        print(f"\nPreço escolhido: {brl(escolhido)}")

        resposta = perguntar(
            f"\nConfirma alterar o preço deste anúncio\nde {brl(linha.preco)} para {brl(escolhido)}?\n\n"
            "[S] Sim\n[N] Não\n[C] Cancelar\n\n> ", {"S", "N", "C"})

        if resposta == "C":
            for pendente in linhas[n - 1:]:
                registrar(resultados, pendente, None, "PENDENTE", "execução cancelada pelo usuário")
            return True
        if resposta == "N":
            registrar(resultados, linha, None, "IGNORADO", "ignorado pelo usuário")
            print("\nStatus: IGNORADO")
            continue

        status, msg = enviar_preco(linha.item_id, escolhido, real)
        registrar(resultados, linha, escolhido, status, msg)
        if status in ("ALTERADO", "SIMULADO"):
            print("\nPreço alterado com sucesso." if status == "ALTERADO" else "\n🧪 Simulação: preço NÃO enviado ao Mercado Livre.")
        else:
            print(f"\n❌ Não foi possível alterar: {msg}")
        if n < total:
            input("\nPressione ENTER para continuar para o próximo anúncio.")
    return False


# ---------------------- resultado ----------------------

def salvar_resultado(resultados, caminho):
    colunas = ["sku", "item_id", "produto", "preco_anterior", "preco_sugerido", "preco_aplicado",
               "status", "mensagem", "data_hora"]
    Path(caminho).parent.mkdir(parents=True, exist_ok=True)
    try:
        pd.DataFrame(resultados, columns=colunas).to_excel(caminho, index=False)
    except PermissionError:
        falhar(f"Não consegui gravar {caminho}. Feche o arquivo no Excel e rode de novo.")


def main():
    inicio = time.monotonic()
    parser = argparse.ArgumentParser(description="Ajuste de preços dos anúncios em alerta (API do Mercado Livre)")
    parser.add_argument("--demo", action="store_true", help=f"usa {PASTA_DEMO}/saida/alertas_revisao_precos_demo.xlsx (sempre simulação)")
    parser.add_argument("--aplicar", action="store_true", help="altera os preços de verdade (padrão: simulação)")
    parser.add_argument("--sim", action="store_true", help="responde S à pergunta inicial (modo automático)")
    parser.add_argument("--max", type=int, default=None, help="processa no máximo N anúncios")
    parser.add_argument("--arquivo", default=None, help=f"planilha de alertas (padrão: {ARQUIVO_ALERTAS})")
    args = parser.parse_args()

    caminho = args.arquivo or (str(PASTA_DEMO / "saida" / "alertas_revisao_precos_demo.xlsx") if args.demo else ARQUIVO_ALERTAS)
    saida = PASTA_DEMO / "saida" / "resultado_precos_demo.xlsx" if args.demo else Path(ARQUIVO_RESULTADO)
    real = args.aplicar and not args.demo
    arquivo_log = configurar_log()
    logger.info("INÍCIO | modo=%s | arquivo=%s", "REAL" if real else "SIMULAÇÃO", caminho)

    df = carregar_alertas(caminho)
    if args.max:
        df = df.head(args.max)

    print("=" * 40)
    print(" ATUALIZAÇÃO DE PREÇOS")
    print("=" * 40)
    print(f"\nArquivo: {caminho}")
    print(f"Anúncios encontrados para revisão: {len(df)}")
    if args.aplicar and args.demo:
        print("ℹ️  --aplicar é ignorado no modo demo.")
    print("Modo: " + ("⚠️  REAL: os preços serão ALTERADOS no Mercado Livre" if real
                      else "🧪 SIMULAÇÃO: nenhum preço será enviado ao Mercado Livre"))

    if df.empty:
        print("\nNenhum anúncio em alerta. Nada a fazer.")
        logger.info("FIM | nenhum anúncio em alerta")
        return

    if real:
        print(f"\nIsto vai alterar o preço de até {len(df)} anúncios reais da conta.")
        if input("Digite CONFIRMAR para continuar: ").strip() != "CONFIRMAR":
            print("Cancelado. Nada foi alterado.")
            logger.info("FIM | cancelado antes de aplicar")
            return

    if args.sim:
        automatico = True
        print("\nModo automático (--sim): preço sugerido em todos os anúncios.")
    else:
        resposta = perguntar("\nDeseja alterar TODOS os anúncios\npara o preço sugerido automaticamente?\n\n"
                             "Digite S para SIM ou N para NÃO:\n> ", {"S", "N"})
        automatico = resposta == "S"

    linhas = list(df.itertuples())
    resultados = []
    if automatico:
        print("\nProcessando:")
        cancelado = modo_automatico(linhas, real, resultados)
    else:
        cancelado = modo_assistido(linhas, real, resultados)

    salvar_resultado(resultados, saida)

    contagem = pd.Series([r["status"] for r in resultados]).value_counts().to_dict()
    alterados = contagem.get("ALTERADO", 0) + contagem.get("SIMULADO", 0)
    duracao = int(time.monotonic() - inicio)
    tempo = f"{duracao // 60:02d}:{duracao % 60:02d}"
    logger.info("FIM | alterados=%s | ignorados=%s | conferir=%s | erros=%s | pendentes=%s | tempo=%s",
                alterados, contagem.get("IGNORADO", 0), contagem.get("CONFERIR", 0),
                contagem.get("ERRO", 0), contagem.get("PENDENTE", 0), tempo)

    print("\n" + "=" * 40)
    print(" EXECUÇÃO ENCERRADA PELO USUÁRIO" if cancelado else " EXECUÇÃO FINALIZADA")
    print("=" * 40)
    print(f"Total de anúncios       : {len(resultados)}")
    print(f"{'Simulados' if not real else 'Alterados':<24}: {alterados}")
    print(f"Ignorados               : {contagem.get('IGNORADO', 0)}")
    print(f"Para conferir           : {contagem.get('CONFERIR', 0)}")
    print(f"Erros                   : {contagem.get('ERRO', 0)}")
    if cancelado:
        print(f"Pendentes               : {contagem.get('PENDENTE', 0)}")
    print(f"Tempo de execução       : {tempo}")

    erros = [r for r in resultados if r["status"] == "ERRO"]
    if erros:
        print("\nFalhas:")
        for r in erros:
            print(f"- SKU {r['sku']} - {r['mensagem']}")
    conferir = [r for r in resultados if r["status"] == "CONFERIR"]
    if conferir:
        print("\nPara conferir (não aplicados):")
        for r in conferir:
            print(f"- SKU {r['sku']}: {brl(r['preco_anterior'])} -> {brl(r['preco_sugerido'])}")
    print(f"\nResultado: {saida}")
    print(f"Log      : {arquivo_log}")
    if real and alterados:
        print("\nDica: rode o auditor_margem.py de novo para conferir as novas margens.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        logger.exception("Erro inesperado")
        raise
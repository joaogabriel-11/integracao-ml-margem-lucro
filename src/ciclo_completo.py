"""Ciclo completo: custos (RPA) -> auditoria de margem -> ajuste de preços.

    fornecedor --(RPA)--> custos.xlsx --(auditor)--> alertas --(API)--> preços

Cada etapa continua sendo um script independente; este arquivo só coordena a ordem.

Uso:
    python src/ciclo_completo.py --demo            # demonstração completa, 1 comando, sem credenciais reais
    python src/ciclo_completo.py --demo --visivel  # idem, mostrando o navegador do robô

    python src/ciclo_completo.py                   # real: robô de custos + auditoria
    python src/ciclo_completo.py --sem-custos      # real: só auditoria (quando não há portal)
    python src/ciclo_completo.py --precos          # real: inclui o ajuste de preços (simulação)
    python src/ciclo_completo.py --precos --aplicar  # real: ajuste de preços de verdade (pede confirmação)

No modo --demo o portal do fornecedor simulado é iniciado e encerrado automaticamente.
"""
import argparse
import os
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent   # este arquivo fica em src/
SRC = Path("src")
PORTAL_URL = os.getenv("PORTAL_URL", "http://127.0.0.1:5000")
DEMO = Path("demo")
DEMO_SAIDA = DEMO / "saida"
DEMO_RELATORIO_CUSTOS = RAIZ / DEMO_SAIDA / "relatorio_custos_demo.xlsx"


# ---------------------- utilitários ----------------------

def titulo(texto):
    print("\n" + "=" * 60)
    print(f" {texto}")
    print("=" * 60)


def executar(etapa, script, *argumentos, env=None):
    """Roda um script do projeto; se falhar, interrompe o ciclo."""
    caminho = SRC / script
    print(f"\n▶ {etapa}\n  $ python {caminho.as_posix()} {' '.join(argumentos)}\n")
    inicio = time.monotonic()
    resultado = subprocess.run([sys.executable, str(caminho), *argumentos], cwd=RAIZ, env=env)
    if resultado.returncode != 0:
        sys.exit(f"\n❌ A etapa '{etapa}' falhou (código {resultado.returncode}). Ciclo interrompido.")
    print(f"\n  ✔ concluída em {time.monotonic() - inicio:.0f}s")


def portal_no_ar():
    try:
        urllib.request.urlopen(f"{PORTAL_URL}/login", timeout=1)
        return True
    except Exception:
        return False


def iniciar_portal():
    """Sobe o portal simulado, a menos que já esteja rodando. Devolve o processo (ou None)."""
    if portal_no_ar():
        print("Portal simulado já estava no ar: será reaproveitado.")
        return None

    app = RAIZ / DEMO / "portal_fornecedor" / "app.py"
    if not app.is_file():
        sys.exit(f"❌ Não encontrei o portal simulado em: {app}\n"
                 "   Confira se a pasta demo/portal_fornecedor/ está no projeto.")

    log = RAIZ / "saida" / "logs" / "portal_demo.log"   # a saída do portal fica registrada aqui
    log.parent.mkdir(parents=True, exist_ok=True)
    with open(log, "w", encoding="utf-8") as saida:
        processo = subprocess.Popen([sys.executable, str(app)], cwd=RAIZ,
                                    stdout=saida, stderr=subprocess.STDOUT)
    for _ in range(40):
        if portal_no_ar():
            print(f"Portal simulado iniciado em {PORTAL_URL}")
            return processo
        if processo.poll() is not None:
            break
        time.sleep(0.25)

    processo.terminate()
    detalhe = log.read_text(encoding="utf-8", errors="replace").strip()[-1500:]
    sys.exit("❌ Não consegui iniciar o portal simulado.\n"
             + (f"\nErro registrado pelo portal:\n{detalhe}\n" if detalhe else "")
             + "\nPossíveis causas:\n"
               "  - Flask não instalado neste Python (pip install flask)\n"
               "  - Porta 5000 ocupada por outro programa (PowerShell: netstat -ano | findstr :5000)\n"
               f"  - Veja o log completo em: {log}")


def pct(v):
    return "-" if v is None or pd.isna(v) else f"{v:.1f}%".replace(".", ",")


def reais(v):
    return "-" if v is None or pd.isna(v) else f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ---------------------- resumo final ----------------------

def resumo_demo(antes_xlsx, depois_xlsx, precos_xlsx):
    antes = pd.read_excel(antes_xlsx, sheet_name="Auditoria completa").set_index("sku")
    depois = pd.read_excel(depois_xlsx, sheet_name="Auditoria completa").set_index("sku")

    titulo("RESULTADO DO CICLO")
    print("\nO que mudou quando os custos do fornecedor foram atualizados:\n")
    print(f"{'SKU':<10} {'Custo (R$)':<19} {'Margem':<17} {'Situação'}")
    print("-" * 62)
    for sku in depois.index:
        a, d = antes.loc[sku], depois.loc[sku]
        print(f"{sku:<10} {reais(a['custo']) + ' → ' + reais(d['custo']):<19} "
              f"{pct(a['margem_pct']) + ' → ' + pct(d['margem_pct']):<17} {a['situacao']} → {d['situacao']}")

    mudancas = pd.read_excel(DEMO_RELATORIO_CUSTOS, sheet_name="Mudanças")
    retidos = mudancas[mudancas["status"].isin(["CONFERIR", "ERRO DE LEITURA"])]
    for r in retidos.itertuples():
        novo = "ilegível" if pd.isna(r.custo_novo) else f"{reais(r.custo_novo)} ({r.variacao_pct:+.1f}%)".replace(".", ",")
        print(f"\n⚠️  {r.sku}: custo do portal {novo} retido para conferência com o fornecedor ({r.motivo}).")

    n_antes = (antes["situacao"] == "ALERTA").sum()
    n_depois = (depois["situacao"] == "ALERTA").sum()
    print(f"\nAnúncios em alerta: {n_antes} (antes) → {n_depois} (depois dos novos custos)")

    precos = pd.read_excel(precos_xlsx)
    if len(precos):
        print("\nAjuste de preço sugerido para os anúncios em alerta (SIMULAÇÃO):\n")
        print(f"{'SKU':<10} {'Preço atual':>12} {'Sugerido':>12}   Status")
        print("-" * 52)
        for r in precos.itertuples():
            print(f"{r.sku:<10} {reais(r.preco_anterior):>12} {reais(r.preco_sugerido):>12}   {r.status}")

    print("\nArquivos gerados em demo/saida/:")
    for nome in ("relatorio_custos_demo.xlsx", "custos_atualizado.xlsx",
                 "alertas_revisao_precos_demo.xlsx", "resultado_precos_demo.xlsx"):
        print(f"  - {nome}")
    print("\nOs arquivos originais do demo/ não foram alterados.")


# ---------------------- ciclos ----------------------

def ciclo_demo(visivel):
    env = {**os.environ,
           "PORTAL_USUARIO": os.environ.get("PORTAL_USUARIO", "demo"),   # credenciais FICTÍCIAS do portal simulado
           "PORTAL_SENHA": os.environ.get("PORTAL_SENHA", "demo123")}
    antes = DEMO_SAIDA / "alertas_antes_demo.xlsx"
    depois = DEMO_SAIDA / "alertas_revisao_precos_demo.xlsx"

    titulo("CICLO COMPLETO (DEMONSTRAÇÃO, dados fictícios)")
    portal = iniciar_portal()
    try:
        executar("1/4  Auditoria ANTES (custos atuais)", "auditor_margem.py", "--demo")
        (RAIZ / DEMO_SAIDA).mkdir(parents=True, exist_ok=True)
        shutil.copy2(RAIZ / depois, RAIZ / antes)

        rpa = ["--demo"] + (["--visivel"] if visivel else [])
        executar("2/4  Robô atualiza os custos a partir do portal do fornecedor", "rpa_custos.py", *rpa, env=env)

        executar("3/4  Auditoria DEPOIS (custos novos)", "auditor_margem.py", "--demo",
                 "--custos", str(DEMO_SAIDA / "custos_atualizado.xlsx"))

        executar("4/4  Ajuste de preços dos anúncios em alerta", "atualizar_precos.py", "--demo", "--sim")
    finally:
        if portal is not None:
            portal.terminate()
            print("\nPortal simulado encerrado.")

    resumo_demo(RAIZ / antes, RAIZ / depois, RAIZ / DEMO_SAIDA / "resultado_precos_demo.xlsx")


def ciclo_real(args):
    titulo("CICLO COMPLETO")
    passo = 0
    total = (0 if args.sem_custos else 1) + 1 + (1 if args.precos else 0)

    if not args.sem_custos:
        passo += 1
        executar(f"{passo}/{total}  Robô atualiza os custos", "rpa_custos.py", *(["--visivel"] if args.visivel else []))
    passo += 1
    executar(f"{passo}/{total}  Auditoria de margem", "auditor_margem.py")
    if args.precos:
        passo += 1
        executar(f"{passo}/{total}  Ajuste de preços", "atualizar_precos.py", *(["--aplicar"] if args.aplicar else []))

    titulo("CICLO FINALIZADO")
    try:
        alertas = pd.read_excel(RAIZ / "saida" / "alertas_revisao_precos.xlsx", sheet_name="Alertas")
        print(f"Anúncios em alerta no relatório: {len(alertas)}  (saida/alertas_revisao_precos.xlsx)")
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser(description="Ciclo completo: custos (RPA) -> auditoria -> preços")
    parser.add_argument("--demo", action="store_true", help="demonstração completa com dados fictícios, em um comando")
    parser.add_argument("--visivel", action="store_true", help="mostra o navegador do robô de custos")
    parser.add_argument("--sem-custos", action="store_true", help="real: pula o robô de custos")
    parser.add_argument("--precos", action="store_true", help="real: inclui o ajuste de preços")
    parser.add_argument("--aplicar", action="store_true", help="real: com --precos, altera os preços de verdade")
    args = parser.parse_args()

    if args.aplicar and not args.precos:
        sys.exit("❌ --aplicar só faz sentido junto com --precos.")

    inicio = time.monotonic()
    if args.demo:
        ciclo_demo(args.visivel)
    else:
        ciclo_real(args)
    print(f"\nTempo total: {time.monotonic() - inicio:.0f}s")


if __name__ == "__main__":
    main()
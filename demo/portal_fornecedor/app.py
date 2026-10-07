"""Portal de fornecedor SIMULADO (todos os dados são fictícios).

Existe só para o RPA de custos ter o que automatizar: login, tabela paginada
com valores no formato brasileiro e botão de próxima página.

Rodar:
    python demo/portal_fornecedor/app.py
Acessar: http://127.0.0.1:5000   (usuário: demo | senha: demo123)
"""
import os
import random
from functools import wraps

from flask import Flask, redirect, render_template, request, session, url_for

app = Flask(__name__)
app.secret_key = os.getenv("PORTAL_SECRET", "chave-de-demonstracao-nao-usar-em-producao")

# Credenciais FICTÍCIAS, só do portal simulado.
USUARIO = os.getenv("PORTAL_USUARIO", "demo")
SENHA = os.getenv("PORTAL_SENHA", "demo123")
POR_PAGINA = 10

# Os 6 primeiros SKUs existem no demo/custos.xlsx, com custos novos:
#   DEMO-001 +4%, DEMO-002 +0,4%, DEMO-003 +45% (suspeito!), DEMO-004 -11,5%,
#   DEMO-005 sem mudança, DEMO-006 +3,2%.
BASE = [
    ("DEMO-001", "Caixa de Direção Hidráulica Modelo A 2005-2012", 301.50),
    ("DEMO-002", "Bomba de Combustível Elétrica Modelo B", 236.00),
    ("DEMO-003", "Coxim do Motor Dianteiro Modelo C", 205.90),
    ("DEMO-004", "Pastilha de Freio Dianteira Modelo D", 88.50),
    ("DEMO-005", "Kit Correia Dentada com Tensor Modelo E", 165.00),
    ("DEMO-006", "Junta da Tampa de Válvulas Modelo F", 64.00),
]
PECAS = [
    "Amortecedor Traseiro", "Disco de Freio", "Filtro de Óleo", "Vela de Ignição",
    "Bieleta da Barra Estabilizadora", "Terminal de Direção", "Pivô da Suspensão",
    "Rolamento de Roda", "Bateria Automotiva", "Sensor de Temperatura",
]
MODELOS = ["Modelo G", "Modelo H", "Modelo J"]


def gerar_produtos():
    rng = random.Random(42)  # determinístico: mesmos dados em toda execução
    produtos = [{"sku": s, "descricao": d, "custo": c} for s, d, c in BASE]
    for i in range(7, 31):
        n = i - 7
        produtos.append({
            "sku": f"DEMO-{i:03d}",
            "descricao": f"{PECAS[n % len(PECAS)]} {MODELOS[n % len(MODELOS)]}",
            "custo": round(rng.uniform(25, 480), 2),
        })
    return produtos


PRODUTOS = gerar_produtos()


@app.template_filter("brl")
def brl(valor):
    """1234.5 -> 'R$ 1.234,50' (formato brasileiro, como num portal real)."""
    texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def login_obrigatorio(rota):
    @wraps(rota)
    def protegida(*args, **kwargs):
        if not session.get("logado"):
            return redirect(url_for("login"))
        return rota(*args, **kwargs)
    return protegida


@app.route("/")
def inicio():
    return redirect(url_for("produtos" if session.get("logado") else "login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    erro = None
    if request.method == "POST":
        if request.form.get("usuario") == USUARIO and request.form.get("senha") == SENHA:
            session["logado"] = True
            return redirect(url_for("produtos"))
        erro = "Usuário ou senha inválidos."
    return render_template("login.html", erro=erro)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/produtos")
@login_obrigatorio
def produtos():
    total = len(PRODUTOS)
    paginas = (total + POR_PAGINA - 1) // POR_PAGINA
    pagina = min(max(request.args.get("pagina", 1, type=int), 1), paginas)
    inicio_ = (pagina - 1) * POR_PAGINA
    return render_template(
        "produtos.html",
        itens=PRODUTOS[inicio_:inicio_ + POR_PAGINA],
        pagina=pagina, paginas=paginas, total=total,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)

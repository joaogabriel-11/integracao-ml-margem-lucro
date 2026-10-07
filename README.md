# Integração Mercado Livre — Margem de Lucro

Sistema desenvolvido para automatizar o cruzamento entre os anúncios do **Mercado Livre** e os produtos de um fornecedor, permitindo analisar custos, preços e margens de lucro.

## Funcionalidade

O sistema utiliza a **API oficial do Mercado Livre** para consultar os anúncios ativos da conta e cruza essas informações com os dados obtidos diretamente do **site do fornecedor da empresa**.

O fluxo principal é:

1. Consulta os anúncios ativos no Mercado Livre.
2. Acessa o portal do fornecedor.
3. Cataloga os produtos encontrados e seus respectivos preços de custo.
4. Cruza os anúncios com os produtos do fornecedor.
5. Calcula custos, preços e margem de lucro.
6. Gera relatórios com os resultados.
7. Opcionalmente, pode **alterar automaticamente os dados dos anúncios no Mercado Livre**, de acordo com as regras configuradas.

O projeto também possui um **modo Demo**, que permite testar todo o funcionamento sem precisar configurar uma conta real do Mercado Livre ou do fornecedor.

## Tecnologias

* Python
* API do Mercado Livre
* Selenium
* Pandas
* Requests
* Automação Web
* OAuth 2.0

## Como executar

### 1. Clone o repositório

```bash
git clone https://github.com/joaogabriel-11/integracao-ml-margem-lucro.git
cd integracao-ml-margem-lucro
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

### 3. Execute a Demo

A versão Demo já possui os dados necessários para teste e pode ser executada normalmente:

```bash
python demo\portal_fornecedor\app.py
```

A Demo permite visualizar o funcionamento do sistema sem precisar configurar credenciais reais.

### 4. Executando a versão real

A versão real requer a configuração das credenciais e parâmetros da integração com o Mercado Livre e com o portal do fornecedor.

Após a configuração, o sistema pode ser executado pelos scripts disponíveis no diretório `src`.

## Estrutura

```text
├── demo/
│   └── portal_fornecedor/
│       └── app.py
├── src/
│   ├── ...
├── requirements.txt
└── README.md
```

## Observação

As alterações nos anúncios do Mercado Livre são **opcionais**. O sistema pode ser utilizado apenas para consulta, cruzamento dos dados e geração de relatórios, sem modificar os anúncios.

---

Desenvolvido por **João Gabriel dos Santos**

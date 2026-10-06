# Auditor de Margem de Lucro: Mercado Livre

Script em Python que consulta os anúncios ativos da conta no Mercado Livre, cruza com a planilha de custos e avisa quais anúncios estão com a margem abaixo do aceitável.

**Cálculo da margem:**

```
Lucro = Preço de venda − Comissão ML − Frete do vendedor − Custo do fornecedor
Margem % = Lucro ÷ Preço de venda × 100
```

> **Quer só ver funcionando?** Pule para a seção [6. Testar sem conta do Mercado Livre (modo demo)](#6-testar-sem-conta-do-mercado-livre-modo-demo): não precisa de credenciais.

---

## 1. Arquivos do projeto

| Arquivo                       | Para que serve                                                                   |
| ----------------------------- | -------------------------------------------------------------------------------- |
| `auditor_margem.py`           | **Script principal.** Gera o relatório de alertas (modo real ou `--demo`)        |
| `custos.xlsx`                 | Planilha com o SKU e o custo de cada peça (no repositório, é o exemplo fictício) |
| `anuncios_ml.csv`             | Anúncios de exemplo (fictícios), usados apenas no modo demo                      |
| `alertas_revisao_precos.xlsx` | Relatório gerado (é sobrescrito a cada execução)                                 |
| `.env`                        | Credenciais do app (App ID, Secret Key e redirect URI)                           |
| `tokens.json`                 | Token de acesso, gerado automaticamente                                          |
| `ml_auth.py`                  | Autenticação e renovação do token                                                |
| `gerar_token.py`              | Gera o token na primeira vez (ou quando precisar reautorizar)                    |
| `testar_api.py`               | Testa se a conexão com o Mercado Livre está funcionando                          |
| `coletar_anuncios.py`         | Funções de coleta de anúncios (usado pelo auditor)                               |
| `testar_frete.py`             | Teste da consulta de frete (opcional)                                            |

---

## 2. Instalação (uma vez só)

```bash
pip install requests python-dotenv pandas openpyxl
pip install --upgrade openpyxl
```

Crie o arquivo `.env` na pasta do projeto:

```env
ML_CLIENT_ID=seu_app_id
ML_CLIENT_SECRET=sua_secret_key
ML_REDIRECT_URI=https://www.google.com
```

> O `ML_REDIRECT_URI` precisa ser **idêntico** ao cadastrado no app do Mercado Livre Developers (sem barra no final).

**Configuração do app no painel do Mercado Livre Developers:**

- Fluxos OAuth: **Authorization Code** e **Refresh Token** marcados
- PKCE: desmarcado
- Permissões: _Leitura_ em Publicação e sincronização, Faturamento e Venda e envios de um produto

---

## 3. Primeira autenticação (uma vez só)

1. Acesse o link abaixo utilizando a conta do Mercado Livre autorizada para o aplicativo.
2. Substitua `SEU_APP_ID` pelo `Client ID` da aplicação antes de acessar o link:

   ```
   https://auth.mercadolivre.com.br/authorization?response_type=code&client_id=SEU_APP_ID&redirect_uri=https://www.google.com
   ```

2. Clique em **Permitir**. Você será levado ao Google, e a URL terá o código:
   `https://www.google.com/?code=TG-xxxxxxxx...`

3. Rode o script e cole a URL inteira (ou só o código `TG-...`):

   ```bash
   python gerar_token.py
   ```

   O resultado esperado é `✅ Token gerado!` e `Tem refresh_token?: True`.

4. Teste a conexão:

   ```bash
   python testar_api.py
   ```

> O código `TG-` é de uso único e expira em poucos minutos. Se der `invalid_grant`, gere outro.
> Depois disso, o token é renovado automaticamente. Só repita esta etapa se o acesso for revogado ou o `tokens.json` for apagado.

---

## 4. Preparar o `custos.xlsx`

Duas colunas, na primeira aba, com cabeçalho na linha 1:

| SKU     | Custo do Fornecedor |
| ------- | ------------------- |
| ABC-123 | 85,90               |
| XYZ-456 | 210,00              |

- O SKU deve ser **igual** ao cadastrado no anúncio do Mercado Livre (maiúsculas e minúsculas não importam).
- O mesmo SKU pode aparecer várias vezes se o custo for igual (ex.: uma peça compatível com vários carros).
- Custo aceito como `85.9`, `85,90` ou `R$ 85,90`.
- Sempre que mudar um custo, **salve o arquivo** e rode o auditor de novo.

---

## 5. Rodar o auditor

Feche o `alertas_revisao_precos.xlsx` (se estiver aberto) e execute:

```bash
python auditor_margem.py
```

No final aparece um resumo:

```
Com margem calculada : 301
Em alerta            : 12
Sem custo/erro       : 0
📄 Arquivo gerado: alertas_revisao_precos.xlsx
```

O arquivo `alertas_revisao_precos.xlsx` é **sobrescrito** a cada execução, e não precisa apagar o anterior.

### O que tem no relatório

| Aba                    | Conteúdo                                                                                  |
| ---------------------- | ----------------------------------------------------------------------------------------- |
| **Alertas**            | Anúncios abaixo da margem mínima, do pior para o melhor, com um preço sugerido aproximado |
| **Sem custo ou erro**  | SKUs que não foram achados no `custos.xlsx` ou anúncios com erro na consulta              |
| **Auditoria completa** | Todos os anúncios com comissão, frete, custo, lucro e margem                              |

---

## 6. Testar sem conta do Mercado Livre (modo demo)

Para ver o auditor funcionando **sem `.env`, sem token e sem chamar a API**, use o modo demo:

```bash
python auditor_margem.py --demo
```

Ele lê o `anuncios_ml.csv` e o `custos.xlsx` da pasta (6 produtos **fictícios**) e gera o relatório `alertas_revisao_precos_demo.xlsx`. O relatório real (`alertas_revisao_precos.xlsx`) não é sobrescrito.

Resultado esperado no terminal:

```
Modo demo: 6 anúncios lidos de anuncios_ml.csv (sem chamar a API).

--- Resumo ---
Com margem calculada : 6
Em alerta            : 4
Sem custo/erro       : 0
📄 Arquivo gerado: alertas_revisao_precos_demo.xlsx
```

Com a margem mínima padrão de 10%, o resultado esperado é:

| SKU      | Preço (R$) | Lucro (R$) | Margem | Resultado                      |
| -------- | ---------- | ---------- | ------ | ------------------------------ |
| DEMO-001 | 520,00     | 103,10     | 19,8%  | OK                             |
| DEMO-002 | 340,00     | 14,00      | 4,1%   | **ALERTA** (margem baixa)      |
| DEMO-003 | 210,00     | 2,40       | 1,1%   | **ALERTA** (margem quase zero) |
| DEMO-004 | 140,00     | −11,20     | −8,0%  | **ALERTA** (prejuízo)          |
| DEMO-005 | 280,00     | 33,30      | 11,9%  | OK                             |
| DEMO-006 | 70,00      | −3,90      | −5,6%  | **ALERTA** (prejuízo)          |

O `custos.xlsx` de exemplo tem uma segunda aba, **Comentários**, com a conta de cada produto e a explicação de cada resultado.

**Como o modo demo funciona:**

- Precisa apenas de `pandas` e `openpyxl` (não usa `requests`, `.env` nem `tokens.json`).
- O `anuncios_ml.csv` do demo tem as colunas `item_id`, `titulo`, `sku`, `preco`, `categoria`, `tipo_anuncio`, `comissao`, `frete_gratis`, `logistica`, `modo_envio` e **`frete`**. No modo real o frete vem da API; no demo ele vem dessa coluna.
- Para usar outros arquivos: `python auditor_margem.py --demo --csv caminho/anuncios.csv --custos caminho/custos.xlsx`

> ⚠️ O `coletar_anuncios.py` grava a coleta real em `anuncios_ml.csv`, o mesmo nome do arquivo de exemplo. Se rodá-lo, você sobrescreve o demo (e o CSV real não tem a coluna `frete`). Mantenha o exemplo em outra pasta ou mude o nome de saída da coleta.

---

## 7. Ajustar a margem mínima

No topo do `auditor_margem.py`:

```python
ARQUIVO_CUSTOS = "custos.xlsx"
COL_SKU = "SKU"
COL_CUSTO = "Custo do Fornecedor"
MARGEM_MIN_PCT = 10.0    # alerta se a margem (%) ficar abaixo disso
LUCRO_MIN_REAIS = 0.0    # alerta se o lucro (R$) ficar abaixo disso
ARQUIVO_SAIDA = "alertas_revisao_precos.xlsx"
LIMITE_TESTE = None      # modo real: use um número (ex.: 20) para testar com poucos anúncios
```

---

## 8. Problemas comuns

| Mensagem / sintoma                                       | O que fazer                                                                               |
| -------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| `invalid_grant` ao gerar o token                         | Código expirado ou já usado, ou `redirect_uri` diferente do cadastrado. Gere outro código |
| `403 PA_UNAUTHORIZED_RESULT_FROM_POLICIES`               | Falta a permissão de envios no app. Libere _Leitura_, salve e reautorize (passo 3)        |
| `Tem refresh_token?: False`                              | Marque **Refresh Token** no app e refaça o passo 3                                        |
| `PermissionError` ao gerar o relatório                   | Feche o `alertas_revisao_precos.xlsx` (ou o `_demo.xlsx`) no Excel                        |
| `Pandas requires version '3.1.5' or newer of 'openpyxl'` | `pip install --upgrade openpyxl`                                                          |
| `Coluna 'SKU' não encontrada`                            | Ajuste `COL_SKU` e `COL_CUSTO` no topo do script para os nomes da sua planilha            |
| `O CSV ... não tem a(s) coluna(s): ['frete']` (demo)     | O `anuncios_ml.csv` foi sobrescrito pela coleta real. Restaure o arquivo de exemplo       |
| Muitos itens em "Sem custo ou erro"                      | O SKU do Excel está diferente do anúncio (hífen, espaço, zero à esquerda)                 |
| `--demo` começou a varrer a API                          | Está rodando o `auditor_margem.py` antigo. Atualize o arquivo para a versão com `--demo`  |

---

## 9. Segurança

- **Nunca compartilhe** o `.env` nem o `tokens.json`. Quem tiver esses arquivos acessa a conta.
- Se usar Git, adicione ao `.gitignore`:

  ```
  .env
  tokens.json
  *.xlsx
  !custos.xlsx
  ```

  A linha `!custos.xlsx` mantém no repositório o `custos.xlsx` **de exemplo (fictício)**. Se você trocar esse arquivo pelos custos reais da empresa, remova essa exceção antes de fazer commit.

- O app tem somente permissão de **leitura**: o script não altera preços nem anúncios.

---

## 10. Limitações

- O preço considerado é o de tabela do anúncio. Promoções ativas não entram no cálculo.
- O frete é a cotação da API (`list_cost`). Confira de vez em quando com o custo real de uma venda recente.
- O preço sugerido é uma estimativa, pois o frete varia conforme o preço.

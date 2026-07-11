# Como rodar em outras maquinas

Este fluxo usa apenas Python padrao. Nao precisa instalar pandas, openpyxl ou outras bibliotecas.

## Arquivos necessarios

Copie estes arquivos para a mesma pasta:

- `ofx_to_csv.py`
- `classify_transactions.py`
- `build_monthly_html_report.py`
- `ofx_web_app.py`
- `classification_rules_completed.csv` ou `classification_rules.csv`

Tambem deixe o arquivo `.ofx` bancario acessivel na maquina.

## Modo recomendado: arrastar OFX no navegador

Rode o app local:

```powershell
python -B ofx_web_app.py
```

O navegador deve abrir automaticamente em:

```text
http://127.0.0.1:8765/
```

Na pagina, arraste o arquivo `.ofx` para a area de upload. O app gera automaticamente:

- CSV bruto;
- Excel classificado;
- HTML mensal agrupado;
- CSV de pendencias, caso existam operacoes novas para avaliar.

Os arquivos gerados ficam isolados por sessao do navegador. Os links de download so funcionam para a sessao que enviou o OFX.

Depois de processar, use o botao `Deletar informacoes e enviar outro arquivo` para limpar a analise atual e enviar outro OFX.

Por padrao, os arquivos temporarios expiram em 60 minutos. Para mudar:

```powershell
python -B ofx_web_app.py --ttl-minutes 15
```

Para disponibilizar em outros computadores da mesma rede:

```powershell
python -B ofx_web_app.py --host 0.0.0.0 --port 8765
```

Para rodar em outra maquina, copie estes arquivos para a mesma pasta e rode o comando acima:

- `ofx_web_app.py`
- `ofx_to_csv.py`
- `classify_transactions.py`
- `build_monthly_html_report.py`
- `classification_rules_completed.csv` ou `classification_rules.csv`

## Modo GitHub Pages: processar no navegador com Pyodide

A pasta `docs/` contem uma versao estatica que pode ser publicada no GitHub Pages. Ela carrega Pyodide pelo CDN e executa os scripts Python dentro do navegador.

Teste local:

```powershell
python -m http.server 8000 --directory docs
```

Abra:

```text
http://127.0.0.1:8000/
```

No GitHub, publique com Actions:

1. Envie as alteracoes para a branch `main`.
2. Abra `Settings > Pages`.
3. Em `Build and deployment`, selecione `GitHub Actions` como fonte.
4. O workflow `Deploy GitHub Pages` publicara a pasta `docs/` automaticamente a cada push.

Se o Pages estiver configurado para publicar a raiz da branch em vez do Actions, o `index.html` da raiz redireciona automaticamente para `docs/`.

Importante: todo arquivo dentro de `docs/` fica publico no GitHub Pages. Antes de publicar, mantenha as regras de classificacao sem nomes de pessoas, CPF/CNPJ, chaves PIX, dados bancarios, emails ou telefones.

## Modo manual: Etapa 1, converter OFX para CSV

```powershell
python ofx_to_csv.py "C:\caminho\do\extrato.ofx"
```

Isso gera um CSV bruto ao lado do OFX.

Tambem e possivel escolher o caminho de saida:

```powershell
python ofx_to_csv.py "C:\caminho\do\extrato.ofx" -o "C:\caminho\do\extrato.csv"
```

## Modo manual: Etapa 2, classificar CSV e gerar Excel

```powershell
python classify_transactions.py "C:\caminho\do\extrato.csv"
```

Isso gera um Excel com quatro abas:

- `extrato_bruto`
- `extrato_classificado`
- `resumo_classificacao`
- `operacoes_a_classificar`

Tambem e possivel escolher o caminho de saida:

```powershell
python classify_transactions.py "C:\caminho\do\extrato.csv" -o "C:\caminho\do\extrato_classificado.xlsx"
```

O script tambem gera um CSV resumido com as operacoes que ainda precisam ser revisadas:

```text
extrato_classificado_operacoes_a_classificar.csv
```

Esse arquivo e o melhor material para enviar de volta no chat quando houver operacoes sem classificacao.

A aba `extrato_classificado` preserva todas as linhas e colunas originais do extrato.
A aba `resumo_classificacao` consolida os valores por classificacao. Quando uma regra usa apenas `memo`, como PIX, a descricao/remetente nao entra no agrupamento.

## Regras de classificacao

As regras ficam no arquivo `classification_rules_completed.csv` ou `classification_rules.csv`.

Formato atual:

```csv
memo;descricao;Precisa descricao;grupo;categoria;subcategoria;tipo_fluxo;observacao
PIX RECEBIDO - OUTRA IF;;Nao;PIX;Pix entrada;;;
CR COMPRAS VISA;SIPAG_Cred._Visa;Sim;Cartao;Credito SIPAG;;;
```

Regras:

- Se `Precisa descricao` for `Sim`, `memo` e `descricao` precisam bater.
- Se `Precisa descricao` for `Nao` ou estiver vazio, uma regra com `memo` ignora a descricao.
- Se apenas `memo` estiver preenchido, a regra usa so o memo.
- Se apenas `descricao` estiver preenchida, a regra usa so a descricao.
- Se `grupo` estiver preenchido e `categoria` estiver vazia, a categoria recebe o mesmo valor do grupo.
- Regras incompletas de debito por titulo/boleto/cobranca/pagamento viram `Despesas > Fornecedores`.
- Regras incompletas de convenio/tributos/telecom/energia/saneamento viram `Despesas > Convenios`.
- Regras incompletas com cheque/talao viram `Cheque > Cheque`.
- Operacoes com `PIX` no memo tem prioridade: recebidas viram `PIX > Pix entrada` e emitidas/realizadas viram `PIX > Pix saida`.
- Se uma transacao nao bater em nenhuma regra, ela entra no HTML como `Outros gastos > Avaliar` ou `Outras entradas > Avaliar`.
- Se bater em uma regra sem grupo/categoria definidos, ela fica como `regra_sem_classificacao`.

## Modo manual: Etapa 3, gerar HTML mensal agrupado

Depois de gerar o Excel classificado, rode:

```powershell
python build_monthly_html_report.py "C:\caminho\do\extrato_classificado.xlsx" --ofx "C:\caminho\do\extrato.ofx"
```

Isso gera um HTML com:

- resumo geral;
- uma tabela para cada combinacao de `grupo` e `categoria`;
- valores mensais somados;
- quantidade de transacoes;
- entradas, saidas e total mensal.
- botao `Ver mais` em cada tabela para abrir as linhas completas daquele grupo/categoria.
- conferencia de saldo com saldo inicial inferido, movimentacoes, saldo final OFX e diferencas.
- movimentacoes novas sem regra lancadas como `Outros gastos > Avaliar` ou `Outras entradas > Avaliar`.

Tambem e possivel escolher o caminho de saida:

```powershell
python build_monthly_html_report.py "C:\caminho\do\extrato_classificado.xlsx" --ofx "C:\caminho\do\extrato.ofx" -o "C:\caminho\do\relatorio_mensal.html"
```

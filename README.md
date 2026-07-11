# Analisador OFX

Aplicacao local em Python para processar extratos bancarios `.ofx`, classificar movimentacoes e gerar relatorio HTML mensal por grupo e categoria.

## Como rodar

```powershell
python -B ofx_web_app.py
```

Abra no navegador:

```text
http://127.0.0.1:8765/
```

Arraste o arquivo `.ofx` na pagina. O app gera:

- CSV bruto;
- Excel classificado;
- HTML mensal agrupado;
- CSV de pendencias, caso haja operacoes novas para avaliar.

## Acesso em rede local

```powershell
python -B ofx_web_app.py --host 0.0.0.0 --port 8765
```

## Segurança dos arquivos

O processador pode ser publicado, mas arquivos bancarios nao devem ser versionados.

O app salva cada upload em uma sessao temporaria isolada e os arquivos expiram automaticamente. Por padrao:

```powershell
python -B ofx_web_app.py --ttl-minutes 60
```

## Versao GitHub Pages

A pasta `docs/` contem uma versao estatica para GitHub Pages. Ela usa Pyodide para rodar os mesmos scripts Python no navegador, sem enviar o OFX para um servidor.

Para testar localmente:

```powershell
python -m http.server 8000 --directory docs
```

Abra:

```text
http://127.0.0.1:8000/
```

Para publicar no GitHub Pages com GitHub Actions:

1. Envie as alteracoes para a branch `main`.
2. No GitHub, abra `Settings > Pages`.
3. Em `Build and deployment`, selecione `GitHub Actions` como fonte.
4. O workflow `Deploy GitHub Pages` publicara a pasta `docs/` automaticamente a cada push.

Se o Pages estiver configurado para publicar a raiz da branch em vez do Actions, o `index.html` da raiz redireciona automaticamente para `docs/`.

## Regras de classificacao

As regras ficam em:

```text
classification_rules_completed.csv
```

Formato esperado:

```csv
memo;descricao;Precisa descricao;grupo;categoria;subcategoria;tipo_fluxo;observacao
PIX RECEBIDO - OUTRA IF;;Nao;PIX;Pix entrada;;;
CR COMPRAS VISA;SIPAG_Cred._Visa;Sim;Cartao;Credito SIPAG;;;
```

## Arquivos principais

- `ofx_web_app.py`: app local com upload via navegador.
- `ofx_to_csv.py`: extrai transacoes do OFX para CSV.
- `classify_transactions.py`: classifica o extrato e gera Excel.
- `build_monthly_html_report.py`: gera HTML mensal agrupado.
- `classification_rules_completed.csv`: base de regras.

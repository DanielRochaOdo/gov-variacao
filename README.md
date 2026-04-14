# Conversor Consignacao (RETORNO e VARIACAO)

Aplicacao web em Flask para converter planilhas Excel (`.xlsx`) em TXT no layout de consignacao.

## Funcionalidades

- Conversao de arquivos `RETORNO`
- Conversao de arquivos `VARIACAO`
- Download imediato do arquivo `.txt` apos processamento
- Interface web leve, responsiva e intuitiva

## Estrutura principal

- `app.py`: servidor web e API de conversao
- `conversores.py`: regras de formatacao dos layouts
- `appRETORNO.py`: utilitario de linha de comando para RETORNO
- `appVARIACAO.py`: utilitario de linha de comando para VARIACAO
- `templates/` e `static/`: frontend

## Rodar localmente

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Acesse: `http://127.0.0.1:5000`

## Deploy na Vercel

1. Instale a CLI:

```bash
npm i -g vercel
```

2. No diretorio do projeto, execute:

```bash
vercel
```

3. Para novo deploy de preview:

```bash
vercel deploy -y
```

Os arquivos de configuracao para a Vercel ja estao prontos:

- `vercel.json`
- `.python-version`
- `.vercelignore`

# Conversor Consignacao (RETORNO, VARIACAO e Bradesco)

Aplicacao web em Flask para converter planilhas Excel (`.xlsx`) em TXT no layout de consignacao.

## Funcionalidades

- Conversao de arquivos `RETORNO`
- Conversao de arquivos `VARIACAO`
- Conversao de arquivos `Pagamento Bradesco > Transferencia` (logica das paginas 16 a 19)
- Conversao de arquivos `Pagamento Bradesco > Boleto` (logica das paginas 23 a 27)
- Entrada somente via planilha Excel (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) para todos os tipos
- Em `Pagamento Bradesco > Transferencia`, quando houver pagamento com valor `0,00`, o download nao e bloqueado e o frontend exibe alerta com os dados bancarios do favorecido
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

Acesse no mesmo computador: `http://127.0.0.1:5000`  
Na rede local: `http://SEU_IP_LOCAL:5000`

Atalhos via npm (opcional):

```bash
npm run dev
# ou
npm run api
```

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

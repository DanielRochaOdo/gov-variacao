from __future__ import annotations

import json
import os
from pathlib import Path
from datetime import datetime
from urllib.parse import quote

from flask import Flask, Response, jsonify, render_template, request, send_file

from conversores import (
    ConversionError,
    TIPO_BRADESCO_BOLETO,
    TIPO_BRADESCO_TRANSFERENCIA,
    TIPO_RETORNO,
    TIPO_VARIACAO,
    gerar_txt_por_tipo,
    listar_alertas_valor_zero_bradesco_transferencia,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 8 MB
_REM_FILE_COUNTERS: dict[str, int] = {}


def _nome_arquivo_rem_bradesco_pix() -> str:
    hoje = datetime.now()
    base = hoje.strftime("%d%m%Y")
    proximo = _REM_FILE_COUNTERS.get(base, 0) + 1
    _REM_FILE_COUNTERS[base] = proximo
    sufixo = "" if proximo == 1 else str(proximo - 1)
    return f"{base}{sufixo}.rem"


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.post("/api/converter")
def converter() -> Response:
    tipo = (request.form.get("tipo") or "").strip().upper()
    arquivo = request.files.get("arquivo")

    tipos_validos = {
        TIPO_RETORNO,
        TIPO_VARIACAO,
        TIPO_BRADESCO_TRANSFERENCIA,
        TIPO_BRADESCO_BOLETO,
    }
    if tipo not in tipos_validos:
        return jsonify(
            {
                "erro": (
                    "Tipo invalido. Selecione RETORNO, VARIACAO, "
                    "Pagamento Bradesco > PIX ou Pagamento Bradesco > Boleto."
                )
            }
        ), 400

    if arquivo is None or arquivo.filename is None or arquivo.filename.strip() == "":
        return jsonify({"erro": "Envie um arquivo antes de converter."}), 400

    nome_arquivo = arquivo.filename.strip()
    extensao = Path(nome_arquivo).suffix.lower()
    formatos_excel = (".xlsx", ".xlsm", ".xltx", ".xltm")
    if extensao not in formatos_excel:
        return jsonify({"erro": "Formato invalido. Envie um arquivo .xlsx."}), 400

    try:
        conteudo_txt = gerar_txt_por_tipo(arquivo.stream, tipo)
    except ConversionError as exc:
        return jsonify({"erro": str(exc)}), 400
    except Exception:
        return jsonify({"erro": "Falha inesperada ao converter o arquivo."}), 500

    if tipo == TIPO_BRADESCO_TRANSFERENCIA:
        nome_saida = _nome_arquivo_rem_bradesco_pix()
    else:
        nome_base = Path(nome_arquivo).stem
        prefixos = {
            TIPO_RETORNO: "retorno_",
            TIPO_VARIACAO: "variacao_",
            TIPO_BRADESCO_BOLETO: "bradesco_boleto_",
        }
        prefixo = prefixos.get(tipo, "arquivo_")
        nome_saida = f"{prefixo}{nome_base}.txt"

    try:
        conteudo_bytes = conteudo_txt.encode("latin-1")
        charset = "iso-8859-1"
    except UnicodeEncodeError:
        # Fallback defensivo para caracteres fora do latin-1.
        conteudo_bytes = conteudo_txt.encode("latin-1", errors="replace")
        charset = "iso-8859-1"

    headers = {"Content-Disposition": f'attachment; filename="{nome_saida}"'}
    if tipo == TIPO_BRADESCO_TRANSFERENCIA:
        alertas = listar_alertas_valor_zero_bradesco_transferencia(conteudo_txt)
        total_alertas = len(alertas)
        if total_alertas > 0:
            max_alertas_header = 50
            alertas_expostos = alertas[:max_alertas_header]
            payload = json.dumps(alertas_expostos, ensure_ascii=False, separators=(",", ":"))
            headers["X-Valor-Zero-Total"] = str(total_alertas)
            headers["X-Valor-Zero-Truncated"] = "1" if total_alertas > max_alertas_header else "0"
            headers["X-Valor-Zero-Details"] = quote(payload, safe="")

    return Response(
        conteudo_bytes,
        content_type=f"text/plain; charset={charset}",
        headers=headers,
    )


@app.get("/api/modelo/bradesco-transferencia")
def baixar_modelo_bradesco_transferencia() -> Response:
    caminho_modelo = Path(__file__).resolve().parent / "modelo_bradesco_transferencia.xlsx"
    if not caminho_modelo.exists():
        return jsonify({"erro": "Modelo base de Bradesco PIX nao encontrado no servidor."}), 404

    return send_file(
        caminho_modelo,
        as_attachment=True,
        download_name="modelo_bradesco_transferencia.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def payload_too_large(_: Exception):
    return jsonify({"erro": "Arquivo muito grande. Limite: 8 MB."}), 413


if __name__ == "__main__":
    app.run(
        host=os.getenv("HOST", "0.0.0.0"),
        port=int(os.getenv("PORT", "5000")),
        debug=os.getenv("DEBUG", "true").lower() in {"1", "true", "yes", "on"},
    )

from __future__ import annotations

from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request

from conversores import ConversionError, gerar_txt_por_tipo

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 8 MB


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.post("/api/converter")
def converter() -> Response:
    tipo = (request.form.get("tipo") or "").strip().upper()
    arquivo = request.files.get("arquivo")

    if tipo not in {"RETORNO", "VARIACAO"}:
        return jsonify({"erro": "Tipo invalido. Selecione RETORNO ou VARIACAO."}), 400

    if arquivo is None or arquivo.filename is None or arquivo.filename.strip() == "":
        return jsonify({"erro": "Envie um arquivo Excel antes de converter."}), 400

    nome_arquivo = arquivo.filename.strip()
    if not nome_arquivo.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return jsonify({"erro": "Formato invalido. Envie um arquivo .xlsx."}), 400

    try:
        conteudo_txt = gerar_txt_por_tipo(arquivo.stream, tipo)
    except ConversionError as exc:
        return jsonify({"erro": str(exc)}), 400
    except Exception:
        return jsonify({"erro": "Falha inesperada ao converter o arquivo."}), 500

    nome_base = Path(nome_arquivo).stem
    prefixo = "retorno_" if tipo == "RETORNO" else "variacao_"
    nome_saida = f"{prefixo}{nome_base}.txt"

    return Response(
        conteudo_txt,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{nome_saida}"'},
    )


@app.errorhandler(413)
def payload_too_large(_: Exception):
    return jsonify({"erro": "Arquivo muito grande. Limite: 8 MB."}), 413


if __name__ == "__main__":
    app.run(debug=True)

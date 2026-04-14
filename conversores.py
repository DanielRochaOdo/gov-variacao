from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BufferedIOBase, BytesIO
import math
import re
from typing import Any, Dict, List, Mapping

from openpyxl import load_workbook


class ConversionError(Exception):
    """Erro de validacao ou formatacao durante a conversao."""


def gerar_txt_por_tipo(excel_file: BufferedIOBase | BytesIO, tipo: str) -> str:
    tipo_normalizado = (tipo or "").strip().upper()
    if tipo_normalizado == "RETORNO":
        return gerar_layout_retorno(excel_file)
    if tipo_normalizado == "VARIACAO":
        return gerar_layout_variacao(excel_file)
    raise ConversionError("Tipo invalido. Use RETORNO ou VARIACAO.")


def gerar_layout_retorno(excel_file: BufferedIOBase | BytesIO) -> str:
    linhas = [
        _formatar_linha_retorno(row)
        for row in _ler_planilha_excel(
            excel_file,
            required_headers=[
                "ano_mes",
                "orgao",
                "matricula",
                "consignataria",
                "valor_parcela",
                "cpf",
                "contrato",
                "nome_servidor",
            ],
            nome_layout="RETORNO",
        )
    ]
    return "\n".join(linhas) + ("\n" if linhas else "")


def gerar_layout_variacao(excel_file: BufferedIOBase | BytesIO) -> str:
    linhas = [
        _formatar_linha_variacao(row)
        for row in _ler_planilha_excel(
            excel_file,
            required_headers=[
                "consignataria",
                "nome_consignataria",
                "instituicao",
                "contrato",
                "nome_servidor",
                "cpf",
                "orgao",
                "matricula",
                "tipo_ajuste",
                "categoria_ajuste",
                "data_inicial",
                "valor_total",
                "qtd_parcelas",
                "valor_parcela",
                "proxima_parcela",
            ],
            nome_layout="VARIACAO",
        )
    ]
    return "\n".join(linhas) + ("\n" if linhas else "")


def _ler_planilha_excel(
    excel_file: BufferedIOBase | BytesIO,
    required_headers: List[str] | None = None,
    nome_layout: str = "",
) -> List[Dict[str, Any]]:
    try:
        if hasattr(excel_file, "seek"):
            excel_file.seek(0)
        workbook = load_workbook(excel_file, data_only=True, read_only=True)
    except Exception as exc:  # pragma: no cover - depende do parser do openpyxl
        raise ConversionError("Nao foi possivel ler o arquivo Excel enviado.") from exc

    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)

    headers_raw = next(rows, None)
    if not headers_raw:
        raise ConversionError("A planilha esta vazia.")

    headers = [_normalizar_header(valor) for valor in headers_raw]
    if not any(headers):
        raise ConversionError("Cabecalho invalido na planilha.")

    if required_headers:
        missing = [col for col in required_headers if col not in headers]
        if missing:
            layout_detectado = _detectar_layout(headers)
            dica = ""
            if layout_detectado and layout_detectado != nome_layout:
                dica = (
                    f" A planilha enviada parece ser do layout {layout_detectado}. "
                    f"Selecione {layout_detectado} para gerar o TXT correto."
                )
            raise ConversionError(
                f"Planilha nao corresponde ao layout {nome_layout}. "
                f"Colunas ausentes: {', '.join(missing)}.{dica}"
            )

    itens: List[Dict[str, Any]] = []
    for values in rows:
        if values is None or all(valor is None or str(valor).strip() == "" for valor in values):
            continue
        registro = {headers[idx]: values[idx] if idx < len(values) else None for idx in range(len(headers))}
        itens.append(registro)

    return itens


def _normalizar_header(valor: Any) -> str:
    if valor is None:
        return ""
    return str(valor).strip().lower()


def _detectar_layout(headers: List[str]) -> str:
    headers_set = set(headers)
    retorno_base = {"ano_mes", "orgao", "matricula", "consignataria", "valor_parcela", "cpf", "contrato", "nome_servidor"}
    variacao_base = {
        "consignataria",
        "nome_consignataria",
        "instituicao",
        "contrato",
        "nome_servidor",
        "cpf",
        "orgao",
        "matricula",
        "tipo_ajuste",
        "categoria_ajuste",
        "data_inicial",
        "valor_total",
        "qtd_parcelas",
        "valor_parcela",
        "proxima_parcela",
    }
    if variacao_base.issubset(headers_set):
        return "VARIACAO"
    if retorno_base.issubset(headers_set):
        return "RETORNO"
    return ""


def _formatar_linha_retorno(row: Mapping[str, Any]) -> str:
    linha = (
        f"{_texto(row.get('ano_mes'))[:6]:<6}"
        f"{_somente_digitos(row.get('orgao')).zfill(3)}"
        f"{_texto(row.get('matricula'))[:8]:<8}"
        f"{_somente_digitos(row.get('consignataria')).zfill(6)}"
        f"{_valor_em_centavos(row.get('valor_parcela'), 11)}"
        "00"
        f"{_somente_digitos(row.get('cpf')).zfill(11)}"
        f"{_texto(row.get('contrato')).zfill(20)}"
        f"{_texto(row.get('nome_servidor'))[:27]:<27}"
    )
    if len(linha) != 94:
        raise ConversionError("Linha RETORNO com tamanho invalido. Esperado: 94.")
    return linha


def _formatar_linha_variacao(row: Mapping[str, Any]) -> str:
    tipo_ajuste = _texto(row.get("tipo_ajuste")).upper()[:1]
    categoria_ajuste = _texto(row.get("categoria_ajuste"))[:1]
    valor_parcela = (
        _valor_em_centavos(row.get("valor_parcela"), 7)
        if tipo_ajuste != "E"
        else "0000000"
    )

    linha = (
        f"{_numero_inteiro(row.get('consignataria'), 6)}"
        f"{_texto(row.get('nome_consignataria'))[:20]:<20}"
        f"{_texto(row.get('instituicao'))[:20]:<20}"
        f"{_numero_inteiro(row.get('contrato'), 15)}"
        f"{_texto(row.get('nome_servidor'))[:30]:<30}"
        f"{_somente_digitos(row.get('cpf')).zfill(11)}"
        f"{_numero_inteiro(row.get('orgao'), 3)}"
        f"{_numero_inteiro(row.get('matricula'), 8)}"
        f"{tipo_ajuste}"
        f"{categoria_ajuste}"
        f"{_texto(row.get('data_inicial')).zfill(8)}"
        f"{_valor_em_centavos(row.get('valor_total'), 10)}"
        f"{_numero_inteiro(row.get('qtd_parcelas'), 3)}"
        f"{valor_parcela}"
        f"{_numero_inteiro(row.get('proxima_parcela'), 3)}"
    )
    if len(linha) != 146:
        raise ConversionError("Linha VARIACAO com tamanho invalido. Esperado: 146.")
    return linha


def _somente_digitos(valor: Any) -> str:
    return re.sub(r"\D", "", _texto(valor))


def _texto(valor: Any) -> str:
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d%m%Y")
    if isinstance(valor, date):
        return valor.strftime("%d%m%Y")
    if isinstance(valor, float):
        if math.isnan(valor):
            return ""
        if valor.is_integer():
            return str(int(valor))
        return format(valor, "f").rstrip("0").rstrip(".")
    texto = str(valor).strip()
    if texto.endswith(".0") and texto.replace(".", "", 1).isdigit():
        return texto[:-2]
    return texto


def _numero_inteiro(valor: Any, largura: int) -> str:
    texto = _texto(valor)
    if texto == "":
        numero = 0
    else:
        texto = _normalizar_decimal(texto)
        try:
            numero = int(Decimal(texto))
        except (InvalidOperation, ValueError) as exc:
            raise ConversionError(f"Valor inteiro invalido: {valor!r}") from exc
    return f"{numero:0{largura}d}"


def _valor_em_centavos(valor: Any, largura: int) -> str:
    texto = _texto(valor)
    if texto == "":
        centavos = 0
    else:
        texto = _normalizar_decimal(texto)
        try:
            centavos = int(Decimal(texto) * 100)
        except (InvalidOperation, ValueError) as exc:
            raise ConversionError(f"Valor monetario invalido: {valor!r}") from exc
    return f"{centavos:0{largura}d}"


def _normalizar_decimal(texto: str) -> str:
    limpo = re.sub(r"[^\d,.\-]", "", texto)
    if "," in limpo and "." in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    elif "," in limpo:
        limpo = limpo.replace(",", ".")
    return limpo

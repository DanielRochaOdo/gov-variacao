from __future__ import annotations

import argparse
from pathlib import Path

from conversores import gerar_layout_retorno


def gerar_layout_retorno_arquivo(path_excel: str, path_saida_txt: str) -> None:
    with open(path_excel, "rb") as excel_file:
        conteudo = gerar_layout_retorno(excel_file)

    with open(path_saida_txt, "w", encoding="utf-8") as txt_file:
        txt_file.write(conteudo)


def main() -> None:
    parser = argparse.ArgumentParser(description="Gera TXT de RETORNO a partir de planilha Excel.")
    parser.add_argument(
        "entrada",
        nargs="?",
        default="exemplo_consignacao_RETORNO.xlsx",
        help="Arquivo Excel de entrada.",
    )
    parser.add_argument(
        "saida",
        nargs="?",
        default="retorno_consignacao.txt",
        help="Arquivo TXT de saida.",
    )
    args = parser.parse_args()

    gerar_layout_retorno_arquivo(args.entrada, args.saida)
    print(f"Arquivo gerado: {Path(args.saida).resolve()}")


if __name__ == "__main__":
    main()

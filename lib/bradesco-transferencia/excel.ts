import ExcelJS from "exceljs";

import { ALIAS_TO_CANONICAL, MODEL_HEADERS, normalizeHeader, REQUIRED_BASE_HEADERS } from "./aliases";
import { BradescoValidationError, CanonicalRecord, ParsedSheet } from "./types";

function readCellAsString(cellValue: unknown): string {
  if (cellValue === null || cellValue === undefined) {
    return "";
  }

  if (cellValue instanceof Date) {
    const day = String(cellValue.getDate()).padStart(2, "0");
    const month = String(cellValue.getMonth() + 1).padStart(2, "0");
    const year = String(cellValue.getFullYear());
    return `${day}/${month}/${year}`;
  }

  if (typeof cellValue === "object" && cellValue !== null) {
    const maybeRichText = cellValue as { richText?: Array<{ text?: string }>; text?: string; result?: unknown };
    if (Array.isArray(maybeRichText.richText)) {
      return maybeRichText.richText.map((item) => item.text ?? "").join("").trim();
    }
    if (typeof maybeRichText.text === "string") {
      return maybeRichText.text.trim();
    }
    if (maybeRichText.result !== undefined) {
      return String(maybeRichText.result).trim();
    }
  }

  return String(cellValue).trim();
}

export async function parseSheetFromBuffer(buffer: ArrayBuffer): Promise<ParsedSheet> {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.load(buffer);
  } catch {
    throw new BradescoValidationError("Nao foi possivel ler a planilha enviada. Use um arquivo .xlsx valido.");
  }

  const worksheet = workbook.getWorksheet("Modelo") ?? workbook.worksheets[0];
  if (!worksheet) {
    throw new BradescoValidationError("A planilha enviada nao possui abas para leitura.");
  }

  const headerRow = worksheet.getRow(1);
  if (!headerRow || headerRow.cellCount === 0) {
    throw new BradescoValidationError("Cabecalho nao encontrado na planilha.");
  }

  const columnMap = new Map<number, string>();
  const unknownHeaders: string[] = [];

  for (let col = 1; col <= headerRow.cellCount; col += 1) {
    const rawHeader = readCellAsString(headerRow.getCell(col).value);
    const normalized = normalizeHeader(rawHeader);
    if (!normalized) {
      continue;
    }

    const canonical = ALIAS_TO_CANONICAL.get(normalized);
    if (canonical) {
      columnMap.set(col, canonical);
    } else {
      unknownHeaders.push(rawHeader);
    }
  }

  const missingBase = REQUIRED_BASE_HEADERS.filter((header) => !Array.from(columnMap.values()).includes(header));
  if (missingBase.length > 0) {
    throw new BradescoValidationError(
      "Planilha fora do padrao para Bradesco Transferencia.",
      [
        `Colunas obrigatorias ausentes: ${missingBase.join(", ")}.`,
        "Baixe o Modelo Base para preencher no formato esperado.",
      ]
    );
  }

  const rows: CanonicalRecord[] = [];
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    if (!row) {
      continue;
    }

    const record: CanonicalRecord = {};
    let hasData = false;

    for (const [col, canonical] of columnMap.entries()) {
      const value = readCellAsString(row.getCell(col).value);
      if (value !== "") {
        hasData = true;
      }
      record[canonical] = value;
    }

    if (!hasData) {
      continue;
    }

    record.__row = String(rowNumber);
    rows.push(record);
  }

  if (rows.length === 0) {
    throw new BradescoValidationError("A planilha nao possui linhas de pagamento para processar.");
  }

  const warnings: string[] = [];
  if (unknownHeaders.length > 0) {
    warnings.push(`Colunas ignoradas: ${unknownHeaders.join(", ")}`);
  }

  return { rows, warnings };
}

export async function generateModelWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Conversor Bradesco";
  workbook.created = new Date();

  const modelSheet = workbook.addWorksheet("Modelo");
  const instructionsSheet = workbook.addWorksheet("Instrucoes");

  modelSheet.addRow(MODEL_HEADERS);
  modelSheet.addRow([
    "EMPRESA_001",
    "2",
    "03187913000157",
    "628608",
    "0564",
    "9",
    "000000009350",
    "5",
    "ODONTOART PLANOS ODONTOLOGICOS LTDA",
    "20",
    "45",
    "237",
    "00448",
    "000000011439",
    "1",
    "CLINICA TESTE PIX CPF",
    "03/05/2026",
    150.25,
    "1",
    "00002122624329",
    "",
    "00000000",
    "PGTO-PIX-0001",
    "",
    "",
    "CE",
    "60000000",
  ]);
  modelSheet.addRow([
    "EMPRESA_001",
    "2",
    "03187913000157",
    "628608",
    "0564",
    "9",
    "000000009350",
    "5",
    "ODONTOART PLANOS ODONTOLOGICOS LTDA",
    "20",
    "45",
    "237",
    "02793",
    "000000075961",
    "9",
    "CLINICA TESTE PIX DADOS BANCARIOS",
    "03/05/2026",
    188.75,
    "1",
    "00002122624329",
    "",
    "00000000",
    "PGTO-PIX-0002",
    "RUA EXEMPLO, 100",
    "FORTALEZA",
    "CE",
    "60000000",
  ]);

  const headerRow = modelSheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FF1D3557" } };
  const mandatoryHeaders = new Set([
    "tipo_inscricao_empresa",
    "numero_inscricao_empresa",
    "convenio",
    "agencia_empresa",
    "conta_empresa",
    "nome_empresa",
    "nome_favorecido",
    "data_pagamento",
    "valor_pagamento",
    "tipo_inscricao_favorecido",
    "numero_inscricao_favorecido",
    "numero_pagamento",
    "banco_favorecido",
    "agencia_favorecido",
    "conta_favorecido",
    "dv_conta_favorecido",
    "tipo_conta_recebedor",
  ]);
  for (let col = 1; col <= MODEL_HEADERS.length; col += 1) {
    const header = MODEL_HEADERS[col - 1];
    const cell = headerRow.getCell(col);
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: mandatoryHeaders.has(header) ? "00FFFF00" : "FFE9F2FF" },
    };
  }

  modelSheet.columns = MODEL_HEADERS.map((header) => {
    const wider = ["nome_favorecido", "nome_empresa", "informacao12", "informacao10", "informacao11"].includes(header);
    return { key: header, width: wider ? 30 : 18 };
  });

  modelSheet.getColumn("data_pagamento").numFmt = "dd/mm/yyyy";
  modelSheet.getColumn("valor_pagamento").numFmt = "R$ #,##0.00";
  modelSheet.getColumn("numero_inscricao_favorecido").numFmt = "@";
  modelSheet.getColumn("numero_inscricao_empresa").numFmt = "@";
  modelSheet.views = [{ state: "frozen", ySplit: 1 }];

  instructionsSheet.columns = [
    { header: "Campo", key: "campo", width: 34 },
    { header: "Regra simples", key: "descricao", width: 90 },
  ];
  instructionsSheet.getRow(1).font = { bold: true };

  const instructions: Array<[string, string]> = [
    ["Visao geral", "Este modelo gera CNAB240 PIX para Bradesco. Campos amarelos pedem preenchimento obrigatorio."],
    ["lote", "Opcional. Use para separar lotes/trailers diferentes (ex.: EMPRESA_A, EMPRESA_B ou 1, 2, 3)."],
    ["tipo_inscricao_empresa", "Obrigatorio. 1=CPF da empresa, 2=CNPJ da empresa."],
    ["numero_inscricao_empresa", "Obrigatorio. CPF/CNPJ da empresa pagadora. Preencha como texto para manter zeros."],
    ["convenio", "Obrigatorio. Codigo do convenio no banco."],
    ["agencia_empresa", "Obrigatorio. Agencia da conta pagadora (sem DV)."],
    ["dv_agencia_empresa", "Digito da agencia da empresa (se houver)."],
    ["conta_empresa", "Obrigatorio. Conta da empresa pagadora (sem DV)."],
    ["dv_conta_empresa", "Digito da conta da empresa (se houver)."],
    ["nome_empresa", "Obrigatorio. Nome da empresa pagadora."],
    ["tipo_servico", "Pode deixar 20."],
    ["forma_lancamento", "Pode deixar 45 (PIX Transferencia)."],
    ["banco_favorecido", "Na forma 05, informar banco do recebedor. Ex.: 237."],
    ["agencia_favorecido", "Na forma 05, obrigatorio: 4 ou 5 digitos (coluna AM). Se vier com 4, o TXT completa com zero a esquerda."],
    ["dv_agencia_favorecido", "Fixo no processamento: 0 (nao preencher coluna no modelo)."],
    ["conta_favorecido", "Na forma 05, obrigatorio: conta do recebedor sem DV."],
    ["dv_conta_favorecido", "Na forma 05, obrigatorio: DV da conta do recebedor."],
    ["nome_favorecido", "Obrigatorio. Nome do recebedor."],
    ["data_pagamento", "Obrigatorio. Data do pagamento (dd/mm/aaaa)."],
    ["valor_pagamento", "Obrigatorio. Valor do pagamento. Ex.: 150,25."],
    ["forma_iniciacao", "Fixo no processamento: 05 (dados bancarios). Nao preencher coluna no modelo."],
    ["tipo_inscricao_favorecido", "Obrigatorio. 1=CPF, 2=CNPJ do favorecido."],
    ["numero_inscricao_favorecido", "Obrigatorio. CPF/CNPJ do favorecido, com zeros a esquerda quando necessario."],
    ["informacao12", "Chave PIX. Obrigatoria nas formas 01, 02 e 04. Na forma 03 pode ficar em branco."],
    ["codigo_ispb", "Opcional."],
    ["numero_pagamento", "Obrigatorio. Controle interno. Vai para Segmento A, posicoes 74 a 93."],
    ["informacao10", "Opcional. Complemento livre (ate 35 caracteres)."],
    ["informacao11", "Opcional. Complemento livre (ate 60 caracteres)."],
    ["estado", "Opcional. UF com 2 letras (ex.: CE)."],
    ["cep", "Opcional. CEP numerico."],
    ["Regras forma 05", "Conversao fixa para forma 05. Agencia favorecido (AM) segue regra de 1 a 5 digitos; DV da agencia favorecido e fixado em 0."],
  ];

  instructions.forEach(([campo, descricao]) => instructionsSheet.addRow({ campo, descricao }));

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

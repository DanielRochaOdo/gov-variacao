import {
  BradescoSummary,
  BradescoValidationError,
  CanonicalRecord,
  CompanyBase,
  LoteGroup,
  ProcessResult,
} from "./types";

const FORMA_INICIACAO_VALIDAS = new Set(["01", "02", "03", "04", "05"]);

function text(value: string | undefined): string {
  return (value ?? "").trim();
}

function onlyDigits(value: string | undefined): string {
  return text(value).replace(/\D/g, "");
}

function cnabAlfa(value: string | undefined, width: number, defaultValue = ""): string {
  const raw = text(value) || defaultValue;
  const sanitized = raw.replace(/[\r\n]/g, " ").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  return sanitized.slice(0, width).padEnd(width, " ");
}

function cnabNum(value: string | number | undefined, width: number, defaultValue = "0"): string {
  const raw = typeof value === "number" ? String(Math.trunc(value)) : onlyDigits(value);
  const fallback = onlyDigits(defaultValue) || "0";
  const digits = raw || fallback;
  return digits.slice(-width).padStart(width, "0");
}

function normalizeDecimal(input: string): string {
  const cleaned = input.replace(/[^\d,.-]/g, "");
  if (cleaned.includes(",") && cleaned.includes(".")) {
    return cleaned.replace(/\./g, "").replace(",", ".");
  }
  if (cleaned.includes(",")) {
    return cleaned.replace(",", ".");
  }
  return cleaned;
}

function toCentavos(value: string | undefined, rowLabel: string): number {
  const raw = text(value);
  if (!raw) {
    throw new BradescoValidationError(`Linha ${rowLabel}: valor_pagamento vazio.`);
  }

  const parsed = Number(normalizeDecimal(raw));
  if (!Number.isFinite(parsed)) {
    throw new BradescoValidationError(`Linha ${rowLabel}: valor_pagamento invalido (${raw}).`);
  }
  return Math.round(parsed * 100);
}

function toQuantidadeMoeda(value: string | undefined): number {
  const raw = text(value);
  if (!raw) {
    return 0;
  }
  const parsed = Number(normalizeDecimal(raw));
  if (!Number.isFinite(parsed)) {
    return 0;
  }
  return Math.round(parsed * 100000);
}

function parseDateToCnab(value: string | undefined, rowLabel: string): string {
  const raw = text(value);
  if (!raw) {
    throw new BradescoValidationError(`Linha ${rowLabel}: data_pagamento vazia.`);
  }

  const numericRaw = Number(raw);
  if (Number.isFinite(numericRaw) && numericRaw >= 20000 && numericRaw <= 90000) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const date = new Date(excelEpoch + Math.trunc(numericRaw) * 86400000);
    const dd = String(date.getUTCDate()).padStart(2, "0");
    const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
    const yyyy = String(date.getUTCFullYear());
    return `${dd}${mm}${yyyy}`;
  }

  const digits = onlyDigits(raw);
  if (digits.length === 8) {
    if (raw.includes("/")) {
      const [dd, mm, yyyy] = raw.split("/");
      if (dd && mm && yyyy && yyyy.length === 4) {
        return `${dd.padStart(2, "0")}${mm.padStart(2, "0")}${yyyy}`;
      }
    }

    const year = Number(digits.slice(0, 4));
    if (year >= 1900) {
      return `${digits.slice(6, 8)}${digits.slice(4, 6)}${digits.slice(0, 4)}`;
    }

    const month = Number(digits.slice(2, 4));
    if (month >= 1 && month <= 12) {
      return digits;
    }
  }

  throw new BradescoValidationError(`Linha ${rowLabel}: data_pagamento invalida (${raw}). Use dd/mm/aaaa.`);
}

function parseNumeroDv(numeroRaw: string | undefined, dvRaw: string | undefined, width: number): { numero: string; dv: string } {
  const numeroTxt = text(numeroRaw);
  let dv = text(dvRaw).toUpperCase().slice(0, 1);

  let numericBase = onlyDigits(numeroTxt);
  const match = numeroTxt.match(/^\s*([0-9]+)\s*[-\s]+\s*([0-9A-Za-z])\s*$/);
  if (match) {
    numericBase = onlyDigits(match[1]);
    if (!dv) {
      dv = match[2].toUpperCase();
    }
  } else if (!dv && numericBase.length === width + 1) {
    dv = numericBase.slice(-1).toUpperCase();
    numericBase = numericBase.slice(0, -1);
  }

  return {
    numero: cnabNum(numericBase, width),
    dv: cnabAlfa(dv, 1),
  };
}

function join240(fields: string[], label: string): string {
  const line = fields.join("");
  if (line.length !== 240) {
    throw new BradescoValidationError(`${label} com tamanho invalido (${line.length}). Esperado 240.`);
  }
  return line;
}

function resolveBase(row: CanonicalRecord, now: Date): CompanyBase {
  const agenciaParsed = parseNumeroDv(row.agencia_empresa, row.dv_agencia_empresa, 5);
  const contaParsed = parseNumeroDv(row.conta_empresa, row.dv_conta_empresa, 12);

  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yyyy = String(now.getFullYear());
  const hh = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");

  const dataGeracao = row.data_geracao ? parseDateToCnab(row.data_geracao, row.__row ?? "?") : `${dd}${mm}${yyyy}`;

  return {
    banco: cnabNum(row.banco, 3, "237"),
    tipoInscricaoEmpresa: cnabNum(row.tipo_inscricao_empresa, 1, "2"),
    numeroInscricaoEmpresa: cnabNum(row.numero_inscricao_empresa, 14),
    convenio: cnabAlfa(row.convenio, 20),
    agenciaEmpresa: agenciaParsed.numero,
    dvAgenciaEmpresa: agenciaParsed.dv,
    contaEmpresa: contaParsed.numero,
    dvContaEmpresa: contaParsed.dv,
    dvAgenciaContaEmpresa: cnabAlfa(row.dv_agencia_conta_empresa, 1),
    nomeEmpresa: cnabAlfa(row.nome_empresa, 30),
    nomeBanco: cnabAlfa(row.nome_banco, 30, "BANCO BRADESCO S.A"),
    codigoRemessaRetorno: cnabNum(row.codigo_remessa_retorno, 1, "1"),
    dataGeracao,
    horaGeracao: cnabNum(row.hora_geracao, 6, `${hh}${mi}${ss}`),
    nsa: cnabNum(row.nsa, 6, "1"),
    densidade: cnabNum(row.densidade, 5, "0"),
    headerPix: "PIX",
  };
}

function buildHeaderArquivo(base: CompanyBase): string {
  return join240(
    [
      base.banco,
      "0000",
      "0",
      " ".repeat(9),
      base.tipoInscricaoEmpresa,
      base.numeroInscricaoEmpresa,
      base.convenio,
      base.agenciaEmpresa,
      base.dvAgenciaEmpresa,
      base.contaEmpresa,
      base.dvContaEmpresa,
      base.dvAgenciaContaEmpresa,
      base.nomeEmpresa,
      base.nomeBanco,
      " ".repeat(10),
      base.codigoRemessaRetorno,
      base.dataGeracao,
      base.horaGeracao,
      base.nsa,
      "089",
      base.densidade,
      base.headerPix,
      " ".repeat(17),
      " ".repeat(20),
      " ".repeat(29),
    ],
    "Header de arquivo"
  );
}

function buildHeaderLote(base: CompanyBase, lote: string, row: CanonicalRecord): string {
  return join240(
    [
      base.banco,
      lote,
      "1",
      "C",
      cnabNum(row.tipo_servico, 2, "20"),
      cnabNum(row.forma_lancamento, 2, "45"),
      "045",
      " ",
      base.tipoInscricaoEmpresa,
      base.numeroInscricaoEmpresa,
      base.convenio,
      base.agenciaEmpresa,
      base.dvAgenciaEmpresa,
      base.contaEmpresa,
      base.dvContaEmpresa,
      base.dvAgenciaContaEmpresa,
      base.nomeEmpresa,
      cnabAlfa(row.mensagem_lote, 40),
      cnabAlfa(row.logradouro, 30),
      cnabNum(row.numero_local, 5),
      cnabAlfa(row.complemento_endereco, 15),
      cnabAlfa(row.cidade, 20),
      cnabNum(row.cep, 5),
      cnabAlfa(row.complemento_cep, 3),
      cnabAlfa(row.estado, 2),
      cnabNum(row.indicativo_forma_pagamento, 2, "1"),
      " ".repeat(6),
      cnabAlfa(row.ocorrencias_header_lote, 10),
    ],
    "Header de lote"
  );
}

function buildSegmentoA(base: CompanyBase, lote: string, sequencial: number, row: CanonicalRecord): string {
  const contaFav = parseNumeroDv(row.conta_favorecido, row.dv_conta_favorecido, 12);
  const agenciaFav = parseNumeroDv(row.agencia_favorecido, row.dv_agencia_favorecido, 5);
  const valor = toCentavos(row.valor_pagamento, row.__row ?? "?");

  return join240(
    [
      base.banco,
      lote,
      "3",
      cnabNum(sequencial, 5),
      "A",
      cnabNum(row.tipo_movimento, 1, "0"),
      cnabNum(row.codigo_instrucao_movimento, 2, "00"),
      cnabNum(row.camara_centralizadora, 3, "009"),
      cnabNum(row.banco_favorecido, 3, row.forma_iniciacao === "05" ? "237" : "000"),
      agenciaFav.numero,
      agenciaFav.dv,
      contaFav.numero,
      contaFav.dv,
      cnabAlfa(row.dv_agencia_conta_favorecido, 1),
      cnabAlfa(row.nome_favorecido, 30),
      cnabAlfa(row.seu_numero, 20),
      parseDateToCnab(row.data_pagamento, row.__row ?? "?"),
      cnabAlfa(row.tipo_moeda, 3, "BRL"),
      cnabNum(toQuantidadeMoeda(row.quantidade_moeda), 15),
      cnabNum(valor, 15),
      cnabAlfa(row.nosso_numero, 20),
      "00000000",
      cnabNum(row.valor_real_efetivacao, 15, "0"),
      cnabAlfa(row.informacao2, 40),
      cnabAlfa(row.codigo_finalidade_doc, 2),
      cnabAlfa(row.codigo_finalidade_ted, 5),
      cnabAlfa(row.codigo_finalidade_complementar, 2),
      " ".repeat(3),
      cnabNum(row.aviso_favorecido, 1, "0"),
      cnabAlfa(row.ocorrencias, 10),
    ],
    "Segmento A"
  );
}

function buildSegmentoB(base: CompanyBase, lote: string, sequencial: number, row: CanonicalRecord): string {
  const forma = cnabNum(row.forma_iniciacao, 2).padStart(3, " ");

  return join240(
    [
      base.banco,
      lote,
      "3",
      cnabNum(sequencial, 5),
      "B",
      forma,
      cnabNum(row.tipo_inscricao_favorecido, 1, "0"),
      cnabNum(row.numero_inscricao_favorecido, 14),
      cnabAlfa(row.informacao10, 35),
      cnabAlfa(row.informacao11, 60),
      cnabAlfa(row.informacao12, 99),
      cnabNum(row.codigo_ug_centralizadora, 6),
      cnabNum(row.codigo_ispb, 8),
    ],
    "Segmento B"
  );
}

function buildTrailerLote(base: CompanyBase, lote: string, qtdRegistros: number, somaValores: number, somaQtdMoeda: number): string {
  return join240(
    [
      base.banco,
      lote,
      "5",
      " ".repeat(9),
      cnabNum(qtdRegistros, 6),
      cnabNum(somaValores, 18),
      cnabNum(somaQtdMoeda, 18),
      "000000",
      " ".repeat(165),
      " ".repeat(10),
    ],
    "Trailer de lote"
  );
}

function buildTrailerArquivo(banco: string, qtdLotes: number, qtdRegistros: number): string {
  return join240(
    [
      cnabNum(banco, 3, "237"),
      "9999",
      "9",
      " ".repeat(9),
      cnabNum(qtdLotes, 6),
      cnabNum(qtdRegistros, 6),
      "000000",
      " ".repeat(205),
    ],
    "Trailer de arquivo"
  );
}

function groupRows(rows: CanonicalRecord[]): LoteGroup[] {
  const hasExplicitGroup = rows.some((row) => text(row.lote) !== "");

  const grouped = new Map<string, CanonicalRecord[]>();
  for (const row of rows) {
    const companyKey = [
      onlyDigits(row.tipo_inscricao_empresa),
      onlyDigits(row.numero_inscricao_empresa),
      onlyDigits(row.agencia_empresa),
      onlyDigits(row.conta_empresa),
      text(row.convenio).toUpperCase(),
    ].join("|");

    const explicit = text(row.lote);
    const key = hasExplicitGroup ? (explicit || companyKey) : companyKey;

    if (!grouped.has(key)) {
      grouped.set(key, []);
    }
    grouped.get(key)?.push(row);
  }

  const result: LoteGroup[] = [];
  const explicitLoteToCode = new Map<string, string>();
  let sequence = 1;

  for (const [key, groupRowsValue] of grouped.entries()) {
    let loteCode: string;

    if (hasExplicitGroup) {
      const raw = text(groupRowsValue[0].lote);
      if (/^\d+$/.test(raw)) {
        loteCode = cnabNum(raw, 4, "1");
      } else {
        const reused = explicitLoteToCode.get(raw);
        if (reused) {
          loteCode = reused;
        } else {
          loteCode = cnabNum(sequence, 4, "1");
          explicitLoteToCode.set(raw, loteCode);
          sequence += 1;
        }
      }
    } else {
      loteCode = cnabNum(sequence, 4, "1");
      sequence += 1;
    }

    result.push({ groupKey: key, lote: loteCode, rows: groupRowsValue });
  }

  if (result.length > 9999) {
    throw new BradescoValidationError("Quantidade de lotes excede o limite CNAB (9999). Divida o arquivo em partes.");
  }

  return result;
}

function validateRows(rows: CanonicalRecord[]): { warnings: string[] } {
  const errors: string[] = [];
  const warnings: string[] = [];

  for (const row of rows) {
    const line = row.__row ?? "?";
    const missing: string[] = [];

    const required = [
      "tipo_inscricao_empresa",
      "numero_inscricao_empresa",
      "convenio",
      "agencia_empresa",
      "conta_empresa",
      "nome_empresa",
      "nome_favorecido",
      "data_pagamento",
      "valor_pagamento",
    ] as const;

    for (const field of required) {
      if (!text(row[field])) {
        missing.push(field);
      }
    }

    if (missing.length > 0) {
      errors.push(`Linha ${line}: campos obrigatorios ausentes (${missing.join(", ")}).`);
      continue;
    }

    // Bradesco PIX neste fluxo sempre usa forma 05 (dados bancarios).
    row.forma_iniciacao = "05";
    const forma = row.forma_iniciacao;
    row.forma_iniciacao = forma;

    if (!FORMA_INICIACAO_VALIDAS.has(forma)) {
      errors.push(`Linha ${line}: forma_iniciacao invalida (${row.forma_iniciacao}). Use 01, 02, 03, 04 ou 05.`);
      continue;
    }

    try {
      const valor = toCentavos(row.valor_pagamento, line);
      if (valor === 0) {
        warnings.push(`Linha ${line}: pagamento com valor 0,00.`);
      }
      parseDateToCnab(row.data_pagamento, line);
    } catch (err) {
      if (err instanceof BradescoValidationError) {
        errors.push(err.message);
      } else {
        errors.push(`Linha ${line}: erro inesperado ao validar valor/data.`);
      }
    }

    if (["01", "02", "04"].includes(forma) && !text(row.informacao12)) {
      errors.push(`Linha ${line}: informacao12 (chave Pix) obrigatoria para forma_iniciacao ${forma}.`);
    }

    if (["03", "05"].includes(forma)) {
      if (!text(row.tipo_inscricao_favorecido) || !text(row.numero_inscricao_favorecido)) {
        errors.push(`Linha ${line}: tipo_inscricao_favorecido e numero_inscricao_favorecido sao obrigatorios para forma_iniciacao ${forma}.`);
      }
    }

    if (forma === "05") {
      row.dv_agencia_favorecido = "0";

      const neededBankData = ["banco_favorecido", "agencia_favorecido", "conta_favorecido"] as const;
      const missingBank = neededBankData.filter((field) => !text(row[field]));
      if (missingBank.length > 0) {
        errors.push(`Linha ${line}: campos ${missingBank.join(", ")} obrigatorios para forma_iniciacao 05.`);
      }

      const agenciaDigits = onlyDigits(row.agencia_favorecido);
      if (agenciaDigits.length < 1 || agenciaDigits.length > 5) {
        errors.push("Linha " + line + ": agencia_favorecido deve ter de 1 a 5 digitos na forma 05. O TXT preenche com zeros a esquerda ate 5.");
      }
    }
  }

  if (errors.length > 0) {
    throw new BradescoValidationError("Foram encontrados erros na planilha.", errors.slice(0, 200));
  }

  return { warnings };
}

function centavosToHuman(value: number): string {
  const abs = Math.abs(value);
  const integer = Math.floor(abs / 100);
  const cents = abs % 100;
  const full = `${integer.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")},${String(cents).padStart(2, "0")}`;
  return value < 0 ? `-${full}` : full;
}

export function buildBradescoTransferencia(rows: CanonicalRecord[], extraWarnings: string[] = []): ProcessResult {
  const { warnings } = validateRows(rows);
  const lotes = groupRows(rows);
  const now = new Date();

  const lines: string[] = [];
  const baseGlobal = resolveBase(lotes[0].rows[0], now);
  lines.push(buildHeaderArquivo(baseGlobal));

  let totalCentavos = 0;
  let totalPagamentos = 0;

  for (const lote of lotes) {
    const base = resolveBase(lote.rows[0], now);
    lines.push(buildHeaderLote(base, lote.lote, lote.rows[0]));

    let sequencial = 1;
    let somaValoresLote = 0;
    let somaQtdMoeda = 0;

    for (const row of lote.rows) {
      lines.push(buildSegmentoA(base, lote.lote, sequencial, row));
      sequencial += 1;
      lines.push(buildSegmentoB(base, lote.lote, sequencial, row));
      sequencial += 1;

      const valorCentavos = toCentavos(row.valor_pagamento, row.__row ?? "?");
      somaValoresLote += valorCentavos;
      somaQtdMoeda += toQuantidadeMoeda(row.quantidade_moeda);
      totalCentavos += valorCentavos;
      totalPagamentos += 1;
    }

    const qtdRegistrosLote = lote.rows.length * 2 + 2;
    lines.push(buildTrailerLote(base, lote.lote, qtdRegistrosLote, somaValoresLote, somaQtdMoeda));
  }

  const qtdRegistrosArquivo = lines.length + 1;
  lines.push(buildTrailerArquivo(baseGlobal.banco, lotes.length, qtdRegistrosArquivo));

  const summary: BradescoSummary = {
    lotes: lotes.length,
    pagamentos: totalPagamentos,
    registros: lines.length,
    valorTotal: centavosToHuman(totalCentavos),
    valorTotalCentavos: totalCentavos,
    avisos: [...extraWarnings, ...warnings],
  };

  return {
    txt: `${lines.join("\r\n")}\r\n`,
    summary,
  };
}

const aliasGroups: Record<string, string[]> = {
  banco: ["banco", "codigo_banco", "cod_banco"],
  tipo_inscricao_empresa: [
    "tipo_inscricao_empresa",
    "tipo inscricao empresa",
    "tipo de inscricao da empresa",
    "tp_inscricao_empresa"
  ],
  numero_inscricao_empresa: [
    "numero_inscricao_empresa",
    "numero inscricao empresa",
    "cnpj_empresa",
    "cpf_cnpj_empresa",
    "documento_empresa"
  ],
  convenio: ["convenio", "codigo_convenio", "cod_convenio"],
  agencia_empresa: ["agencia_empresa", "agencia pagadora", "agencia"],
  dv_agencia_empresa: ["dv_agencia_empresa", "digito_agencia_empresa", "dv_agencia"],
  conta_empresa: ["conta_empresa", "conta pagadora", "conta"],
  dv_conta_empresa: ["dv_conta_empresa", "digito_conta_empresa", "dv_conta"],
  dv_agencia_conta_empresa: ["dv_agencia_conta_empresa"],
  nome_empresa: ["nome_empresa", "empresa", "nome pagador", "nome_pagador"],
  nome_banco: ["nome_banco", "banco_nome"],
  data_geracao: ["data_geracao", "data geracao"],
  hora_geracao: ["hora_geracao", "hora geracao"],
  nsa: ["nsa", "sequencia_arquivo", "numero_sequencial_arquivo"],
  densidade: ["densidade"],

  lote: ["lote", "grupo", "trailer", "id_lote", "id lote", "lote_id"],

  tipo_servico: ["tipo_servico", "tipo servico"],
  forma_lancamento: ["forma_lancamento", "forma lancamento", "forma_pagamento"],
  mensagem_lote: ["mensagem_lote", "mensagem lote"],
  logradouro: ["logradouro", "endereco"],
  numero_local: ["numero_local", "numero endereco"],
  complemento_endereco: ["complemento_endereco", "complemento"],
  cidade: ["cidade", "municipio"],
  cep: ["cep"],
  complemento_cep: ["complemento_cep"],
  estado: ["estado", "uf"],
  indicativo_forma_pagamento: ["indicativo_forma_pagamento"],
  ocorrencias_header_lote: ["ocorrencias_header_lote"],

  banco_favorecido: ["banco_favorecido", "banco favorecido", "cod_banco_favorecido"],
  agencia_favorecido: ["agencia_favorecido", "agencia favorecido"],
  dv_agencia_favorecido: ["dv_agencia_favorecido", "digito_agencia_favorecido"],
  conta_favorecido: ["conta_favorecido", "conta favorecido"],
  dv_conta_favorecido: ["dv_conta_favorecido", "digito_conta_favorecido"],
  dv_agencia_conta_favorecido: ["dv_agencia_conta_favorecido"],
  nome_favorecido: ["nome_favorecido", "nome favorecido", "favorecido", "beneficiario"],
  seu_numero: ["seu_numero", "seu numero", "identificador_pagamento", "numero_pagamento", "numero pagamento"],
  data_pagamento: ["data_pagamento", "data pagamento", "vencimento", "data_pgto"],
  tipo_moeda: ["tipo_moeda", "moeda"],
  quantidade_moeda: ["quantidade_moeda", "qtd_moeda"],
  valor_pagamento: ["valor_pagamento", "valor", "valor pagamento", "valor_pix", "valor transfer�ncia", "valor transferencia"],
  nosso_numero: ["nosso_numero", "nosso numero"],
  data_real_efetivacao: ["data_real_efetivacao"],
  valor_real_efetivacao: ["valor_real_efetivacao"],
  informacao2: ["informacao2"],
  codigo_finalidade_doc: ["codigo_finalidade_doc"],
  codigo_finalidade_ted: ["codigo_finalidade_ted"],
  codigo_finalidade_complementar: ["codigo_finalidade_complementar"],
  aviso_favorecido: ["aviso_favorecido"],
  ocorrencias: ["ocorrencias"],

  forma_iniciacao: ["forma_iniciacao", "forma iniciacao", "tipo_chave_pix", "tipo chave pix"],
  tipo_inscricao_favorecido: ["tipo_inscricao_favorecido", "tipo inscricao favorecido"],
  numero_inscricao_favorecido: ["numero_inscricao_favorecido", "documento_favorecido", "cpf_cnpj_favorecido"],
  informacao10: ["informacao10", "info10", "logradouro_favorecido", "endereco_favorecido"],
  informacao11: ["informacao11", "info11", "cidade_favorecido"],
  informacao12: ["informacao12", "info12", "chave_pix", "chave pix", "chave"],
  codigo_ug_centralizadora: ["codigo_ug_centralizadora"],
  codigo_ispb: ["codigo_ispb", "ispb"],
};

function normalizeHeader(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

export const ALIAS_TO_CANONICAL = new Map<string, string>();

for (const [canonical, aliases] of Object.entries(aliasGroups)) {
  ALIAS_TO_CANONICAL.set(normalizeHeader(canonical), canonical);
  for (const alias of aliases) {
    ALIAS_TO_CANONICAL.set(normalizeHeader(alias), canonical);
  }
}

export const MODEL_HEADERS = [
  "lote",
  "tipo_inscricao_empresa",
  "numero_inscricao_empresa",
  "convenio",
  "agencia_empresa",
  "dv_agencia_empresa",
  "conta_empresa",
  "dv_conta_empresa",
  "nome_empresa",
  "tipo_servico",
  "forma_lancamento",
  "banco_favorecido",
  "agencia_favorecido",
  "dv_agencia_favorecido",
  "conta_favorecido",
  "dv_conta_favorecido",
  "nome_favorecido",
  "data_pagamento",
  "valor_pagamento",
  "forma_iniciacao",
  "tipo_inscricao_favorecido",
  "numero_inscricao_favorecido",
  "informacao12",
  "codigo_ispb",
  "numero_pagamento",
  "informacao10",
  "informacao11",
  "estado",
  "cep"
] as const;

export const REQUIRED_BASE_HEADERS = [
  "tipo_inscricao_empresa",
  "numero_inscricao_empresa",
  "convenio",
  "agencia_empresa",
  "conta_empresa",
  "nome_empresa",
  "nome_favorecido",
  "seu_numero",
  "data_pagamento",
  "valor_pagamento",
  "forma_iniciacao"
] as const;

export { normalizeHeader };

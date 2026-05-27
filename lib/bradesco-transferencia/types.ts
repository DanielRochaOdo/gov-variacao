export class BradescoValidationError extends Error {
  readonly details: string[];

  constructor(message: string, details: string[] = []) {
    super(message);
    this.name = "BradescoValidationError";
    this.details = details;
  }
}

export type CanonicalRecord = Record<string, string>;

export interface ProcessResult {
  txt: string;
  summary: BradescoSummary;
}

export interface BradescoSummary {
  lotes: number;
  pagamentos: number;
  registros: number;
  valorTotal: string;
  valorTotalCentavos: number;
  avisos: string[];
}

export interface ParsedSheet {
  rows: CanonicalRecord[];
  warnings: string[];
}

export interface LoteGroup {
  groupKey: string;
  lote: string;
  rows: CanonicalRecord[];
}

export interface CompanyBase {
  banco: string;
  tipoInscricaoEmpresa: string;
  numeroInscricaoEmpresa: string;
  convenio: string;
  agenciaEmpresa: string;
  dvAgenciaEmpresa: string;
  contaEmpresa: string;
  dvContaEmpresa: string;
  dvAgenciaContaEmpresa: string;
  nomeEmpresa: string;
  nomeBanco: string;
  codigoRemessaRetorno: string;
  dataGeracao: string;
  horaGeracao: string;
  nsa: string;
  densidade: string;
  headerPix: string;
}

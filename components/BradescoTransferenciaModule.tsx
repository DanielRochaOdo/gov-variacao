"use client";

import { FormEvent, useMemo, useState } from "react";

type Summary = {
  lotes: number;
  pagamentos: number;
  registros: number;
  valorTotal: string;
  valorTotalCentavos: number;
  avisos: string[];
};

function decodeSummaryHeader(base64Value: string | null): Summary | null {
  if (!base64Value) {
    return null;
  }

  try {
    const json = atob(base64Value);
    return JSON.parse(json) as Summary;
  } catch {
    return null;
  }
}

export function BradescoTransferenciaModule() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>("");
  const [details, setDetails] = useState<string[]>([]);
  const [success, setSuccess] = useState<string>("");
  const [summary, setSummary] = useState<Summary | null>(null);

  const canSubmit = useMemo(() => !!file && !loading, [file, loading]);

  async function onSubmit(event: FormEvent<HTMLFormElement>): Promise<void> {
    event.preventDefault();
    if (!file) {
      setError("Selecione uma planilha .xlsx antes de gerar.");
      return;
    }

    setLoading(true);
    setError("");
    setDetails([]);
    setSuccess("");
    setSummary(null);

    try {
      const formData = new FormData();
      formData.append("arquivo", file);

      const response = await fetch("/api/bradesco-transferencia/processar", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = (await response.json().catch(() => null)) as
          | { erro?: string; detalhes?: string[] }
          | null;

        setError(payload?.erro ?? "Falha ao gerar remessa.");
        setDetails(payload?.detalhes ?? []);
        return;
      }

      const blob = await response.blob();
      const cd = response.headers.get("content-disposition") ?? "";
      const match = cd.match(/filename="?([^\"]+)"?/i);
      const fileName = match?.[1] ?? "remessa_bradesco.rem";

      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(link.href);

      const parsedSummary = decodeSummaryHeader(response.headers.get("x-odonto-bradesco-summary"));
      setSummary(parsedSummary);
      setSuccess("Remessa gerada com sucesso.");
    } catch {
      setError("Falha de rede ao chamar o processamento da remessa.");
    } finally {
      setLoading(false);
    }
  }

  async function onDownloadModel(): Promise<void> {
    setError("");

    try {
      const response = await fetch("/api/bradesco-transferencia/modelo", { method: "GET" });
      if (!response.ok) {
        setError("Nao foi possivel baixar o modelo base agora.");
        return;
      }

      const blob = await response.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "modelo_bradesco_transferencia_cnab240_pix.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(link.href);
    } catch {
      setError("Falha de rede ao baixar o modelo base.");
    }
  }

  return (
    <main className="container">
      <section className="card">
        <h1>Pagamento Bradesco &gt; PIX (CNAB240)</h1>
        <p>
          Envie a planilha <code>.xlsx</code>, gere o arquivo de remessa e receba o resumo de lotes, pagamentos,
          registros e valor total.
        </p>

        <form onSubmit={onSubmit} className="form">
          <label htmlFor="arquivo">Planilha (.xlsx)</label>
          <input
            id="arquivo"
            name="arquivo"
            type="file"
            accept=".xlsx"
            onChange={(event) => setFile(event.target.files?.[0] ?? null)}
          />

          <div className="actions">
            <button type="submit" disabled={!canSubmit}>
              {loading ? "Gerando..." : "Gerar Remessa Bradesco"}
            </button>
            <button type="button" onClick={onDownloadModel} className="secondary" disabled={loading}>
              Baixar Modelo Base
            </button>
          </div>
        </form>

        {error ? <p className="error">{error}</p> : null}
        {details.length > 0 ? (
          <ul className="details">
            {details.map((detail) => (
              <li key={detail}>{detail}</li>
            ))}
          </ul>
        ) : null}
        {success ? <p className="success">{success}</p> : null}

        {summary ? (
          <div className="summary">
            <h2>Resumo</h2>
            <p>Lotes: {summary.lotes}</p>
            <p>Pagamentos: {summary.pagamentos}</p>
            <p>Registros: {summary.registros}</p>
            <p>Valor Total: R$ {summary.valorTotal}</p>
            {summary.avisos.length > 0 ? (
              <>
                <h3>Avisos</h3>
                <ul>
                  {summary.avisos.map((warning) => (
                    <li key={warning}>{warning}</li>
                  ))}
                </ul>
              </>
            ) : null}
          </div>
        ) : null}
      </section>
    </main>
  );
}

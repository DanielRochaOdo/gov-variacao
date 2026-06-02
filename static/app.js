const form = document.getElementById("form-conversao");
const statusNode = document.getElementById("status");
const submitButton = document.getElementById("botao-converter");
const fileInput = document.getElementById("arquivo");
const typeInput = document.getElementById("tipo");
const modelTransferButton = document.getElementById("botao-modelo-transferencia");
const principalUploadBlock = document.getElementById("bloco-upload-principal");
const dropzone = document.getElementById("dropzone");
const selectedFileNameNode = document.getElementById("arquivo-selecionado");
const previewResumoNode = document.getElementById("preview-resumo");
const conciliacaoCard = document.getElementById("card-conciliacao-pix");
const conciliacaoForm = document.getElementById("form-conciliacao-pix");
const conciliacaoStatusNode = document.getElementById("status-conciliacao");
const conciliacaoSubmitButton = document.getElementById("botao-conciliar");
const excelLogInput = document.getElementById("arquivo-excel-log");
const remFilesInput = document.getElementById("arquivos-rem");
const excelLogDropzone = document.getElementById("dropzone-excel-log");
const remDropzone = document.getElementById("dropzone-rem");
const excelLogSelectedNode = document.getElementById("arquivo-excel-log-selecionado");
const remSelectedNode = document.getElementById("arquivos-rem-selecionados");
const ACCEPT_EXCEL = ".xlsx,.xlsm,.xltx,.xltm";
let droppedFile = null;
let droppedExcelLogFile = null;
let droppedRemFiles = [];

function setStatus(message, kind) {
  statusNode.textContent = message || "";
  statusNode.classList.remove("ok", "err", "warn");
  if (kind) statusNode.classList.add(kind);
}

function limparPreviewResumo() {
  if (!previewResumoNode) return;
  previewResumoNode.innerHTML = "<p class=\"preview-empty\">O resumo da remessa aparece aqui apos a conversao.</p>";
  previewResumoNode.classList.remove("ok", "err", "warn");
}

function formatarCentavosParaBRL(centavos) {
  const valor = Number(centavos || 0) / 100;
  return valor.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function renderizarPreviewResumo(response) {
  if (!previewResumoNode) return;
  const contasRaw = response.headers.get("X-Preview-Contas");
  const valorCentavosRaw = response.headers.get("X-Preview-Valor-Centavos");
  if (!contasRaw || !valorCentavosRaw) {
    limparPreviewResumo();
    return;
  }

  const contas = Number.parseInt(contasRaw, 10);
  const valorCentavos = Number.parseInt(valorCentavosRaw, 10);
  if (!Number.isFinite(contas) || !Number.isFinite(valorCentavos)) {
    limparPreviewResumo();
    return;
  }

  previewResumoNode.innerHTML = [
    "<div class=\"preview-content\">",
    "<p class=\"preview-title\">Preview da remessa</p>",
    "<div class=\"preview-metrics\">",
    `<p class=\"preview-metric\">Numero de contas: <strong>${contas}</strong></p>`,
    `<p class=\"preview-metric\">Somatorio total: <strong>R$ ${formatarCentavosParaBRL(valorCentavos)}</strong></p>`,
    "</div>",
    "</div>",
  ].join("");
  previewResumoNode.classList.add("ok");
}

function atualizarFiltroArquivo() {
  fileInput.accept = ACCEPT_EXCEL;
}

function arquivoAtualSelecionado() {
  return droppedFile || fileInput.files?.[0] || null;
}

function atualizarNomeArquivo(file) {
  if (!selectedFileNameNode) return;
  selectedFileNameNode.textContent = file ? `Arquivo selecionado: ${file.name}` : "Nenhum arquivo selecionado";
}

function setStatusConciliacao(message, kind) {
  if (!conciliacaoStatusNode) return;
  conciliacaoStatusNode.textContent = message || "";
  conciliacaoStatusNode.classList.remove("ok", "err", "warn");
  if (kind) conciliacaoStatusNode.classList.add(kind);
}

function arquivoExcelLogAtual() {
  return droppedExcelLogFile || excelLogInput?.files?.[0] || null;
}

function arquivosRemAtuais() {
  const files = droppedRemFiles.length > 0 ? droppedRemFiles : Array.from(remFilesInput?.files || []);
  return files.filter(Boolean);
}

function atualizarNomeArquivoExcelLog(file) {
  if (!excelLogSelectedNode) return;
  excelLogSelectedNode.textContent = file ? `Arquivo selecionado: ${file.name}` : "Nenhum arquivo selecionado";
}

function atualizarNomeArquivosRem(files) {
  if (!remSelectedNode) return;
  if (!files || files.length === 0) {
    remSelectedNode.textContent = "Nenhum arquivo selecionado";
    return;
  }
  const nomes = files.map((file) => file.name);
  remSelectedNode.textContent = `Arquivos selecionados (${files.length}): ${nomes.join(", ")}`;
}

function definirArquivoPorDrop(file) {
  droppedFile = file;
  atualizarNomeArquivo(file);
}

function atualizarVisibilidadeModeloTransferencia() {
  if (!modelTransferButton) return;
  const mostrar = typeInput.value === "BRADESCO_TRANSFERENCIA";
  modelTransferButton.classList.toggle("hidden", !mostrar);
}

function atualizarVisibilidadeConciliacao() {
  if (!conciliacaoCard) return;
  const mostrar = typeInput.value === "PIX_BRADESCO_CONCILIACAO";
  conciliacaoCard.classList.toggle("hidden", !mostrar);
  if (principalUploadBlock) {
    principalUploadBlock.classList.toggle("hidden", mostrar);
  }
  if (submitButton) {
    submitButton.classList.toggle("hidden", mostrar);
  }
}

function lerAlertasValorZero(response) {
  const totalRaw = response.headers.get("X-Valor-Zero-Total");
  if (!totalRaw) {
    return { total: 0, truncated: false, itens: [] };
  }

  const total = Number.parseInt(totalRaw, 10);
  if (!Number.isFinite(total) || total <= 0) {
    return { total: 0, truncated: false, itens: [] };
  }

  const truncated = response.headers.get("X-Valor-Zero-Truncated") === "1";
  const detalhes = response.headers.get("X-Valor-Zero-Details") || "";

  try {
    const itens = JSON.parse(decodeURIComponent(detalhes));
    return { total, truncated, itens: Array.isArray(itens) ? itens : [] };
  } catch (_error) {
    return { total, truncated, itens: [] };
  }
}

async function baixarRespostaComoArquivo(response, fallbackName) {
  const blob = await response.blob();
  const header = response.headers.get("Content-Disposition") || "";
  const match = header.match(/filename="?([^"]+)"?/i);
  const fileName = match?.[1] || fallbackName;
  const url = URL.createObjectURL(blob);

  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  URL.revokeObjectURL(url);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const file = arquivoAtualSelecionado();

  if (!file) {
    setStatus("Selecione um arquivo antes de continuar.", "err");
    return;
  }

  if (!typeInput.value) {
    setStatus("Selecione o tipo do layout antes de converter.", "err");
    return;
  }
  if (typeInput.value === "PIX_BRADESCO_CONCILIACAO") {
    setStatus("Para esse tipo, use o formulario de conciliacao logo abaixo.", "warn");
    return;
  }

  submitButton.disabled = true;
  limparPreviewResumo();
  setStatus("Convertendo arquivo...", null);

  try {
    const formData = new FormData();
    formData.append("tipo", typeInput.value);
    formData.append("arquivo", file);

    const response = await fetch("/api/converter", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      let errorMessage = "Nao foi possivel converter o arquivo.";
      try {
        const payload = await response.json();
        errorMessage = payload.erro || errorMessage;
      } catch (_error) {
        // Ignora erro de parsing JSON e usa mensagem padrao.
      }
      setStatus(errorMessage, "err");
      return;
    }

    const alertaValorZero = lerAlertasValorZero(response);
    const fallbackName = typeInput.value === "BRADESCO_TRANSFERENCIA" ? "remessa_bradesco.rem" : "arquivo_convertido.txt";
    await baixarRespostaComoArquivo(response, fallbackName);
    renderizarPreviewResumo(response);
    if (alertaValorZero.total > 0) {
      const itens = alertaValorZero.itens.map((item) => {
        const registro = item.registro_lote || "00000";
        const banco = item.banco_favorecido || "000";
        const agencia = item.agencia_favorecido || "-";
        const conta = item.conta_favorecido || "-";
        const favorecido = item.nome_favorecido || "Favorecido nao informado";
        return `Reg ${registro}: Banco ${banco} Ag ${agencia} C/C ${conta} (${favorecido})`;
      });

      let mensagem = `Conversao concluida com alerta: ${alertaValorZero.total} pagamento(s) com valor 0,00.`;
      if (itens.length > 0) {
        mensagem += ` ${itens.join(" | ")}`;
      }
      if (alertaValorZero.truncated) {
        mensagem += " Lista parcial exibida.";
      }
      setStatus(mensagem, "warn");
    } else {
      setStatus("Conversao concluida com sucesso.", "ok");
    }
  } catch (_error) {
    setStatus("Falha de conexao. Tente novamente.", "err");
  } finally {
    submitButton.disabled = false;
  }
});

typeInput.addEventListener("change", () => {
  atualizarFiltroArquivo();
  atualizarVisibilidadeModeloTransferencia();
  atualizarVisibilidadeConciliacao();
});

fileInput.addEventListener("change", () => {
  droppedFile = null;
  const file = fileInput.files?.[0];
  atualizarNomeArquivo(file || null);
  if (!file) return;

  const name = file.name.toLowerCase();
  if (name.includes("bradesco") && (name.includes("transfer") || name.includes("pix")) && !typeInput.value) {
    typeInput.value = "BRADESCO_TRANSFERENCIA";
  } else if (name.includes("bradesco") && name.includes("boleto") && !typeInput.value) {
    typeInput.value = "BRADESCO_BOLETO";
  } else if (name.includes("variacao") && !typeInput.value) {
    typeInput.value = "VARIACAO";
  } else if (name.includes("retorno") && !typeInput.value) {
    typeInput.value = "RETORNO";
  }
  atualizarVisibilidadeModeloTransferencia();
});

if (dropzone) {
  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      fileInput.click();
    }
  });

  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");

    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    definirArquivoPorDrop(file);
  });
}

if (conciliacaoForm) {
  conciliacaoForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const excelLog = arquivoExcelLogAtual();
    const remFiles = arquivosRemAtuais();
    if (!excelLog) {
      setStatusConciliacao("Selecione o arquivo XLSX do log de erros.", "err");
      return;
    }
    if (remFiles.length === 0) {
      setStatusConciliacao("Selecione ao menos um arquivo .rem.", "err");
      return;
    }

    conciliacaoSubmitButton.disabled = true;
    setStatusConciliacao("Conciliando arquivos...", null);

    try {
      const formData = new FormData();
      formData.append("arquivo_excel", excelLog);
      remFiles.forEach((file) => formData.append("arquivos_rem", file));

      const response = await fetch("/api/pix-bradesco-conciliacao", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let errorMessage = "Nao foi possivel conciliar os arquivos.";
        try {
          const payload = await response.json();
          errorMessage = payload.erro || errorMessage;
        } catch (_error) {
          // Usa mensagem padrao.
        }
        setStatusConciliacao(errorMessage, "err");
        return;
      }

      await baixarRespostaComoArquivo(response, "duplicados_pix_bradesco.rem");
      setStatusConciliacao("Conciliacao concluida com sucesso.", "ok");
    } catch (_error) {
      setStatusConciliacao("Falha de conexao. Tente novamente.", "err");
    } finally {
      conciliacaoSubmitButton.disabled = false;
    }
  });
}

if (excelLogInput) {
  excelLogInput.addEventListener("change", () => {
    droppedExcelLogFile = null;
    atualizarNomeArquivoExcelLog(excelLogInput.files?.[0] || null);
  });
}

if (remFilesInput) {
  remFilesInput.addEventListener("change", () => {
    droppedRemFiles = [];
    atualizarNomeArquivosRem(Array.from(remFilesInput.files || []));
  });
}

function registrarDropzoneArquivoUnico(dropzoneNode, inputNode, onFile) {
  if (!dropzoneNode || !inputNode) return;
  dropzoneNode.addEventListener("click", () => inputNode.click());
  dropzoneNode.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      inputNode.click();
    }
  });
  dropzoneNode.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzoneNode.classList.add("dragover");
  });
  dropzoneNode.addEventListener("dragleave", () => dropzoneNode.classList.remove("dragover"));
  dropzoneNode.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzoneNode.classList.remove("dragover");
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    onFile(file);
  });
}

function registrarDropzoneMultiplosArquivos(dropzoneNode, inputNode, onFiles) {
  if (!dropzoneNode || !inputNode) return;
  dropzoneNode.addEventListener("click", () => inputNode.click());
  dropzoneNode.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      inputNode.click();
    }
  });
  dropzoneNode.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzoneNode.classList.add("dragover");
  });
  dropzoneNode.addEventListener("dragleave", () => dropzoneNode.classList.remove("dragover"));
  dropzoneNode.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzoneNode.classList.remove("dragover");
    const files = Array.from(event.dataTransfer?.files || []);
    if (files.length === 0) return;
    onFiles(files);
  });
}

registrarDropzoneArquivoUnico(excelLogDropzone, excelLogInput, (file) => {
  droppedExcelLogFile = file;
  atualizarNomeArquivoExcelLog(file);
});

registrarDropzoneMultiplosArquivos(remDropzone, remFilesInput, (files) => {
  droppedRemFiles = files;
  atualizarNomeArquivosRem(files);
});

atualizarFiltroArquivo();
atualizarVisibilidadeModeloTransferencia();
atualizarVisibilidadeConciliacao();
atualizarNomeArquivo(null);
atualizarNomeArquivoExcelLog(null);
atualizarNomeArquivosRem([]);

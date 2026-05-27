const form = document.getElementById("form-conversao");
const statusNode = document.getElementById("status");
const submitButton = document.getElementById("botao-converter");
const fileInput = document.getElementById("arquivo");
const typeInput = document.getElementById("tipo");
const modelTransferButton = document.getElementById("botao-modelo-transferencia");
const dropzone = document.getElementById("dropzone");
const selectedFileNameNode = document.getElementById("arquivo-selecionado");
const ACCEPT_EXCEL = ".xlsx,.xlsm,.xltx,.xltm";
let droppedFile = null;

function setStatus(message, kind) {
  statusNode.textContent = message || "";
  statusNode.classList.remove("ok", "err", "warn");
  if (kind) statusNode.classList.add(kind);
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

function definirArquivoPorDrop(file) {
  droppedFile = file;
  atualizarNomeArquivo(file);
}

function atualizarVisibilidadeModeloTransferencia() {
  if (!modelTransferButton) return;
  const mostrar = typeInput.value === "BRADESCO_TRANSFERENCIA";
  modelTransferButton.classList.toggle("hidden", !mostrar);
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

  submitButton.disabled = true;
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

atualizarFiltroArquivo();
atualizarVisibilidadeModeloTransferencia();
atualizarNomeArquivo(null);

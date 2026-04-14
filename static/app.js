const form = document.getElementById("form-conversao");
const statusNode = document.getElementById("status");
const submitButton = document.getElementById("botao-converter");
const fileInput = document.getElementById("arquivo");
const typeInput = document.getElementById("tipo");

function setStatus(message, kind) {
  statusNode.textContent = message || "";
  statusNode.classList.remove("ok", "err");
  if (kind) statusNode.classList.add(kind);
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
  const file = fileInput.files?.[0];

  if (!file) {
    setStatus("Selecione um arquivo .xlsx antes de continuar.", "err");
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

    await baixarRespostaComoArquivo(response, "arquivo_convertido.txt");
    setStatus("Conversao concluida com sucesso.", "ok");
  } catch (_error) {
    setStatus("Falha de conexao. Tente novamente.", "err");
  } finally {
    submitButton.disabled = false;
  }
});

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  const name = file.name.toLowerCase();
  if (name.includes("variacao") && !typeInput.value) {
    typeInput.value = "VARIACAO";
  } else if (name.includes("retorno") && !typeInput.value) {
    typeInput.value = "RETORNO";
  }
});

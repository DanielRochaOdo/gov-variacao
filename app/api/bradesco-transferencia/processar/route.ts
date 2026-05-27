import { NextResponse } from "next/server";

import { buildBradescoTransferencia } from "@/lib/bradesco-transferencia/cnab";
import { parseSheetFromBuffer } from "@/lib/bradesco-transferencia/excel";
import { BradescoValidationError } from "@/lib/bradesco-transferencia/types";

export const runtime = "nodejs";
const generatedFileCounters = new Map<string, number>();

function toBase64Json(payload: unknown): string {
  const json = JSON.stringify(payload);
  return Buffer.from(json, "utf-8").toString("base64");
}

function buildRemFileName(now: Date): string {
  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yyyy = String(now.getFullYear());
  const base = `${dd}${mm}${yyyy}`;
  const next = (generatedFileCounters.get(base) ?? 0) + 1;
  generatedFileCounters.set(base, next);
  const suffix = next === 1 ? "" : String(next - 1);
  return `${base}${suffix}.rem`;
}

export async function POST(request: Request): Promise<NextResponse> {
  try {
    const formData = await request.formData();
    const file = formData.get("arquivo") ?? formData.get("file");

    if (!(file instanceof File)) {
      return NextResponse.json({ erro: "Envie uma planilha .xlsx no campo arquivo." }, { status: 400 });
    }

    const extension = file.name.split(".").pop()?.toLowerCase();
    if (extension !== "xlsx") {
      return NextResponse.json({ erro: "Formato invalido. Envie um arquivo .xlsx." }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const parsed = await parseSheetFromBuffer(arrayBuffer);
    const result = buildBradescoTransferencia(parsed.rows, parsed.warnings);

    const fileName = buildRemFileName(new Date());
    const summaryBase64 = toBase64Json(result.summary);

    return new NextResponse(new Uint8Array(Buffer.from(result.txt, "latin1")), {
      status: 200,
      headers: {
        "content-type": "text/plain; charset=iso-8859-1",
        "content-disposition": `attachment; filename="${fileName}"`,
        "x-odonto-bradesco-summary": summaryBase64,
      },
    });
  } catch (err) {
    if (err instanceof BradescoValidationError) {
      return NextResponse.json(
        {
          erro: err.message,
          detalhes: err.details,
        },
        { status: 400 }
      );
    }

    return NextResponse.json(
      {
        erro: "Falha inesperada ao processar a remessa Bradesco Transferencia.",
      },
      { status: 500 }
    );
  }
}

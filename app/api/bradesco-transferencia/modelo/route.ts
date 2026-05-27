import { NextResponse } from "next/server";

import { generateModelWorkbook } from "@/lib/bradesco-transferencia/excel";

export const runtime = "nodejs";

export async function GET(): Promise<NextResponse> {
  const modelBuffer = await generateModelWorkbook();
  const fileName = "modelo_bradesco_transferencia_cnab240_pix.xlsx";

  return new NextResponse(new Uint8Array(modelBuffer), {
    status: 200,
    headers: {
      "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "content-disposition": `attachment; filename="${fileName}"`,
      "cache-control": "no-store",
    },
  });
}

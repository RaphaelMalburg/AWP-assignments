import { NextResponse } from "next/server";
import { getStocks } from "@/lib/stocks";

export async function GET(): Promise<NextResponse> {
  const { data, source } = await getStocks();
  return NextResponse.json({ success: true, data, source });
}

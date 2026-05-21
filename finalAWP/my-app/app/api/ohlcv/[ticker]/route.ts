import { NextRequest, NextResponse } from "next/server";

const INTERVAL_SECONDS: Record<string, number> = {
  "1m":  60,
  "5m":  300,
  "15m": 900,
  "60m": 3600,
  "1d":  86400,
  "1wk": 604800,
};

const VALID_INTERVALS = new Set(Object.keys(INTERVAL_SECONDS));

export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ ticker: string }> }
): Promise<NextResponse> {
  const { ticker } = await params;
  const { searchParams } = request.nextUrl;

  const rawInterval = searchParams.get("interval") ?? "60m";
  const interval = VALID_INTERVALS.has(rawInterval) ? rawInterval : "60m";
  const before = searchParams.get("before");

  const COUNT = 400;
  const intervalSec = INTERVAL_SECONDS[interval];
  const durationSec = COUNT * intervalSec;

  const period2 = before ? parseInt(before, 10) : Math.floor(Date.now() / 1000);
  const period1 = period2 - durationSec;

  const symbol = ticker.toUpperCase();
  const url = [
    `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}`,
    `?interval=${interval}`,
    `&period1=${period1}`,
    `&period2=${period2}`,
    `&includePrePost=false`,
  ].join("");

  try {
    const res = await fetch(url, {
      headers: { "User-Agent": "CryptoStockDashboard/1.0" },
      cache: "no-store",
    });
    if (!res.ok) throw new Error(`Yahoo returned ${res.status}`);

    const json = await res.json();
    const result = json?.chart?.result?.[0];
    if (!result) throw new Error("Empty chart result");

    const timestamps: number[] = result.timestamp ?? [];
    const q = result.indicators?.quote?.[0] ?? {};

    const candles = timestamps
      .map((t: number, i: number) => ({
        time:   t,
        open:   typeof q.open?.[i]   === "number" ? (q.open[i]   as number) : null,
        high:   typeof q.high?.[i]   === "number" ? (q.high[i]   as number) : null,
        low:    typeof q.low?.[i]    === "number" ? (q.low[i]    as number) : null,
        close:  typeof q.close?.[i]  === "number" ? (q.close[i]  as number) : null,
        volume: typeof q.volume?.[i] === "number" ? (q.volume[i] as number) : 0,
      }))
      .filter((c) => c.open !== null && c.close !== null && c.high !== null && c.low !== null);

    return NextResponse.json({ success: true, candles });
  } catch (err) {
    return NextResponse.json(
      { success: false, candles: [], error: String(err) },
      { status: 500 }
    );
  }
}

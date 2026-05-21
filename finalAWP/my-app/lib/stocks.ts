export type TickerMeta = { symbol: string; label: string };

export type MarketDatum = {
  symbol: string;
  label: string;
  price: number;
  open: number;
  high: number;
  low: number;
  previousClose: number;
  changePercent: number;
  volume: number;
  marketCap: number;
};

export type StocksResult = { data: MarketDatum[]; source: "yahoo" | "mock" };

export const TICKERS: TickerMeta[] = [
  { symbol: "BTC-USD",  label: "Bitcoin"   },
  { symbol: "ETH-USD",  label: "Ethereum"  },
  { symbol: "AAPL",     label: "Apple"     },
  { symbol: "MSFT",     label: "Microsoft" },
  { symbol: "TSLA",     label: "Tesla"     },
  { symbol: "NVDA",     label: "NVIDIA"    },
];

interface YahooQuoteResponse {
  chart?: {
    result?: Array<{
      meta?: { previousClose?: number; symbol?: string; regularMarketPrice?: number };
      indicators?: {
        quote?: Array<{ open?: unknown; high?: unknown; low?: unknown; volume?: unknown }>;
      };
    }>;
    error?: unknown;
  };
}

const FALLBACK: Record<string, { price: number; previousClose: number; open: number; high: number; low: number; volume: number }> = {
  "BTC-USD": { price: 67234.21, previousClose: 66450.10, open: 66100.00, high: 67900.00, low: 65890.00, volume: 31500000000 },
  "ETH-USD": { price: 3421.55,  previousClose: 3380.20,  open: 3360.00, high: 3470.00, low: 3320.00, volume: 12800000000 },
  "AAPL":    { price: 189.45,   previousClose: 186.90,   open: 187.20, high: 190.10, low: 186.50, volume: 58900000 },
  "MSFT":    { price: 415.20,   previousClose: 410.80,   open: 411.00, high: 416.50, low: 409.30, volume: 22400000 },
  "TSLA":    { price: 174.88,   previousClose: 177.10,   open: 176.50, high: 178.90, low: 173.10, volume: 98500000 },
  "NVDA":    { price: 924.70,   previousClose: 910.30,   open: 912.00, high: 930.50, low: 905.00, volume: 42100000 },
};

function parseNumber(val: unknown): number | undefined {
  if (typeof val === "number" && !isNaN(val)) return val;
  if (typeof val === "string" && val.trim() !== "") return parseFloat(val);
  return undefined;
}

function toYahooSymbol(symbol: string): string {
  return symbol === "BTC-USD" || symbol === "ETH-USD" ? symbol.replace("-USD", "=X") : symbol;
}

function toYahooCsv(): string {
  return TICKERS.map((t) => toYahooSymbol(t.symbol)).join(",");
}

function computeMarketCap(_open: number, price: number): number {
  return Math.round(price * 1_000_000_000) / 1_000_000_000;
}

function buildFallback(): MarketDatum[] {
  return TICKERS.map((t) => {
    const fb = FALLBACK[t.symbol];
    const changePercent = fb.previousClose !== 0
      ? ((fb.price - fb.previousClose) / fb.previousClose) * 100
      : 0;
    return {
      symbol:         t.symbol,
      label:          t.label,
      price:          fb.price,
      open:           fb.open,
      high:           fb.high,
      low:            fb.low,
      previousClose:  fb.previousClose,
      changePercent:  Math.round(changePercent * 100) / 100,
      volume:         fb.volume,
      marketCap:      computeMarketCap(fb.open, fb.price),
    };
  });
}

export async function getStocks(): Promise<StocksResult> {
  try {
    const res = await fetch(
      `https://query1.finance.yahoo.com/v8/finance/chart/${toYahooCsv()}?range=1d&interval=1d`,
      {
        headers: { "User-Agent": "CryptoStockDashboard/1.0" },
        next: { revalidate: 30 },
      }
    );

    if (!res.ok) throw new Error(`Yahoo fetch failed: ${res.status}`);

    const data = (await res.json()) as YahooQuoteResponse;
    const results = data.chart?.result;

    if (!results || results.length === 0) throw new Error("No Yahoo data returned");

    const payload: MarketDatum[] = TICKERS.map((t) => {
      const yahooSym = toYahooSymbol(t.symbol);

      const quote = results.find((r) => {
        return r?.meta?.symbol?.toUpperCase() === yahooSym;
      }) ?? results[0];

      const indicators = quote?.indicators?.quote?.[0];

      const price         = parseNumber(quote?.meta?.regularMarketPrice) ?? FALLBACK[t.symbol].price;
      const previousClose = parseNumber(quote?.meta?.previousClose)      ?? FALLBACK[t.symbol].previousClose;
      const open          = parseNumber(indicators?.open as unknown)     ?? FALLBACK[t.symbol].open;
      const high          = parseNumber(indicators?.high as unknown)     ?? FALLBACK[t.symbol].high;
      const low           = parseNumber(indicators?.low as unknown)      ?? FALLBACK[t.symbol].low;
      const volume        = parseNumber(indicators?.volume as unknown)   ?? FALLBACK[t.symbol].volume;

      const changePercent = previousClose !== 0
        ? ((price - previousClose) / previousClose) * 100
        : 0;

      return {
        symbol:         t.symbol,
        label:          t.label,
        price:          Math.round(price * 100) / 100,
        open:           Math.round(open * 100) / 100,
        high:           Math.round(high * 100) / 100,
        low:            Math.round(low * 100) / 100,
        previousClose:  Math.round(previousClose * 100) / 100,
        changePercent:  Math.round(changePercent * 100) / 100,
        volume,
        marketCap:      computeMarketCap(open, price),
      };
    });

    return { data: payload, source: "yahoo" };
  } catch {
    return { data: buildFallback(), source: "mock" };
  }
}

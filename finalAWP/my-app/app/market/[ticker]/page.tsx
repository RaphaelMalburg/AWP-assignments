import Link from "next/link";
import { notFound } from "next/navigation";
import CandlestickChart from "@/components/candlestick-chart";

type MarketDatum = {
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

type FetchResult = { success: true; data: MarketDatum[] };

const TICKERS = ["BTC-USD", "ETH-USD", "AAPL", "MSFT", "TSLA", "NVDA"];

async function fetchStocks(): Promise<FetchResult> {
  const res = await fetch(`${process.env.NEXT_PUBLIC_BASE_URL ?? "http://localhost:3000"}/api/stocks`, {
    cache: "no-store",
  });
  if (!res.ok) throw new Error("Failed to fetch stocks");
  return res.json();
}

type TickerPageProps = { params: Promise<{ ticker: string }> };

const UpArrow  = () => <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="currentColor"><path d="M12 4l-8 8h5v8h6v-8h5z"/></svg>;
const DownArrow = () => <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="currentColor"><path d="M12 20l8-8h-5V4H9v8H4z"/></svg>;

function fmtCurrency(n: number): string {
  if (n >= 1e9) return `$${(n / 1e9).toFixed(2)}B`;
  if (n >= 1e6) return `$${(n / 1e6).toFixed(2)}M`;
  if (n >= 1e3) return `$${(n / 1e3).toFixed(2)}K`;
  return `$${n.toFixed(2)}`;
}

function fmtVolume(n: number): string {
  if (n >= 1e9) return `${(n / 1e9).toFixed(2)}B`;
  if (n >= 1e6) return `${(n / 1e6).toFixed(2)}M`;
  if (n >= 1e3) return `${(n / 1e3).toFixed(2)}K`;
  return n.toString();
}

export async function generateMetadata({ params }: TickerPageProps): Promise<{ title: string; description: string }> {
  const { ticker } = await params;
  return {
    title: `${ticker.toUpperCase()} — Market Detail`,
    description: `Live price and metrics for ${ticker.toUpperCase()}`,
  };
}

export default async function TickerDetailPage({ params }: TickerPageProps) {
  const { ticker } = await params;
  const upper = ticker.toUpperCase();

  if (!TICKERS.includes(upper)) notFound();

  const { data } = await fetchStocks();
  const stock = data.find((d) => d.symbol === upper);

  if (!stock) notFound();

  const isPositive = stock.changePercent >= 0;
  const sign       = isPositive ? "+" : "";
  const accent     = isPositive ? "text-emerald-400"    : "text-rose-400";
  const barColor   = isPositive ? "bg-emerald-500" : "bg-rose-500";
  const rowAccent  = isPositive ? "border-emerald-500/20" : "border-rose-500/20";

  const daysHighPct = ((stock.high - stock.low) / stock.low) * 100;
  const dayRange    = ((stock.price - stock.low) / (stock.high - stock.low)) * 100;

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <nav className="border-b border-zinc-800/60 bg-zinc-950/80 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <span className="text-xs font-medium text-zinc-500 uppercase tracking-widest">
            Market Detail
          </span>
          <Link
            href="/market"
            className="text-xs font-medium text-zinc-400 hover:text-zinc-100 transition-colors flex items-center gap-1.5"
          >
            <svg className="w-3 h-3" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M19 12H5M12 19l-7-7 7-7" />
            </svg>
            Back to Dashboard
          </Link>
        </div>
      </nav>

      <main className="max-w-5xl mx-auto px-6 py-10">
        <div className="mb-10">
          <div className="flex items-center gap-3 mb-2">
            <h1 className="text-3xl font-semibold tracking-tight text-white">{stock.symbol}</h1>
            <span className="text-sm text-zinc-500 font-light">{stock.label}</span>
          </div>
          <div className="flex items-baseline gap-4 mt-1">
            <span className="text-4xl font-bold text-white tracking-tight">${stock.price.toLocaleString()}</span>
            <span className={`inline-flex items-center gap-1 text-lg font-medium ${accent}`}>
              {isPositive ? <UpArrow /> : <DownArrow />}
              {sign}{stock.changePercent.toFixed(2)}%
            </span>
          </div>
        </div>

        <div className="mb-10">
          <CandlestickChart ticker={upper} />
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-px bg-zinc-800/60 border border-zinc-800/60 mb-10">
          {[
            { label: "Open",           value: `$${stock.open.toLocaleString()}` },
            { label: "Previous Close", value: `$${stock.previousClose.toLocaleString()}` },
            { label: "Day High",       value: `$${stock.high.toLocaleString()}` },
            { label: "Day Low",        value: `$${stock.low.toLocaleString()}` },
          ].map((m) => (
            <div key={m.label} className="bg-zinc-950 px-5 py-4">
              <p className="text-[10px] uppercase tracking-widest text-zinc-500 font-medium mb-1">{m.label}</p>
              <p className="text-base font-semibold text-zinc-100">{m.value}</p>
            </div>
          ))}
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="border border-zinc-800/60 bg-zinc-950 px-6 py-5">
            <h3 className="text-xs uppercase tracking-widest text-zinc-500 font-medium mb-4">Daily Range</h3>
            <div className="flex items-center justify-between text-sm mb-2">
              <span className="text-zinc-400">Low</span>
              <span className="text-zinc-200 font-semibold">${stock.low.toLocaleString()}</span>
            </div>
            <div className="flex items-center justify-between text-sm mb-3">
              <span className="text-zinc-400">High</span>
              <span className="text-zinc-200 font-semibold">${stock.high.toLocaleString()}</span>
            </div>
            <div className="relative h-2 bg-zinc-800 rounded-full overflow-hidden">
              <div className={`absolute left-0 top-0 h-full rounded-full transition-all duration-500 ${barColor}`}
                   style={{ width: `${Math.min(Math.max(dayRange, 0), 100)}%` }}
              />
            </div>
            <p className="text-[10px] text-zinc-600 mt-2">{daysHighPct.toFixed(2)}% daily spread</p>
          </div>

          <div className={`border ${rowAccent} bg-zinc-950 px-6 py-5 ring-1 ring-inset`}>
            <h3 className="text-xs uppercase tracking-widest text-zinc-500 font-medium mb-4">Performance Snapshot</h3>
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-sm text-zinc-400">Δ vs Prev. Close</span>
                <span className={`text-sm font-semibold ${accent}`}>{sign}{stock.changePercent.toFixed(2)}%</span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-sm text-zinc-400">Day Change</span>
                <span className={`text-sm font-semibold ${accent}`}>${Math.abs(stock.price - stock.previousClose).toFixed(2)}</span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-sm text-zinc-400">Est. Market Cap</span>
                <span className="text-sm font-semibold text-zinc-200">{fmtCurrency(stock.marketCap)}</span>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-sm text-zinc-400">24h Volume</span>
                <span className="text-sm font-semibold text-zinc-200">{fmtVolume(stock.volume)}</span>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

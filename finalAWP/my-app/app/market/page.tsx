import Link from "next/link";
import { notFound } from "next/navigation";
import FilterInput from "@/components/filter-input";
import { getStocks, type MarketDatum } from "@/lib/stocks";

function fmtCurrency(n: number): string {
  if (n >= 1e12) return `$${(n / 1e12).toFixed(2)}T`;
  if (n >= 1e9)  return `$${(n / 1e9).toFixed(2)}B`;
  if (n >= 1e6)  return `$${(n / 1e6).toFixed(2)}M`;
  if (n >= 1e3)  return `$${(n / 1e3).toFixed(2)}K`;
  return `$${n.toFixed(2)}`;
}

function fmtVolume(n: number): string {
  if (n >= 1e12) return `${(n / 1e12).toFixed(2)}T`;
  if (n >= 1e9)  return `${(n / 1e9).toFixed(2)}B`;
  if (n >= 1e6)  return `${(n / 1e6).toFixed(2)}M`;
  if (n >= 1e3)  return `${(n / 1e3).toFixed(2)}K`;
  return n.toString();
}

const UpArrow  = () => <svg className="w-3 h-3" viewBox="0 0 24 24" fill="currentColor"><path d="M12 4l-8 8h5v8h6v-8h5z"/></svg>;
const DownArrow = () => <svg className="w-3 h-3" viewBox="0 0 24 24" fill="currentColor"><path d="M12 20l8-8h-5V4H9v8H4z"/></svg>;
const ArrowRight = () => <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12h14M12 5l7 7-7 7"/></svg>;

export default async function MarketPage() {
  let data: MarketDatum[];
  try {
    const result = await getStocks();
    data = result.data;
  } catch {
    notFound();
  }

  const positiveCount = data.filter((d) => d.changePercent >= 0).length;

  const MARKET_LABEL: Record<string, string> = {
    "BTC-USD": "Crypto",
    "ETH-USD": "Crypto",
    "AAPL":    "Equity",
    "MSFT":    "Equity",
    "TSLA":    "Equity",
    "NVDA":    "Equity",
  };

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="max-w-6xl mx-auto px-6 py-12">
        <header className="mb-10">
          <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-6">
            <div>
              <h1 className="text-2xl font-semibold tracking-tight text-white">Market Overview</h1>
              <p className="text-sm text-zinc-500 mt-1">
                Live prices &middot; {positiveCount}/{data.length} assets positive
              </p>
            </div>
            <FilterInput initialData={data} />
          </div>
        </header>

        <div className="border border-zinc-800/60 divide-y divide-zinc-800/40">
          {/* Header row */}
          <div className="grid grid-cols-12 gap-4 px-5 py-2.5 text-[10px] uppercase tracking-widest text-zinc-600 font-semibold select-none">
            <div className="col-span-3">Asset</div>
            <div className="col-span-3 text-right">Price</div>
            <div className="col-span-2 text-right">24h Change</div>
            <div className="col-span-2 text-right">Volume</div>
            <div className="col-span-2 text-right">Market Cap</div>
          </div>

          {/* Data rows */}
          {data.map((stock) => {
            const isPositive = stock.changePercent >= 0;
            const sign       = isPositive ? "+" : "";
            const accent     = isPositive ? "text-emerald-400"  : "text-rose-400";
            const badgeBg    = isPositive ? "bg-emerald-500/10" : "bg-rose-500/10";
            const rowHover   = isPositive
              ? "group-hover:bg-emerald-500/[0.04]"
              : "group-hover:bg-rose-500/[0.04]";

            return (
              <Link
                key={stock.symbol}
                href={`/market/${stock.symbol}`}
                className={`
                  group grid grid-cols-12 gap-4 px-5 py-3.5 items-center
                  transition-colors duration-150 border border-transparent hover:border-zinc-800/40
                  ${rowHover}
                `}
              >
                <div className="col-span-3 min-w-0">
                  <div className="flex items-center gap-2.5 min-w-0">
                    <span className={`inline-block w-1.5 h-1.5 rounded-full shrink-0 ${badgeBg}`} />
                    <div className="min-w-0">
                      <p className="text-sm font-semibold text-white truncate">{stock.symbol}</p>
                      <p className="text-[11px] text-zinc-500 truncate">{stock.label} · {MARKET_LABEL[stock.symbol]}</p>
                    </div>
                  </div>
                </div>

                <div className="col-span-3 text-right">
                  <p className="text-sm font-medium text-zinc-200 tabular-nums">
                    {stock.price < 1
                      ? `$${stock.price.toFixed(4)}`
                      : stock.price < 100
                        ? `$${stock.price.toFixed(2)}`
                        : `$${stock.price.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
                    }
                  </p>
                </div>

                <div className="col-span-2 text-right">
                  <span className={`inline-flex items-center gap-1 text-sm font-semibold tabular-nums ${accent}`}>
                    {isPositive ? <UpArrow /> : <DownArrow />}
                    {sign}{stock.changePercent.toFixed(2)}%
                  </span>
                </div>

                <div className="col-span-2 text-right">
                  <p className="text-sm text-zinc-400 tabular-nums">{fmtVolume(stock.volume)}</p>
                </div>

                <div className="col-span-2 text-right flex items-center justify-end gap-2">
                  <p className="text-sm text-zinc-400 tabular-nums">{fmtCurrency(stock.marketCap)}</p>
                  <ArrowRight />
                </div>
              </Link>
            );
          })}
        </div>
      </div>
    </div>
  );
}

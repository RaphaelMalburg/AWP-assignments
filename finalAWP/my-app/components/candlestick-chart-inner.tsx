"use client";

import {
  createChart,
  CandlestickSeries,
  HistogramSeries,
  LineSeries,
  ColorType,
  CrosshairMode,
  LineStyle,
  type UTCTimestamp,
} from "lightweight-charts";
import { useCallback, useEffect, useRef, useState } from "react";

// ── Types ────────────────────────────────────────────────────────────────────

type Candle = {
  time: number;
  open: number;
  high: number;
  low: number;
  close: number;
  volume: number;
};

type TF = { label: string; interval: string };
type IndKey = "vol" | "ma20" | "ma50" | "bb";

// ── Constants ─────────────────────────────────────────────────────────────────

const TIMEFRAMES: TF[] = [
  { label: "1m",  interval: "1m"  },
  { label: "5m",  interval: "5m"  },
  { label: "15m", interval: "15m" },
  { label: "1H",  interval: "60m" },
  { label: "1D",  interval: "1d"  },
  { label: "1W",  interval: "1wk" },
];

const IND_CONFIG: { key: IndKey; label: string }[] = [
  { key: "vol",  label: "Vol"   },
  { key: "ma20", label: "MA 20" },
  { key: "ma50", label: "MA 50" },
  { key: "bb",   label: "BB 20" },
];

const GREEN  = "#10b981";
const RED    = "#f43f5e";
const CYAN   = "#22d3ee";
const AMBER  = "#f59e0b";
const INDIGO = "#818cf8";

function candleColor(c: Candle) { return c.close >= c.open ? GREEN : RED; }

// ── Indicator math ────────────────────────────────────────────────────────────

function calcSMA(candles: Candle[], period: number) {
  const out: { time: UTCTimestamp; value: number }[] = [];
  for (let i = period - 1; i < candles.length; i++) {
    let sum = 0;
    for (let j = i - period + 1; j <= i; j++) sum += candles[j].close;
    out.push({ time: candles[i].time as UTCTimestamp, value: sum / period });
  }
  return out;
}

function calcBB(candles: Candle[], period = 20, mult = 2) {
  const upper: { time: UTCTimestamp; value: number }[] = [];
  const lower: { time: UTCTimestamp; value: number }[] = [];
  for (let i = period - 1; i < candles.length; i++) {
    let sum = 0;
    for (let j = i - period + 1; j <= i; j++) sum += candles[j].close;
    const ma = sum / period;
    let variance = 0;
    for (let j = i - period + 1; j <= i; j++) variance += (candles[j].close - ma) ** 2;
    const std = Math.sqrt(variance / period);
    const t = candles[i].time as UTCTimestamp;
    upper.push({ time: t, value: ma + mult * std });
    lower.push({ time: t, value: ma - mult * std });
  }
  return { upper, lower };
}

// ── Component ─────────────────────────────────────────────────────────────────

export default function CandlestickChartInner({ ticker }: { ticker: string }) {
  // Chart instance refs
  const containerRef    = useRef<HTMLDivElement>(null);
  const chartRef        = useRef<ReturnType<typeof createChart> | null>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const candleSeriesRef = useRef<any>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const volumeSeriesRef = useRef<any>(null);

  // Indicator series refs (created/destroyed on toggle)
  const indSeriesRef = useRef<{
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ma20?: any; ma50?: any; bb_upper?: any; bb_lower?: any;
  }>({});

  // Mutable state via refs (used inside effects/timers without stale closure issues)
  const candlesRef       = useRef<Candle[]>([]);
  const loadingOlderRef  = useRef(false);
  const pollInFlightRef  = useRef(false);
  const tfIntervalRef    = useRef("60m");
  const activeIndRef     = useRef<Set<IndKey>>(new Set(["vol"]));

  // UI state
  const [tf, setTf]                 = useState<TF>(TIMEFRAMES[3]);
  const [chartReady, setChartReady] = useState(false);
  const [status, setStatus]         = useState<"loading" | "ready" | "error">("loading");
  const [activeInd, setActiveInd]   = useState<Set<IndKey>>(new Set(["vol"]));

  // Keep ref in sync with state (so poll closure reads current value)
  useEffect(() => { activeIndRef.current = activeInd; }, [activeInd]);

  // ── Fetch ─────────────────────────────────────────────────────────────────

  const fetchCandles = useCallback(
    async (interval: string, before?: number): Promise<Candle[]> => {
      const qs = new URLSearchParams({ interval });
      if (before !== undefined) qs.set("before", String(before));
      const res = await fetch(`/api/ohlcv/${ticker}?${qs}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      if (!json.success) throw new Error(json.error ?? "API error");
      return json.candles as Candle[];
    },
    [ticker]
  );

  // ── Indicator sync ────────────────────────────────────────────────────────

  const syncIndicators = useCallback((candles: Candle[], active: Set<IndKey>) => {
    const chart = chartRef.current;
    const vol   = volumeSeriesRef.current;
    if (!chart || !vol) return;

    // Volume visibility
    vol.applyOptions({ visible: active.has("vol") });

    // MA 20
    if (active.has("ma20")) {
      if (!indSeriesRef.current.ma20) {
        indSeriesRef.current.ma20 = chart.addSeries(LineSeries, {
          color: CYAN, lineWidth: 1, priceLineVisible: false, lastValueVisible: false,
        });
      }
      indSeriesRef.current.ma20.setData(calcSMA(candles, 20));
    } else if (indSeriesRef.current.ma20) {
      chart.removeSeries(indSeriesRef.current.ma20);
      indSeriesRef.current.ma20 = undefined;
    }

    // MA 50
    if (active.has("ma50")) {
      if (!indSeriesRef.current.ma50) {
        indSeriesRef.current.ma50 = chart.addSeries(LineSeries, {
          color: AMBER, lineWidth: 1, priceLineVisible: false, lastValueVisible: false,
        });
      }
      indSeriesRef.current.ma50.setData(calcSMA(candles, 50));
    } else if (indSeriesRef.current.ma50) {
      chart.removeSeries(indSeriesRef.current.ma50);
      indSeriesRef.current.ma50 = undefined;
    }

    // Bollinger Bands
    if (active.has("bb")) {
      if (!indSeriesRef.current.bb_upper) {
        indSeriesRef.current.bb_upper = chart.addSeries(LineSeries, {
          color: INDIGO, lineWidth: 1, lineStyle: LineStyle.Dashed,
          priceLineVisible: false, lastValueVisible: false,
        });
        indSeriesRef.current.bb_lower = chart.addSeries(LineSeries, {
          color: INDIGO, lineWidth: 1, lineStyle: LineStyle.Dashed,
          priceLineVisible: false, lastValueVisible: false,
        });
      }
      const { upper, lower } = calcBB(candles);
      indSeriesRef.current.bb_upper.setData(upper);
      indSeriesRef.current.bb_lower.setData(lower);
    } else if (indSeriesRef.current.bb_upper) {
      chart.removeSeries(indSeriesRef.current.bb_upper);
      chart.removeSeries(indSeriesRef.current.bb_lower);
      indSeriesRef.current.bb_upper = undefined;
      indSeriesRef.current.bb_lower = undefined;
    }
  }, []);

  // ── Chart init (once) ─────────────────────────────────────────────────────

  useEffect(() => {
    if (!containerRef.current) return;

    const chart = createChart(containerRef.current, {
      width:  containerRef.current.clientWidth,
      height: 460,
      layout: {
        background: { type: ColorType.Solid, color: "#09090b" },
        textColor: "#71717a",
        fontSize: 11,
        fontFamily: "ui-monospace, SFMono-Regular, monospace",
      },
      grid: {
        vertLines: { color: "#18181b" },
        horzLines: { color: "#18181b" },
      },
      crosshair: { mode: CrosshairMode.Normal },
      rightPriceScale: { borderColor: "#27272a" },
      timeScale: {
        borderColor: "#27272a",
        timeVisible: true,
        secondsVisible: false,
        rightOffset: 12,
        barSpacing: 8,
      },
    });

    const candleSeries = chart.addSeries(CandlestickSeries, {
      upColor: GREEN, downColor: RED,
      borderUpColor: GREEN, borderDownColor: RED,
      wickUpColor: GREEN, wickDownColor: RED,
    });

    const volumeSeries = chart.addSeries(HistogramSeries, {
      priceFormat: { type: "volume" },
      priceScaleId: "vol",
    });
    chart.priceScale("vol").applyOptions({ scaleMargins: { top: 0.82, bottom: 0 } });

    chartRef.current        = chart;
    candleSeriesRef.current = candleSeries;
    volumeSeriesRef.current = volumeSeries;

    const ro = new ResizeObserver(() => {
      if (containerRef.current) chart.applyOptions({ width: containerRef.current.clientWidth });
    });
    ro.observe(containerRef.current);
    setChartReady(true);

    return () => {
      ro.disconnect();
      chart.remove();
      chartRef.current = candleSeriesRef.current = volumeSeriesRef.current = null;
      indSeriesRef.current = {};
      setChartReady(false);
    };
  }, []);

  // ── Load data on TF change ────────────────────────────────────────────────

  useEffect(() => {
    if (!chartReady) return;
    tfIntervalRef.current = tf.interval;
    setStatus("loading");
    candlesRef.current = [];

    // Remove indicator series so they get recreated with correct data
    const chart = chartRef.current;
    if (chart) {
      for (const s of [indSeriesRef.current.ma20, indSeriesRef.current.ma50,
                        indSeriesRef.current.bb_upper, indSeriesRef.current.bb_lower]) {
        if (s) chart.removeSeries(s);
      }
      indSeriesRef.current = {};
    }

    fetchCandles(tf.interval)
      .then((candles) => {
        candlesRef.current = candles;

        const cs = candleSeriesRef.current;
        const vs = volumeSeriesRef.current;
        if (!cs || !vs) return;

        cs.setData(candles.map((c) => ({ time: c.time as UTCTimestamp, open: c.open, high: c.high, low: c.low, close: c.close })));
        vs.setData(candles.map((c) => ({ time: c.time as UTCTimestamp, value: c.volume, color: candleColor(c) })));

        syncIndicators(candles, activeIndRef.current);
        chartRef.current?.timeScale().fitContent();
        setStatus("ready");
      })
      .catch(() => setStatus("error"));
  }, [chartReady, tf, fetchCandles, syncIndicators]);

  // ── Re-sync indicators when toggles change ────────────────────────────────

  useEffect(() => {
    if (!chartReady || candlesRef.current.length === 0) return;
    syncIndicators(candlesRef.current, activeInd);
  }, [chartReady, activeInd, syncIndicators]);

  // ── Scroll-left pagination ────────────────────────────────────────────────

  useEffect(() => {
    if (!chartReady || !chartRef.current) return;

    const handler = (range: { from: number; to: number } | null) => {
      if (!range || loadingOlderRef.current || candlesRef.current.length === 0) return;
      if (range.from > 10) return;

      loadingOlderRef.current = true;
      const oldest = candlesRef.current[0].time;

      fetchCandles(tfIntervalRef.current, oldest)
        .then((older) => {
          if (older.length === 0) return;
          const seen = new Set<number>();
          const merged = [...older, ...candlesRef.current]
            .filter((c) => { if (seen.has(c.time)) return false; seen.add(c.time); return true; })
            .sort((a, b) => a.time - b.time);
          candlesRef.current = merged;

          const cs = candleSeriesRef.current;
          const vs = volumeSeriesRef.current;
          if (!cs || !vs) return;

          cs.setData(merged.map((c) => ({ time: c.time as UTCTimestamp, open: c.open, high: c.high, low: c.low, close: c.close })));
          vs.setData(merged.map((c) => ({ time: c.time as UTCTimestamp, value: c.volume, color: candleColor(c) })));
          syncIndicators(merged, activeIndRef.current);
        })
        .finally(() => { loadingOlderRef.current = false; });
    };

    const ts = chartRef.current.timeScale();
    ts.subscribeVisibleLogicalRangeChange(handler);
    return () => { ts.unsubscribeVisibleLogicalRangeChange(handler); };
  }, [chartReady, tf, fetchCandles, syncIndicators]);

  // ── 15-second live polling ────────────────────────────────────────────────

  useEffect(() => {
    if (!chartReady) return;

    const poll = async () => {
      if (pollInFlightRef.current) return;
      const cs = candleSeriesRef.current;
      const vs = volumeSeriesRef.current;
      if (!cs || !vs || candlesRef.current.length === 0) return;

      pollInFlightRef.current = true;
      try {
        const fetched = await fetchCandles(tfIntervalRef.current);
        if (fetched.length === 0) return;

        const current  = candlesRef.current;
        const lastTime = current[current.length - 1].time;

        // All candles at or after the current last bar
        const relevant = fetched.filter((c) => c.time >= lastTime);
        if (relevant.length === 0) return;

        // Rebuild ref: drop the old last bar, append all relevant ones
        const base = current.filter((c) => c.time < lastTime);
        candlesRef.current = [...base, ...relevant];

        // Push each candle to the chart series in order
        for (const c of relevant) {
          cs.update({ time: c.time as UTCTimestamp, open: c.open, high: c.high, low: c.low, close: c.close });
          vs.update({ time: c.time as UTCTimestamp, value: c.volume, color: candleColor(c) });
        }

        // Re-sync indicators with updated full data set
        syncIndicators(candlesRef.current, activeIndRef.current);
      } catch { /* silent — retries in 15s */ }
      finally { pollInFlightRef.current = false; }
    };

    const id = setInterval(poll, 15_000);
    return () => clearInterval(id);
  }, [chartReady, fetchCandles, syncIndicators]);

  // ── Toggle indicator ──────────────────────────────────────────────────────

  function toggleInd(key: IndKey) {
    setActiveInd((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="rounded-xl border border-zinc-800/60 overflow-hidden bg-zinc-950">

      {/* ── Toolbar ─────────────────────────────────────── */}
      <div className="flex flex-wrap items-center gap-x-1 gap-y-1.5 px-4 py-2.5 border-b border-zinc-800/60 bg-zinc-900/30">

        {/* Timeframe */}
        <span className="text-[10px] uppercase tracking-widest text-zinc-600 font-medium mr-0.5">TF</span>
        {TIMEFRAMES.map((t) => (
          <button
            key={t.label}
            onClick={() => setTf(t)}
            className={`px-2.5 py-1 rounded text-xs font-mono transition-colors ${
              tf.label === t.label
                ? "bg-zinc-700 text-zinc-100"
                : "text-zinc-500 hover:text-zinc-200 hover:bg-zinc-800"
            }`}
          >
            {t.label}
          </button>
        ))}

        {/* Divider */}
        <div className="w-px h-4 bg-zinc-800 mx-1.5" />

        {/* Indicators */}
        <span className="text-[10px] uppercase tracking-widest text-zinc-600 font-medium mr-0.5">Ind</span>
        {IND_CONFIG.map(({ key, label }) => {
          const active = activeInd.has(key);
          const colors: Record<IndKey, string> = {
            vol:  "text-zinc-300 border-zinc-500",
            ma20: "text-cyan-400 border-cyan-500",
            ma50: "text-amber-400 border-amber-500",
            bb:   "text-indigo-400 border-indigo-500",
          };
          return (
            <button
              key={key}
              onClick={() => toggleInd(key)}
              className={`px-2.5 py-1 rounded border text-xs font-mono transition-colors ${
                active
                  ? `bg-zinc-800 ${colors[key]}`
                  : "border-transparent text-zinc-600 hover:text-zinc-400 hover:bg-zinc-800/50"
              }`}
            >
              {label}
            </button>
          );
        })}

        {/* Status */}
        <div className="ml-auto flex items-center gap-2 text-[10px] font-mono">
          {status === "loading" && <span className="text-zinc-600 animate-pulse">Loading…</span>}
          {status === "error"   && <span className="text-rose-500">Failed to load</span>}
          {status === "ready"   && (
            <span className="text-zinc-600 flex items-center gap-1.5">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
              Live · 15s
            </span>
          )}
        </div>
      </div>

      {/* ── Indicator legend ─────────────────────────────── */}
      {(activeInd.has("ma20") || activeInd.has("ma50") || activeInd.has("bb")) && (
        <div className="flex items-center gap-4 px-4 py-1.5 border-b border-zinc-800/40 bg-zinc-900/10 text-[10px] font-mono">
          {activeInd.has("ma20") && <span className="flex items-center gap-1.5"><span className="w-5 h-px bg-cyan-400 inline-block" />MA 20</span>}
          {activeInd.has("ma50") && <span className="flex items-center gap-1.5"><span className="w-5 h-px bg-amber-400 inline-block" />MA 50</span>}
          {activeInd.has("bb")   && <span className="flex items-center gap-1.5"><span className="w-5 h-px bg-indigo-400 inline-block border-dashed" />BB 20</span>}
        </div>
      )}

      {/* ── Chart ───────────────────────────────────────── */}
      <div ref={containerRef} style={{ height: 460 }} />
    </div>
  );
}
